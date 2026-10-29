#!/usr/bin/env node

// ─────────────────────────────────────────────────────────────────
// Phase 3: classifyBlocks
// Reads phase2 JSON from stdin, writes classified JSON to stdout
// Usage: echo '<json>' | node classify_blocks.js <tzOffsetHours>
// ─────────────────────────────────────────────────────────────────

const BOUT_MAX_LEN   = 3;
const BLOCK_SEP      = 3600;   // 1 hour gap → new block
const CONF_BOUT      = 0.6;
const CONF_FIRMWARE  = 1.0;

const tzOffsetHours  = parseFloat(process.argv[2] ?? '0');
const tzOffsetSec    = Math.round(tzOffsetHours * 3600);

// "2026.04.18 16:52:01" → epoch seconds (treated as UTC)
const parseDateStr = (d) => {
    if (!d) return 0;
    const iso = d.replace(/\./g, (m, offset) => offset < 10 ? '-' : '.').replace(' ', 'T') + 'Z';
    return Math.floor(new Date(iso).getTime() / 1000);
};

const localHour = (epoch) => new Date((epoch + tzOffsetSec) * 1000).getUTCHours();

// Read all stdin
let raw = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => raw += chunk);
process.stdin.on('end', () => {
    let rows;
    try {
        rows = JSON.parse(raw);
    } catch (e) {
        process.stderr.write('classify_blocks: failed to parse stdin JSON: ' + e.message + '\n');
        process.exit(1);
    }

    if (!rows || rows.length === 0) {
        process.stdout.write(JSON.stringify([]));
        process.exit(0);
    }

    // ── Step 1: Assign bout IDs (consecutive same-quality runs) ──
    const withBouts = rows.map((row, i) => {
        const prev = rows[i - 1];
        const isBoutStart = !prev
            || row.quality !== prev.quality
            || parseDateStr(row.date) - parseDateStr(prev.date) > 65
            || row.quality === 5
            || (prev && prev.quality === 5);
        return { ...row, is_bout_start: isBoutStart ? 1 : 0 };
    });

    let boutId = 0;
    const withBoutIds = withBouts.map(row => {
        if (row.is_bout_start) boutId++;
        return { ...row, bout_id: boutId };
    });

    // ── Step 2: Consolidate short deep bouts (≤3 rows) → light ──
    const boutLengths   = {};
    const boutQualities = {};
    for (const row of withBoutIds) {
        boutLengths[row.bout_id]   = (boutLengths[row.bout_id]   || 0) + 1;
        boutQualities[row.bout_id] = row.quality;
    }

    const consolidated = withBoutIds.map(row => {
        const len           = boutLengths[row.bout_id];
        const qual          = boutQualities[row.bout_id];
        const prevBoutQual  = boutQualities[row.bout_id - 1];
        const nextBoutQual  = boutQualities[row.bout_id + 1];
        // Only consolidate deep sleep (1) micro-fragments — never REM (3)
        const isMergeable   = qual === 1
            && len <= BOUT_MAX_LEN
            && (prevBoutQual === 2 || nextBoutQual === 2);

        return {
            ...row,
            quality:    isMergeable ? 2 : row.quality,
            source:     isMergeable ? 'bout_consolidated' : row.source,
            confidence: isMergeable ? CONF_BOUT : row.confidence,
        };
    });

    // ── Step 3: Assign block IDs (gap > BLOCK_SEP → new block) ──
    let blockId = 0;
    const withBlocks = consolidated.map((row, i) => {
        const prev = consolidated[i - 1];
        if (!prev || parseDateStr(row.date) - parseDateStr(prev.date) > BLOCK_SEP) blockId++;
        return { ...row, block_id: blockId };
    });

    // ── Step 4: Find primary block (most sleep mins in NIGHT hours) ──
    const blockSleepMins  = {};
    const blockNightMins  = {};
    const blockFwMins     = {};
    const sessionSizes    = {};
    const sessionFwMins   = {};

    for (const row of withBlocks) {
        if ([1, 2, 3].includes(row.quality)) {
            blockSleepMins[row.block_id] = (blockSleepMins[row.block_id] || 0) + 1;
            const h = localHour(parseDateStr(row.date));
            if (h >= 20 || h < 10)
                blockNightMins[row.block_id] = (blockNightMins[row.block_id] || 0) + 1;
        }
        if (row.source === 'firmware') {
            blockFwMins[row.block_id]     = (blockFwMins[row.block_id]     || 0) + 1;
            sessionFwMins[row.session_id] = (sessionFwMins[row.session_id] || 0) + 1;
        }
        sessionSizes[row.session_id] = (sessionSizes[row.session_id] || 0) + 1;
    }

    const primaryScore   = Object.keys(blockNightMins).length ? blockNightMins : blockSleepMins;
    const primaryBlockId = Object.entries(primaryScore)
        .sort(([idA, a], [idB, b]) => b - a || Number(idA) - Number(idB))[0]?.[0];

    // ── Step 5: Assign block_type and clean up internal fields ──
    const result = withBlocks.map(({ is_bout_start, bout_id, block_id, session_id, ...row }) => ({
        ...row,
        block_type:
            sessionSizes[session_id] < 30 && !sessionFwMins[session_id]
                ? 'microsleep'
                : (blockFwMins[block_id] || 0) === 0 && (blockSleepMins[block_id] || 0) < 30
                    ? 'microsleep'
                    : String(block_id) === primaryBlockId
                        ? 'primary'
                        : 'secondary',
    }));

    process.stdout.write(JSON.stringify(result));
});
