# Timon Data Processor

A command-line tool that imports raw Timon health data from JSON, converts timezones, runs a full sleep processing pipeline, and provides interactive analysis with gap detection.

---

## Requirements

- Python 3.10+
- The following packages are **auto-installed** on first run if missing:
  - `pandas`
  - `pyarrow`

No manual `pip install` needed.

---

## Setup

Place `process_timon.py` and your input JSON file in the same directory:

```
timon_cli/
├── process_timon.py
└── kk.json
```

---

## Running

```bash
python3 process_timon.py <your_file.json>
```

**Example:**

```bash
python3 process_timon.py kk.json
```

---

## What It Does — Step by Step

### Step 1 — Import JSON → Parquet

Reads the input JSON file. Each top-level key becomes its own `.parquet` file saved in `./output/`.

```
./output/
├── sleep.parquet
├── heartrate.parquet
├── hrv_table.parquet
├── spo2.parquet
├── temperature_table.parquet
└── activitydetails.parquet
```

> Nested objects and arrays are serialised to JSON strings so Parquet can handle them.

---

### Step 2 — Select Timezone

You will be prompted to pick your timezone. This is used to convert all `date` fields from UTC into local time.

```
1) IST  – India Standard Time   (UTC+5:30)
2) PST  – Pacific Standard Time (UTC-8:00)
3) BST  – British Summer Time   (UTC+1:00)
```

---

### Step 3 — Date Conversion

All parquet files are re-saved with dates converted to your chosen timezone into `./output_corrected/`.

```
./output_corrected/
├── sleep_corrected.parquet
├── heartrate_corrected.parquet
└── ...
```

---

### Step 4 — Sleep Processing Pipeline

Runs a three-phase sleep analysis over all nights found in the data:

| Phase | What it does |
|-------|-------------|
| Phase 1 | Detects sleep onset and wake boundaries using HR, HRV, SpO2, temperature, and activity signals |
| Phase 2 | Corrects the night window — fills envelope gaps and recovers long gaps using HR data |
| Phase 3 | Classifies records into blocks (`primary`, `secondary`, `microsleep`) and labels sleep quality |

**Quality values:**

| Value | Meaning |
|-------|---------|
| 1 | Deep sleep |
| 2 | Light sleep |
| 3 | REM |
| 4 | Awake |
| 5 | Wake |

**Output files** (saved to `./output_corrected/`):

```
sleep_processed.parquet   ← full processed dataset
sleep_processed.json      ← same data as JSON (human-readable)
```

---

### Step 5 — Analysis Menu

After processing completes, an interactive analysis menu appears. It loops so you can run multiple analyses without re-running the script.

```
🔍 Analysis
═══════════════════════════════════════════════
  Which dataset would you like to analyse?

  1) Sleep
  0) Skip / Exit
```

#### Sleep Analysis

Select **1**, then enter a date in `YYYY.MM.DD` format:

```
Enter the date to analyse (YYYY.MM.DD): 2026.03.18
```

**Window queried:**

The script fetches all sleep records in the 24-hour window ending at 18:00 on the chosen date (local time):

```
Previous day 18:00:00  →  Chosen day 17:59:59  (local time)
```

For example, entering `2026.03.18` with IST selected queries:

```
Local : 2026.03.17 18:00:00 → 2026.03.18 17:59:59
UTC   : 2026.03.17 12:30:00 → 2026.03.18 12:29:59
```

**Output includes:**

- Total records found in the window
- First and last record timestamps
- Total sleep span (minutes / hours)
- Quality breakdown by sleep stage
- **Gap detection** — sleep records are expected every 60 seconds. Any consecutive gap longer than 60 s is flagged:

```
⚠️  3 gap(s) detected:

[  1] 2026.03.17 23:14:00  →  2026.03.17 23:22:00
       Gap : 480 s  |  Missing ≈ 7.0 min

[  2] 2026.03.18 02:01:00  →  2026.03.18 02:04:00
       Gap : 180 s  |  Missing ≈ 2.0 min

Total missing data : ≈ 9.0 min across 2 gap(s)
```

---

## Output Directory Structure

After a full run:

```
timon_cli/
├── process_timon.py
├── kk.json
├── output/                          ← raw parquet (one per table)
│   ├── sleep.parquet
│   ├── heartrate.parquet
│   └── ...
└── output_corrected/                ← timezone-corrected + processed
    ├── sleep_corrected.parquet
    ├── heartrate_corrected.parquet
    ├── ...
    ├── sleep_processed.parquet
    └── sleep_processed.json
```

---

## Expected Input JSON Format

The input file must be a JSON object where each key is a table name and the value is an array of records:

```json
{
  "sleep": [
    { "date": "2026.03.18 01:00:00", "quality": 2, "start": "2026.03.17 23:00:00", "unitLength": 1 },
    ...
  ],
  "heartrate": [
    { "date": "2026.03.18 01:00:00", "singleHR": 58 },
    ...
  ],
  "hrv_table": [...],
  "spo2": [...],
  "temperature_table": [...],
  "activitydetails": [...]
}
```

Date strings must be in `YYYY.MM.DD HH:MM:SS` format (UTC).

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `pyarrow.lib.ArrowNotImplementedError: Cannot write struct type` | Update to the latest `process_timon.py` — this is fixed |
| `❌ File not found` | Check the filename passed as the argument matches the actual file |
| `⚠️ No sleep data — skipping` | The JSON `sleep` key is missing or empty |
| `⚠️ No sleep records found in this window` | No records exist for that date — try an adjacent date |
| Analysis shows large gaps | Normal if the device was off-wrist; gaps > 1 min are flagged by design |
