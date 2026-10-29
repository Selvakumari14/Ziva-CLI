#include <stdarg.h>
#include <stdbool.h>
#include <stdint.h>
#include <stdlib.h>

void rust_string_free(char *s);

char *nativeInitTimon(const char *storage_path, uint32_t bucket_interval, const char *username);

char *nativeCreateDatabase(const char *db_name);

char *nativeCreateTable(const char *db_name, const char *table_name, const char *schema);

char *nativeListDatabases(void);

char *nativeListTables(const char *db_name);

char *nativeDeleteDatabase(const char *db_name);

char *nativeDeleteTable(const char *db_name, const char *table_name);

char *nativeInsert(const char *db_name, const char *table_name, const char *json_data);

char *nativeQuery(const char *db_name,
                  const char *sql_query,
                  const char *username,
                  int32_t limit_partitions);

char *nativePreloadTables(const char *db_name, const char *table_names_json, const char *username);

char *nativeInitBucket(const char *bucket_endpoint,
                       const char *bucket_name,
                       const char *access_key_id,
                       const char *secret_access_key,
                       const char *bucket_region);

char *nativeCloudSyncParquet(const char *db_name,
                             const char *table_name,
                             const char *date_range_json,
                             const char *username);

char *nativeCloudSinkParquet(const char *db_name, const char *table_name);

char *nativeCloudFetchParquet(const char *username,
                              const char *db_name,
                              const char *table_name,
                              const char *date_range_json);

char *nativeCloudFetchParquetBatch(const char *usernames_json,
                                   const char *db_names_json,
                                   const char *table_names_json,
                                   const char *date_range_json);

char *nativeGetAllSyncMetadata(const char *db_name);

char *nativeGetSyncMetadata(const char *db_name, const char *table_name);
