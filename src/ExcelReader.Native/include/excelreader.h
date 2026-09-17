#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

#define XL_OK                 0
#define XL_EOF               -1
#define XL_BUFFER_TOO_SMALL  -2
#define XL_INVALID_HANDLE    -3
#define XL_INVALID_ARGUMENT  -4
#define XL_ERROR             -5
#define XL_STATUS_PASSWORD_REQUIRED  (-6) /* the workbook is encrypted and no password was supplied */
#define XL_STATUS_PASSWORD_INCORRECT (-7) /* the supplied password did not match the workbook's verifier */

#define XL_ABI_VERSION 5

#define XL_FORMAT_AUTO  0  /* sniffs XLS/XLSX/XLSB; does NOT detect CSV */
#define XL_FORMAT_XLS   1
#define XL_FORMAT_XLSX  2
#define XL_FORMAT_XLSB  3
#define XL_FORMAT_CSV   4  /* must be requested explicitly */

#define XL_CELL_EMPTY   0
#define XL_CELL_STRING  1
#define XL_CELL_NUMBER  2
#define XL_CELL_DATE    3
#define XL_CELL_BOOL    4
#define XL_CELL_FORMULA 5
#define XL_CELL_ERROR   6

#define XL_MAX_COLUMN_SPECS      16384
#define XL_MAX_COLUMN_NAME_BYTES 131068

typedef struct xl_workbook xl_workbook;

typedef struct xl_typed_reader xl_typed_reader;

typedef struct xl_row_cell {
    int32_t column;
    int32_t type;
    int32_t value_len;
    const uint8_t* value;
} xl_row_cell;

typedef struct xl_row {
    int32_t cell_count;
    xl_row_cell* cells;
} xl_row;

int32_t xl_open_file(const uint8_t* path, int32_t path_len, int32_t format, xl_workbook** out_handle);

int32_t xl_open_memory(const uint8_t* data, int32_t data_len, int32_t format, xl_workbook** out_handle);

#define XL_OPT_DEFAULT 0
#define XL_OPT_FALSE   1
#define XL_OPT_TRUE    2

typedef struct xl_open_options {
    int32_t struct_size;

    int32_t csv_sniff_dialect;        
    int32_t csv_delimiter;            
    int32_t csv_quote;                
    int32_t csv_detect_bom;           
    int32_t csv_max_cell_bytes;       
    int32_t csv_intern_strings;       

    int64_t max_total_decompressed_bytes; 
    int32_t max_cell_bytes;               
    int64_t max_shared_string_bytes;      
    int32_t max_zip_entries;              
    int32_t prefetch_decompression;       
    int32_t intern_strings;               

    const uint8_t* password;
    int32_t password_len;
} xl_open_options;

int32_t xl_open_file_ex(const uint8_t* path, int32_t path_len, int32_t format,
                        const xl_open_options* options, xl_workbook** out_handle);

int32_t xl_open_memory_ex(const uint8_t* data, int32_t data_len, int32_t format,
                          const xl_open_options* options, xl_workbook** out_handle);

int32_t xl_close(xl_workbook* handle);

int32_t xl_sheet_count(xl_workbook* handle, int32_t* out_count);

int32_t xl_sheet_name(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_len);

int32_t xl_sheet_name_at(xl_workbook* handle, int32_t index, uint8_t* buffer, int32_t capacity, int32_t* out_len);

int32_t xl_move_to_sheet(xl_workbook* handle, int32_t index);

int32_t xl_is_date1904(xl_workbook* handle, int32_t* out_flag);

int32_t xl_next_row(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

int32_t xl_read_all_blob(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

typedef struct xl_rows {
    int32_t row_count;
    xl_row* rows;
} xl_rows;

int32_t xl_read_all_decoded(xl_workbook* handle, xl_rows* out_rows);

void xl_free_rows(xl_rows* rows);

#define XL_T_STRING    0
#define XL_T_I64       1
#define XL_T_F64       2
#define XL_T_BOOL      3
#define XL_T_DATE      4   /* days since 1970-01-01, int32 */
#define XL_T_TIME      5   /* microseconds since midnight, int64 */
#define XL_T_TIMESTAMP 6   /* microseconds since 1970-01-01T00:00:00Z, int64 */

typedef struct xl_column_spec {
    const uint8_t* const* names; 
    const int32_t* name_lens;    
    int32_t name_count;
    int32_t index;                
    int32_t type;                 
    int32_t nullable;             
} xl_column_spec;

typedef struct xl_column {
    int32_t type;
    int64_t length;
    const void* values;
    const uint8_t* validity;
    const uint8_t* data;
    int64_t data_len;
} xl_column;

typedef struct xl_table {
    int32_t column_count;
    int64_t row_count;
    xl_column* columns;
} xl_table;

int32_t xl_parse_typed(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                       int32_t header_row, xl_table* out_table);

void xl_free_table(xl_table* table);


int32_t xl_typed_reader_open(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                             int32_t header_row, int64_t max_rows, xl_typed_reader** out_reader);

int32_t xl_typed_reader_next(xl_typed_reader* reader, xl_table* out_table);

void xl_typed_reader_close(xl_typed_reader* reader);


typedef struct xl_write_options {
    int32_t struct_size;
    int32_t sheet_name_len;
    const uint8_t* sheet_name;    

    int32_t csv_delimiter;        
    int32_t csv_quote;            

    int32_t date1904;             

    int32_t use_shared_strings;   
} xl_write_options;

int32_t xl_write_typed(const uint8_t* path, int32_t path_len,
                       int32_t format,
                       const xl_column_spec* specs,
                       const xl_table* table,
                       const xl_write_options* options);

typedef struct xl_buffer {
    uint8_t* data;
    int64_t len;
} xl_buffer;

int32_t xl_write_typed_to_memory(int32_t format,
                                 const xl_column_spec* specs,
                                 const xl_table* table,
                                 const xl_write_options* options,
                                 xl_buffer* out_buffer);

void xl_free_buffer(xl_buffer* buffer);

int32_t xl_encrypt_package(const uint8_t* package_path, int32_t package_path_len,
                           const uint8_t* destination_path, int32_t destination_path_len,
                           const uint8_t* password, int32_t password_len);

typedef struct xl_inferred_schema {
    xl_column_spec* columns;
    int32_t column_count;
} xl_inferred_schema;

int32_t xl_infer_schema(xl_workbook* handle, int32_t header_row, int32_t sample_size, xl_inferred_schema* out_schema);

void xl_free_schema(xl_inferred_schema* schema);

typedef struct xl_writer_handle xl_writer_handle;

int32_t xl_open_write_handle(const uint8_t* path, int32_t path_len, int32_t format,
                             const xl_write_options* options, xl_writer_handle** out_handle);

int32_t xl_open_write_handle_to_memory(int32_t format, const xl_write_options* options,
                                       xl_writer_handle** out_handle);

int32_t xl_start_sheet(xl_writer_handle* handle, const uint8_t* name, int32_t name_len);

int32_t xl_start_row(xl_writer_handle* handle);

int32_t xl_write_string(xl_writer_handle* handle, const uint8_t* value, int32_t value_len);
int32_t xl_write_int64(xl_writer_handle* handle, int64_t value);
int32_t xl_write_float64(xl_writer_handle* handle, double value);
int32_t xl_write_bool(xl_writer_handle* handle, int32_t value);
int32_t xl_write_date(xl_writer_handle* handle, int32_t days_since_epoch);
int32_t xl_write_time(xl_writer_handle* handle, int64_t microseconds_since_midnight);
int32_t xl_write_timestamp(xl_writer_handle* handle, int64_t microseconds_since_epoch);
int32_t xl_write_null(xl_writer_handle* handle, int32_t type);

int32_t xl_end_row(xl_writer_handle* handle);
int32_t xl_end_sheet(xl_writer_handle* handle);
int32_t xl_close_write_handle(xl_writer_handle* handle);

int32_t xl_write_handle_bytes(xl_writer_handle* handle, xl_buffer* out_buffer);

/* Parallel CSV aggregation.
 *
 * The library partitions the source, seeds one accumulator per partition, calls `accumulate` for
 * every record on worker threads, then folds the partitions together in source order with
 * `combine`. The surviving accumulator is written to `out_state` and becomes the caller's to free.
 *
 * Ownership: every state the library seeds, the library frees with `free_state` - except the one
 * returned in `out_state`. `combine` folds `next` into `acc` and must NOT free `next`; the library
 * does that. `seed` must return a distinct pointer on every call. A NULL state is legal and is
 * never passed to `free_state`.
 *
 * `seed` is called at least once per partition and may be called again for a partition that has to
 * be re-read, so a seeded state can be discarded without ever being combined. Peak live states is
 * proportional to the partition count, not to `degree_of_parallelism`: a 40 GB source at 16 threads
 * holds roughly 640 states at once.
 *
 * Threading: `accumulate` runs concurrently on worker threads, each on its own state. `user_data`
 * is shared by all of them and must be read-only or internally synchronized. No callback ever runs
 * on the calling thread, not even when the source is too small to partition. A callback must not
 * call any xl_* function for the same aggregation, and must not let an exception or panic escape
 * into the library - wrap Rust shims in catch_unwind and mark C++ shims noexcept.
 *
 * A row's cell pointers are valid only for the duration of the `accumulate` call that received
 * them. Copy anything you keep.
 *
 * Returning nonzero from `accumulate` aborts the run: `out_state` is not written, every state is
 * freed, and that code is returned verbatim. Sibling workers only notice the abort every few
 * thousand records, so `accumulate` must tolerate being called again after it has returned nonzero.
 * Use positive codes; every code the library returns is <= 0.
 */
typedef int32_t (*xl_csv_seed_fn)      (void** out_state, void* user_data);
typedef int32_t (*xl_csv_accumulate_fn)(void* state, const xl_row* row, void* user_data);
typedef int32_t (*xl_csv_combine_fn)   (void* acc, void* next, void* user_data);
typedef void    (*xl_csv_free_state_fn)(void* state, void* user_data);

typedef struct xl_csv_aggregation {
    int32_t struct_size;
    xl_csv_seed_fn       seed;
    xl_csv_accumulate_fn accumulate;
    xl_csv_combine_fn    combine;
    xl_csv_free_state_fn free_state;
    void*                user_data;
} xl_csv_aggregation;

typedef struct xl_csv_parallel_options {
    int32_t struct_size;
    int32_t degree_of_parallelism;  /* 0 = processor count, 1 = sequential */
    int32_t header_row;             /* 1-based; 0 = no header */
    int32_t delimiter;              /* 0 = default (',') */
    int32_t quote;                  /* 0 = default ('"') */
    int32_t detect_bom;             /* XL_OPT_DEFAULT / XL_OPT_FALSE / XL_OPT_TRUE */
    int32_t max_cell_bytes;         /* 0 = default */
} xl_csv_parallel_options;

int32_t xl_csv_aggregate_file(const uint8_t* path, int32_t path_len,
                              const xl_csv_aggregation* agg,
                              const xl_csv_parallel_options* options,
                              void** out_state);

/* `data` must stay valid and unmodified until this call returns; the library does not copy it. */
int32_t xl_csv_aggregate_memory(const uint8_t* data, int32_t data_len,
                                const xl_csv_aggregation* agg,
                                const xl_csv_parallel_options* options,
                                void** out_state);

int32_t xl_last_error(uint8_t* buffer, int32_t capacity, int32_t* out_len);

const uint8_t* xl_last_error_ptr(int32_t* out_len);

int32_t xl_abi_version(void);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_H */
