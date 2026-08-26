/* ExcelReader C ABI - reading only.
 * Every function returns an XL_* status code. All strings are UTF-8 with an explicit length.
 * All integers in the row blob are little-endian int32.
 * No handle is thread-safe; use one handle per thread. */
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

/* ABI revision of this header. Bumped on any change to a struct layout, a status code, or the
 * meaning of an existing function; adding a new function does not bump it. A caller should refuse
 * to proceed if xl_abi_version() does not equal the XL_ABI_VERSION it was compiled against. */
#define XL_ABI_VERSION 3

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

/* Ceilings on what xl_parse_typed/xl_parse_arrow accept, taken from Excel's own limits: a request
 * naming more columns than a sheet holds (A..XFD), or a header name longer than a cell holds
 * (32,767 characters, at UTF-8's 4-byte worst case), cannot describe a real workbook. Both counts
 * size an allocation and drive a read over YOUR memory, so anything past them is rejected with
 * XL_INVALID_ARGUMENT rather than trusted. */
#define XL_MAX_COLUMN_SPECS      16384
#define XL_MAX_COLUMN_NAME_BYTES 131068

/* Opaque workbook handle. Never dereferenced by the caller; only ever passed back to xl_*. */
typedef struct xl_workbook xl_workbook;

/* One UTF-8 cell value returned by xl_read_all_decoded. value_len excludes the trailing NUL;
 * use it when the value could contain an embedded NUL. */
typedef struct xl_row_cell {
    int32_t column;
    int32_t type;
    int32_t value_len;
    const uint8_t* value;
} xl_row_cell;

/* A decoded row, as returned inside xl_rows by xl_read_all_decoded. The cells array and every
 * cell->value live in one allocation owned by the row; they stay valid until xl_free_rows frees
 * the enclosing xl_rows and must never be freed individually. Values are NUL-terminated in
 * addition to carrying value_len. */
typedef struct xl_row {
    int32_t cell_count;
    xl_row_cell* cells;
} xl_row;

/* Copies the path; the caller may free it on return. */
int32_t xl_open_file(const uint8_t* path, int32_t path_len, int32_t format, xl_workbook** out_handle);

/* Copies the data; the caller may free it on return. */
int32_t xl_open_memory(const uint8_t* data, int32_t data_len, int32_t format, xl_workbook** out_handle);

/* Every boolean-shaped xl_open_options field below uses one of these three states, never a plain
 * 0/1 - several of them default to true, so a bare 0 would be ambiguous between "off" and "use the
 * library default". */
#define XL_OPT_DEFAULT 0
#define XL_OPT_FALSE   1
#define XL_OPT_TRUE    2

/* Optional overrides for xl_open_file_ex/xl_open_memory_ex. Every numeric field is 0 for "use the
 * library default"; struct_size must equal sizeof(xl_open_options) exactly - a caller built against
 * a different header version gets XL_INVALID_ARGUMENT with detail in xl_last_error, not a silently
 * truncated/misread struct. This mirrors ExcelReader.Core.Reader.CsvReaderOptions/ExcelReaderOptions -
 * there is no field here that does not correspond to a real, already-existing Core option. */
typedef struct xl_open_options {
    int32_t struct_size;

    /* CSV only (format == XL_FORMAT_CSV); ignored for every other format. */
    int32_t csv_sniff_dialect;        /* XL_OPT_*; TRUE infers delimiter/quote/encoding from a leading
                                        * sample (see Excel.SniffCsvDialect* in Core) before opening -
                                        * csv_delimiter/csv_quote below are then ignored */
    int32_t csv_delimiter;            /* byte value 1-255, 0 = default ',' */
    int32_t csv_quote;                /* byte value 1-255, 0 = default '"' */
    int32_t csv_detect_bom;           /* XL_OPT_*; default TRUE */
    int32_t csv_max_cell_bytes;       /* 0 = default (32 MiB) */
    int32_t csv_intern_strings;       /* XL_OPT_*; default FALSE */

    /* XLS/XLSX/XLSB only; ignored for CSV. */
    int64_t max_total_decompressed_bytes; /* 0 = default (512 MiB) */
    int32_t max_cell_bytes;               /* 0 = default (32 MiB) */
    int64_t max_shared_string_bytes;      /* 0 = default (128 MiB) */
    int32_t max_zip_entries;              /* 0 = default (65536) */
    int32_t prefetch_decompression;       /* XL_OPT_*; default FALSE */
    int32_t intern_strings;               /* XL_OPT_*; default FALSE */
} xl_open_options;

/* `options` may be NULL, which is identical to xl_open_file. Copies the path; the caller may free it
 * on return. XL_INVALID_ARGUMENT for an unrecognized struct_size or an out-of-range field, detail in
 * xl_last_error. */
int32_t xl_open_file_ex(const uint8_t* path, int32_t path_len, int32_t format,
                        const xl_open_options* options, xl_workbook** out_handle);

/* `options` may be NULL, which is identical to xl_open_memory. Copies the data; the caller may free
 * it on return. Same validation as xl_open_file_ex. */
int32_t xl_open_memory_ex(const uint8_t* data, int32_t data_len, int32_t format,
                          const xl_open_options* options, xl_workbook** out_handle);

/* A handle is valid until this call succeeds. Once xl_close succeeds, the handle value is retired
 * permanently: any later call with it (including a second xl_close on the same value) returns
 * XL_INVALID_HANDLE, never undefined behavior. The caller should still null its own copy of the
 * pointer immediately after a successful xl_close - good practice, just no longer load-bearing for
 * memory safety. */
int32_t xl_close(xl_workbook* handle);

int32_t xl_sheet_count(xl_workbook* handle, int32_t* out_count);

/* On XL_BUFFER_TOO_SMALL, *out_len holds the required byte count. */
int32_t xl_sheet_name(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_len);

/* Name of the sheet at `index`, without changing the current sheet or disturbing row enumeration.
 * On XL_BUFFER_TOO_SMALL, *out_len holds the required byte count. XL_INVALID_ARGUMENT for a
 * negative index; an index >= xl_sheet_count is XL_ERROR with detail in xl_last_error. */
int32_t xl_sheet_name_at(xl_workbook* handle, int32_t index, uint8_t* buffer, int32_t capacity, int32_t* out_len);

/* Resets row enumeration to the start of the selected sheet, dropping any row held pending from a
 * prior XL_BUFFER_TOO_SMALL on xl_next_row - that row is not replayed after the sheet changes. */
int32_t xl_move_to_sheet(xl_workbook* handle, int32_t index);

int32_t xl_is_date1904(xl_workbook* handle, int32_t* out_flag);

/* Writes one row as:
 *     int32 cell_count
 *     repeated: int32 column, int32 type, int32 value_len, uint8 value[value_len]
 * Returns XL_EOF at the end of the sheet. On XL_BUFFER_TOO_SMALL, *out_written holds the required
 * byte count and the row is held until the next call - no row is ever lost.
 *
 * `capacity`/`*out_written` are int32_t: a single row whose blob would exceed INT32_MAX bytes cannot
 * be returned through this function. That is XL_ERROR (detail in xl_last_error), not silent
 * truncation. No real spreadsheet row approaches this size; xl_parse_typed uses int64_t lengths and
 * is columnar (and faster) if you ever need to plan around it regardless. */
int32_t xl_next_row(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

/* Writes every remaining row of the current sheet into `buffer` as:
 *     int32 row_count
 *     repeated row_count times:
 *         int32 row_length          (byte count of this row's blob, excluding this field)
 *         <row blob>                (int32 cell_count, then cell_count * {column,type,value_len,value})
 *
 * Returns XL_OK with *out_written set to the bytes used. On XL_BUFFER_TOO_SMALL, *out_written holds
 * a sufficient (not necessarily minimal) required capacity and NO rows have been lost - the
 * accumulated result is held until the next call, so the caller can retry with a bigger buffer.
 * An empty remainder is XL_OK with row_count == 0; XL_EOF is never returned here.
 *
 * `capacity`/`*out_written` are int32_t: a sheet whose ENTIRE accumulated blob would exceed
 * INT32_MAX bytes cannot be returned through this function - that is XL_ERROR (detail in
 * xl_last_error), not silent truncation or a wrapped/negative count. xl_parse_typed uses int64_t
 * lengths, is columnar, and is markedly faster - prefer it for a sheet anywhere near this size. */
int32_t xl_read_all_blob(xl_workbook* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

/* A decoded sheet returned by xl_read_all_decoded. Rows remain valid until xl_free_rows. */
typedef struct xl_rows {
    int32_t row_count;
    xl_row* rows;
} xl_rows;

/* Decodes every remaining row of the current sheet in one call, avoiding one native round-trip
 * per row. XL_EOF is never returned here - an empty remainder comes back as XL_OK with
 * row_count == 0. The caller owns the returned allocation and must call xl_free_rows. */
int32_t xl_read_all_decoded(xl_workbook* handle, xl_rows* out_rows);

/* Releases a result returned by xl_read_all_decoded and resets it to zero. Safe on a zeroed value. */
void xl_free_rows(xl_rows* rows);

#define XL_T_STRING    0
#define XL_T_I64       1
#define XL_T_F64       2
#define XL_T_BOOL      3
#define XL_T_DATE      4   /* days since 1970-01-01, int32 */
#define XL_T_TIME      5   /* microseconds since midnight, int64 */
#define XL_T_TIMESTAMP 6   /* microseconds since 1970-01-01T00:00:00Z, int64 */

/* Describes one output column of xl_parse_typed. Resolve by header name (names != NULL, matched
 * case-insensitively and trimmed against the header row, trying each candidate in `names` in order
 * and stopping at the first match) or by physical column index (name_count == 0). */
typedef struct xl_column_spec {
    const uint8_t* const* names; /* candidate header texts, UTF-8, in priority order; NULL when
                                   * name_count == 0 to match by index instead */
    const int32_t* name_lens;    /* one length per entry in `names` */
    int32_t name_count;
    int32_t index;                /* zero-based column index, used when name_count == 0 */
    int32_t type;                 /* XL_T_* */
    int32_t nullable;             /* 0 = a failed conversion is XL_ERROR; 1 = it becomes null (validity bit 0) */
} xl_column_spec;

/* One output column. `values` is the only allocation this column owns directly:
 *   - XL_T_STRING: an int32 offsets array (length+1 entries) followed immediately in the SAME
 *     allocation by the UTF-8 data blob - `data` is an interior pointer into `values`, and freeing it
 *     separately is a double free. `data_len` is the blob's byte length.
 *   - every other type: a dense array of `length` elements at that type's native width (int64 for
 *     XL_T_I64/TIME/TIMESTAMP, double for XL_T_F64, int32 for XL_T_DATE, uint8 0/1 for XL_T_BOOL).
 *     `data`/`data_len` are unused (NULL/0).
 * `validity` is a SEPARATE allocation: an LSB-first bitmap, 1 = valid, bit i = row i. NULL when no
 * value in the column was ever null (nothing to indicate). */
typedef struct xl_column {
    int32_t type;
    int64_t length;
    const void* values;
    const uint8_t* validity;
    const uint8_t* data;
    int64_t data_len;
} xl_column;

/* Result of xl_parse_typed. `columns` is one allocation of `column_count` xl_column values, in the
 * same order as the specs passed in; every column's own `values`/`validity` are separate allocations. */
typedef struct xl_table {
    int32_t column_count;
    int64_t row_count;
    xl_column* columns;
} xl_table;

/* Schema-driven columnar read of the WHOLE current sheet, from its first row - independent of, and
 * never disturbing, the row cursor xl_next_row/xl_read_all_blob share on `handle`.
 * `header_row` is the 1-based row number used to resolve name-based specs (rows before it are
 * skipped, and it is never itself yielded as a data row); 0 means "no header" and every spec must be
 * index-based. XL_INVALID_ARGUMENT for a bad spec (unknown type, negative index with no name, a
 * blank name, a name-based spec with header_row == 0), an unmatched header name, a `spec_count`
 * outside 1..XL_MAX_COLUMN_SPECS, or a `name_len` past XL_MAX_COLUMN_NAME_BYTES; XL_ERROR (detail in
 * xl_last_error) for a non-nullable column whose value failed to convert, or a sheet with fewer than
 * header_row rows. The caller owns the returned table and must call xl_free_table.
 *
 * Only read *out_table when the call returns XL_OK; on failure it is zeroed, and xl_free_table on a
 * zeroed table is a no-op. */
int32_t xl_parse_typed(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                       int32_t header_row, xl_table* out_table);

/* Releases a result returned by xl_parse_typed and resets it to zero. Safe on a zeroed value. */
void xl_free_table(xl_table* table);

/* ---- Writing --------------------------------------------------------------------------------- */

/* Optional overrides for xl_write_typed. Every numeric field is 0 for "use the library default";
 * struct_size must equal sizeof(xl_write_options) exactly, same contract as xl_open_options. */
typedef struct xl_write_options {
    int32_t struct_size;
    int32_t sheet_name_len;
    const uint8_t* sheet_name;    /* UTF-8; NULL = "Sheet1". Excel's rules: 1-31 characters, none of
                                    * : \ / ? * [ ] . Ignored for XL_FORMAT_CSV, and ignored entirely
                                    * by xl_open_write_handle - each sheet is named by its own
                                    * xl_start_sheet call instead. */

    /* CSV only; ignored for every other format. */
    int32_t csv_delimiter;        /* byte value 1-255, 0 = default ',' */
    int32_t csv_quote;            /* byte value 1-255, 0 = default '"' */

    /* XLS/XLSB only; ignored for XLSX and CSV. */
    int32_t date1904;             /* XL_OPT_*; default FALSE */

    /* XLSX/XLSB only; ignored for XLS and CSV. */
    int32_t use_shared_strings;   /* XL_OPT_*; default FALSE */
} xl_write_options;

/* Writes `table` to `path` as a single sheet, then closes the file. One-shot: no writer handle exists
 * before or after this call, and nothing in `specs`, `table` or `options` is retained.
 *
 * `specs` is a parallel array of table->column_count entries supplying each column's NAME. Only
 * name/name_len/type are read - `index` is ignored (columns are written left to right in array order)
 * and `nullable` is ignored (a column is nullable iff its xl_column.validity is non-NULL). spec.type
 * must equal the matching column.type, or XL_INVALID_ARGUMENT: the redundancy is checked rather than
 * silently resolved, so the two can never disagree.
 *
 * Column names are all-or-nothing: either every spec has name == NULL (no header row) or every spec
 * has a non-NULL name (one header row is written first). A mix is XL_INVALID_ARGUMENT.
 *
 * `format` must be XL_FORMAT_XLS/XLSX/XLSB/CSV. XL_FORMAT_AUTO is XL_INVALID_ARGUMENT: sniffing reads
 * a file's existing signature bytes, and a file being created has none.
 *
 * `options` may be NULL, meaning every default.
 *
 * Every buffer reachable from `specs`, `table` and `options` is BORROWED for the duration of the call
 * and never freed by this library. Unlike the xl_table xl_parse_typed produces, an INPUT column's
 * `data` need not be interior to its `values` allocation - the two are read independently.
 *
 * XL_T_DATE/TIME/TIMESTAMP columns are written with a number format attached (the builtin date style
 * for DATE, "hh:mm:ss" for TIME, "yyyy-mm-dd hh:mm:ss" for TIMESTAMP), so Excel shows a date rather
 * than a serial number. CSV ignores styling.
 *
 * On any failure the destination file may exist and be incomplete; the caller owns cleaning it up.
 * Detail is in xl_last_error. */
int32_t xl_write_typed(const uint8_t* path, int32_t path_len,
                       int32_t format,
                       const xl_column_spec* specs,
                       const xl_table* table,
                       const xl_write_options* options);

/* An owned block of memory returned by xl_write_typed_to_memory or xl_write_handle_bytes. The
 * caller must release it with xl_free_buffer - same convention as xl_table/xl_inferred_schema. */
typedef struct xl_buffer {
    uint8_t* data;
    int64_t len;
} xl_buffer;

/* Same as xl_write_typed, except there is no path: the workbook is built in memory and returned as
 * out_buffer instead of being written to disk. Every other parameter and validation rule is
 * identical.
 *
 * Only read *out_buffer when the call returns XL_OK; on failure it is zeroed, and xl_free_buffer on
 * a zeroed buffer is a no-op. */
int32_t xl_write_typed_to_memory(int32_t format,
                                 const xl_column_spec* specs,
                                 const xl_table* table,
                                 const xl_write_options* options,
                                 xl_buffer* out_buffer);

/* Releases a buffer returned by xl_write_typed_to_memory or xl_write_handle_bytes and resets it to
 * zero. Safe on a zeroed value. */
void xl_free_buffer(xl_buffer* buffer);

/* Result of xl_infer_schema: one xl_column_spec per column the sheet appears to have, in ascending
 * column order. Each spec's `type`/`nullable` is a guess from the sampled cells' own XL_CELL_* tags
 * (no text sniffing) and can be handed straight to xl_parse_typed/xl_parse_arrow; `name` is set from
 * the header row, or NULL (resolve by `index`) when header_row == 0, the header cell was blank, or
 * the column was never seen in the header at all. `columns` is one allocation of `column_count`
 * xl_column_spec values; each non-NULL `name` is its own separate allocation. */
typedef struct xl_inferred_schema {
    xl_column_spec* columns;
    int32_t column_count;
} xl_inferred_schema;

/* Guesses a xl_parse_typed/xl_parse_arrow schema by sampling the WHOLE current sheet, from its first
 * row - independent of, and never disturbing, the row cursor xl_next_row/xl_read_all_blob share
 * on `handle`. `header_row` has the same meaning as in xl_parse_typed (0 = no
 * header). `sample_size` bounds how many rows after the header are inspected; a column is guessed
 * XL_T_STRING with nullable = 1 when its sampled cells mix kinds, are all XL_CELL_FORMULA/ERROR, or
 * were never populated. XL_T_I64 vs XL_T_F64 is decided by whether every sampled numeric cell parses
 * as an integer. `nullable` is 1 when any sampled row left the column empty (including a row
 * narrower than the widest one seen). This is a guess over a sample, not a guarantee - always check
 * it fits before trusting it against the full sheet.
 *
 * XL_INVALID_ARGUMENT for a negative header_row, a non-positive sample_size, or a sheet with fewer
 * than header_row rows. The caller owns the returned schema and must call xl_free_schema.
 *
 * Only read *out_schema when the call returns XL_OK; on failure it is zeroed, and xl_free_schema on a
 * zeroed schema is a no-op. */
int32_t xl_infer_schema(xl_workbook* handle, int32_t header_row, int32_t sample_size, xl_inferred_schema* out_schema);

/* Releases a result returned by xl_infer_schema and resets it to zero. Safe on a zeroed value. */
void xl_free_schema(xl_inferred_schema* schema);

/* Streaming writer handle: one sheet and one row open at a time, written directly to disk as each
 * call arrives instead of building an xl_table in memory first (see xl_write_typed for that
 * alternative). The required call order is:
 *
 *   xl_open_write_handle
 *     xl_start_sheet
 *       xl_start_row
 *         xl_write_string / xl_write_int64 / xl_write_float64 / xl_write_bool /
 *         xl_write_date / xl_write_time / xl_write_timestamp / xl_write_null   (one call per cell,
 *                                                                               left to right)
 *       xl_end_row                                                            (repeat per row)
 *     xl_end_sheet                                                            (repeat per sheet)
 *   xl_close_write_handle
 *
 * Calling one of these out of order (e.g. a cell write before xl_start_row, or xl_start_sheet
 * again before xl_end_sheet) returns XL_ERROR with a message from xl_last_error/xl_last_error_ptr;
 * the handle itself is still usable afterward - fix the call order and continue, or give up and
 * call xl_close_write_handle to discard it.
 *
 * xl_close_write_handle must be called exactly once to produce a valid file: it implicitly closes
 * any row/sheet still open and writes the workbook's trailing structure (the zip central directory
 * for XLSX/XLSB, the BIFF EOF record for XLS). It always releases the handle - including on
 * XL_ERROR - so *handle must not be used again after calling it, successful or not.
 *
 * xl_open_write_handle_to_memory is the same handle, backed by an in-memory buffer instead of a
 * file: every xl_start_sheet/xl_start_row/xl_write_xxx/xl_end_row/xl_end_sheet/xl_close_write_handle
 * call above works identically on it. Call xl_write_handle_bytes to read the buffer out - it
 * implicitly finishes the workbook's trailing structure the same way xl_close_write_handle does
 * (so it is safe to call whether or not every sheet/row was explicitly ended), but unlike
 * xl_close_write_handle it does NOT release the handle: call xl_close_write_handle afterward, same
 * as for a file-backed handle. Calling xl_write_handle_bytes on a file-backed handle (one opened by
 * xl_open_write_handle) is XL_INVALID_ARGUMENT. */
typedef struct xl_writer_handle xl_writer_handle;

/* Creates path (truncating it if it already exists) and returns a handle for it. options may be
 * NULL for every default, same convention as xl_write_typed's options parameter. format must be
 * one of XL_FORMAT_XLS/XLSX/XLSB/CSV (XL_FORMAT_AUTO is rejected - see xl_write_typed). */
int32_t xl_open_write_handle(const uint8_t* path, int32_t path_len, int32_t format,
                             const xl_write_options* options, xl_writer_handle** out_handle);

/* Same as xl_open_write_handle, except there is no path: the handle is backed by an in-memory
 * buffer, read out with xl_write_handle_bytes - see the call-order note above. */
int32_t xl_open_write_handle_to_memory(int32_t format, const xl_write_options* options,
                                       xl_writer_handle** out_handle);

/* Starts a new sheet named name (UTF-8, name_len bytes). Must not be called again before the
 * current sheet, if any, has been ended with xl_end_sheet. */
int32_t xl_start_sheet(xl_writer_handle* handle, const uint8_t* name, int32_t name_len);

/* Starts a new row on the current sheet. Must not be called again before the current row, if any,
 * has been ended with xl_end_row. */
int32_t xl_start_row(xl_writer_handle* handle);

/* Writes the next cell of the current row as text, or a blank cell when value is NULL. */
int32_t xl_write_string(xl_writer_handle* handle, const uint8_t* value, int32_t value_len);
/* Writes the next cell of the current row as an integer. */
int32_t xl_write_int64(xl_writer_handle* handle, int64_t value);
/* Writes the next cell of the current row as a floating-point number. */
int32_t xl_write_float64(xl_writer_handle* handle, double value);
/* Writes the next cell of the current row as a boolean (0/nonzero). */
int32_t xl_write_bool(xl_writer_handle* handle, int32_t value);
/* Writes the next cell of the current row as a date. See XL_T_DATE for the wire format. */
int32_t xl_write_date(xl_writer_handle* handle, int32_t days_since_epoch);
/* Writes the next cell of the current row as a time of day. See XL_T_TIME for the wire format. */
int32_t xl_write_time(xl_writer_handle* handle, int64_t microseconds_since_midnight);
/* Writes the next cell of the current row as a date/time. See XL_T_TIMESTAMP for the wire format. */
int32_t xl_write_timestamp(xl_writer_handle* handle, int64_t microseconds_since_epoch);
/* Writes the next cell of the current row as a blank cell of the given XL_T_* type. */
int32_t xl_write_null(xl_writer_handle* handle, int32_t type);

/* Ends the current row, started by xl_start_row. */
int32_t xl_end_row(xl_writer_handle* handle);
/* Ends the current sheet, started by xl_start_sheet. Must not be called with a row still open. */
int32_t xl_end_sheet(xl_writer_handle* handle);
/* Finishes and releases handle - see the call-order note above. */
int32_t xl_close_write_handle(xl_writer_handle* handle);

/* Reads back everything written so far to a handle opened by xl_open_write_handle_to_memory - see
 * the call-order note above for what this does and does not do to handle. XL_INVALID_ARGUMENT if
 * handle was opened by xl_open_write_handle instead.
 *
 * Only read *out_buffer when the call returns XL_OK; on failure it is zeroed, and xl_free_buffer on
 * a zeroed buffer is a no-op. */
int32_t xl_write_handle_bytes(xl_writer_handle* handle, xl_buffer* out_buffer);

/* Last error on the CALLING thread. */
int32_t xl_last_error(uint8_t* buffer, int32_t capacity, int32_t* out_len);

/* Borrowed pointer to the calling thread's last error message, UTF-8, NOT NUL-terminated - use
 * *out_len. Returns NULL with *out_len == 0 when there is no error. The pointer stays valid until
 * the next ExcelReader call ON THIS THREAD. Copy it if you need it longer. xl_last_error is kept
 * for callers that prefer to own the buffer. */
const uint8_t* xl_last_error_ptr(int32_t* out_len);

/* ABI revision of the loaded library. See XL_ABI_VERSION above. */
int32_t xl_abi_version(void);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_H */
