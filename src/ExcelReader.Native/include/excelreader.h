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

/* One UTF-8 cell value returned by xl_next_row_decoded. value_len excludes the trailing NUL;
 * use it when the value could contain an embedded NUL. */
typedef struct xl_row_cell {
    int32_t column;
    int32_t type;
    int32_t value_len;
    const uint8_t* value;
} xl_row_cell;

/* A decoded row returned by xl_next_row_decoded. Its contents remain valid until xl_free_row. */
typedef struct xl_row {
    int32_t cell_count;
    xl_row_cell* cells;
} xl_row;

/* Copies the path; the caller may free it on return. */
int32_t xl_open_file(const uint8_t* path, int32_t path_len, int32_t format, void** out_handle);

/* Copies the data; the caller may free it on return. */
int32_t xl_open_memory(const uint8_t* data, int32_t data_len, int32_t format, void** out_handle);

/* A handle is valid until this call succeeds. Calling anything on a handle afterwards - including a
 * second xl_close on the same value - is undefined behavior; the caller must null its own copy of
 * the pointer immediately after a successful xl_close so it can never be reused or double-closed. */
int32_t xl_close(void* handle);

int32_t xl_sheet_count(void* handle, int32_t* out_count);

/* On XL_BUFFER_TOO_SMALL, *out_len holds the required byte count. */
int32_t xl_sheet_name(void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_len);

/* Resets row enumeration to the start of the selected sheet, dropping any row held pending from a
 * prior XL_BUFFER_TOO_SMALL on xl_next_row - that row is not replayed after the sheet changes. */
int32_t xl_move_to_sheet(void* handle, int32_t index);

int32_t xl_is_date1904(void* handle, int32_t* out_flag);

/* Writes one row as:
 *     int32 cell_count
 *     repeated: int32 column, int32 type, int32 value_len, uint8 value[value_len]
 * Returns XL_EOF at the end of the sheet. On XL_BUFFER_TOO_SMALL, *out_written holds the required
 * byte count and the row is held until the next call - no row is ever lost. */
int32_t xl_next_row(void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

/* Reads one row into C-friendly structs. Returns XL_EOF at the end of the sheet. The caller owns
 * the returned allocation and must call xl_free_row, including after partially processing a row. */
int32_t xl_next_row_decoded(void* handle, xl_row* out_row);

/* Releases a row returned by xl_next_row_decoded and resets it to zero. Safe to call on a zeroed row. */
void xl_free_row(xl_row* row);

/* A decoded sheet returned by xl_read_all_decoded. Rows remain valid until xl_free_rows. */
typedef struct xl_rows {
    int32_t row_count;
    xl_row* rows;
} xl_rows;

/* Decodes every remaining row of the current sheet in one call, avoiding one native round-trip
 * per row. XL_EOF is never returned here - an empty remainder comes back as XL_OK with
 * row_count == 0. The caller owns the returned allocation and must call xl_free_rows. */
int32_t xl_read_all_decoded(void* handle, xl_rows* out_rows);

/* Releases a result returned by xl_read_all_decoded and resets it to zero. Safe on a zeroed value. */
void xl_free_rows(xl_rows* rows);

/* Last error on the CALLING thread. */
int32_t xl_last_error(uint8_t* buffer, int32_t capacity, int32_t* out_len);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_H */
