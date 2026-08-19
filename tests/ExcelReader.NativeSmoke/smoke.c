/* Real C consumer of the ExcelReader ABI, run in CI on Windows/Linux/macOS (see
 * .github/workflows/python.yml). Two jobs:
 *
 *   1. Compile-time _STATIC_ASSERTs (below) pin every ABI struct's layout against the C standard's
 *      own natural-alignment rules, on whatever compiler builds this file. Catches an accidental
 *      field reorder/insertion/padding change in excelreader.h itself.
 *   2. The runtime checks in main() call the real published library and assert on real values.
 *      Layer 1 alone cannot prove excelreader.h and the C# side (NativeColumn.cs, NativeRow.cs, ...)
 *      agree — only running real data through the real exports can. A mismatch there produces
 *      garbage values, and these assertions fail on the values, not on a crash.
 *
 * The library is loaded dynamically (LoadLibrary/dlopen) rather than linked at build time. This is
 * deliberate, not a shortcut: NativeAOT's publish output ships no `ExcelReader.Native.lib` import
 * library on Windows, so a normal `target_link_libraries` against the DLL does not work with MSVC
 * out of the box (see docs/NATIVE_HARDENING_PLAN.md, task C, H3). Dynamic loading works identically
 * on all three platforms and needs nothing beyond the shared library file itself.
 */
#include "excelreader.h"
#include "excelreader_arrow.h"

#include <stdint.h>
#include <stddef.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#ifdef _WIN32
#include <windows.h>
typedef HMODULE xl_lib_handle;
#else
#include <dlfcn.h>
typedef void* xl_lib_handle;
#endif

/* ---- Layer 1: struct layout static asserts --------------------------------------------------- */

#define XL_STATIC_ASSERT(cond, name) typedef char xl_static_assert_##name[(cond) ? 1 : -1]

/* Every published RID (win-x64, linux-x64, osx-x64, osx-arm64) is a 64-bit target with 8-byte
 * pointers and standard natural-alignment struct layout (no #pragma pack anywhere in the header) -
 * the offsets below follow directly from that, on any conforming compiler. */
XL_STATIC_ASSERT(sizeof(void*) == 8, pointer_is_64_bit);

XL_STATIC_ASSERT(offsetof(xl_row_cell, column) == 0, row_cell_column);
XL_STATIC_ASSERT(offsetof(xl_row_cell, type) == 4, row_cell_type);
XL_STATIC_ASSERT(offsetof(xl_row_cell, value_len) == 8, row_cell_value_len);
XL_STATIC_ASSERT(offsetof(xl_row_cell, value) == 16, row_cell_value);
XL_STATIC_ASSERT(sizeof(xl_row_cell) == 24, row_cell_size);

XL_STATIC_ASSERT(offsetof(xl_row, cell_count) == 0, row_cell_count);
XL_STATIC_ASSERT(offsetof(xl_row, cells) == 8, row_cells);
XL_STATIC_ASSERT(sizeof(xl_row) == 16, row_size);

XL_STATIC_ASSERT(offsetof(xl_rows, row_count) == 0, rows_row_count);
XL_STATIC_ASSERT(offsetof(xl_rows, rows) == 8, rows_rows);
XL_STATIC_ASSERT(sizeof(xl_rows) == 16, rows_size);

XL_STATIC_ASSERT(offsetof(xl_column_spec, names) == 0, column_spec_names);
XL_STATIC_ASSERT(offsetof(xl_column_spec, name_lens) == 8, column_spec_name_lens);
XL_STATIC_ASSERT(offsetof(xl_column_spec, name_count) == 16, column_spec_name_count);
XL_STATIC_ASSERT(offsetof(xl_column_spec, index) == 20, column_spec_index);
XL_STATIC_ASSERT(offsetof(xl_column_spec, type) == 24, column_spec_type);
XL_STATIC_ASSERT(offsetof(xl_column_spec, nullable) == 28, column_spec_nullable);
XL_STATIC_ASSERT(sizeof(xl_column_spec) == 32, column_spec_size);

XL_STATIC_ASSERT(offsetof(xl_write_options, struct_size) == 0, write_options_struct_size);
XL_STATIC_ASSERT(offsetof(xl_write_options, sheet_name_len) == 4, write_options_sheet_name_len);
XL_STATIC_ASSERT(offsetof(xl_write_options, sheet_name) == 8, write_options_sheet_name);
XL_STATIC_ASSERT(offsetof(xl_write_options, csv_delimiter) == 16, write_options_csv_delimiter);
XL_STATIC_ASSERT(offsetof(xl_write_options, csv_quote) == 20, write_options_csv_quote);
XL_STATIC_ASSERT(offsetof(xl_write_options, date1904) == 24, write_options_date1904);
XL_STATIC_ASSERT(offsetof(xl_write_options, use_shared_strings) == 28, write_options_use_shared_strings);
XL_STATIC_ASSERT(sizeof(xl_write_options) == 32, write_options_size);

XL_STATIC_ASSERT(offsetof(xl_column, type) == 0, column_type);
XL_STATIC_ASSERT(offsetof(xl_column, length) == 8, column_length);
XL_STATIC_ASSERT(offsetof(xl_column, values) == 16, column_values);
XL_STATIC_ASSERT(offsetof(xl_column, validity) == 24, column_validity);
XL_STATIC_ASSERT(offsetof(xl_column, data) == 32, column_data);
XL_STATIC_ASSERT(offsetof(xl_column, data_len) == 40, column_data_len);
XL_STATIC_ASSERT(sizeof(xl_column) == 48, column_size);

XL_STATIC_ASSERT(offsetof(xl_table, column_count) == 0, table_column_count);
XL_STATIC_ASSERT(offsetof(xl_table, row_count) == 8, table_row_count);
XL_STATIC_ASSERT(offsetof(xl_table, columns) == 16, table_columns);
XL_STATIC_ASSERT(sizeof(xl_table) == 24, table_size);

XL_STATIC_ASSERT(offsetof(xl_inferred_schema, columns) == 0, inferred_schema_columns);
XL_STATIC_ASSERT(offsetof(xl_inferred_schema, column_count) == 8, inferred_schema_column_count);
XL_STATIC_ASSERT(sizeof(xl_inferred_schema) == 16, inferred_schema_size);

XL_STATIC_ASSERT(offsetof(struct ArrowSchema, format) == 0, arrow_schema_format);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, name) == 8, arrow_schema_name);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, metadata) == 16, arrow_schema_metadata);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, flags) == 24, arrow_schema_flags);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, n_children) == 32, arrow_schema_n_children);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, children) == 40, arrow_schema_children);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, dictionary) == 48, arrow_schema_dictionary);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, release) == 56, arrow_schema_release);
XL_STATIC_ASSERT(offsetof(struct ArrowSchema, private_data) == 64, arrow_schema_private_data);
XL_STATIC_ASSERT(sizeof(struct ArrowSchema) == 72, arrow_schema_size);

XL_STATIC_ASSERT(offsetof(struct ArrowArray, length) == 0, arrow_array_length);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, null_count) == 8, arrow_array_null_count);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, offset) == 16, arrow_array_offset);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, n_buffers) == 24, arrow_array_n_buffers);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, n_children) == 32, arrow_array_n_children);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, buffers) == 40, arrow_array_buffers);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, children) == 48, arrow_array_children);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, dictionary) == 56, arrow_array_dictionary);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, release) == 64, arrow_array_release);
XL_STATIC_ASSERT(offsetof(struct ArrowArray, private_data) == 72, arrow_array_private_data);
XL_STATIC_ASSERT(sizeof(struct ArrowArray) == 80, arrow_array_size);

/* ---- Dynamic loading -------------------------------------------------------------------------- */

static xl_lib_handle load_library(const char* path)
{
#ifdef _WIN32
    return LoadLibraryA(path);
#else
    return dlopen(path, RTLD_NOW);
#endif
}

static void* load_symbol(xl_lib_handle lib, const char* name)
{
#ifdef _WIN32
    return (void*)GetProcAddress(lib, name);
#else
    return dlsym(lib, name);
#endif
}

typedef int32_t (*xl_abi_version_fn)(void);
typedef int32_t (*xl_open_file_fn)(const uint8_t*, int32_t, int32_t, xl_workbook**);
typedef int32_t (*xl_open_file_ex_fn)(const uint8_t*, int32_t, int32_t, const xl_open_options*, xl_workbook**);
typedef int32_t (*xl_close_fn)(xl_workbook*);
typedef int32_t (*xl_sheet_count_fn)(xl_workbook*, int32_t*);
typedef int32_t (*xl_sheet_name_fn)(xl_workbook*, uint8_t*, int32_t, int32_t*);
typedef int32_t (*xl_sheet_name_at_fn)(xl_workbook*, int32_t, uint8_t*, int32_t, int32_t*);
typedef int32_t (*xl_is_date1904_fn)(xl_workbook*, int32_t*);
typedef int32_t (*xl_next_row_fn)(xl_workbook*, uint8_t*, int32_t, int32_t*);
typedef int32_t (*xl_read_all_blob_fn)(xl_workbook*, uint8_t*, int32_t, int32_t*);
typedef int32_t (*xl_next_row_decoded_fn)(xl_workbook*, xl_row*);
typedef void (*xl_free_row_fn)(xl_row*);
typedef int32_t (*xl_read_all_decoded_fn)(xl_workbook*, xl_rows*);
typedef void (*xl_free_rows_fn)(xl_rows*);
typedef int32_t (*xl_parse_typed_fn)(xl_workbook*, const xl_column_spec*, int32_t, int32_t, xl_table*);
typedef void (*xl_free_table_fn)(xl_table*);
typedef int32_t (*xl_infer_schema_fn)(xl_workbook*, int32_t, int32_t, xl_inferred_schema*);
typedef void (*xl_free_schema_fn)(xl_inferred_schema*);
typedef const uint8_t* (*xl_last_error_ptr_fn)(int32_t*);
typedef int32_t (*xl_parse_arrow_fn)(xl_workbook*, const xl_column_spec*, int32_t, int32_t, struct ArrowArray*, struct ArrowSchema*);
typedef int32_t (*xl_write_typed_fn)(const uint8_t*, int32_t, int32_t, const xl_column_spec*,
                                     const xl_table*, const xl_write_options*);

typedef struct
{
    xl_abi_version_fn abi_version;
    xl_open_file_fn open_file;
    xl_open_file_ex_fn open_file_ex;
    xl_close_fn close_;
    xl_sheet_count_fn sheet_count;
    xl_sheet_name_fn sheet_name;
    xl_sheet_name_at_fn sheet_name_at;
    xl_is_date1904_fn is_date1904;
    xl_next_row_fn next_row;
    xl_read_all_blob_fn read_all_blob;
    xl_next_row_decoded_fn next_row_decoded;
    xl_free_row_fn free_row;
    xl_read_all_decoded_fn read_all_decoded;
    xl_free_rows_fn free_rows;
    xl_parse_typed_fn parse_typed;
    xl_free_table_fn free_table;
    xl_infer_schema_fn infer_schema;
    xl_free_schema_fn free_schema;
    xl_last_error_ptr_fn last_error_ptr;
    xl_parse_arrow_fn parse_arrow;
    xl_write_typed_fn write_typed;
} api_t;

#define BIND(field, type, name)                                                                     \
    do                                                                                              \
    {                                                                                                \
        api->field = (type)load_symbol(lib, name);                                                  \
        if (!api->field)                                                                             \
        {                                                                                            \
            fprintf(stderr, "FAIL: missing export %s\n", name);                                      \
            return 0;                                                                                \
        }                                                                                             \
    } while (0)

static int bind_all(xl_lib_handle lib, api_t* api)
{
    BIND(abi_version, xl_abi_version_fn, "xl_abi_version");
    BIND(open_file, xl_open_file_fn, "xl_open_file");
    BIND(open_file_ex, xl_open_file_ex_fn, "xl_open_file_ex");
    BIND(close_, xl_close_fn, "xl_close");
    BIND(sheet_count, xl_sheet_count_fn, "xl_sheet_count");
    BIND(sheet_name, xl_sheet_name_fn, "xl_sheet_name");
    BIND(sheet_name_at, xl_sheet_name_at_fn, "xl_sheet_name_at");
    BIND(is_date1904, xl_is_date1904_fn, "xl_is_date1904");
    BIND(next_row, xl_next_row_fn, "xl_next_row");
    BIND(read_all_blob, xl_read_all_blob_fn, "xl_read_all_blob");
    BIND(next_row_decoded, xl_next_row_decoded_fn, "xl_next_row_decoded");
    BIND(free_row, xl_free_row_fn, "xl_free_row");
    BIND(read_all_decoded, xl_read_all_decoded_fn, "xl_read_all_decoded");
    BIND(free_rows, xl_free_rows_fn, "xl_free_rows");
    BIND(parse_typed, xl_parse_typed_fn, "xl_parse_typed");
    BIND(free_table, xl_free_table_fn, "xl_free_table");
    BIND(infer_schema, xl_infer_schema_fn, "xl_infer_schema");
    BIND(free_schema, xl_free_schema_fn, "xl_free_schema");
    BIND(last_error_ptr, xl_last_error_ptr_fn, "xl_last_error_ptr");
    BIND(parse_arrow, xl_parse_arrow_fn, "xl_parse_arrow");
    BIND(write_typed, xl_write_typed_fn, "xl_write_typed");
    return 1;
}

/* ---- Layer 2: runtime checks against the real library ---------------------------------------- */

#define CHECK(cond, msg)                                                                            \
    do                                                                                              \
    {                                                                                                \
        if (!(cond))                                                                                 \
        {                                                                                             \
            fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__);                          \
            return 1;                                                                                 \
        }                                                                                              \
    } while (0)

/* Fills `spec` as a single-candidate name-based spec, using `name_slot`/`len_slot` as the
 * one-element backing storage `spec->names`/`spec->name_lens` point into — that storage must
 * outlive every use of `spec` (the caller declares it in the same or an outer scope). */
static void set_spec_name1(xl_column_spec* spec, const uint8_t** name_slot, int32_t* len_slot, const char* text)
{
    *name_slot = (const uint8_t*)text;
    *len_slot = (int32_t)strlen(text);
    spec->names = name_slot;
    spec->name_lens = len_slot;
    spec->name_count = 1;
}

static int32_t open_fixture(const api_t* api, const char* fixture, xl_workbook** out_handle)
{
    size_t path_len = strlen(fixture);
    return api->open_file((const uint8_t*)fixture, (int32_t)path_len, XL_FORMAT_XLSB, out_handle);
}

static int test_abi_version(const api_t* api)
{
    CHECK(api->abi_version() == XL_ABI_VERSION, "xl_abi_version() must equal the header's XL_ABI_VERSION");
    return 0;
}

static int test_open_missing_file_reports_an_error(const api_t* api)
{
    xl_workbook* handle = NULL;
    const char* missing = "this-file-does-not-exist.xlsb";
    int32_t status = api->open_file((const uint8_t*)missing, (int32_t)strlen(missing), XL_FORMAT_XLSB, &handle);
    CHECK(status != XL_OK, "opening a missing file must not report XL_OK");
    CHECK(handle == NULL, "a failed open must not hand back a handle");

    int32_t error_len = 0;
    const uint8_t* message = api->last_error_ptr(&error_len);
    CHECK(message != NULL && error_len > 0, "xl_last_error_ptr must report detail for the failed open");
    return 0;
}

static int test_sheets_and_flags(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file on the fixture must succeed");

    int32_t sheet_count = 0;
    CHECK(api->sheet_count(handle, &sheet_count) == XL_OK, "xl_sheet_count must succeed");
    CHECK(sheet_count == 1, "RealExcel.xlsb has exactly one sheet");

    uint8_t name_buffer[64];
    int32_t name_len = 0;
    CHECK(api->sheet_name(handle, name_buffer, 0, &name_len) == XL_BUFFER_TOO_SMALL,
          "xl_sheet_name with capacity 0 must report XL_BUFFER_TOO_SMALL");
    CHECK(name_len > 0, "xl_sheet_name must report the required length on XL_BUFFER_TOO_SMALL");
    CHECK((size_t)name_len < sizeof(name_buffer), "test fixture's sheet name must fit the local buffer");
    CHECK(api->sheet_name(handle, name_buffer, (int32_t)sizeof(name_buffer), &name_len) == XL_OK,
          "xl_sheet_name must succeed once the buffer is big enough");
    CHECK(name_len == 9 && memcmp(name_buffer, "Planilha1", 9) == 0, "sheet name must be Planilha1");

    CHECK(api->sheet_name_at(handle, 0, name_buffer, (int32_t)sizeof(name_buffer), &name_len) == XL_OK,
          "xl_sheet_name_at(0) must succeed");
    CHECK(name_len == 9 && memcmp(name_buffer, "Planilha1", 9) == 0, "xl_sheet_name_at(0) must match xl_sheet_name");

    int32_t date1904 = -1;
    CHECK(api->is_date1904(handle, &date1904) == XL_OK, "xl_is_date1904 must succeed");
    CHECK(date1904 == 0, "RealExcel.xlsb does not use the 1904 date system");

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed on a live handle");
    return 0;
}

static int test_next_row_blob_and_growth(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    uint8_t tiny[1];
    int32_t written = 0;
    int32_t status = api->next_row(handle, tiny, 0, &written);
    CHECK(status == XL_BUFFER_TOO_SMALL, "xl_next_row with capacity 0 must report XL_BUFFER_TOO_SMALL");
    CHECK(written > 0, "xl_next_row must report the required size on XL_BUFFER_TOO_SMALL");

    uint8_t* buffer = (uint8_t*)malloc((size_t)written);
    CHECK(buffer != NULL, "test allocation must succeed");
    int32_t capacity = written;
    status = api->next_row(handle, buffer, capacity, &written);
    CHECK(status == XL_OK, "retrying xl_next_row with the reported size must succeed - no row is lost");

    int32_t cell_count = 0;
    memcpy(&cell_count, buffer, sizeof(int32_t));
    CHECK(cell_count == 18, "RealExcel.xlsb's header row has 18 columns");

    int32_t column = 0, type = 0, value_len = 0;
    memcpy(&column, buffer + 4, sizeof(int32_t));
    memcpy(&type, buffer + 8, sizeof(int32_t));
    memcpy(&value_len, buffer + 12, sizeof(int32_t));
    CHECK(column == 0, "first cell's column index must be 0");
    CHECK(type == XL_CELL_STRING, "the header row's cells are strings");
    CHECK(value_len == 7 && memcmp(buffer + 16, "Coluna1", 7) == 0, "first header cell must read Coluna1");
    free(buffer);

    /* Drain the rest (100 data rows) and confirm the total, then confirm XL_EOF at the end. */
    int row_count = 1;
    uint8_t scratch[4096];
    for (;;)
    {
        status = api->next_row(handle, scratch, (int32_t)sizeof(scratch), &written);
        if (status == XL_EOF)
        {
            break;
        }
        CHECK(status == XL_OK, "xl_next_row must succeed for every remaining row of this fixture");
        row_count++;
    }
    CHECK(row_count == 101, "RealExcel.xlsb has 101 rows (1 header + 100 data)");

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_next_row_decoded(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    xl_row row;
    memset(&row, 0, sizeof(row));
    CHECK(api->next_row_decoded(handle, &row) == XL_OK, "xl_next_row_decoded must succeed for the header row");
    CHECK(row.cell_count == 18, "the header row has 18 cells");
    CHECK(row.cells[0].value_len == 7 && memcmp(row.cells[0].value, "Coluna1", 7) == 0,
          "first decoded cell must read Coluna1");
    api->free_row(&row);

    int row_count = 1;
    for (;;)
    {
        xl_row next;
        memset(&next, 0, sizeof(next));
        int32_t status = api->next_row_decoded(handle, &next);
        if (status == XL_EOF)
        {
            break;
        }
        CHECK(status == XL_OK, "xl_next_row_decoded must succeed for every remaining row");
        api->free_row(&next);
        row_count++;
    }
    CHECK(row_count == 101, "xl_next_row_decoded must see the same 101 rows as xl_next_row");

    /* Documented safe on a zeroed value. */
    xl_row zeroed;
    memset(&zeroed, 0, sizeof(zeroed));
    api->free_row(&zeroed);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_read_all_blob_and_decoded(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    /* static, not stack-local: 1 MiB blows past MSVC's default 1 MiB thread stack reserve
     * and faults with a stack overflow in Release builds. */
    static uint8_t buffer[1 << 20];
    int32_t written = 0;
    CHECK(api->read_all_blob(handle, buffer, (int32_t)sizeof(buffer), &written) == XL_OK,
          "xl_read_all_blob must succeed with a generously sized buffer");
    int32_t row_count = 0;
    memcpy(&row_count, buffer, sizeof(int32_t));
    CHECK(row_count == 101, "xl_read_all_blob must report all 101 rows");

    /* The sheet is now fully drained. A second call must be XL_OK with row_count == 0, never XL_EOF -
     * xl_read_all_blob never returns XL_EOF, by contract. */
    CHECK(api->read_all_blob(handle, buffer, (int32_t)sizeof(buffer), &written) == XL_OK,
          "a drained xl_read_all_blob call must still be XL_OK");
    memcpy(&row_count, buffer, sizeof(int32_t));
    CHECK(row_count == 0, "a drained sheet's xl_read_all_blob must report zero rows");

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");

    /* Fresh handle for xl_read_all_decoded, so this is not entangled with the blob drain above. */
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");
    xl_rows rows;
    memset(&rows, 0, sizeof(rows));
    CHECK(api->read_all_decoded(handle, &rows) == XL_OK, "xl_read_all_decoded must succeed");
    CHECK(rows.row_count == 101, "xl_read_all_decoded must report all 101 rows");
    api->free_rows(&rows);

    xl_rows drained;
    memset(&drained, 0, sizeof(drained));
    CHECK(api->read_all_decoded(handle, &drained) == XL_OK, "a drained xl_read_all_decoded call must still be XL_OK");
    CHECK(drained.row_count == 0, "a drained sheet's xl_read_all_decoded must report zero rows");
    api->free_rows(&drained);

    /* Documented safe on a zeroed value. */
    xl_rows zeroed;
    memset(&zeroed, 0, sizeof(zeroed));
    api->free_rows(&zeroed);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_open_file_ex(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    size_t path_len = strlen(fixture);

    /* NULL options must behave exactly like xl_open_file. */
    CHECK(api->open_file_ex((const uint8_t*)fixture, (int32_t)path_len, XL_FORMAT_XLSB, NULL, &handle) == XL_OK,
          "xl_open_file_ex with NULL options must succeed like xl_open_file");
    int32_t sheet_count = 0;
    CHECK(api->sheet_count(handle, &sheet_count) == XL_OK && sheet_count == 1,
          "a workbook opened via xl_open_file_ex(NULL) must behave normally");
    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");

    /* A wrong struct_size must be rejected before anything else is inspected. */
    xl_open_options bad_options;
    memset(&bad_options, 0, sizeof(bad_options));
    bad_options.struct_size = 999999;
    handle = NULL;
    int32_t status = api->open_file_ex((const uint8_t*)fixture, (int32_t)path_len, XL_FORMAT_XLSB, &bad_options, &handle);
    CHECK(status == XL_INVALID_ARGUMENT, "a wrong xl_open_options.struct_size must be XL_INVALID_ARGUMENT");
    CHECK(handle == NULL, "a rejected xl_open_file_ex must not hand back a handle");
    return 0;
}

static int build_specs(xl_column_spec* specs, const uint8_t** name_ptrs, int32_t* name_lens)
{
    memset(specs, 0, 3 * sizeof(xl_column_spec));
    set_spec_name1(&specs[0], &name_ptrs[0], &name_lens[0], "Coluna1");
    specs[0].type = XL_T_STRING;
    specs[0].nullable = 1;
    set_spec_name1(&specs[1], &name_ptrs[1], &name_lens[1], "Coluna2");
    specs[1].type = XL_T_DATE;
    specs[1].nullable = 1;
    set_spec_name1(&specs[2], &name_ptrs[2], &name_lens[2], "Coluna3");
    specs[2].type = XL_T_I64;
    specs[2].nullable = 1;
    return 3;
}

/* The counts xl_parse_typed/xl_parse_arrow take are the only numbers a C caller hands over that
 * size an allocation AND drive a read across this process's memory. This is the only layer that can
 * test that guard: those entry points are [UnmanagedCallersOnly], so no managed test can invoke them
 * (the predicate itself is unit-tested in NativeApiTests).
 *
 * Every call below passes a ONE-element spec array while claiming more. Before the bound existed
 * these walked off the end of `specs` and sized an array from the claimed count — so a regression
 * here does not fail an assertion, it takes the process down, which CI reports just as loudly. */
static int test_parse_rejects_hostile_counts(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    xl_column_spec one_spec;
    memset(&one_spec, 0, sizeof(one_spec));
    const uint8_t* one_spec_name;
    int32_t one_spec_name_len;
    set_spec_name1(&one_spec, &one_spec_name, &one_spec_name_len, "Coluna1");
    one_spec.type = XL_T_STRING;

    xl_table table;
    memset(&table, 0, sizeof(table));

    CHECK(api->parse_typed(handle, &one_spec, INT32_MAX, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject spec_count == INT32_MAX");
    CHECK(table.columns == NULL, "a rejected xl_parse_typed must leave the out table zeroed");

    CHECK(api->parse_typed(handle, &one_spec, XL_MAX_COLUMN_SPECS + 1, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject spec_count one past XL_MAX_COLUMN_SPECS");
    CHECK(api->parse_typed(handle, &one_spec, 0, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject spec_count == 0");
    CHECK(api->parse_typed(handle, &one_spec, -1, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject a negative spec_count");

    /* A plausible count with an implausible name_len: the bound has to cover both, since name_len is
     * what becomes a read length over the caller's string. */
    xl_column_spec wide_name = one_spec;
    int32_t wide_name_len = XL_MAX_COLUMN_NAME_BYTES + 1;
    wide_name.name_lens = &wide_name_len;
    CHECK(api->parse_typed(handle, &wide_name, 1, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject a name_len past XL_MAX_COLUMN_NAME_BYTES");

    wide_name_len = -1;
    wide_name.name_lens = &wide_name_len;
    CHECK(api->parse_typed(handle, &wide_name, 1, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject a negative name_len");

    /* xl_parse_arrow decodes the same specs through the same path, so it needs the same guard. */
    struct ArrowArray array;
    struct ArrowSchema schema;
    memset(&array, 0, sizeof(array));
    memset(&schema, 0, sizeof(schema));
    CHECK(api->parse_arrow(handle, &one_spec, INT32_MAX, 1, &array, &schema) == XL_INVALID_ARGUMENT,
          "xl_parse_arrow must reject spec_count == INT32_MAX");
    CHECK(array.release == NULL && schema.release == NULL,
          "a rejected xl_parse_arrow must leave both out params releasable-as-no-op");

    /* A blank name would otherwise trim to "" and match the first empty header cell. */
    xl_column_spec blank_name;
    memset(&blank_name, 0, sizeof(blank_name));
    const uint8_t* blank_name_name;
    int32_t blank_name_len;
    set_spec_name1(&blank_name, &blank_name_name, &blank_name_len, "   ");
    blank_name.type = XL_T_STRING;
    CHECK(api->parse_typed(handle, &blank_name, 1, 1, &table) == XL_INVALID_ARGUMENT,
          "xl_parse_typed must reject a blank column name");

    /* The handle must still be usable: every rejection above is an argument error, not a fault that
     * leaves the workbook in a broken state. */
    xl_column_spec specs[3];
    const uint8_t* name_ptrs[3];
    int32_t name_lens[3];
    build_specs(specs, name_ptrs, name_lens);
    CHECK(api->parse_typed(handle, specs, 3, 1, &table) == XL_OK,
          "a valid xl_parse_typed after the rejections must still succeed");
    CHECK(table.row_count == 100, "the recovered parse must still return all 100 data rows");
    api->free_table(&table);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_parse_typed_and_cursor_independence(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    /* Advance the shared row cursor past the header before calling xl_parse_typed, then confirm
     * xl_parse_typed did not disturb it: the next xl_next_row call below must still see the FIRST
     * data row ("Valor1", not something further along), exactly as the header documents. */
    uint8_t scratch[4096];
    int32_t written = 0;
    CHECK(api->next_row(handle, scratch, (int32_t)sizeof(scratch), &written) == XL_OK,
          "reading the header row via xl_next_row must succeed");

    xl_column_spec specs[3];
    const uint8_t* name_ptrs[3];
    int32_t name_lens[3];
    build_specs(specs, name_ptrs, name_lens);

    xl_table table;
    memset(&table, 0, sizeof(table));
    CHECK(api->parse_typed(handle, specs, 3, 1, &table) == XL_OK, "xl_parse_typed must succeed");
    CHECK(table.column_count == 3, "xl_parse_typed must return exactly the requested columns");
    CHECK(table.row_count == 100, "xl_parse_typed must return all 100 data rows, independent of the row cursor");

    xl_column string_column = table.columns[0];
    CHECK(string_column.validity == NULL, "a column with no nulls must report a NULL validity bitmap");
    const int32_t* offsets = (const int32_t*)string_column.values;
    const uint8_t* data = string_column.data;
    CHECK(offsets[1] - offsets[0] == 6 && memcmp(data + offsets[0], "Valor1", 6) == 0,
          "the STRING column's first value must be Valor1");

    xl_column date_column = table.columns[1];
    const int32_t* dates = (const int32_t*)date_column.values;
    CHECK(dates[0] == 20454, "the DATE column's first value must be 20454 days since 1970-01-01 (Excel serial 46023)");

    xl_column int_column = table.columns[2];
    const int64_t* ints = (const int64_t*)int_column.values;
    CHECK(ints[0] == 1, "the I64 column's first value must be 1");

    api->free_table(&table);

    /* The row cursor must still be positioned right after the header row. */
    CHECK(api->next_row(handle, scratch, (int32_t)sizeof(scratch), &written) == XL_OK,
          "xl_next_row after xl_parse_typed must still succeed");
    int32_t cell_count = 0;
    memcpy(&cell_count, scratch, sizeof(int32_t));
    int32_t value_len = 0;
    memcpy(&value_len, scratch + 12, sizeof(int32_t));
    CHECK(cell_count == 18 && value_len == 6 && memcmp(scratch + 16, "Valor1", 6) == 0,
          "xl_parse_typed must not have disturbed the xl_next_row cursor - this must be the first data row");

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static void release_arrow_schema(struct ArrowSchema* schema)
{
    if (schema->release)
    {
        schema->release(schema);
    }
}

static void release_arrow_array(struct ArrowArray* array)
{
    if (array->release)
    {
        array->release(array);
    }
}

static int test_parse_arrow(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    xl_column_spec specs[3];
    const uint8_t* name_ptrs[3];
    int32_t name_lens[3];
    build_specs(specs, name_ptrs, name_lens);

    struct ArrowArray array;
    struct ArrowSchema schema;
    memset(&array, 0, sizeof(array));
    memset(&schema, 0, sizeof(schema));
    CHECK(api->parse_arrow(handle, specs, 3, 1, &array, &schema) == XL_OK, "xl_parse_arrow must succeed");

    CHECK(strcmp(schema.format, "+s") == 0, "the top-level Arrow schema must be a struct (\"+s\")");
    CHECK(schema.n_children == 3, "xl_parse_arrow must export exactly the requested columns");
    CHECK(strcmp(schema.children[0]->format, "u") == 0, "the STRING column's Arrow format must be \"u\"");
    CHECK(strcmp(schema.children[0]->name, "Coluna1") == 0, "the first child schema must be named Coluna1");
    CHECK(strcmp(schema.children[1]->format, "tdD") == 0, "the DATE column's Arrow format must be \"tdD\"");
    CHECK(strcmp(schema.children[2]->format, "l") == 0, "the I64 column's Arrow format must be \"l\"");

    CHECK(array.length == 100, "xl_parse_arrow must report 100 rows");
    CHECK(array.n_children == 3, "xl_parse_arrow's array must have one child per column");

    struct ArrowArray* string_array = array.children[0];
    const int32_t* offsets = (const int32_t*)string_array->buffers[1];
    const char* data = (const char*)string_array->buffers[2];
    CHECK(offsets[1] - offsets[0] == 6 && memcmp(data + offsets[0], "Valor1", 6) == 0,
          "the exported STRING array's first value must be Valor1");

    struct ArrowArray* int_array = array.children[2];
    const int64_t* ints = (const int64_t*)int_array->buffers[1];
    CHECK(ints[0] == 1, "the exported I64 array's first value must be 1");

    /* The real Arrow consumer contract: call release yourself. Not xl_free_table. */
    release_arrow_array(&array);
    release_arrow_schema(&schema);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_infer_schema(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    /* Advance the shared row cursor past the header, same setup as
     * test_parse_typed_and_cursor_independence, to prove xl_infer_schema reads from the sheet's
     * first row independent of it. */
    uint8_t scratch[4096];
    int32_t written = 0;
    CHECK(api->next_row(handle, scratch, (int32_t)sizeof(scratch), &written) == XL_OK,
          "reading the header row via xl_next_row must succeed");

    xl_inferred_schema schema;
    memset(&schema, 0, sizeof(schema));
    CHECK(api->infer_schema(handle, 1, 100, &schema) == XL_OK, "xl_infer_schema must succeed");
    CHECK(schema.column_count == 18, "RealExcel.xlsb's header row has 18 columns");

    xl_column_spec coluna1 = schema.columns[0];
    CHECK(coluna1.name_count == 1 && coluna1.name_lens[0] == 7 && memcmp(coluna1.names[0], "Coluna1", 7) == 0, "column 0 must be named Coluna1");
    CHECK(coluna1.type == XL_T_STRING, "Coluna1 must be guessed as XL_T_STRING");
    CHECK(coluna1.index == 0, "column 0's index must be 0 regardless of its name");

    xl_column_spec coluna2 = schema.columns[1];
    CHECK(coluna2.name_count == 1 && coluna2.name_lens[0] == 7 && memcmp(coluna2.names[0], "Coluna2", 7) == 0, "column 1 must be named Coluna2");
    CHECK(coluna2.type == XL_T_DATE, "Coluna2 must be guessed as XL_T_DATE");

    xl_column_spec coluna3 = schema.columns[2];
    CHECK(coluna3.name_count == 1 && coluna3.name_lens[0] == 7 && memcmp(coluna3.names[0], "Coluna3", 7) == 0, "column 2 must be named Coluna3");
    CHECK(coluna3.type == XL_T_I64, "Coluna3 must be guessed as XL_T_I64 - every sampled value is a whole number");

    /* The row cursor must still be positioned right after the header row. */
    CHECK(api->next_row(handle, scratch, (int32_t)sizeof(scratch), &written) == XL_OK,
          "xl_next_row after xl_infer_schema must still succeed");
    int32_t value_len = 0;
    memcpy(&value_len, scratch + 12, sizeof(int32_t));
    CHECK(value_len == 6 && memcmp(scratch + 16, "Valor1", 6) == 0,
          "xl_infer_schema must not have disturbed the xl_next_row cursor - this must be the first data row");

    /* The whole point of the shape match: an inferred schema is directly usable by xl_parse_typed.
     * Each spec's name pointer is only valid until xl_free_schema runs, so parse_typed must be
     * called first - copying the specs does not copy the name bytes they point to. */
    xl_column_spec first_three[3];
    memcpy(first_three, schema.columns, 3 * sizeof(xl_column_spec));

    xl_table table;
    memset(&table, 0, sizeof(table));
    CHECK(api->parse_typed(handle, first_three, 3, 1, &table) == XL_OK,
          "an inferred schema must be directly usable by xl_parse_typed");
    api->free_schema(&schema);
    CHECK(table.row_count == 100, "xl_parse_typed with the inferred schema must return all 100 data rows");
    api->free_table(&table);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_infer_schema_rejects_bad_arguments(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    xl_inferred_schema schema;
    memset(&schema, 0, sizeof(schema));
    CHECK(api->infer_schema(handle, -1, 100, &schema) == XL_INVALID_ARGUMENT,
          "xl_infer_schema must reject a negative header_row");
    CHECK(schema.columns == NULL, "a rejected xl_infer_schema must leave the out schema zeroed");

    CHECK(api->infer_schema(handle, 1, 0, &schema) == XL_INVALID_ARGUMENT,
          "xl_infer_schema must reject sample_size == 0");
    CHECK(api->infer_schema(handle, 1, -1, &schema) == XL_INVALID_ARGUMENT,
          "xl_infer_schema must reject a negative sample_size");

    /* Documented safe on a zeroed value. */
    xl_inferred_schema zeroed;
    memset(&zeroed, 0, sizeof(zeroed));
    api->free_schema(&zeroed);

    CHECK(api->close_(handle) == XL_OK, "xl_close must succeed");
    return 0;
}

static int test_double_close_is_rejected(const api_t* api, const char* fixture)
{
    xl_workbook* handle = NULL;
    CHECK(open_fixture(api, fixture, &handle) == XL_OK, "xl_open_file must succeed");

    CHECK(api->close_(handle) == XL_OK, "the first xl_close on a live handle must succeed");
    CHECK(api->close_(handle) == XL_INVALID_HANDLE,
          "a second xl_close on the same value must be XL_INVALID_HANDLE, never undefined behavior");

    int32_t sheet_count = 0;
    CHECK(api->sheet_count(handle, &sheet_count) == XL_INVALID_HANDLE,
          "any call on a closed handle must be XL_INVALID_HANDLE, not a crash or stale data");
    return 0;
}

/* Writes a two-row table through xl_write_typed, reads it back with the existing read exports, and
 * asserts on the values. Layer 1's static asserts prove the header's own layout; only running real
 * data through both directions proves excelreader.h, the C# structs and the writer agree. */
static int test_write_typed(const api_t* api)
{
    const char* out_path = "excelreader_smoke_write.csv";
    int64_t qty[2] = { 3, 7 };
    xl_column column;
    xl_table table;
    xl_column_spec spec;
    xl_write_options options;
    xl_workbook* handle = NULL;
    xl_table read_back;
    int32_t status;

    memset(&column, 0, sizeof(column));
    column.type = XL_T_I64;
    column.length = 2;
    column.values = qty;

    memset(&table, 0, sizeof(table));
    table.column_count = 1;
    table.row_count = 2;
    table.columns = &column;

    memset(&spec, 0, sizeof(spec));
    const uint8_t* spec_name;
    int32_t spec_name_len;
    set_spec_name1(&spec, &spec_name, &spec_name_len, "qty");
    spec.type = XL_T_I64;

    memset(&options, 0, sizeof(options));
    options.struct_size = (int32_t)sizeof(xl_write_options);

    status = api->write_typed((const uint8_t*)out_path, (int32_t)strlen(out_path),
                              XL_FORMAT_CSV, &spec, &table, &options);
    CHECK(status == XL_OK, "xl_write_typed should write a CSV");

    status = api->open_file((const uint8_t*)out_path, (int32_t)strlen(out_path), XL_FORMAT_CSV, &handle);
    CHECK(status == XL_OK, "the written CSV should reopen");

    memset(&read_back, 0, sizeof(read_back));
    status = api->parse_typed(handle, &spec, 1, 1, &read_back);
    CHECK(status == XL_OK, "the written CSV should parse back");
    CHECK(read_back.row_count == 2, "the written CSV should hold two rows");
    CHECK(((const int64_t*)read_back.columns[0].values)[0] == 3, "row 0 should round-trip as 3");
    CHECK(((const int64_t*)read_back.columns[0].values)[1] == 7, "row 1 should round-trip as 7");

    api->free_table(&read_back);
    api->close_(handle);
    remove(out_path);
    return 0;
}

int main(int argc, char** argv)
{
    const char* library_path = argc > 1 ? argv[1] : EXCELREADER_LIB_PATH_DEFAULT;
    const char* fixture_path = argc > 2 ? argv[2] : EXCELREADER_FIXTURE_PATH_DEFAULT;

    xl_lib_handle lib = load_library(library_path);
    if (!lib)
    {
        fprintf(stderr, "FAIL: could not load the ExcelReader.Native library at %s\n", library_path);
        return 1;
    }

    api_t api;
    memset(&api, 0, sizeof(api));
    if (!bind_all(lib, &api))
    {
        return 1;
    }

    int failures = 0;
    failures += test_abi_version(&api);
    failures += test_open_missing_file_reports_an_error(&api);
    failures += test_sheets_and_flags(&api, fixture_path);
    failures += test_next_row_blob_and_growth(&api, fixture_path);
    failures += test_next_row_decoded(&api, fixture_path);
    failures += test_read_all_blob_and_decoded(&api, fixture_path);
    failures += test_open_file_ex(&api, fixture_path);
    failures += test_parse_rejects_hostile_counts(&api, fixture_path);
    failures += test_parse_typed_and_cursor_independence(&api, fixture_path);
    failures += test_parse_arrow(&api, fixture_path);
    failures += test_infer_schema(&api, fixture_path);
    failures += test_infer_schema_rejects_bad_arguments(&api, fixture_path);
    failures += test_double_close_is_rejected(&api, fixture_path);
    failures += test_write_typed(&api);

    if (failures == 0)
    {
        printf("OK: all ExcelReader.Native smoke checks passed (%s)\n", fixture_path);
        return 0;
    }
    fprintf(stderr, "FAIL: %d smoke check group(s) failed\n", failures);
    return 1;
}
