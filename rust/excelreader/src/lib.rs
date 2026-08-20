//! Raw FFI bindings to ExcelReader's C ABI (`excelreader.h`). The `extern "C"` block below is
//! declared directly in this crate root (not a submodule), so `crate::xl_open_file_ex` etc. resolve
//! from anywhere in the crate. See the `workbook` module for the safe wrapper built on top of this.

mod error;
mod options;
mod temporal;
pub mod workbook;

pub use error::Error;
pub use options::OpenOptions;
pub use temporal::{Date, Time, Timestamp};

use std::os::raw::{c_int, c_void};

pub const XL_OK: i32 = 0;
pub const XL_EOF: i32 = -1;
pub const XL_BUFFER_TOO_SMALL: i32 = -2;
pub const XL_INVALID_HANDLE: i32 = -3;
pub const XL_INVALID_ARGUMENT: i32 = -4;
pub const XL_ERROR: i32 = -5;

/// ABI revision this crate is compiled against. `Workbook::open` refuses to proceed when the loaded
/// library's `xl_abi_version()` disagrees - see `workbook::check_abi_version`.
pub const XL_ABI_VERSION: i32 = 2;

pub const XL_T_STRING: i32 = 0;
pub const XL_T_I64: i32 = 1;
pub const XL_T_F64: i32 = 2;
pub const XL_T_BOOL: i32 = 3;
/// Days since 1970-01-01, stored as `i32`.
pub const XL_T_DATE: i32 = 4;
/// Microseconds since midnight, stored as `i64`.
pub const XL_T_TIME: i32 = 5;
/// Microseconds since 1970-01-01T00:00:00Z, stored as `i64`.
pub const XL_T_TIMESTAMP: i32 = 6;

pub const XL_FORMAT_AUTO: i32 = 0;
pub const XL_FORMAT_XLS: i32 = 1;
pub const XL_FORMAT_XLSX: i32 = 2;
pub const XL_FORMAT_XLSB: i32 = 3;
pub const XL_FORMAT_CSV: i32 = 4;

/// Every boolean-shaped `xl_open_options` field uses one of these three states, never a plain 0/1 -
/// several of them default to true, so a bare 0 would be ambiguous between "off" and "use the
/// library default".
pub const XL_OPT_DEFAULT: i32 = 0;
pub const XL_OPT_FALSE: i32 = 1;
pub const XL_OPT_TRUE: i32 = 2;

/// Opaque handle - never dereferenced by Rust, only passed back to `xl_*` functions.
#[repr(C)]
pub struct XlWorkbook {
    _private: [u8; 0],
}

/// Mirrors `xl_open_options`. `struct_size` must be set to `size_of::<XlOpenOptions>()`; build one
/// through [`OpenOptions`] rather than by hand.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XlOpenOptions {
    pub struct_size: i32,
    pub csv_sniff_dialect: i32,
    pub csv_delimiter: i32,
    pub csv_quote: i32,
    pub csv_detect_bom: i32,
    pub csv_max_cell_bytes: i32,
    pub csv_intern_strings: i32,
    pub max_total_decompressed_bytes: i64,
    pub max_cell_bytes: i32,
    pub max_shared_string_bytes: i64,
    pub max_zip_entries: i32,
    pub prefetch_decompression: i32,
    pub intern_strings: i32,
}

#[repr(C)]
pub struct XlColumnSpec {
    pub names: *const *const u8,
    pub name_lens: *const i32,
    pub name_count: i32,
    pub index: i32,
    pub r#type: i32,
    pub nullable: i32,
}

#[repr(C)]
pub struct XlColumn {
    pub r#type: i32,
    pub length: i64,
    pub values: *const c_void,
    pub validity: *const u8,
    pub data: *const u8,
    pub data_len: i64,
}

#[repr(C)]
pub struct XlTable {
    pub column_count: i32,
    pub row_count: i64,
    pub columns: *mut XlColumn,
}

/// Mirrors `xl_inferred_schema`. Owned by the native library until `xl_free_schema`.
#[repr(C)]
pub struct XlInferredSchema {
    pub columns: *mut XlColumnSpec,
    pub column_count: i32,
}

extern "C" {
    pub fn xl_abi_version() -> c_int;

    pub fn xl_open_file_ex(
        path: *const u8,
        path_len: i32,
        format: i32,
        options: *const XlOpenOptions,
        out_handle: *mut *mut XlWorkbook,
    ) -> c_int;

    pub fn xl_open_memory_ex(
        data: *const u8,
        data_len: i32,
        format: i32,
        options: *const XlOpenOptions,
        out_handle: *mut *mut XlWorkbook,
    ) -> c_int;

    pub fn xl_close(handle: *mut XlWorkbook) -> c_int;

    pub fn xl_sheet_count(handle: *mut XlWorkbook, out_count: *mut i32) -> c_int;

    pub fn xl_sheet_name(
        handle: *mut XlWorkbook,
        buffer: *mut u8,
        capacity: i32,
        out_len: *mut i32,
    ) -> c_int;

    pub fn xl_sheet_name_at(
        handle: *mut XlWorkbook,
        index: i32,
        buffer: *mut u8,
        capacity: i32,
        out_len: *mut i32,
    ) -> c_int;

    pub fn xl_move_to_sheet(handle: *mut XlWorkbook, index: i32) -> c_int;

    pub fn xl_is_date1904(handle: *mut XlWorkbook, out_flag: *mut i32) -> c_int;

    pub fn xl_parse_typed(
        handle: *mut XlWorkbook,
        specs: *const XlColumnSpec,
        spec_count: i32,
        header_row: i32,
        out_table: *mut XlTable,
    ) -> c_int;

    pub fn xl_free_table(table: *mut XlTable);

    pub fn xl_infer_schema(
        handle: *mut XlWorkbook,
        header_row: i32,
        sample_size: i32,
        out_schema: *mut XlInferredSchema,
    ) -> c_int;

    pub fn xl_free_schema(schema: *mut XlInferredSchema);

    pub fn xl_last_error_ptr(out_len: *mut i32) -> *const u8;
}
