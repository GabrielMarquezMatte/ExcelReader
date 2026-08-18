//! Raw FFI bindings to ExcelReader's C ABI (`excelreader.h`), phase 1: opening a workbook and
//! `xl_parse_typed` only. The `extern "C"` block below is declared directly in this crate root
//! (not a submodule), so `crate::xl_open_file_ex` etc. resolve from anywhere in the crate. See the
//! `workbook` module for the safe wrapper built on top of this.

mod error;
pub mod workbook;

pub use error::Error;

use std::os::raw::{c_int, c_void};

pub const XL_OK: i32 = 0;
pub const XL_ABI_VERSION: i32 = 1;

pub const XL_T_STRING: i32 = 0;
pub const XL_T_I64: i32 = 1;
pub const XL_T_F64: i32 = 2;
pub const XL_T_BOOL: i32 = 3;
pub const XL_T_DATE: i32 = 4;
pub const XL_T_TIME: i32 = 5;
pub const XL_T_TIMESTAMP: i32 = 6;

pub const XL_FORMAT_AUTO: i32 = 0;
pub const XL_FORMAT_XLS: i32 = 1;
pub const XL_FORMAT_XLSX: i32 = 2;
pub const XL_FORMAT_XLSB: i32 = 3;
pub const XL_FORMAT_CSV: i32 = 4;

/// Opaque handle - never dereferenced by Rust, only passed back to `xl_*` functions.
#[repr(C)]
pub struct XlWorkbook {
    _private: [u8; 0],
}

#[repr(C)]
pub struct XlColumnSpec {
    pub name: *const u8,
    pub name_len: i32,
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

extern "C" {
    pub fn xl_abi_version() -> c_int;

    pub fn xl_open_file_ex(
        path: *const u8,
        path_len: i32,
        format: i32,
        options: *const c_void, // NULL in phase 1 - OpenOptions is out of scope
        out_handle: *mut *mut XlWorkbook,
    ) -> c_int;

    pub fn xl_close(handle: *mut XlWorkbook) -> c_int;

    pub fn xl_parse_typed(
        handle: *mut XlWorkbook,
        specs: *const XlColumnSpec,
        spec_count: i32,
        header_row: i32,
        out_table: *mut XlTable,
    ) -> c_int;

    pub fn xl_free_table(table: *mut XlTable);

    pub fn xl_last_error_ptr(out_len: *mut i32) -> *const u8;
}
