//! Raw FFI bindings to ExcelReader's C ABI (`excelreader.h`). The `extern "C"` block below is
//! declared directly in this crate root (not a submodule), so `crate::xl_open_file_ex` etc. resolve
//! from anywhere in the crate. See the `workbook` module for the safe wrapper built on top of this.

mod error;
mod options;
mod temporal;
pub mod workbook;
pub mod writer;
pub mod writer_handle;
pub mod rows;

#[cfg(feature = "arrow")]
pub mod arrow;

pub use error::Error;
pub use rows::{AllRows, CellIter, CellRef, CellType, DecodedRows, RowCursor, RowRef};
pub use options::{OpenOptions, OpenOptionsRaw, WriteOptions};
pub use temporal::{Date, Time, Timestamp};

use std::os::raw::{c_int, c_void};

pub const XL_OK: i32 = 0;
pub const XL_EOF: i32 = -1;
pub const XL_BUFFER_TOO_SMALL: i32 = -2;
pub const XL_INVALID_HANDLE: i32 = -3;
pub const XL_INVALID_ARGUMENT: i32 = -4;
pub const XL_ERROR: i32 = -5;
/// The workbook is encrypted and no password was supplied.
pub const XL_STATUS_PASSWORD_REQUIRED: i32 = -6;
/// The supplied password did not match the workbook's verifier.
pub const XL_STATUS_PASSWORD_INCORRECT: i32 = -7;

pub const XL_CELL_EMPTY: i32 = 0;
pub const XL_CELL_STRING: i32 = 1;
pub const XL_CELL_NUMBER: i32 = 2;
pub const XL_CELL_DATE: i32 = 3;
pub const XL_CELL_BOOL: i32 = 4;
pub const XL_CELL_FORMULA: i32 = 5;
pub const XL_CELL_ERROR: i32 = 6;

/// ABI revision this crate is compiled against. `Workbook::open` refuses to proceed when the loaded
/// library's `xl_abi_version()` disagrees - see `workbook::check_abi_version`.
pub const XL_ABI_VERSION: i32 = 4;

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

    /// Password for an encrypted OOXML workbook, as UTF-8 bytes. NULL (with `password_len` 0) means
    /// "not encrypted, or fail with `XL_STATUS_PASSWORD_REQUIRED`". Not NUL-terminated -
    /// `password_len` is authoritative, since a password may contain any byte. The pointer need only
    /// remain valid for the duration of the call - see [`OpenOptionsRaw`](crate::options::OpenOptionsRaw).
    pub password: *const u8,
    pub password_len: i32,
}

/// Mirrors `xl_write_options`. Field ORDER is the C struct's, not a tidied-up version of it: with
/// `repr(C)` the 4 bytes of padding after `sheet_name_len` land exactly where a C compiler puts
/// them, giving the 32-byte x64 layout `tests/ExcelReader.NativeSmoke/smoke.c` static-asserts.
/// Build one through [`WriteOptions`] rather than by hand.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XlWriteOptions {
    pub struct_size: i32,
    pub sheet_name_len: i32,
    pub sheet_name: *const u8,
    pub csv_delimiter: i32,
    pub csv_quote: i32,
    pub date1904: i32,
    pub use_shared_strings: i32,
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

/// Mirrors `xl_buffer`: an owned block of native memory returned by `xl_write_typed_to_memory` or
/// `xl_write_handle_bytes`. Released with `xl_free_buffer` - see `writer::buffer_to_vec`, the one
/// place this crate touches the raw struct directly.
#[repr(C)]
pub struct XlBuffer {
    pub data: *mut u8,
    pub len: i64,
}

/// Opaque streaming writer handle - never dereferenced by Rust, only passed back to `xl_*`
/// functions. See [`writer_handle::WriterHandle`] for the safe wrapper.
#[repr(C)]
pub struct XlWriterHandle {
    _private: [u8; 0],
}

/// Mirrors `xl_row_cell`. Named `cell_type` because `type` is a Rust keyword.
#[repr(C)]
#[derive(Debug)]
pub struct XlRowCell {
    pub column: i32,
    pub cell_type: i32,
    pub value_len: i32,
    pub value: *const u8,
}

/// Mirrors `xl_row`.
#[repr(C)]
pub struct XlRow {
    pub cell_count: i32,
    pub cells: *mut XlRowCell,
}

/// Mirrors `xl_rows`.
#[repr(C)]
pub struct XlRows {
    pub row_count: i32,
    pub rows: *mut XlRow,
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

    /// Same schema-driven parse as `xl_parse_typed`, exported as one top-level Arrow struct
    /// array/schema. `out_array`/`out_schema` are `struct ArrowArray*`/`struct ArrowSchema*` from
    /// the Arrow C Data Interface, typed here as `c_void` because arrow-rs's own `#[repr(C)]`
    /// `FFI_ArrowArray`/`FFI_ArrowSchema` are ABI-identical to them - redeclaring the spec structs
    /// would be a second source of truth for a fixed, versioned ABI.
    ///
    /// On `XL_OK` the caller owns both and releases each through its OWN `release` callback, never
    /// through `xl_free_table`. On any other status both outputs are left untouched.
    pub fn xl_parse_arrow(
        handle: *mut XlWorkbook,
        specs: *const XlColumnSpec,
        spec_count: i32,
        header_row: i32,
        out_array: *mut c_void,
        out_schema: *mut c_void,
    ) -> c_int;

    pub fn xl_free_table(table: *mut XlTable);

    pub fn xl_infer_schema(
        handle: *mut XlWorkbook,
        header_row: i32,
        sample_size: i32,
        out_schema: *mut XlInferredSchema,
    ) -> c_int;

    pub fn xl_free_schema(schema: *mut XlInferredSchema);
    pub fn xl_write_typed(
        path: *const u8,
        path_len: i32,
        format: i32,
        specs: *const XlColumnSpec,
        table: *const XlTable,
        options: *const XlWriteOptions,
    ) -> c_int;

    /// Same as `xl_write_typed`, except the result is returned as `out_buffer` instead of being
    /// written to a path. Only read `*out_buffer` when the call returns `XL_OK`; on failure it is
    /// zeroed, and `xl_free_buffer` on a zeroed buffer is a no-op.
    pub fn xl_write_typed_to_memory(
        format: i32,
        specs: *const XlColumnSpec,
        table: *const XlTable,
        options: *const XlWriteOptions,
        out_buffer: *mut XlBuffer,
    ) -> c_int;

    /// Releases a buffer returned by `xl_write_typed_to_memory` or `xl_write_handle_bytes` and
    /// resets it to zero. Safe on a zeroed value.
    pub fn xl_free_buffer(buffer: *mut XlBuffer);

    /// Reads the plaintext XLSX/XLSB package at `package_path` and writes its agile-encrypted
    /// (ECMA-376 4.4) counterpart to `destination_path`, overwriting an existing file. The result
    /// opens with the same password via `xl_open_file_ex`'s `xl_open_options::password`.
    pub fn xl_encrypt_package(
        package_path: *const u8,
        package_path_len: i32,
        destination_path: *const u8,
        destination_path_len: i32,
        password: *const u8,
        password_len: i32,
    ) -> c_int;

    // ---- Streaming writer handle: see writer_handle::WriterHandle for the call-order contract. ----

    pub fn xl_open_write_handle(
        path: *const u8,
        path_len: i32,
        format: i32,
        options: *const XlWriteOptions,
        out_handle: *mut *mut XlWriterHandle,
    ) -> c_int;

    pub fn xl_open_write_handle_to_memory(
        format: i32,
        options: *const XlWriteOptions,
        out_handle: *mut *mut XlWriterHandle,
    ) -> c_int;

    pub fn xl_start_sheet(handle: *mut XlWriterHandle, name: *const u8, name_len: i32) -> c_int;
    pub fn xl_start_row(handle: *mut XlWriterHandle) -> c_int;

    pub fn xl_write_string(handle: *mut XlWriterHandle, value: *const u8, value_len: i32) -> c_int;
    pub fn xl_write_int64(handle: *mut XlWriterHandle, value: i64) -> c_int;
    pub fn xl_write_float64(handle: *mut XlWriterHandle, value: f64) -> c_int;
    pub fn xl_write_bool(handle: *mut XlWriterHandle, value: i32) -> c_int;
    pub fn xl_write_date(handle: *mut XlWriterHandle, days_since_epoch: i32) -> c_int;
    pub fn xl_write_time(handle: *mut XlWriterHandle, micros_since_midnight: i64) -> c_int;
    pub fn xl_write_timestamp(handle: *mut XlWriterHandle, micros_since_epoch: i64) -> c_int;
    pub fn xl_write_null(handle: *mut XlWriterHandle, r#type: i32) -> c_int;

    pub fn xl_end_row(handle: *mut XlWriterHandle) -> c_int;
    pub fn xl_end_sheet(handle: *mut XlWriterHandle) -> c_int;
    pub fn xl_close_write_handle(handle: *mut XlWriterHandle) -> c_int;

    /// Reads back everything written so far to a handle opened by `xl_open_write_handle_to_memory`.
    /// `XL_INVALID_ARGUMENT` for one from `xl_open_write_handle`. Only read `*out_buffer` when the
    /// call returns `XL_OK`; on failure it is zeroed, and `xl_free_buffer` on a zeroed buffer is a
    /// no-op.
    pub fn xl_write_handle_bytes(handle: *mut XlWriterHandle, out_buffer: *mut XlBuffer) -> c_int;

    pub fn xl_next_row(
        handle: *mut XlWorkbook,
        buffer: *mut u8,
        capacity: i32,
        out_written: *mut i32,
    ) -> c_int;
    pub fn xl_read_all_decoded(handle: *mut XlWorkbook, out_rows: *mut XlRows) -> c_int;
    pub fn xl_free_rows(rows: *mut XlRows);

    pub fn xl_read_all_blob(
        handle: *mut XlWorkbook,
        buffer: *mut u8,
        capacity: i32,
        out_written: *mut i32,
    ) -> c_int;

    pub fn xl_last_error_ptr(out_len: *mut i32) -> *const u8;
}
