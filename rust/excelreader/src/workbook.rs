use crate::ffi::{
    self, XlColumn, XlColumnSpec, XlInferredSchema, XlOpenOptions, XlTable, XlWorkbook,
    XL_BUFFER_TOO_SMALL, XL_EOF, XL_OK,
};
use crate::{
    AllRows, Date, Error, ErrorKind, OpenOptions, ReadOptions, RowCursor, Time, Timestamp,
};
use std::marker::PhantomData;
use std::path::Path;

pub(crate) fn last_error(code: i32) -> Error {
    unsafe {
        let mut len: i32 = 0;
        let ptr = ffi::xl_last_error_ptr(&mut len);
        let message = if ptr.is_null() || len <= 0 {
            "unknown error".to_string()
        } else {
            let bytes = std::slice::from_raw_parts(ptr, len as usize);
            String::from_utf8_lossy(bytes).into_owned()
        };
        Error::from_status(code, message)
    }
}

pub(crate) fn check(code: i32) -> Result<(), Error> {
    if code == XL_OK {
        Ok(())
    } else {
        Err(last_error(code))
    }
}

/// The ABI takes every path as UTF-8 bytes.
pub(crate) fn path_str(path: &Path) -> Result<&str, Error> {
    path.to_str().ok_or_else(|| {
        Error::new(
            ErrorKind::InvalidArgument,
            format!("path {} is not valid UTF-8", path.display()),
        )
    })
}

/// The ABI takes every length as `int32`, so an oversized input is rejected rather than truncated.
pub(crate) fn len_i32(len: usize, what: &str) -> Result<i32, Error> {
    i32::try_from(len).map_err(|_| {
        Error::new(
            ErrorKind::InvalidArgument,
            format!("{what} is {len} bytes, past the ABI's int32 length limit"),
        )
    })
}

/// Copies a native `XlBuffer` into an owned `Vec<u8>` and releases the native allocation via
/// `xl_free_buffer`. `buffer.data` may be null (an empty result), which `from_raw_parts` cannot
/// take even for a zero-length slice.
pub(crate) fn buffer_to_vec(mut buffer: ffi::XlBuffer) -> Vec<u8> {
    let bytes = if buffer.data.is_null() || buffer.len <= 0 {
        Vec::new()
    } else {
        unsafe { std::slice::from_raw_parts(buffer.data, buffer.len as usize).to_vec() }
    };
    unsafe {
        ffi::xl_free_buffer(&mut buffer);
    }
    bytes
}

/// Verifies the loaded shared library speaks the ABI revision this crate was compiled against.
///
/// The native binary is resolved at build time from a GitHub release asset (or from
/// `EXCELREADER_NATIVE_LIB_DIR`), so it is entirely possible to end up linking a library built from
/// a different ABI revision than the one this crate's `#[repr(C)]` structs mirror. Proceeding past a
/// mismatch would mean reading native memory through the wrong layout, so every constructor calls
/// this first.
///
/// The result is cached: it cannot change for the lifetime of the process, and every open would
/// otherwise pay an FFI call for it.
pub(crate) fn check_abi_version() -> Result<(), Error> {
    use std::sync::OnceLock;
    static CHECKED: OnceLock<Result<(), Error>> = OnceLock::new();

    CHECKED
        .get_or_init(|| {
            let loaded = unsafe { ffi::xl_abi_version() };
            if loaded == ffi::XL_ABI_VERSION {
                warm_up();
                Ok(())
            } else {
                Err(Error::new(
                    ErrorKind::Other,
                    format!(
                        "ExcelReader native library reports ABI version {loaded}, but this crate \
                         was built against {}. Update the crate and the native library together.",
                        ffi::XL_ABI_VERSION
                    ),
                ))
            }
        })
        .clone()
}

/// Runs one small read and one small aggregation, so the native runtime finishes starting up on a
/// single thread.
///
/// Several threads whose FIRST call into the library lands in the runtime's class-constructor
/// machinery at the same moment can crash it: a thread comes back from a spin in `Thread.Yield` on
/// a corrupted return address, which ends the process with no message (SIGABRT on Linux, SIGSEGV on
/// macOS). It is not this crate's bug - a plain C host doing the same thing crashes too, just far
/// less often - but this crate makes it likely, because a Rust test binary calls in from a fresh
/// thread per test. [`check_abi_version`]'s `OnceLock` already serializes the first call, so doing
/// the heavy initialization here means later threads find the constructors already run.
///
/// Failures are ignored on purpose: this is a warm-up, and a real call will report the same problem
/// properly.
fn warm_up() {
    const CSV: &[u8] = b"a,b\n1,2\n3,4\n";

    let mut handle: *mut XlWorkbook = std::ptr::null_mut();
    let status = unsafe {
        ffi::xl_open_memory(
            CSV.as_ptr(),
            CSV.len() as i32,
            ffi::XL_FORMAT_CSV,
            std::ptr::null(),
            &mut handle,
        )
    };
    if status == XL_OK && !handle.is_null() {
        unsafe { ffi::xl_close(handle) };
    }

    struct CountRows(i64);
    impl crate::CsvAccumulator for CountRows {
        fn accumulate(&mut self, _row: crate::RowRef<'_>) -> Result<(), i32> {
            self.0 += 1;
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            self.0 += other.0;
            Ok(())
        }
    }

    let _ = crate::aggregate::aggregate_unchecked(
        crate::CsvInput::Bytes(CSV),
        &OpenOptions::new(),
        0,
        2,
        || CountRows(0),
    );
}

/// An open workbook. Safe to share between threads: every read goes through a [`Sheet`] and opens
/// its own cursor, so reads on one workbook never disturb one another.
pub struct Workbook {
    handle: *mut XlWorkbook,
}

// SAFETY: ABI v6 documents the workbook handle as usable from several threads at once.
unsafe impl Send for Workbook {}
unsafe impl Sync for Workbook {}

impl Workbook {
    /// Opens `path`, sniffing the format and using every library default. Sniffing does not detect
    /// CSV: open one with [`OpenOptions::format`] set to [`Format::Csv`](crate::Format::Csv).
    pub fn open(path: impl AsRef<Path>) -> Result<Workbook, Error> {
        OpenOptions::new().open(path)
    }

    /// Whether the workbook uses the 1904 date system - needed to interpret raw Excel serial dates.
    pub fn is_date1904(&self) -> Result<bool, Error> {
        let mut flag: i32 = 0;
        check(unsafe { ffi::xl_is_date1904(self.handle, &mut flag) })?;
        Ok(flag != 0)
    }

    /// The sheet at `index`. An index outside the workbook is an error here, not at the first read.
    pub fn sheet(&self, index: i32) -> Result<Sheet<'_>, Error> {
        let mut visibility: i32 = 0;
        check(unsafe { ffi::xl_sheet_visibility_at(self.handle, index, &mut visibility) })?;
        Ok(Sheet {
            workbook: self,
            index,
        })
    }

    /// Every sheet, in workbook order. Opens nothing.
    pub fn sheets(&self) -> Result<Vec<Sheet<'_>>, Error> {
        let mut count: i32 = 0;
        check(unsafe { ffi::xl_sheet_count(self.handle, &mut count) })?;
        Ok((0..count)
            .map(|index| Sheet {
                workbook: self,
                index,
            })
            .collect())
    }

    /// The sheet called `name`, compared without regard to case. `Ok(None)` when no sheet matches.
    pub fn sheet_by_name(&self, name: &str) -> Result<Option<Sheet<'_>>, Error> {
        let mut index: i32 = -1;
        check(unsafe {
            ffi::xl_sheet_index(self.handle, name.as_ptr(), name.len() as i32, &mut index)
        })?;
        Ok((index >= 0).then_some(Sheet {
            workbook: self,
            index,
        }))
    }

    /// Shared two-pass buffer dance for the `xl_*` functions that write a UTF-8 name into a caller
    /// buffer and report the required capacity through `XL_BUFFER_TOO_SMALL`.
    fn fill_string(
        &self,
        call: impl Fn(*mut XlWorkbook, *mut u8, i32, *mut i32) -> i32,
    ) -> Result<String, Error> {
        let mut buffer = [0u8; 128];
        let mut len: i32 = 0;
        let mut status = call(
            self.handle,
            buffer.as_mut_ptr(),
            buffer.len() as i32,
            &mut len,
        );
        if status != XL_BUFFER_TOO_SMALL {
            check(status)?;
            let buffer_slice = &buffer[..len.max(0) as usize];
            return str::from_utf8(buffer_slice)
                .map(|s| s.to_string())
                .map_err(|e| non_utf8_name(&e));
        }
        let mut vec_buffer = vec![0u8; len.max(0) as usize];
        status = call(
            self.handle,
            vec_buffer.as_mut_ptr(),
            vec_buffer.len() as i32,
            &mut len,
        );
        check(status)?;
        vec_buffer.truncate(len.max(0) as usize);
        String::from_utf8(vec_buffer).map_err(|e| non_utf8_name(&e))
    }
}

fn non_utf8_name(error: &dyn std::fmt::Display) -> Error {
    Error::new(
        ErrorKind::Other,
        format!("native library returned a non-UTF-8 name: {error}"),
    )
}

impl OpenOptions {
    /// Opens the workbook at `path` with these options.
    pub fn open(&self, path: impl AsRef<Path>) -> Result<Workbook, Error> {
        let path = path_str(path.as_ref())?;
        let path_len = len_i32(path.len(), "path")?;
        self.open_raw(|format, options, handle| unsafe {
            ffi::xl_open_file(path.as_ptr(), path_len, format, options, handle)
        })
    }

    /// Opens a workbook held in memory. The native library copies `data`, so it need not outlive
    /// this call.
    pub fn open_bytes(&self, data: &[u8]) -> Result<Workbook, Error> {
        let data_len = len_i32(data.len(), "buffer")?;
        self.open_raw(|format, options, handle| unsafe {
            ffi::xl_open_memory(data.as_ptr(), data_len, format, options, handle)
        })
    }

    /// The tail every open shares: the ABI check, lowering these options, and turning the status
    /// and handle into a `Workbook`. `open` runs only after the ABI check passes.
    pub(crate) fn open_raw(
        &self,
        open: impl FnOnce(i32, *const XlOpenOptions, *mut *mut XlWorkbook) -> i32,
    ) -> Result<Workbook, Error> {
        check_abi_version()?;
        let raw = self.to_raw();
        let mut handle: *mut XlWorkbook = std::ptr::null_mut();
        check(open(self.format_raw(), raw.as_ptr(), &mut handle))?;
        Ok(Workbook { handle })
    }
}

/// How a sheet is shown in the application's tab bar.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

/// One sheet of a [`Workbook`]: the workbook and an index. Holds no native resource, so it is
/// `Copy` and free to pass to another thread; each read method opens what it needs.
#[derive(Clone, Copy)]
pub struct Sheet<'w> {
    workbook: &'w Workbook,
    index: i32,
}

impl<'w> Sheet<'w> {
    /// Zero-based position in the workbook.
    #[must_use]
    pub fn index(&self) -> i32 {
        self.index
    }

    /// The sheet's name.
    pub fn name(&self) -> Result<String, Error> {
        self.workbook
            .fill_string(|handle, buffer, capacity, out_len| unsafe {
                ffi::xl_sheet_name_at(handle, self.index, buffer, capacity, out_len)
            })
    }

    /// Whether the sheet is visible, hidden or very hidden. An error if the native side reports a
    /// value this crate does not know.
    pub fn visibility(&self) -> Result<SheetVisibility, Error> {
        let mut raw: i32 = 0;
        check(unsafe { ffi::xl_sheet_visibility_at(self.workbook.handle, self.index, &mut raw) })?;
        match raw {
            ffi::XL_SHEET_VISIBLE => Ok(SheetVisibility::Visible),
            ffi::XL_SHEET_HIDDEN => Ok(SheetVisibility::Hidden),
            ffi::XL_SHEET_VERY_HIDDEN => Ok(SheetVisibility::VeryHidden),
            other => Err(Error::new(
                ErrorKind::Other,
                format!("native library returned an unknown sheet visibility {other}"),
            )),
        }
    }

    /// A row-at-a-time reader over this sheet, from its first row. Each call opens an independent
    /// cursor, so several may be alive at once, on this thread or on others.
    pub fn rows(&self) -> Result<RowCursor<'w>, Error> {
        RowCursor::open(self.workbook.handle, self.index)
    }

    /// Every row of this sheet, read in one native call into a buffer the result owns outright.
    pub fn read_all(&self) -> Result<AllRows, Error> {
        let cursor = self.rows()?;
        AllRows::read(cursor.raw())
    }

    /// Schema-driven columnar parse of this sheet into `T`s, matching C++'s `xl::parse_sheet<T>`.
    /// The fastest way to read a sheet whose columns you know: it converts on the native side and
    /// never formats a cell to text.
    ///
    /// Uses [`ReadOptions`]'s header row and parallelism; a negative parallelism is an error.
    pub fn parse<T: ExcelMapper>(&self, options: &ReadOptions) -> Result<TableView<T>, Error> {
        let arena = build_specs::<T>();
        let mut table = empty_table();
        check(unsafe {
            ffi::xl_parse_typed(
                self.workbook.handle,
                self.index,
                arena.specs.as_ptr(),
                arena.specs.len() as i32,
                options.header_row,
                options.parallelism,
                &mut table,
            )
        })?;
        TableView::new(table, arena.bindings, "xl_parse_typed")
    }

    /// [`parse`](Self::parse) delivered [`ReadOptions`]'s batch size rows at a time; a negative
    /// batch size is an error. The header row is consumed once, here, not re-read per batch.
    pub fn parse_batches<T: ExcelMapper>(
        &self,
        options: &ReadOptions,
    ) -> Result<TypedBatches<'w, T>, Error> {
        let arena = build_specs::<T>();
        let mut reader: *mut ffi::XlTypedReader = std::ptr::null_mut();
        check(unsafe {
            ffi::xl_typed_reader_open(
                self.workbook.handle,
                self.index,
                arena.specs.as_ptr(),
                arena.specs.len() as i32,
                options.header_row,
                options.batch_size,
                &mut reader,
            )
        })?;
        Ok(TypedBatches {
            reader,
            bindings: arena.bindings,
            done: false,
            _workbook: PhantomData,
        })
    }

    /// Guesses a [`parse`](Self::parse) schema by sampling this sheet.
    ///
    /// `header_row` has the same meaning as in `parse`; `sample_size` bounds how many rows after
    /// the header are inspected. With `parse_text`, cells that hold text are typed too - every CSV
    /// field, or numbers stored as text: integers, decimals, `true`/`false` and ISO dates or
    /// date-times, when the text has exactly that shape (leading-zero codes such as `00123` stay
    /// text). This is a guess over a sample, not a guarantee.
    pub fn infer_schema(
        &self,
        header_row: i32,
        sample_size: i32,
        parse_text: bool,
    ) -> Result<Vec<InferredColumn>, Error> {
        let flags = if parse_text {
            ffi::XL_INFER_PARSE_TEXT
        } else {
            0
        };
        let mut schema = XlInferredSchema {
            columns: std::ptr::null_mut(),
            column_count: 0,
        };
        check(unsafe {
            ffi::xl_infer_schema(
                self.workbook.handle,
                self.index,
                header_row,
                sample_size,
                flags,
                &mut schema,
            )
        })?;

        let columns = unsafe { copy_inferred(&schema) };
        unsafe { ffi::xl_free_schema(&mut schema) };
        columns
    }

    #[cfg(feature = "arrow")]
    pub(crate) fn handle(&self) -> *mut XlWorkbook {
        self.workbook.handle
    }
}

/// The sheet's position only - reading its name would be a native call inside `Debug`.
impl std::fmt::Debug for Sheet<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Sheet").field("index", &self.index).finish()
    }
}

/// A column's value type, shared by reading ([`InferredColumn`]) and writing ([`ColumnData`]).
///
/// [`ColumnData`]: crate::ColumnData
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ColumnType {
    String,
    I64,
    F64,
    Bool,
    /// Days since 1970-01-01.
    Date,
    /// Microseconds since midnight.
    Time,
    /// Microseconds since 1970-01-01T00:00:00Z.
    Timestamp,
}

impl ColumnType {
    pub(crate) fn raw(self) -> i32 {
        match self {
            ColumnType::String => ffi::XL_T_STRING,
            ColumnType::I64 => ffi::XL_T_I64,
            ColumnType::F64 => ffi::XL_T_F64,
            ColumnType::Bool => ffi::XL_T_BOOL,
            ColumnType::Date => ffi::XL_T_DATE,
            ColumnType::Time => ffi::XL_T_TIME,
            ColumnType::Timestamp => ffi::XL_T_TIMESTAMP,
        }
    }

    fn from_raw(value: i32) -> Option<ColumnType> {
        match value {
            ffi::XL_T_STRING => Some(ColumnType::String),
            ffi::XL_T_I64 => Some(ColumnType::I64),
            ffi::XL_T_F64 => Some(ColumnType::F64),
            ffi::XL_T_BOOL => Some(ColumnType::Bool),
            ffi::XL_T_DATE => Some(ColumnType::Date),
            ffi::XL_T_TIME => Some(ColumnType::Time),
            ffi::XL_T_TIMESTAMP => Some(ColumnType::Timestamp),
            _ => None,
        }
    }
}

/// Deep-copies a native-owned inferred schema into owned Rust values, so the caller can free the
/// native allocation immediately.
///
/// # Safety
/// `schema` must be one `xl_infer_schema` returned `XL_OK` for, not yet passed to `xl_free_schema`.
unsafe fn copy_inferred(schema: &XlInferredSchema) -> Result<Vec<InferredColumn>, Error> {
    if schema.columns.is_null() || schema.column_count <= 0 {
        return Ok(Vec::new());
    }
    let specs = std::slice::from_raw_parts(schema.columns, schema.column_count as usize);
    specs
        .iter()
        .map(|spec| {
            let column_type = ColumnType::from_raw(spec.r#type).ok_or_else(|| {
                Error::new(
                    ErrorKind::Other,
                    format!(
                        "native library inferred an unknown column type {}",
                        spec.r#type
                    ),
                )
            })?;
            let name = if spec.name_count <= 0 || spec.names.is_null() || spec.name_lens.is_null() {
                None
            } else {
                let name_ptr = *spec.names;
                let name_len = *spec.name_lens;
                if name_ptr.is_null() || name_len <= 0 {
                    None
                } else {
                    let bytes = std::slice::from_raw_parts(name_ptr, name_len as usize);
                    Some(String::from_utf8_lossy(bytes).into_owned())
                }
            };
            Ok(InferredColumn {
                name,
                index: spec.index,
                column_type,
                nullable: spec.nullable != 0,
            })
        })
        .collect()
}

/// One column guessed by [`Sheet::infer_schema`].
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct InferredColumn {
    /// Header text, or `None` when the column must be resolved by [`index`](Self::index) instead.
    pub name: Option<String>,
    pub index: i32,
    pub column_type: ColumnType,
    pub nullable: bool,
}

impl Drop for Workbook {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe {
                ffi::xl_close(self.handle);
            }
            self.handle = std::ptr::null_mut();
        }
    }
}

/// Deliberately opaque: the only state here is a handle whose value means nothing outside the
/// native library, and printing it would invite treating it as an identity it is not.
impl std::fmt::Debug for Workbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Workbook")
            .field("open", &!self.handle.is_null())
            .finish()
    }
}

/// One field <-> column binding for `T`, produced by `#[derive(ExcelMapper)]`.
pub struct ColumnBinding<T> {
    /// Candidate header names, in priority order — the first one present in the header row wins.
    pub names: &'static [&'static str],
    pub column_type: ColumnType,
    pub assign: fn(&mut T, &XlColumn, i64),
}

// Hand-written: `#[derive(Clone)]` would add a `T: Clone` bound the user's row struct need not meet.
impl<T> Clone for ColumnBinding<T> {
    fn clone(&self) -> Self {
        *self
    }
}

impl<T> Copy for ColumnBinding<T> {}

/// Implemented by any struct [`Sheet::parse`] can populate. Derive it with
/// `#[derive(ExcelMapper)]`, which also implements [`ExcelWriter`](crate::ExcelWriter).
pub trait ExcelMapper: Sized {
    #[doc(hidden)]
    fn bindings() -> Vec<ColumnBinding<Self>>;
}

pub use excelreader_derive::ExcelMapper;

fn is_valid(col: &XlColumn, row: i64) -> bool {
    if col.validity.is_null() {
        return true;
    }
    unsafe {
        let byte = *col.validity.offset((row / 8) as isize);
        (byte & (1 << (row % 8))) != 0
    }
}

/// Guards every `column_*` accessor below.
///
/// These are safe `pub` functions that index into raw native buffers, so this check is what makes
/// them sound: without it a `row` past the column returns whatever sits after the allocation. The
/// type half catches the other class of mistake - reading an `XL_T_F64` buffer as `i64` is a
/// bindings bug, not recoverable input.
#[inline]
fn check_access(col: &XlColumn, row: i64, expected_type: i32, accessor: &str) {
    assert_eq!(
        col.r#type, expected_type,
        "{accessor} called on a column of type {}",
        col.r#type
    );
    assert!(
        row >= 0 && row < col.length,
        "{accessor}: row {row} is out of bounds for a column of length {}",
        col.length
    );
}

/// Reads column `col` at `row` as a `&str`.
///
/// The ABI documents every string it returns as UTF-8, but this validates rather than trusting it:
/// the bytes cross an FFI boundary from a binary resolved at build time, and `from_utf8_unchecked`
/// on a value that turns out not to be UTF-8 is undefined behavior, not merely a wrong answer.
///
/// # Panics
/// If `col` is not a string column, if `row` is out of bounds, or if the native library returned
/// bytes that are not valid UTF-8 (an ABI contract violation).
pub fn column_str(col: &XlColumn, row: i64) -> &str {
    check_access(col, row, ffi::XL_T_STRING, "column_str");
    unsafe {
        let offsets = col.values as *const i32;
        let start = *offsets.offset(row as isize);
        let end = *offsets.offset(row as isize + 1);
        let bytes =
            std::slice::from_raw_parts(col.data.offset(start as isize), (end - start) as usize);
        std::str::from_utf8(bytes)
            .expect("native library returned a non-UTF-8 string, violating the ABI contract")
    }
}

/// # Panics
/// If `col` is not an `i64` column, or `row` is out of bounds.
pub fn column_i64(col: &XlColumn, row: i64) -> i64 {
    check_access(col, row, ffi::XL_T_I64, "column_i64");
    unsafe { *(col.values as *const i64).offset(row as isize) }
}

/// # Panics
/// If `col` is not an `f64` column, or `row` is out of bounds.
pub fn column_f64(col: &XlColumn, row: i64) -> f64 {
    check_access(col, row, ffi::XL_T_F64, "column_f64");
    unsafe { *(col.values as *const f64).offset(row as isize) }
}

/// # Panics
/// If `col` is not a bool column, or `row` is out of bounds.
pub fn column_bool(col: &XlColumn, row: i64) -> bool {
    check_access(col, row, ffi::XL_T_BOOL, "column_bool");
    unsafe { *(col.values as *const u8).offset(row as isize) != 0 }
}

/// # Panics
/// If `col` is not a date column, or `row` is out of bounds.
pub fn column_date(col: &XlColumn, row: i64) -> Date {
    check_access(col, row, ffi::XL_T_DATE, "column_date");
    Date {
        days_since_epoch: unsafe { *(col.values as *const i32).offset(row as isize) },
    }
}

/// # Panics
/// If `col` is not a time column, or `row` is out of bounds.
pub fn column_time(col: &XlColumn, row: i64) -> Time {
    check_access(col, row, ffi::XL_T_TIME, "column_time");
    Time {
        micros_since_midnight: unsafe { *(col.values as *const i64).offset(row as isize) },
    }
}

/// # Panics
/// If `col` is not a timestamp column, or `row` is out of bounds.
pub fn column_timestamp(col: &XlColumn, row: i64) -> Timestamp {
    check_access(col, row, ffi::XL_T_TIMESTAMP, "column_timestamp");
    Timestamp {
        micros_since_epoch: unsafe { *(col.values as *const i64).offset(row as isize) },
    }
}

fn empty_table() -> XlTable {
    XlTable {
        column_count: 0,
        row_count: 0,
        columns: std::ptr::null_mut(),
    }
}

/// Owns a typed parse result. Frees the native table on `Drop`. `iter()` builds one `T` per call
/// to `next()` from the columnar buffers - no upfront `Vec<T>` allocation beyond what the native
/// parse itself already made.
pub struct TableView<T: ExcelMapper> {
    table: XlTable,
    bindings: Vec<ColumnBinding<T>>,
    _marker: PhantomData<T>,
}

impl<T: ExcelMapper> TableView<T> {
    /// Takes ownership of `table`, freeing it and failing if its column count disagrees with the
    /// bindings it will be read through.
    fn new(
        mut table: XlTable,
        bindings: Vec<ColumnBinding<T>>,
        source: &str,
    ) -> Result<Self, Error> {
        if table.column_count as usize != bindings.len() {
            let column_count = table.column_count;
            unsafe { ffi::xl_free_table(&mut table) };
            return Err(Error::new(
                ErrorKind::Other,
                format!(
                    "{source} returned {column_count} columns for {} specs",
                    bindings.len()
                ),
            ));
        }
        Ok(TableView {
            table,
            bindings,
            _marker: PhantomData,
        })
    }

    pub fn len(&self) -> i64 {
        self.table.row_count
    }

    pub fn is_empty(&self) -> bool {
        self.table.row_count == 0
    }

    /// Builds `T` from row `row`, or `None` when `row` falls outside `0..len()`.
    ///
    /// Returning `Option` rather than `T` is what keeps this sound: the columnar buffers are raw
    /// native allocations, so an unchecked row would read past them from safe code.
    pub fn get(&self, row: i64) -> Option<T>
    where
        T: Default,
    {
        if row < 0 || row >= self.len() {
            return None;
        }
        let mut instance = T::default();
        let columns = unsafe {
            std::slice::from_raw_parts(self.table.columns, self.table.column_count as usize)
        };
        for (col, binding) in columns.iter().zip(self.bindings.iter()) {
            if is_valid(col, row) {
                (binding.assign)(&mut instance, col, row);
            }
        }
        Some(instance)
    }

    pub fn iter(&self) -> TableViewIter<'_, T>
    where
        T: Default,
    {
        TableViewIter { view: self, row: 0 }
    }
}

impl<T: ExcelMapper> Drop for TableView<T> {
    fn drop(&mut self) {
        unsafe {
            ffi::xl_free_table(&mut self.table);
        }
    }
}

// SAFETY: plain native allocations, read only through &self, released by xl_free_table, which any
// thread may call.
unsafe impl<T: ExcelMapper> Send for TableView<T> {}
unsafe impl<T: ExcelMapper> Sync for TableView<T> {}

/// Shape only. Rendering the rows would mean materializing every `T`, which is the one thing this
/// type exists to avoid.
impl<T: ExcelMapper> std::fmt::Debug for TableView<T> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TableView")
            .field("rows", &self.table.row_count)
            .field("columns", &self.table.column_count)
            .finish_non_exhaustive()
    }
}

/// Iterator over the batches of one chunked typed read. Closes the native reader on `Drop`.
pub struct TypedBatches<'a, T: ExcelMapper> {
    reader: *mut ffi::XlTypedReader,
    bindings: Vec<ColumnBinding<T>>,
    done: bool,
    _workbook: PhantomData<&'a Workbook>,
}

// SAFETY: a typed reader owns its native enumerator; ABI v6 lets it be used from any one thread
// at a time.
unsafe impl<T: ExcelMapper> Send for TypedBatches<'_, T> {}

impl<T: ExcelMapper> Iterator for TypedBatches<'_, T> {
    type Item = Result<TableView<T>, Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        let mut table = empty_table();
        let status = unsafe { ffi::xl_typed_reader_next(self.reader, &mut table) };
        if status == XL_EOF {
            self.done = true;
            return None;
        }
        if status != XL_OK {
            self.done = true;
            return Some(Err(last_error(status)));
        }

        // ponytail: one small Vec of function pointers cloned per batch, not per row.
        let view = TableView::new(table, self.bindings.clone(), "xl_typed_reader_next");
        self.done = view.is_err();
        Some(view)
    }
}

impl<T: ExcelMapper> Drop for TypedBatches<'_, T> {
    fn drop(&mut self) {
        unsafe { ffi::xl_typed_reader_close(self.reader) };
    }
}

/// Progress only - the reader handle is opaque. Needed for `expect_err` on a `Result<Self, _>`.
impl<T: ExcelMapper> std::fmt::Debug for TypedBatches<'_, T> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TypedBatches")
            .field("columns", &self.bindings.len())
            .field("done", &self.done)
            .finish_non_exhaustive()
    }
}

pub struct TableViewIter<'a, T: ExcelMapper> {
    view: &'a TableView<T>,
    row: i64,
}

impl<T: ExcelMapper + Default> Iterator for TableViewIter<'_, T> {
    type Item = T;

    fn next(&mut self) -> Option<T> {
        let item = self.view.get(self.row)?;
        self.row += 1;
        Some(item)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = (self.view.len() - self.row).max(0) as usize;
        (remaining, Some(remaining))
    }
}

impl<T: ExcelMapper + Default> ExactSizeIterator for TableViewIter<'_, T> {}

/// Keeps the per-column name pointer/length vectors alive for as long as the `XlColumnSpec` array
/// that points into them. The two `_name_*` fields are never read - dropping them early would
/// leave `specs` holding dangling pointers, which is the whole reason they are stored here.
///
/// Also carries the `T::bindings()` this arena was built from, so callers that need both the flat
/// spec array (for the FFI call) and the typed bindings (for result-column lookups) can get both
/// from a single `T::bindings()` call instead of computing it twice.
pub(crate) struct SpecArena<T: ExcelMapper> {
    pub(crate) specs: Vec<XlColumnSpec>,
    pub(crate) bindings: Vec<ColumnBinding<T>>,
    _name_ptrs: Vec<Vec<*const u8>>,
    _name_lens: Vec<Vec<i32>>,
}

/// Lowers `T`'s ExcelMapper bindings into the flat `xl_column_spec` array both `xl_parse_typed` and
/// `xl_parse_arrow` take - their column-spec input is identical.
pub(crate) fn build_specs<T: ExcelMapper>() -> SpecArena<T> {
    let bindings = T::bindings();
    let name_ptrs: Vec<Vec<*const u8>> = bindings
        .iter()
        .map(|b| b.names.iter().map(|n| n.as_ptr()).collect())
        .collect();
    let name_lens: Vec<Vec<i32>> = bindings
        .iter()
        .map(|b| b.names.iter().map(|n| n.len() as i32).collect())
        .collect();
    let specs: Vec<XlColumnSpec> = bindings
        .iter()
        .enumerate()
        .map(|(i, b)| XlColumnSpec {
            names: name_ptrs[i].as_ptr(),
            name_lens: name_lens[i].as_ptr(),
            name_count: b.names.len() as i32,
            index: 0,
            r#type: b.column_type.raw(),
            nullable: 1,
        })
        .collect();
    SpecArena {
        specs,
        bindings,
        _name_ptrs: name_ptrs,
        _name_lens: name_lens,
    }
}
