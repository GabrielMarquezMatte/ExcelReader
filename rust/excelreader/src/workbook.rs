use crate::ffi::{
    self, native_slice, XlColumn, XlColumnSpec, XlInferredSchema, XlOpenOptions, XlTable,
    XlWorkbook, XL_BUFFER_TOO_SMALL, XL_EOF, XL_OK,
};
use crate::{
    AllRows, Column, ColumnData, Date, Error, ErrorKind, OpenOptions, ReadOptions, RowCursor,
    RowRef, Time, Timestamp,
};
use std::marker::PhantomData;
use std::path::Path;

pub(crate) fn last_error(code: i32) -> Error {
    unsafe {
        let mut len: i32 = 0;
        let bytes = native_slice(ffi::xl_last_error_ptr(&mut len), len);
        let message = if bytes.is_empty() {
            "unknown error".to_string()
        } else {
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
/// `xl_free_buffer`. `buffer.data` may be null (an empty result).
pub(crate) fn buffer_to_vec(mut buffer: ffi::XlBuffer) -> Vec<u8> {
    let bytes = unsafe { native_slice(buffer.data, buffer.len) }.to_vec();
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

    /// Every Excel table, in workbook order.
    pub fn tables(&self) -> Result<Vec<TableInfo>, Error> {
        let mut count: i32 = 0;
        check(unsafe { ffi::xl_table_count(self.handle, &mut count) })?;
        (0..count).map(|index| self.table_info(index)).collect()
    }

    fn table_info(&self, index: i32) -> Result<TableInfo, Error> {
        let mut raw = ffi::XlTableInfo::default();
        check(unsafe { ffi::xl_table_info_at(self.handle, index, &mut raw) })?;
        let name = self.fill_string(|handle, buffer, capacity, out_len| unsafe {
            ffi::xl_table_name_at(handle, index, buffer, capacity, out_len)
        })?;
        Ok(TableInfo {
            name,
            index,
            sheet: raw.sheet,
            first_row: raw.first_row,
            last_row: raw.last_row,
            first_column: raw.first_column,
            last_column: raw.last_column,
            header_row_count: raw.header_row_count,
            totals_row_count: raw.totals_row_count,
        })
    }

    /// Table `index` read as a sheet: every [`Sheet`] read covers only the table's range, header
    /// row first, with the column indexes of the sheet holding it. Its [`Sheet::name`] and
    /// [`Sheet::visibility`] are that sheet's. An index outside the workbook is an error here.
    pub fn table(&self, index: i32) -> Result<Sheet<'_>, Error> {
        let mut raw = ffi::XlTableInfo::default();
        check(unsafe { ffi::xl_table_info_at(self.handle, index, &mut raw) })?;
        Ok(Sheet {
            workbook: self,
            index: ffi::xl_table_sheet(index),
        })
    }

    /// The table called `name`, compared without regard to case, read as in [`table`](Self::table).
    /// `Ok(None)` when no table matches.
    pub fn table_by_name(&self, name: &str) -> Result<Option<Sheet<'_>>, Error> {
        let mut index: i32 = -1;
        check(unsafe {
            ffi::xl_table_index(self.handle, name.as_ptr(), len_i32(name.len(), "name")?, &mut index)
        })?;
        Ok((index >= 0).then_some(Sheet {
            workbook: self,
            index: ffi::xl_table_sheet(index),
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

/// One Excel table's place in a workbook. Rows and columns are zero-based and include the header
/// and totals rows. Read its cells through [`Workbook::table`].
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct TableInfo {
    pub name: String,
    /// Position among the workbook's tables, as [`Workbook::table`] takes it.
    pub index: i32,
    /// Index of the sheet holding the table.
    pub sheet: i32,
    pub first_row: i32,
    pub last_row: i32,
    pub first_column: i32,
    pub last_column: i32,
    /// 1, or 0 when the table's header row is turned off.
    pub header_row_count: i32,
    /// 1 when the table shows a totals row, otherwise 0.
    pub totals_row_count: i32,
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
    /// Zero-based position in the workbook. A sheet from [`Workbook::table`] reports the table's
    /// encoded index, which only read calls understand.
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

    /// Calls `f` once per row on the calling thread, without copying the row into a cursor buffer.
    ///
    /// Each row's cells borrow the library's buffers for that call only. A cell may carry text, a
    /// number, or both: [`CellRef::text`](crate::CellRef::text) and
    /// [`CellRef::number`](crate::CellRef::number) cover all three. A nonzero `Err(code)` stops the
    /// walk (`Err(0)` is `XL_OK` and does not); a positive code surfaces as
    /// [`ErrorKind::Callback`](crate::ErrorKind::Callback). Other calls
    /// on this workbook are allowed from inside `f`. A panic in `f` stops the walk and is resumed
    /// here.
    pub fn for_each_row<F>(&self, f: F) -> Result<(), Error>
    where
        F: FnMut(RowRef<'_>) -> Result<(), i32>,
    {
        crate::foreach::for_each_row(self.workbook.handle, self.index, f)
    }

    /// Folds this sheet into one `A` across several threads.
    ///
    /// `header_row` is 1-based, `0` meaning no header. `degree_of_parallelism` is `0` for every
    /// processor, `1` sequential, `n` up to `n` threads. The CSV dialect, password and limits come
    /// from the options the workbook was opened with.
    ///
    /// `seed` is called once per chunk, on worker threads; see [`Accumulator`](crate::Accumulator)
    /// for the contract. `seed`, `accumulate` and `combine` never run on the calling thread; each
    /// accumulator is dropped on the calling thread once the run has finished. The call blocks
    /// until the run finishes, so `seed` may borrow locals. A panic inside any callback aborts the
    /// run and is resumed here once the library has freed every state.
    pub fn aggregate<A, F>(
        &self,
        header_row: i32,
        degree_of_parallelism: i32,
        seed: F,
    ) -> Result<A, Error>
    where
        A: crate::Accumulator,
        F: Fn() -> A + Sync,
    {
        crate::aggregate::aggregate_sheet(
            self.workbook.handle,
            self.index,
            header_row,
            degree_of_parallelism,
            seed,
        )
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

    /// [`parse`](Self::parse) over a schema chosen at run time, such as one
    /// [`infer_schema`](Self::infer_schema) returned. A column with a name is matched by it, one
    /// without by its index; an empty schema is an error.
    pub fn parse_with_schema(
        &self,
        schema: &[InferredColumn],
        options: &ReadOptions,
    ) -> Result<Table, Error> {
        let specs = schema_specs(schema);
        let mut table = empty_table();
        check(unsafe {
            ffi::xl_parse_typed(
                self.workbook.handle,
                self.index,
                specs.specs.as_ptr(),
                specs.specs.len() as i32,
                options.header_row,
                options.parallelism,
                &mut table,
            )
        })?;
        Table::new(table, schema, "xl_parse_typed")
    }

    /// [`parse_with_schema`](Self::parse_with_schema) delivered [`ReadOptions`]'s batch size rows
    /// at a time, as [`parse_batches`](Self::parse_batches) does. `schema` is copied, so it need not
    /// outlive the batches.
    pub fn parse_batches_with_schema(
        &self,
        schema: &[InferredColumn],
        options: &ReadOptions,
    ) -> Result<TableBatches<'w>, Error> {
        Ok(TableBatches {
            batches: BatchReader::open(self, &schema_specs(schema).specs, options)?,
            schema: schema.to_vec(),
        })
    }

    /// [`parse`](Self::parse) delivered [`ReadOptions`]'s batch size rows at a time; a negative
    /// batch size is an error. The header row is consumed once, here, not re-read per batch.
    pub fn parse_batches<T: ExcelMapper>(
        &self,
        options: &ReadOptions,
    ) -> Result<TypedBatches<'w, T>, Error> {
        let arena = build_specs::<T>();
        Ok(TypedBatches {
            batches: BatchReader::open(self, &arena.specs, options)?,
            bindings: arena.bindings,
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
    native_slice(schema.columns, schema.column_count)
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
                let bytes = native_slice(*spec.names, *spec.name_lens);
                (!bytes.is_empty()).then(|| String::from_utf8_lossy(bytes).into_owned())
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
    pub assign: fn(&mut T, ColumnRef<'_>, i64),
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

/// A column of a table the native library returned. Only this crate can build one, which is what
/// lets the safe `column_*` accessors trust its pointers: a hand-built `XlColumn` never reaches them.
///
/// ```compile_fail
/// #![forbid(unsafe_code)]
/// let column = excelreader::__private::XlColumn {
///     r#type: 1, length: 1,
///     values: std::ptr::null(), validity: std::ptr::null(),
///     data: std::ptr::null(), data_len: 0,
/// };
/// excelreader::__private::column_i64(excelreader::__private::ColumnRef(&column), 0);
/// ```
#[derive(Clone, Copy)]
pub struct ColumnRef<'a>(&'a XlColumn);

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
pub fn column_str<'a>(col: ColumnRef<'a>, row: i64) -> &'a str {
    let col = col.0;
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
pub fn column_i64(col: ColumnRef<'_>, row: i64) -> i64 {
    let col = col.0;
    check_access(col, row, ffi::XL_T_I64, "column_i64");
    unsafe { *(col.values as *const i64).offset(row as isize) }
}

/// # Panics
/// If `col` is not an `f64` column, or `row` is out of bounds.
pub fn column_f64(col: ColumnRef<'_>, row: i64) -> f64 {
    let col = col.0;
    check_access(col, row, ffi::XL_T_F64, "column_f64");
    unsafe { *(col.values as *const f64).offset(row as isize) }
}

/// # Panics
/// If `col` is not a bool column, or `row` is out of bounds.
pub fn column_bool(col: ColumnRef<'_>, row: i64) -> bool {
    let col = col.0;
    check_access(col, row, ffi::XL_T_BOOL, "column_bool");
    unsafe { *(col.values as *const u8).offset(row as isize) != 0 }
}

/// # Panics
/// If `col` is not a date column, or `row` is out of bounds.
pub fn column_date(col: ColumnRef<'_>, row: i64) -> Date {
    let col = col.0;
    check_access(col, row, ffi::XL_T_DATE, "column_date");
    Date {
        days_since_epoch: unsafe { *(col.values as *const i32).offset(row as isize) },
    }
}

/// # Panics
/// If `col` is not a time column, or `row` is out of bounds.
pub fn column_time(col: ColumnRef<'_>, row: i64) -> Time {
    let col = col.0;
    check_access(col, row, ffi::XL_T_TIME, "column_time");
    Time {
        micros_since_midnight: unsafe { *(col.values as *const i64).offset(row as isize) },
    }
}

/// # Panics
/// If `col` is not a timestamp column, or `row` is out of bounds.
pub fn column_timestamp(col: ColumnRef<'_>, row: i64) -> Timestamp {
    let col = col.0;
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
                (binding.assign)(&mut instance, ColumnRef(col), row);
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

/// Owns a [`Sheet::parse_with_schema`] result. Frees the native table on `Drop`.
///
/// Its columns are the [`Column`]s [`write_columns`](crate::write_columns) takes, borrowed from
/// this table. A column matched by index has no name, and writing needs every column named or none.
pub struct Table {
    table: XlTable,
    names: Vec<Option<String>>,
}

impl Table {
    /// Takes ownership of `table`, freeing it and failing if it does not have one column of a known
    /// type and of `row_count` rows per `schema` entry.
    fn new(mut table: XlTable, schema: &[InferredColumn], source: &str) -> Result<Self, Error> {
        let columns = unsafe { native_slice(table.columns, table.column_count) };
        let consistent = table.row_count >= 0
            && columns.len() == schema.len()
            && columns
                .iter()
                .all(|c| c.length == table.row_count && ColumnType::from_raw(c.r#type).is_some());
        if !consistent {
            unsafe { ffi::xl_free_table(&mut table) };
            return Err(Error::new(
                ErrorKind::Other,
                format!(
                    "{source} returned a table that does not match its {} specs",
                    schema.len()
                ),
            ));
        }
        Ok(Table {
            table,
            names: schema.iter().map(|c| c.name.clone()).collect(),
        })
    }

    pub fn len(&self) -> i64 {
        self.table.row_count
    }

    pub fn is_empty(&self) -> bool {
        self.table.row_count == 0
    }

    pub fn num_columns(&self) -> usize {
        self.names.len()
    }

    /// Column `index`, or `None` past the last one.
    pub fn column(&self, index: usize) -> Option<Column<'_>> {
        let name = self.names.get(index)?;
        let raw = unsafe { &*self.table.columns.add(index) };
        Some(Column {
            name: name.as_deref(),
            data: unsafe { column_data(raw) },
            validity: (!raw.validity.is_null())
                .then(|| unsafe { native_slice(raw.validity, (raw.length as usize).div_ceil(8)) }),
        })
    }

    /// Every column, in schema order.
    pub fn columns(&self) -> Vec<Column<'_>> {
        (0..self.num_columns())
            .filter_map(|i| self.column(i))
            .collect()
    }
}

/// # Safety
/// `raw` must be a column of a native table whose type [`Table::new`] checked.
unsafe fn column_data(raw: &XlColumn) -> ColumnData<'_> {
    let rows = raw.length as usize;
    match ColumnType::from_raw(raw.r#type) {
        Some(ColumnType::String) => ColumnData::Str {
            offsets: native_slice(raw.values.cast::<i32>(), rows + 1),
            data: native_slice(raw.data, raw.data_len as usize),
        },
        Some(ColumnType::I64) => ColumnData::I64(native_slice(raw.values.cast(), rows)),
        Some(ColumnType::F64) => ColumnData::F64(native_slice(raw.values.cast(), rows)),
        Some(ColumnType::Bool) => ColumnData::Bool(native_slice(raw.values.cast(), rows)),
        Some(ColumnType::Date) => ColumnData::Date(native_slice(raw.values.cast(), rows)),
        Some(ColumnType::Time) => ColumnData::Time(native_slice(raw.values.cast(), rows)),
        Some(ColumnType::Timestamp) => ColumnData::Timestamp(native_slice(raw.values.cast(), rows)),
        None => unreachable!("Table::new rejects unknown column types"),
    }
}

impl Drop for Table {
    fn drop(&mut self) {
        unsafe {
            ffi::xl_free_table(&mut self.table);
        }
    }
}

// SAFETY: as for TableView.
unsafe impl Send for Table {}
unsafe impl Sync for Table {}

impl std::fmt::Debug for Table {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Table")
            .field("rows", &self.table.row_count)
            .field("columns", &self.names)
            .finish_non_exhaustive()
    }
}

/// The native half of a chunked typed read: one `xl_typed_reader`, closed on `Drop`.
struct BatchReader<'a> {
    reader: *mut ffi::XlTypedReader,
    done: bool,
    _workbook: PhantomData<&'a Workbook>,
}

// SAFETY: a typed reader owns its native enumerator; ABI v6 lets it be used from any one thread
// at a time.
unsafe impl Send for BatchReader<'_> {}

impl<'a> BatchReader<'a> {
    fn open(
        sheet: &Sheet<'a>,
        specs: &[XlColumnSpec],
        options: &ReadOptions,
    ) -> Result<Self, Error> {
        let mut reader: *mut ffi::XlTypedReader = std::ptr::null_mut();
        check(unsafe {
            ffi::xl_typed_reader_open(
                sheet.workbook.handle,
                sheet.index,
                specs.as_ptr(),
                specs.len() as i32,
                options.header_row,
                options.batch_size,
                &mut reader,
            )
        })?;
        Ok(BatchReader {
            reader,
            done: false,
            _workbook: PhantomData,
        })
    }

    /// Hands the next batch's table to `wrap`, which takes ownership of it. Stops after the first
    /// error, native or from `wrap`.
    fn next_with<R>(
        &mut self,
        wrap: impl FnOnce(XlTable) -> Result<R, Error>,
    ) -> Option<Result<R, Error>> {
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

        let batch = wrap(table);
        self.done = batch.is_err();
        Some(batch)
    }
}

impl Drop for BatchReader<'_> {
    fn drop(&mut self) {
        unsafe { ffi::xl_typed_reader_close(self.reader) };
    }
}

/// Iterator over the batches of one chunked typed read. Closes the native reader on `Drop`.
pub struct TypedBatches<'a, T: ExcelMapper> {
    batches: BatchReader<'a>,
    bindings: Vec<ColumnBinding<T>>,
}

impl<T: ExcelMapper> Iterator for TypedBatches<'_, T> {
    type Item = Result<TableView<T>, Error>;

    fn next(&mut self) -> Option<Self::Item> {
        let bindings = &self.bindings;
        // ponytail: one small Vec of function pointers cloned per batch, not per row.
        self.batches
            .next_with(|table| TableView::new(table, bindings.clone(), "xl_typed_reader_next"))
    }
}

/// Progress only - the reader handle is opaque. Needed for `expect_err` on a `Result<Self, _>`.
impl<T: ExcelMapper> std::fmt::Debug for TypedBatches<'_, T> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TypedBatches")
            .field("columns", &self.bindings.len())
            .field("done", &self.batches.done)
            .finish_non_exhaustive()
    }
}

/// [`TypedBatches`]' counterpart for [`Sheet::parse_batches_with_schema`], yielding [`Table`]s.
pub struct TableBatches<'a> {
    batches: BatchReader<'a>,
    schema: Vec<InferredColumn>,
}

impl Iterator for TableBatches<'_> {
    type Item = Result<Table, Error>;

    fn next(&mut self) -> Option<Self::Item> {
        let schema = &self.schema;
        self.batches
            .next_with(|table| Table::new(table, schema, "xl_typed_reader_next"))
    }
}

impl std::fmt::Debug for TableBatches<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TableBatches")
            .field("columns", &self.schema.len())
            .field("done", &self.batches.done)
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

/// [`SpecArena`] for a schema chosen at run time: `specs` points into `schema`'s names and into the
/// two `_name_*` vectors, so all three must outlive the native call.
pub(crate) struct SchemaSpecs<'a> {
    pub(crate) specs: Vec<XlColumnSpec>,
    _name_ptrs: Vec<*const u8>,
    _name_lens: Vec<i32>,
    _schema: PhantomData<&'a [InferredColumn]>,
}

/// Lowers a runtime schema into `xl_column_spec`s: a column with a name is matched by it, one
/// without by its index.
pub(crate) fn schema_specs(schema: &[InferredColumn]) -> SchemaSpecs<'_> {
    let name_ptrs: Vec<*const u8> = schema
        .iter()
        .map(|c| c.name.as_deref().map_or(std::ptr::null(), str::as_ptr))
        .collect();
    let name_lens: Vec<i32> = schema
        .iter()
        .map(|c| c.name.as_deref().map_or(0, |n| n.len() as i32))
        .collect();
    let specs: Vec<XlColumnSpec> = schema
        .iter()
        .enumerate()
        .map(|(i, c)| XlColumnSpec {
            names: &name_ptrs[i],
            name_lens: &name_lens[i],
            name_count: i32::from(c.name.is_some()),
            index: c.index,
            r#type: c.column_type.raw(),
            nullable: i32::from(c.nullable),
        })
        .collect();
    SchemaSpecs {
        specs,
        _name_ptrs: name_ptrs,
        _name_lens: name_lens,
        _schema: PhantomData,
    }
}
