use crate::{
    Date, Error, OpenOptions, Time, Timestamp, XlColumn, XlColumnSpec, XlInferredSchema, XlTable,
    XlWorkbook, XL_BUFFER_TOO_SMALL, XL_ERROR, XL_FORMAT_AUTO, XL_OK, XL_T_BOOL, XL_T_DATE,
    XL_T_F64, XL_T_I64, XL_T_STRING, XL_T_TIME, XL_T_TIMESTAMP,
};
use std::marker::PhantomData;

pub(crate) fn last_error(code: i32) -> Error {
    unsafe {
        let mut len: i32 = 0;
        let ptr = crate::xl_last_error_ptr(&mut len);
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

/// Copies a native `XlBuffer` into an owned `Vec<u8>` and releases the native allocation via
/// `xl_free_buffer` - shared by `writer::write_columns_to_memory`/`write_sheet_to_memory` and
/// `writer_handle::WriterHandle::bytes`, the two places `xl_write_typed_to_memory`/
/// `xl_write_handle_bytes` hand back an owned buffer. `buffer.data` may be null (an empty result),
/// which `from_raw_parts` cannot take - `slice::from_raw_parts` requires a non-null, well-aligned
/// pointer even for a zero-length slice.
pub(crate) fn buffer_to_vec(mut buffer: crate::XlBuffer) -> Vec<u8> {
    let bytes = if buffer.data.is_null() || buffer.len <= 0 {
        Vec::new()
    } else {
        unsafe { std::slice::from_raw_parts(buffer.data, buffer.len as usize).to_vec() }
    };
    unsafe {
        crate::xl_free_buffer(&mut buffer);
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
/// The result is cached: it cannot change for the lifetime of the process, and every
/// `Workbook::open` would otherwise pay an FFI call for it.
pub(crate) fn check_abi_version() -> Result<(), Error> {
    use std::sync::OnceLock;
    static CHECKED: OnceLock<Result<(), Error>> = OnceLock::new();

    CHECKED
        .get_or_init(|| {
            let loaded = unsafe { crate::xl_abi_version() };
            if loaded == crate::XL_ABI_VERSION {
                Ok(())
            } else {
                Err(Error::from_status(
                    XL_ERROR,
                    format!(
                        "ExcelReader native library reports ABI version {loaded}, but this crate \
                         was built against {}. Update the crate and the native library together.",
                        crate::XL_ABI_VERSION
                    ),
                ))
            }
        })
        .clone()
}

/// An open workbook. Not thread-safe - use one per thread, same contract as the C ABI. (The raw
/// handle makes this type neither `Send` nor `Sync`, so the compiler enforces that for you.)
pub struct Workbook {
    handle: *mut XlWorkbook,
}

impl Workbook {
    /// Opens `path`, sniffing the format and using every library default.
    ///
    /// Sniffing does NOT detect CSV - open one with [`open_with`](Self::open_with) and
    /// [`XL_FORMAT_CSV`](crate::XL_FORMAT_CSV).
    pub fn open(path: &str) -> Result<Workbook, Error> {
        Self::open_with(path, XL_FORMAT_AUTO, None)
    }

    /// Opens `path` with an explicit format and optional [`OpenOptions`].
    pub fn open_with(
        path: &str,
        format: i32,
        options: Option<&OpenOptions>,
    ) -> Result<Workbook, Error> {
        check_abi_version()?;
        // `raw` (and the password bytes it may borrow out of `options`) outlives the call below -
        // it is a local binding in this function's scope, not a temporary - and the native side
        // copies the path before returning.
        let raw = options.map(OpenOptions::to_raw);
        let raw_ptr = raw
            .as_ref()
            .map_or(std::ptr::null(), crate::options::OpenOptionsRaw::as_ptr);
        let mut handle: *mut XlWorkbook = std::ptr::null_mut();
        let status = unsafe {
            crate::xl_open_file_ex(
                path.as_ptr(),
                path.len() as i32,
                format,
                raw_ptr,
                &mut handle,
            )
        };
        check(status)?;
        Ok(Workbook { handle })
    }

    /// In-memory equivalent of [`open_with`](Self::open_with): `data` is copied by the native
    /// library, so it need not outlive this call.
    pub fn open_memory(
        data: &[u8],
        format: i32,
        options: Option<&OpenOptions>,
    ) -> Result<Workbook, Error> {
        check_abi_version()?;
        // Same lifetime shape as `open_with` above: `raw` is a local binding that outlives the FFI
        // call, so a password borrowed from `options` never dangles.
        let raw = options.map(OpenOptions::to_raw);
        let raw_ptr = raw
            .as_ref()
            .map_or(std::ptr::null(), crate::options::OpenOptionsRaw::as_ptr);
        let mut handle: *mut XlWorkbook = std::ptr::null_mut();
        let status = unsafe {
            crate::xl_open_memory_ex(
                data.as_ptr(),
                data.len() as i32,
                format,
                raw_ptr,
                &mut handle,
            )
        };
        check(status)?;
        Ok(Workbook { handle })
    }

    /// Number of sheets in the workbook.
    pub fn sheet_count(&self) -> Result<i32, Error> {
        let mut count: i32 = 0;
        check(unsafe { crate::xl_sheet_count(self.handle, &mut count) })?;
        Ok(count)
    }

    /// Name of the currently selected sheet.
    pub fn sheet_name(&self) -> Result<String, Error> {
        self.fill_string(|handle, buffer, capacity, out_len| unsafe {
            crate::xl_sheet_name(handle, buffer, capacity, out_len)
        })
    }

    /// Name of the sheet at `index`, without changing the current sheet or disturbing row
    /// enumeration.
    pub fn sheet_name_at(&self, index: i32) -> Result<String, Error> {
        self.fill_string(|handle, buffer, capacity, out_len| unsafe {
            crate::xl_sheet_name_at(handle, index, buffer, capacity, out_len)
        })
    }

    /// Every sheet name, in workbook order.
    pub fn sheet_names(&self) -> Result<Vec<String>, Error> {
        (0..self.sheet_count()?)
            .map(|index| self.sheet_name_at(index))
            .collect()
    }

    /// Selects the sheet at `index`, resetting row enumeration to its first row.
    ///
    /// Takes `&mut self` because it moves the cursor every subsequent read shares.
    pub fn move_to_sheet(&mut self, index: i32) -> Result<(), Error> {
        check(unsafe { crate::xl_move_to_sheet(self.handle, index) })
    }

    /// Whether the workbook uses the 1904 date system - needed to interpret raw Excel serial dates.
    pub fn is_date1904(&self) -> Result<bool, Error> {
        let mut flag: i32 = 0;
        check(unsafe { crate::xl_is_date1904(self.handle, &mut flag) })?;
        Ok(flag != 0)
    }

    /// A row-at-a-time reader over the current sheet.
    ///
    /// Takes `&mut self` because it advances the row cursor every read on this handle shares, the
    /// same reason [`move_to_sheet`](Self::move_to_sheet) does.
    pub fn rows(&mut self) -> crate::rows::RowCursor<'_> {
        crate::rows::RowCursor::new(self.handle)
    }

    /// Guesses a [`parse_sheet`] schema by sampling the current sheet.
    ///
    /// `header_row` has the same meaning as in [`parse_sheet`] (0 = no header); `sample_size`
    /// bounds how many rows after the header are inspected. This is a guess over a sample, not a
    /// guarantee - always check it fits before trusting it against the full sheet. Takes `&self`:
    /// the native call samples independently of the shared row cursor and never disturbs it.
    pub fn infer_schema(
        &self,
        header_row: i32,
        sample_size: i32,
    ) -> Result<Vec<InferredColumn>, Error> {
        let mut schema = XlInferredSchema {
            columns: std::ptr::null_mut(),
            column_count: 0,
        };
        check(unsafe {
            crate::xl_infer_schema(self.handle, header_row, sample_size, &mut schema)
        })?;

        // From here the schema is native-owned and must reach xl_free_schema. Nothing between this
        // point and the free can fail - `copy_inferred` only reads through pointers the ABI
        // guarantees - so a plain sequential free needs no drop guard.
        let columns = unsafe { copy_inferred(&schema) };
        unsafe { crate::xl_free_schema(&mut schema) };
        Ok(columns)
    }

    /// The raw handle, for sibling modules (e.g. `arrow::parse_arrow`) that need to call an
    /// `xl_*` function this struct has no wrapper for yet. Not part of the crate's public surface -
    /// `pub(crate)`, not `pub`.
    pub(crate) fn handle(&self) -> *mut XlWorkbook {
        self.handle
    }

    /// Shared two-pass buffer dance for the `xl_*` functions that write a UTF-8 name into a caller
    /// buffer and report the required capacity through `XL_BUFFER_TOO_SMALL`.
    fn fill_string(
        &self,
        call: impl Fn(*mut XlWorkbook, *mut u8, i32, *mut i32) -> i32,
    ) -> Result<String, Error> {
        // One sized attempt first: Excel caps sheet names at 31 characters, so 128 bytes clears
        // even the 4-byte-per-character worst case and the retry never runs in practice.
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
                .map_err(|e| {
                    Error::from_status(
                        XL_ERROR,
                        format!("native library returned a non-UTF-8 name: {e}"),
                    )
                });
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
        String::from_utf8(vec_buffer).map_err(|e| {
            Error::from_status(
                XL_ERROR,
                format!("native library returned a non-UTF-8 name: {e}"),
            )
        })
    }
}

/// Deep-copies a native-owned inferred schema into owned Rust values, so the caller can free the
/// native allocation immediately.
///
/// # Safety
/// `schema` must be one `xl_infer_schema` returned `XL_OK` for, not yet passed to `xl_free_schema`.
unsafe fn copy_inferred(schema: &XlInferredSchema) -> Vec<InferredColumn> {
    if schema.columns.is_null() || schema.column_count <= 0 {
        return Vec::new();
    }
    let specs = std::slice::from_raw_parts(schema.columns, schema.column_count as usize);
    specs
        .iter()
        .map(|spec| InferredColumn {
            // A guessed name is exactly `name_len` bytes with no NUL terminator, and is NULL
            // whenever the column had no usable header cell.
            name: if spec.name_count <= 0 || spec.names.is_null() || spec.name_lens.is_null() {
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
            },
            index: spec.index,
            column_type: spec.r#type,
            nullable: spec.nullable != 0,
        })
        .collect()
}

/// One column guessed by [`Workbook::infer_schema`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct InferredColumn {
    /// Header text, or `None` when the column must be resolved by [`index`](Self::index) instead.
    pub name: Option<String>,
    pub index: i32,
    /// One of the `XL_T_*` constants.
    pub column_type: i32,
    pub nullable: bool,
}

impl Drop for Workbook {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe {
                crate::xl_close(self.handle);
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

/// One field <-> column binding for `T`. Construct via `ExcelMapper::bindings()`.
pub struct ColumnBinding<T> {
    /// Candidate header names, in priority order — the first one present in the header row wins.
    pub names: &'static [&'static str],
    pub xl_type: i32,
    pub assign: fn(&mut T, &XlColumn, i64),
}

/// Implemented by any struct `parse_sheet` can populate from a `parse_typed` result. Implement
/// `bindings()` by hand (mirroring `xl::ExcelMapper<T>` on the C++ side), or derive it with
/// `#[derive(ExcelMapper)]` (see `excelreader_derive`).
pub trait ExcelMapper: Sized {
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
/// type half catches the other class of mistake - reading an `XL_T_F64` buffer as `i64` is an
/// `ExcelMapper::bindings()` bug, not recoverable input.
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
/// Validation is a linear scan over bytes already in cache.
///
/// # Panics
/// If `col` is not an `XL_T_STRING` column, if `row` is out of bounds, or if the native library
/// returned bytes that are not valid UTF-8 (an ABI contract violation).
pub fn column_str(col: &XlColumn, row: i64) -> &str {
    check_access(col, row, XL_T_STRING, "column_str");
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
/// If `col` is not an `XL_T_I64` column, or `row` is out of bounds.
pub fn column_i64(col: &XlColumn, row: i64) -> i64 {
    check_access(col, row, XL_T_I64, "column_i64");
    unsafe { *(col.values as *const i64).offset(row as isize) }
}

/// # Panics
/// If `col` is not an `XL_T_F64` column, or `row` is out of bounds.
pub fn column_f64(col: &XlColumn, row: i64) -> f64 {
    check_access(col, row, XL_T_F64, "column_f64");
    unsafe { *(col.values as *const f64).offset(row as isize) }
}

/// # Panics
/// If `col` is not an `XL_T_BOOL` column, or `row` is out of bounds.
pub fn column_bool(col: &XlColumn, row: i64) -> bool {
    check_access(col, row, XL_T_BOOL, "column_bool");
    unsafe { *(col.values as *const u8).offset(row as isize) != 0 }
}

/// # Panics
/// If `col` is not an `XL_T_DATE` column, or `row` is out of bounds.
pub fn column_date(col: &XlColumn, row: i64) -> Date {
    check_access(col, row, XL_T_DATE, "column_date");
    Date::new(unsafe { *(col.values as *const i32).offset(row as isize) })
}

/// # Panics
/// If `col` is not an `XL_T_TIME` column, or `row` is out of bounds.
pub fn column_time(col: &XlColumn, row: i64) -> Time {
    check_access(col, row, XL_T_TIME, "column_time");
    Time::new(unsafe { *(col.values as *const i64).offset(row as isize) })
}

/// # Panics
/// If `col` is not an `XL_T_TIMESTAMP` column, or `row` is out of bounds.
pub fn column_timestamp(col: &XlColumn, row: i64) -> Timestamp {
    check_access(col, row, XL_T_TIMESTAMP, "column_timestamp");
    Timestamp::new(unsafe { *(col.values as *const i64).offset(row as isize) })
}

/// Owns a `parse_typed` result. Frees the native table on `Drop`. `iter()` builds one `T` per call
/// to `next()` from the columnar buffers - no upfront `Vec<T>` allocation beyond what
/// `xl_parse_typed` itself already made.
pub struct TableView<T: ExcelMapper> {
    table: XlTable,
    bindings: Vec<ColumnBinding<T>>,
    _marker: PhantomData<T>,
}

impl<T: ExcelMapper> TableView<T> {
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
            crate::xl_free_table(&mut self.table);
        }
    }
}

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

pub struct TableViewIter<'a, T: ExcelMapper> {
    view: &'a TableView<T>,
    row: i64,
}

impl<T: ExcelMapper + Default> Iterator for TableViewIter<'_, T> {
    type Item = T;

    fn next(&mut self) -> Option<T> {
        // `get` re-checks the bound, so this cannot walk off the end even if `row` were wrong.
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
            r#type: b.xl_type,
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

/// Schema-driven columnar parse of the current sheet, matching C++'s `xl::parse_sheet<T>`.
///
/// Takes `&mut Workbook` because the parse consumes the workbook's shared row cursor.
pub fn parse_sheet<T: ExcelMapper>(
    workbook: &mut Workbook,
    header_row: i32,
) -> Result<TableView<T>, Error> {
    let arena = build_specs::<T>();
    let bindings = arena.bindings;
    let mut table = XlTable {
        column_count: 0,
        row_count: 0,
        columns: std::ptr::null_mut(),
    };
    unsafe {
        let status = crate::xl_parse_typed(
            workbook.handle,
            arena.specs.as_ptr(),
            arena.specs.len() as i32,
            header_row,
            &mut table,
        );
        if status != XL_OK {
            return Err(last_error(status));
        }
    }

    // xl_parse_typed returns one column per spec, in spec order - but the zip in `get` would
    // silently drop trailing bindings if that ever stopped holding, quietly leaving those fields at
    // their default rather than failing. Check it once here instead of per row.
    if table.column_count as usize != bindings.len() {
        let column_count = table.column_count;
        unsafe { crate::xl_free_table(&mut table) };
        return Err(Error::from_status(
            XL_ERROR,
            format!(
                "xl_parse_typed returned {column_count} columns for {} specs",
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
