use crate::{
    Error, XlColumn, XlColumnSpec, XlTable, XlWorkbook, XL_FORMAT_AUTO, XL_OK, XL_T_BOOL,
    XL_T_F64, XL_T_I64, XL_T_STRING,
};
use std::marker::PhantomData;

fn last_error(code: i32) -> Error {
    unsafe {
        let mut len: i32 = 0;
        let ptr = crate::xl_last_error_ptr(&mut len);
        let message = if ptr.is_null() || len <= 0 {
            "unknown error".to_string()
        } else {
            let bytes = std::slice::from_raw_parts(ptr, len as usize);
            String::from_utf8_lossy(bytes).into_owned()
        };
        Error { code, message }
    }
}

/// An open workbook. Not thread-safe - use one per thread, same contract as the C ABI.
pub struct Workbook {
    handle: *mut XlWorkbook,
}

impl Workbook {
    pub fn open(path: &str) -> Result<Workbook, Error> {
        unsafe {
            let mut handle: *mut XlWorkbook = std::ptr::null_mut();
            let status = crate::xl_open_file_ex(
                path.as_ptr(),
                path.len() as i32,
                XL_FORMAT_AUTO,
                std::ptr::null(),
                &mut handle,
            );
            if status != XL_OK {
                return Err(last_error(status));
            }
            Ok(Workbook { handle })
        }
    }
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

/// One field <-> column binding for `T`. Construct via `ExcelMapper::bindings()`.
pub struct ColumnBinding<T> {
    pub name: &'static str,
    pub xl_type: i32,
    pub assign: fn(&mut T, &XlColumn, i64),
}

/// Implemented by any struct `parse_sheet` can populate from a `parse_typed` result. No derive
/// macro in v1 - implement `bindings()` by hand, mirroring `xl::ExcelMapper<T>` on the C++ side.
pub trait ExcelMapper: Sized {
    fn bindings() -> Vec<ColumnBinding<Self>>;
}

fn is_valid(col: &XlColumn, row: i64) -> bool {
    if col.validity.is_null() {
        return true;
    }
    unsafe {
        let byte = *col.validity.offset((row / 8) as isize);
        (byte & (1 << (row % 8))) != 0
    }
}

/// Reads column `col` at `row` as a `&str`. Only valid for `XL_T_STRING` columns - panics
/// otherwise, since a type mismatch here is an `ExcelMapper::bindings()` bug, not recoverable input.
pub fn column_str(col: &XlColumn, row: i64) -> &str {
    assert_eq!(col.r#type, XL_T_STRING, "column_str called on a non-string column");
    unsafe {
        let offsets = col.values as *const i32;
        let start = *offsets.offset(row as isize);
        let end = *offsets.offset(row as isize + 1);
        let bytes = std::slice::from_raw_parts(col.data.offset(start as isize), (end - start) as usize);
        std::str::from_utf8_unchecked(bytes)
    }
}

pub fn column_i64(col: &XlColumn, row: i64) -> i64 {
    assert_eq!(col.r#type, XL_T_I64, "column_i64 called on a non-I64 column");
    unsafe { *(col.values as *const i64).offset(row as isize) }
}

pub fn column_f64(col: &XlColumn, row: i64) -> f64 {
    assert_eq!(col.r#type, XL_T_F64, "column_f64 called on a non-F64 column");
    unsafe { *(col.values as *const f64).offset(row as isize) }
}

pub fn column_bool(col: &XlColumn, row: i64) -> bool {
    assert_eq!(col.r#type, XL_T_BOOL, "column_bool called on a non-BOOL column");
    unsafe { *(col.values as *const u8).offset(row as isize) != 0 }
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

    pub fn get(&self, row: i64) -> T
    where
        T: Default,
    {
        let mut instance = T::default();
        let columns = unsafe { std::slice::from_raw_parts(self.table.columns, self.table.column_count as usize) };
        for (col, binding) in columns.iter().zip(self.bindings.iter()) {
            if is_valid(col, row) {
                (binding.assign)(&mut instance, col, row);
            }
        }
        instance
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

pub struct TableViewIter<'a, T: ExcelMapper> {
    view: &'a TableView<T>,
    row: i64,
}

impl<'a, T: ExcelMapper + Default> Iterator for TableViewIter<'a, T> {
    type Item = T;
    fn next(&mut self) -> Option<T> {
        if self.row >= self.view.len() {
            return None;
        }
        let item = self.view.get(self.row);
        self.row += 1;
        Some(item)
    }
}

/// Schema-driven columnar parse of the current sheet, matching C++'s `xl::parse_sheet<T>`.
pub fn parse_sheet<T: ExcelMapper>(workbook: &Workbook, header_row: i32) -> Result<TableView<T>, Error> {
    let bindings = T::bindings();
    let specs: Vec<XlColumnSpec> = bindings
        .iter()
        .map(|b| XlColumnSpec {
            name: b.name.as_ptr(),
            name_len: b.name.len() as i32,
            index: 0,
            r#type: b.xl_type,
            nullable: 1,
        })
        .collect();

    let mut table = XlTable {
        column_count: 0,
        row_count: 0,
        columns: std::ptr::null_mut(),
    };
    unsafe {
        let status = crate::xl_parse_typed(
            workbook.handle,
            specs.as_ptr(),
            specs.len() as i32,
            header_row,
            &mut table,
        );
        if status != XL_OK {
            return Err(last_error(status));
        }
    }
    Ok(TableView {
        table,
        bindings,
        _marker: PhantomData,
    })
}
