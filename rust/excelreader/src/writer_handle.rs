//! Row-by-row streaming write, through the ABI's `xl_writer_handle`.
//!
//! [`WriterHandle`] is the streaming counterpart to [`crate::write_columns`]/[`crate::write_sheet`]:
//! instead of materializing a whole table up front, it writes directly as each call arrives - one
//! sheet and one row open at a time. Call order mirrors the C ABI (see `xl_writer_handle` in
//! `excelreader.h`): `create`/`new`, then per sheet `start_sheet`..`end_sheet`, each containing
//! `start_row`..`end_row` with one `write_*` call per cell in between, left to right, or one
//! `write_columns`/`write_sheet` per whole sheet, then `finish`. A call out of order returns an `Err` rather than corrupting output, and the handle
//! stays usable afterward - fix the call order and continue, or let `Drop` discard it.

use crate::ffi::{self, XlWriterHandle};
use crate::workbook::{buffer_to_vec, check, check_abi_version, len_i32, path_str};
use crate::writer::{raw_table, OwnedColumn};
use crate::{
    Column, ColumnType, Date, Error, ExcelWriter, TableOptions, Time, Timestamp, WriteOptions,
};
use std::io::Write;
use std::path::Path;

/// A streaming writer - the row-by-row counterpart to [`Workbook`](crate::Workbook).
///
/// Not thread-safe - use one per thread, same contract as the C ABI. (The raw handle makes this
/// type neither `Send` nor `Sync`, so the compiler enforces that for you.)
///
/// Call [`finish`](Self::finish) to complete the workbook and observe any error. Only `finish`
/// commits: dropping a handle without finishing it - an early `?` return, a panic - discards
/// everything written, leaves an existing file at the path as it was, and writes nothing to an
/// `io::Write` destination.
pub struct WriterHandle<'w> {
    handle: *mut XlWriterHandle,
    writer: Option<Box<dyn Write + 'w>>,
}

impl WriterHandle<'static> {
    /// Writes `path`, replacing an existing file only when [`finish`](Self::finish) succeeds. The
    /// format is [`WriteOptions::format`], or inferred from the extension when that is
    /// [`Format::Auto`](crate::Format::Auto).
    pub fn create(path: impl AsRef<Path>, options: &WriteOptions) -> Result<Self, Error> {
        check_abi_version()?;
        let path = path.as_ref();
        let format = options.format_for(Some(path))?;
        let path = path_str(path)?;
        let path_len = len_i32(path.len(), "path")?;
        let raw = options.to_raw();
        let mut handle: *mut XlWriterHandle = std::ptr::null_mut();
        check(unsafe {
            ffi::xl_open_write_handle(path.as_ptr(), path_len, format, &raw, &mut handle)
        })?;
        Ok(WriterHandle {
            handle,
            writer: None,
        })
    }
}

impl<'w> WriterHandle<'w> {
    /// Builds the workbook in memory and writes it to `writer`, once, in [`finish`](Self::finish).
    /// [`WriteOptions::format`] must be set - there is no extension to infer one from.
    pub fn new(writer: impl Write + 'w, options: &WriteOptions) -> Result<Self, Error> {
        check_abi_version()?;
        let format = options.format_for(None)?;
        let raw = options.to_raw();
        let mut handle: *mut XlWriterHandle = std::ptr::null_mut();
        check(unsafe { ffi::xl_open_write_handle_to_memory(format, &raw, &mut handle) })?;
        Ok(WriterHandle {
            handle,
            writer: Some(Box::new(writer)),
        })
    }

    /// Starts a new sheet named `name`. Must not be called again before the current sheet, if
    /// any, has been ended with [`end_sheet`](Self::end_sheet).
    pub fn start_sheet(&mut self, name: &str) -> Result<(), Error> {
        let name_len = len_i32(name.len(), "sheet name")?;
        check(unsafe { ffi::xl_start_sheet(self.handle, name.as_ptr(), name_len) })
    }

    /// Starts a new row on the current sheet. Must not be called again before the current row, if
    /// any, has been ended with [`end_row`](Self::end_row).
    pub fn start_row(&mut self) -> Result<(), Error> {
        check(unsafe { ffi::xl_start_row(self.handle) })
    }

    /// Writes the next cell of the current row as text, or a blank cell for `None`.
    pub fn write_str(&mut self, value: Option<&str>) -> Result<(), Error> {
        let (ptr, len) = match value {
            Some(text) => (text.as_ptr(), len_i32(text.len(), "text")?),
            None => (std::ptr::null(), 0),
        };
        check(unsafe { ffi::xl_write_string(self.handle, ptr, len) })
    }

    /// Writes the next cell of the current row as an integer, or a blank cell for `None`.
    pub fn write_i64(&mut self, value: Option<i64>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_int64(self.handle, v) }),
            None => self.write_null(ColumnType::I64),
        }
    }

    /// Writes the next cell of the current row as a floating-point number, or a blank cell for
    /// `None`.
    pub fn write_f64(&mut self, value: Option<f64>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_float64(self.handle, v) }),
            None => self.write_null(ColumnType::F64),
        }
    }

    /// Writes the next cell of the current row as a boolean, or a blank cell for `None`.
    pub fn write_bool(&mut self, value: Option<bool>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_bool(self.handle, i32::from(v)) }),
            None => self.write_null(ColumnType::Bool),
        }
    }

    /// Writes the next cell of the current row as a date, or a blank cell for `None`.
    pub fn write_date(&mut self, value: Option<Date>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_date(self.handle, v.days_since_epoch) }),
            None => self.write_null(ColumnType::Date),
        }
    }

    /// Writes the next cell of the current row as a time of day, or a blank cell for `None`.
    pub fn write_time(&mut self, value: Option<Time>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_time(self.handle, v.micros_since_midnight) }),
            None => self.write_null(ColumnType::Time),
        }
    }

    /// Writes the next cell of the current row as a date/time, or a blank cell for `None`.
    pub fn write_timestamp(&mut self, value: Option<Timestamp>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { ffi::xl_write_timestamp(self.handle, v.micros_since_epoch) }),
            None => self.write_null(ColumnType::Timestamp),
        }
    }

    fn write_null(&mut self, column_type: ColumnType) -> Result<(), Error> {
        check(unsafe { ffi::xl_write_null(self.handle, column_type.raw()) })
    }

    /// Ends the current row, started by [`start_row`](Self::start_row).
    pub fn end_row(&mut self) -> Result<(), Error> {
        check(unsafe { ffi::xl_end_row(self.handle) })
    }

    /// Ends the current sheet, started by [`start_sheet`](Self::start_sheet). Must not be called
    /// with a row still open.
    pub fn end_sheet(&mut self) -> Result<(), Error> {
        check(unsafe { ffi::xl_end_sheet(self.handle) })
    }

    /// Writes `columns` as a whole sheet named `name`, as [`write_columns`](crate::write_columns)
    /// does, in one call instead of one per cell; with `table`, as an Excel table (XLSX/XLSB only,
    /// every column named). Uses the parallelism set in the handle's [`WriteOptions`]; its sheet
    /// name does not apply. Fails while a sheet started with [`start_sheet`](Self::start_sheet) is
    /// open.
    pub fn write_columns(
        &mut self,
        name: &str,
        columns: &[Column<'_>],
        table: Option<&TableOptions>,
    ) -> Result<(), Error> {
        let name_len = len_i32(name.len(), "name")?;
        let raw = raw_table(columns)?;
        let options = table.map(TableOptions::to_raw).transpose()?;
        check(unsafe {
            ffi::xl_write_sheet(
                self.handle,
                name.as_ptr(),
                name_len,
                raw.specs.as_ptr(),
                &raw.table,
                options.as_ref().map_or(std::ptr::null(), |o| o),
            )
        })
    }

    /// Writes `rows` as a whole sheet named `name`, as [`write_sheet`](crate::write_sheet) does,
    /// on top of [`write_columns`](Self::write_columns).
    pub fn write_sheet<T: ExcelWriter>(
        &mut self,
        name: &str,
        rows: &[T],
        table: Option<&TableOptions>,
    ) -> Result<(), Error> {
        let owned = T::to_columns(rows)?;
        let borrowed: Vec<Column<'_>> = owned.iter().map(OwnedColumn::as_column).collect();
        self.write_columns(name, &borrowed, table)
    }

    /// Completes the workbook and releases the handle. For a handle from [`new`](Self::new), the
    /// whole workbook is then written to its `io::Write` in one `write_all`.
    pub fn finish(mut self) -> Result<(), Error> {
        let handle = std::mem::replace(&mut self.handle, std::ptr::null_mut());
        let bytes = self.writer.is_some().then(|| {
            let mut buffer = ffi::XlBuffer {
                data: std::ptr::null_mut(),
                len: 0,
            };
            check(unsafe { ffi::xl_write_handle_bytes(handle, &mut buffer) })
                .map(|()| buffer_to_vec(buffer))
        });
        let closed = check(unsafe { ffi::xl_close_write_handle(handle) });
        let bytes = bytes.transpose()?;
        closed?;
        if let (Some(bytes), Some(writer)) = (bytes, self.writer.as_mut()) {
            writer
                .write_all(&bytes)
                .and_then(|()| writer.flush())
                .map_err(|e| Error::io(&e))?;
        }
        Ok(())
    }
}

impl Drop for WriterHandle<'_> {
    fn drop(&mut self) {
        if self.handle.is_null() {
            return;
        }
        unsafe {
            ffi::xl_abort_write_handle(self.handle);
        }
        self.handle = std::ptr::null_mut();
    }
}

/// Deliberately opaque - same reasoning as `Workbook`'s `Debug` impl: the only state here is a
/// handle whose value means nothing outside the native library.
impl std::fmt::Debug for WriterHandle<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WriterHandle").finish_non_exhaustive()
    }
}
