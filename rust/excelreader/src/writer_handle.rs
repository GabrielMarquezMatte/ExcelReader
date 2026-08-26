//! Row-by-row streaming write, through the ABI's `xl_writer_handle`.
//!
//! [`WriterHandle`] is the streaming counterpart to [`crate::writer::write_columns`]/
//! [`crate::writer::write_sheet`]: instead of materializing a whole table up front, it writes
//! directly as each call arrives - one sheet and one row open at a time. Call order mirrors the C
//! ABI (see `xl_writer_handle` in `excelreader.h`): `open`/`open_with`/`open_memory`, then per
//! sheet `start_sheet`..`end_sheet`, each containing `start_row`..`end_row` with one `write_*`
//! call per cell in between, left to right. A call out of order returns an `Err` rather than
//! corrupting output, and the handle stays usable afterward - fix the call order and continue, or
//! let `Drop` discard it.

use crate::workbook::{buffer_to_vec, check, check_abi_version};
use crate::writer::format_from_path;
use crate::{
    Date, Error, Time, Timestamp, WriteOptions, XlBuffer, XlWriterHandle, XL_T_BOOL, XL_T_DATE,
    XL_T_F64, XL_T_I64, XL_T_TIME, XL_T_TIMESTAMP,
};

/// A streaming writer handle - the row-by-row counterpart to
/// [`Workbook`](crate::workbook::Workbook).
///
/// Not thread-safe - use one per thread, same contract as the C ABI. (The raw handle makes this
/// type neither `Send` nor `Sync`, so the compiler enforces that for you.)
///
/// Dropping a `WriterHandle` closes and releases it (`xl_close_write_handle`), silently discarding
/// any error - same convention as [`Workbook`](crate::workbook::Workbook)'s `Drop`. Call
/// [`bytes`](Self::bytes) (memory-backed) or reopen the path (file-backed) to observe the actual
/// result; do not rely on the drop for that.
pub struct WriterHandle {
    handle: *mut XlWriterHandle,
}

impl WriterHandle {
    /// Creates `path` (truncating it if it already exists), inferring the format from its
    /// extension via [`format_from_path`](crate::writer::format_from_path).
    pub fn open(path: &str, options: Option<&WriteOptions>) -> Result<Self, Error> {
        Self::open_with(path, format_from_path(path), options)
    }

    /// Creates `path` with an explicit format. `format` must be one of `XL_FORMAT_XLS`/`XLSX`/
    /// `XLSB`/`CSV` - [`XL_FORMAT_AUTO`](crate::XL_FORMAT_AUTO) is an error, the same as
    /// [`write_columns`](crate::writer::write_columns).
    pub fn open_with(
        path: &str,
        format: i32,
        options: Option<&WriteOptions>,
    ) -> Result<Self, Error> {
        check_abi_version()?;
        let raw = options.map(WriteOptions::to_raw);
        let raw_ptr = crate::options::ptr_or_null(&raw);
        let mut handle: *mut XlWriterHandle = std::ptr::null_mut();
        // `raw` outlives the call below, and the native side copies the path before returning.
        let status = unsafe {
            crate::xl_open_write_handle(
                path.as_ptr(),
                path.len() as i32,
                format,
                raw_ptr,
                &mut handle,
            )
        };
        check(status)?;
        Ok(WriterHandle { handle })
    }

    /// In-memory equivalent of [`open_with`](Self::open_with): read the result back with
    /// [`bytes`](Self::bytes). `format` is always required here - there is no path to infer one
    /// from.
    pub fn open_memory(format: i32, options: Option<&WriteOptions>) -> Result<Self, Error> {
        check_abi_version()?;
        let raw = options.map(WriteOptions::to_raw);
        let raw_ptr = crate::options::ptr_or_null(&raw);
        let mut handle: *mut XlWriterHandle = std::ptr::null_mut();
        let status = unsafe { crate::xl_open_write_handle_to_memory(format, raw_ptr, &mut handle) };
        check(status)?;
        Ok(WriterHandle { handle })
    }

    /// Starts a new sheet named `name`. Must not be called again before the current sheet, if
    /// any, has been ended with [`end_sheet`](Self::end_sheet).
    pub fn start_sheet(&mut self, name: &str) -> Result<(), Error> {
        check(unsafe { crate::xl_start_sheet(self.handle, name.as_ptr(), name.len() as i32) })
    }

    /// Starts a new row on the current sheet. Must not be called again before the current row, if
    /// any, has been ended with [`end_row`](Self::end_row).
    pub fn start_row(&mut self) -> Result<(), Error> {
        check(unsafe { crate::xl_start_row(self.handle) })
    }

    /// Writes the next cell of the current row as text, or a blank cell for `None`.
    pub fn write_str(&mut self, value: Option<&str>) -> Result<(), Error> {
        let (ptr, len) = value.map_or((std::ptr::null(), 0), |text| {
            (text.as_ptr(), text.len() as i32)
        });
        check(unsafe { crate::xl_write_string(self.handle, ptr, len) })
    }

    /// Writes the next cell of the current row as an integer, or a blank cell for `None`.
    pub fn write_i64(&mut self, value: Option<i64>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { crate::xl_write_int64(self.handle, v) }),
            None => self.write_null(XL_T_I64),
        }
    }

    /// Writes the next cell of the current row as a floating-point number, or a blank cell for
    /// `None`.
    pub fn write_f64(&mut self, value: Option<f64>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { crate::xl_write_float64(self.handle, v) }),
            None => self.write_null(XL_T_F64),
        }
    }

    /// Writes the next cell of the current row as a boolean, or a blank cell for `None`.
    pub fn write_bool(&mut self, value: Option<bool>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { crate::xl_write_bool(self.handle, i32::from(v)) }),
            None => self.write_null(XL_T_BOOL),
        }
    }

    /// Writes the next cell of the current row as a date, or a blank cell for `None`.
    pub fn write_date(&mut self, value: Option<Date>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { crate::xl_write_date(self.handle, v.days_since_epoch) }),
            None => self.write_null(XL_T_DATE),
        }
    }

    /// Writes the next cell of the current row as a time of day, or a blank cell for `None`.
    pub fn write_time(&mut self, value: Option<Time>) -> Result<(), Error> {
        match value {
            Some(v) => check(unsafe { crate::xl_write_time(self.handle, v.micros_since_midnight) }),
            None => self.write_null(XL_T_TIME),
        }
    }

    /// Writes the next cell of the current row as a date/time, or a blank cell for `None`.
    pub fn write_timestamp(&mut self, value: Option<Timestamp>) -> Result<(), Error> {
        match value {
            Some(v) => {
                check(unsafe { crate::xl_write_timestamp(self.handle, v.micros_since_epoch) })
            }
            None => self.write_null(XL_T_TIMESTAMP),
        }
    }

    /// Writes a blank cell of the given `XL_T_*` type directly. Every `write_*` method above
    /// already does this for `None`; this is for a caller building a cell from something other
    /// than an `Option<T>`.
    pub fn write_null(&mut self, xl_type: i32) -> Result<(), Error> {
        check(unsafe { crate::xl_write_null(self.handle, xl_type) })
    }

    /// Ends the current row, started by [`start_row`](Self::start_row).
    pub fn end_row(&mut self) -> Result<(), Error> {
        check(unsafe { crate::xl_end_row(self.handle) })
    }

    /// Ends the current sheet, started by [`start_sheet`](Self::start_sheet). Must not be called
    /// with a row still open.
    pub fn end_sheet(&mut self) -> Result<(), Error> {
        check(unsafe { crate::xl_end_sheet(self.handle) })
    }

    /// Reads back everything written so far - only valid for a handle from
    /// [`open_memory`](Self::open_memory); `XL_INVALID_ARGUMENT` for one from
    /// [`open`](Self::open)/[`open_with`](Self::open_with). Ends the workbook's trailing structure
    /// if that has not already happened, but does NOT release the handle: it stays open (and
    /// closeable) exactly like a file-backed one - `Drop` still closes it afterward.
    pub fn bytes(&mut self) -> Result<Vec<u8>, Error> {
        let mut buffer = XlBuffer {
            data: std::ptr::null_mut(),
            len: 0,
        };
        let status = unsafe { crate::xl_write_handle_bytes(self.handle, &mut buffer) };
        check(status)?;
        Ok(buffer_to_vec(buffer))
    }
}

impl Drop for WriterHandle {
    fn drop(&mut self) {
        if self.handle.is_null() {
            return;
        }
        unsafe {
            crate::xl_close_write_handle(self.handle);
        }
        self.handle = std::ptr::null_mut();
    }
}

/// Deliberately opaque - same reasoning as `Workbook`'s `Debug` impl: the only state here is a
/// handle whose value means nothing outside the native library.
impl std::fmt::Debug for WriterHandle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WriterHandle").finish_non_exhaustive()
    }
}
