//! Workbooks opened from bytes the caller serves: a positional [`Source`] or any [`Read`].

use crate::ffi::{self, XlSource, XlStream};
use crate::{Error, ErrorKind, OpenOptions, Workbook};
use std::any::Any;
use std::ffi::c_void;
use std::io::{self, Read};
use std::panic::{catch_unwind, AssertUnwindSafe};

/// Random-access bytes for [`OpenOptions::open_source`]. The library calls `read_at` from several
/// threads at once, which is why the trait requires `Send + Sync`.
pub trait Source: Send + Sync + 'static {
    /// Total size in bytes. Called once, while the workbook opens.
    fn size(&self) -> io::Result<u64>;

    /// Fills `buf` with the bytes starting at `offset` and returns how many it wrote. Returning 0
    /// before the end is a failure.
    fn read_at(&self, offset: u64, buf: &mut [u8]) -> io::Result<usize>;
}

impl OpenOptions {
    /// Opens a workbook over `source`, which the library owns from here on and drops once nothing
    /// opened from the workbook needs it. Every sheet can be read in parallel.
    pub fn open_source<S: Source>(&self, source: S) -> Result<Workbook, Error> {
        let size = source
            .size()
            .map_err(|error| Error::new(ErrorKind::InvalidArgument, error.to_string()))?;
        let length = i64::try_from(size).map_err(|_| {
            Error::new(
                ErrorKind::InvalidArgument,
                "the source size does not fit in an i64",
            )
        })?;
        self.open_raw(|format, options, handle| {
            let raw = XlSource {
                struct_size: std::mem::size_of::<XlSource>() as i32,
                user_data: Box::into_raw(Box::new(source)).cast(),
                length,
                read_at: Some(read_at::<S>),
                release: Some(release::<S>),
            };
            unsafe { ffi::xl_open_source(&raw, format, options, handle) }
        })
    }

    /// Opens a workbook over `reader`, which cannot seek. A CSV is read as it arrives, and again
    /// from the start until 16 MiB of it were read; an XLSX, XLSB or XLS is read whole first. The
    /// library owns `reader` from here on.
    pub fn open_reader<R: Read + Send + 'static>(&self, reader: R) -> Result<Workbook, Error> {
        self.open_raw(|format, options, handle| {
            let raw = XlStream {
                struct_size: std::mem::size_of::<XlStream>() as i32,
                user_data: Box::into_raw(Box::new(reader)).cast(),
                read: Some(read::<R>),
                release: Some(release::<R>),
            };
            unsafe { ffi::xl_open_stream(&raw, format, options, handle) }
        })
    }
}

unsafe extern "C" fn read_at<S: Source>(
    user_data: *mut c_void,
    offset: i64,
    buf: *mut u8,
    len: i64,
) -> i64 {
    let source = &*user_data.cast::<S>();
    let buf = std::slice::from_raw_parts_mut(buf, len as usize);
    report(catch_unwind(AssertUnwindSafe(|| {
        source.read_at(offset as u64, buf)
    })))
}

unsafe extern "C" fn read<R: Read>(user_data: *mut c_void, buf: *mut u8, len: i64) -> i64 {
    let reader = &mut *user_data.cast::<R>();
    let buf = std::slice::from_raw_parts_mut(buf, len as usize);
    report(catch_unwind(AssertUnwindSafe(|| loop {
        match reader.read(buf) {
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            other => return other,
        }
    })))
}

unsafe extern "C" fn release<T>(user_data: *mut c_void) {
    let _ = catch_unwind(AssertUnwindSafe(|| {
        drop(Box::from_raw(user_data.cast::<T>()))
    }));
}

fn report(outcome: Result<io::Result<usize>, Box<dyn Any + Send>>) -> i64 {
    let message = match outcome {
        Ok(Ok(count)) => return count as i64,
        Ok(Err(error)) => error.to_string(),
        Err(_) => "source panicked".to_string(),
    };
    unsafe { ffi::xl_set_source_error(message.as_ptr(), message.len() as i32) };
    -1
}
