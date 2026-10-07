//! Arrow C Data Interface import, behind the `arrow` cargo feature.
//!
//! The native side already produces one top-level Arrow struct array whose children are the
//! columns, so the only work here is handing arrow-rs the FFI pair and letting it take ownership.

use std::ffi::c_void;
use std::marker::PhantomData;

use arrow::array::{RecordBatch, StructArray};
use arrow::datatypes::SchemaRef;
use arrow::error::ArrowError;
use arrow::ffi::{from_ffi, FFI_ArrowArray, FFI_ArrowSchema};
use arrow::ffi_stream::{ArrowArrayStreamReader, FFI_ArrowArrayStream};
use arrow::record_batch::RecordBatchReader;

use crate::workbook::{build_specs, check, ExcelMapper, Sheet, Workbook};
use crate::{Error, XL_ERROR};

/// Schema-driven parse of one sheet into an Arrow [`RecordBatch`], using the same
/// `#[derive(ExcelMapper)]` mapping as [`crate::workbook::parse_sheet`].
///
/// `header_row` has the same meaning as in `parse_sheet` (0 = no header).
pub fn parse_arrow<T: ExcelMapper>(sheet: Sheet<'_>, header_row: i32) -> Result<RecordBatch, Error> {
    let arena = build_specs::<T>();

    let mut array = FFI_ArrowArray::empty();
    let mut schema = FFI_ArrowSchema::empty();

    check(unsafe {
        crate::xl_parse_arrow(
            sheet.handle(),
            sheet.index(),
            arena.specs.as_ptr(),
            arena.specs.len() as i32,
            header_row,
            1,
            &mut array as *mut FFI_ArrowArray as *mut c_void,
            &mut schema as *mut FFI_ArrowSchema as *mut c_void,
        )
    })?;

    let data = unsafe { from_ffi(array, &schema) }.map_err(|e| {
        Error::from_status(
            XL_ERROR,
            format!("importing the native Arrow array failed: {e}"),
        )
    })?;

    Ok(RecordBatch::from(StructArray::from(data)))
}

/// Batched counterpart to [`parse_arrow`], one [`RecordBatch`] at a time.
///
/// Wraps arrow-rs's `'static` reader only to carry the `'a` borrow, which keeps the
/// workbook alive for at least as long as the stream. Unlike [`crate::workbook::TypedChunks`] it does
/// not fuse after an error, so break on the first `Err` rather than spinning on the latched one.
pub struct ArrowChunks<'a> {
    inner: ArrowArrayStreamReader,
    _workbook: PhantomData<&'a Workbook>,
}

impl Iterator for ArrowChunks<'_> {
    type Item = Result<RecordBatch, ArrowError>;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}

impl RecordBatchReader for ArrowChunks<'_> {
    fn schema(&self) -> SchemaRef {
        self.inner.schema()
    }
}

/// `parse_arrow` delivered a batch at a time. `batch_size` is rows per batch, 0 unbounded,
/// negative an error. Any number of streams may be open on one workbook.
pub fn parse_arrow_stream<T: ExcelMapper>(
    sheet: Sheet<'_>,
    header_row: i32,
    batch_size: i64,
) -> Result<ArrowChunks<'_>, Error> {
    let arena = build_specs::<T>();

    let mut stream = FFI_ArrowArrayStream::empty();

    check(unsafe {
        crate::xl_parse_arrow_stream(
            sheet.handle(),
            sheet.index(),
            arena.specs.as_ptr(),
            arena.specs.len() as i32,
            header_row,
            batch_size,
            &mut stream as *mut FFI_ArrowArrayStream as *mut c_void,
        )
    })?;

    let inner = ArrowArrayStreamReader::try_new(stream).map_err(|e| {
        Error::from_status(
            XL_ERROR,
            format!("importing the native Arrow stream failed: {e}"),
        )
    })?;

    Ok(ArrowChunks {
        inner,
        _workbook: PhantomData,
    })
}
