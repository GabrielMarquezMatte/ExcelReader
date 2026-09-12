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

// `check` is `pub(crate)` in workbook.rs (not error.rs) - visible from this sibling module because
// pub(crate) means "crate-wide", not "same file".
use crate::workbook::{build_specs, check, ExcelMapper, Workbook};
use crate::{Error, XL_ERROR};

/// Schema-driven parse of the current sheet into an Arrow [`RecordBatch`], using the same
/// `#[derive(ExcelMapper)]` mapping as [`crate::workbook::parse_sheet`].
///
/// `header_row` has the same meaning as in `parse_sheet` (0 = no header). Takes `&mut Workbook`
/// because the parse consumes the workbook's shared row cursor.
pub fn parse_arrow<T: ExcelMapper>(
    workbook: &mut Workbook,
    header_row: i32,
) -> Result<RecordBatch, Error> {
    let arena = build_specs::<T>();

    // Both start with a null `release`, which the Arrow spec defines as "owns nothing" - so if the
    // call below fails and leaves them untouched, dropping them is a no-op and nothing leaks.
    let mut array = FFI_ArrowArray::empty();
    let mut schema = FFI_ArrowSchema::empty();

    check(unsafe {
        crate::xl_parse_arrow(
            workbook.handle(),
            arena.specs.as_ptr(),
            arena.specs.len() as i32,
            header_row,
            &mut array as *mut FFI_ArrowArray as *mut c_void,
            &mut schema as *mut FFI_ArrowSchema as *mut c_void,
        )
    })?;

    // from_ffi consumes `array` by value: arrow-rs now owns it and will invoke its release callback
    // when the resulting ArrayData is dropped. `schema` stays owned here and releases on drop at
    // the end of this function, which is correct - the two are released independently.
    let data = unsafe { from_ffi(array, &schema) }.map_err(|e| {
        Error::from_status(
            XL_ERROR,
            format!("importing the native Arrow array failed: {e}"),
        )
    })?;

    Ok(RecordBatch::from(StructArray::from(data)))
}

/// Batched counterpart to [`parse_arrow`]: the same schema-driven read, one [`RecordBatch`] at a
/// time, so peak memory is one batch rather than one sheet.
///
/// The lifetime is the whole reason this wraps arrow-rs's reader instead of returning it directly:
/// `ArrowArrayStreamReader` is `'static`, and handing one back would drop the compile-time proof
/// that the workbook is untouched while the stream lives.
///
/// Unlike [`crate::workbook::TypedChunks`], this does NOT fuse after an error: iteration follows
/// arrow-rs's semantics, since the point of this type is to be a drop-in `RecordBatchReader`. The
/// ABI latches a failure, so a `for` loop that ignores the error and keeps pulling can spin on it —
/// break on the first `Err`.
pub struct ArrowChunks<'a> {
    inner: ArrowArrayStreamReader,
    _workbook: PhantomData<&'a mut Workbook>,
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

/// `parse_arrow` delivered a batch at a time. `header_row` means what it does there (0 = no
/// header); `batch_size` is rows per batch, 0 is unbounded and negative is an error.
///
/// Takes `&mut Workbook` for the same reason [`crate::workbook::Workbook::typed_chunks`] does: the
/// stream borrows the workbook's single row cursor, and the ABI serves one chunked read per
/// workbook.
pub fn parse_arrow_stream<T: ExcelMapper>(
    workbook: &mut Workbook,
    header_row: i32,
    batch_size: i64,
) -> Result<ArrowChunks<'_>, Error> {
    let arena = build_specs::<T>();

    // Starts with a null `release` - "owns nothing" per the Arrow spec - so an early failure leaves
    // nothing to clean up.
    let mut stream = FFI_ArrowArrayStream::empty();

    check(unsafe {
        crate::xl_parse_arrow_stream(
            workbook.handle(),
            arena.specs.as_ptr(),
            arena.specs.len() as i32,
            header_row,
            batch_size,
            &mut stream as *mut FFI_ArrowArrayStream as *mut c_void,
        )
    })?;

    // try_new takes the struct by value, so ownership of the release callback moves exactly once -
    // no path where both this function and arrow-rs believe they own it.
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
