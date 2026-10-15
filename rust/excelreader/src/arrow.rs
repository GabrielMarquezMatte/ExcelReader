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

use crate::ffi;
use crate::workbook::{build_specs, check};
use crate::{Error, ErrorKind, ExcelMapper, ReadOptions, Sheet, Workbook};

impl<'w> Sheet<'w> {
    /// [`parse`](Self::parse) into one Arrow [`RecordBatch`], with the same `#[derive(ExcelMapper)]`
    /// mapping and [`ReadOptions`].
    pub fn to_arrow<T: ExcelMapper>(&self, options: &ReadOptions) -> Result<RecordBatch, Error> {
        let arena = build_specs::<T>();

        let mut array = FFI_ArrowArray::empty();
        let mut schema = FFI_ArrowSchema::empty();

        check(unsafe {
            ffi::xl_parse_arrow(
                self.handle(),
                self.index(),
                arena.specs.as_ptr(),
                arena.specs.len() as i32,
                options.header_row,
                options.parallelism,
                &mut array as *mut FFI_ArrowArray as *mut c_void,
                &mut schema as *mut FFI_ArrowSchema as *mut c_void,
            )
        })?;

        let data = unsafe { from_ffi(array, &schema) }.map_err(|e| {
            Error::new(
                ErrorKind::Other,
                format!("importing the native Arrow array failed: {e}"),
            )
        })?;

        Ok(RecordBatch::from(StructArray::from(data)))
    }

    /// [`to_arrow`](Self::to_arrow) delivered [`ReadOptions`]'s batch size rows at a time; a
    /// negative batch size is an error. Any number of streams may be open on one workbook.
    pub fn to_arrow_stream<T: ExcelMapper>(
        &self,
        options: &ReadOptions,
    ) -> Result<ArrowBatches<'w>, Error> {
        let arena = build_specs::<T>();

        let mut stream = FFI_ArrowArrayStream::empty();

        check(unsafe {
            ffi::xl_parse_arrow_stream(
                self.handle(),
                self.index(),
                arena.specs.as_ptr(),
                arena.specs.len() as i32,
                options.header_row,
                options.batch_size,
                &mut stream as *mut FFI_ArrowArrayStream as *mut c_void,
            )
        })?;

        let inner = ArrowArrayStreamReader::try_new(stream).map_err(|e| {
            Error::new(
                ErrorKind::Other,
                format!("importing the native Arrow stream failed: {e}"),
            )
        })?;

        Ok(ArrowBatches {
            inner,
            _workbook: PhantomData,
        })
    }
}

/// Batches of [`Sheet::to_arrow_stream`], one [`RecordBatch`] at a time.
///
/// Wraps arrow-rs's `'static` reader only to carry the `'a` borrow, which keeps the workbook alive
/// for at least as long as the stream. Unlike [`TypedBatches`](crate::TypedBatches) it does not fuse
/// after an error, so break on the first `Err` rather than spinning on the latched one.
pub struct ArrowBatches<'a> {
    inner: ArrowArrayStreamReader,
    _workbook: PhantomData<&'a Workbook>,
}

impl Iterator for ArrowBatches<'_> {
    type Item = Result<RecordBatch, ArrowError>;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}

impl RecordBatchReader for ArrowBatches<'_> {
    fn schema(&self) -> SchemaRef {
        self.inner.schema()
    }
}
