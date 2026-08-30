//! Arrow C Data Interface import, behind the `arrow` cargo feature.
//!
//! The native side already produces one top-level Arrow struct array whose children are the
//! columns, so the only work here is handing arrow-rs the FFI pair and letting it take ownership.

use std::ffi::c_void;

use arrow::array::{RecordBatch, StructArray};
use arrow::ffi::{from_ffi, FFI_ArrowArray, FFI_ArrowSchema};

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
