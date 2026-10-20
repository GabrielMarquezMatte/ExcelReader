//! Push-style row reads over `xl_rows_foreach`.
//!
//! The library calls back on the calling thread, so the closure needs no `Send` bound and the
//! panic slot no lock. A panic is caught before it can unwind into the library and resumed once
//! `xl_rows_foreach` has returned.

use crate::aggregate::{status_error, PANIC_STATUS};
use crate::ffi::{self, XlCellView, XlWorkbook, XL_OK};
use crate::{Error, RowRef};
use std::any::Any;
use std::ffi::c_void;
use std::panic::{catch_unwind, resume_unwind, AssertUnwindSafe};

struct Walk<'f, F> {
    f: &'f mut F,
    panic: Option<Box<dyn Any + Send>>,
}

/// # Safety
/// `user_data` is the `Walk` handed to `xl_rows_foreach`, and `cells` is valid for this call only.
unsafe extern "C" fn row_shim<F>(cells: *const XlCellView, count: i32, user_data: *mut c_void) -> i32
where
    F: FnMut(RowRef<'_>) -> Result<(), i32>,
{
    let walk = unsafe { &mut *(user_data as *mut Walk<'_, F>) };
    match catch_unwind(AssertUnwindSafe(|| {
        (walk.f)(unsafe { RowRef::from_views(cells, count) })
    })) {
        Ok(Ok(())) => XL_OK,
        Ok(Err(code)) => code,
        Err(payload) => {
            walk.panic = Some(payload);
            PANIC_STATUS
        }
    }
}

pub(crate) fn for_each_row<F>(handle: *mut XlWorkbook, sheet: i32, mut f: F) -> Result<(), Error>
where
    F: FnMut(RowRef<'_>) -> Result<(), i32>,
{
    crate::workbook::check_abi_version()?;
    let mut walk = Walk {
        f: &mut f,
        panic: None,
    };
    let status = unsafe {
        ffi::xl_rows_foreach(
            handle,
            sheet,
            Some(row_shim::<F>),
            &mut walk as *mut Walk<'_, F> as *mut c_void,
        )
    };
    if status == PANIC_STATUS {
        if let Some(payload) = walk.panic.take() {
            resume_unwind(payload);
        }
    }
    if status != XL_OK {
        return Err(status_error(status));
    }
    Ok(())
}
