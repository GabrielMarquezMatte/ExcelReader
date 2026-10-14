//! Parallel CSV aggregation over the C ABI's `xl_csv_aggregate_*` functions.
//!
//! The library partitions the source, seeds one accumulator per partition, calls
//! [`CsvAccumulator::accumulate`] for every record on worker threads, then folds the partitions
//! together with [`CsvAccumulator::combine`]. This module owns the four `extern "C"` shims that
//! bridge that contract to safe Rust: it boxes each seeded accumulator, catches any panic before it
//! can unwind into the library, and resumes that panic on the calling thread once the run has
//! finished freeing every state.

use crate::ffi::{self, XlCsvAggregation, XlRow, XL_OK};
use crate::workbook::{self, len_i32, path_str};
use crate::{Error, ErrorKind, OpenOptions, RowRef};
use std::any::Any;
use std::ffi::c_void;
use std::panic::{catch_unwind, resume_unwind, AssertUnwindSafe};
use std::path::{Path, PathBuf};
use std::sync::Mutex;

/// The status a shim returns to signal "a callback panicked". The library treats it like any other
/// caller code: from `accumulate` it may be discarded along with the partition instead of ending the
/// run, while from `seed` or `combine` it always ends the run, freeing every state and returning the
/// status verbatim. Either way, the wrapper then resumes the stored payload instead of turning it
/// into an [`Error`]. Positive, so it can never collide with a library status (every one of those is
/// `<= 0`) - but it CAN collide with a caller's own `Err(i32::MAX)`. `finish` does not disambiguate
/// the two by status; it keys on whether the panic slot actually holds a payload, so an ordinary
/// `Err(i32::MAX)` from `accumulate` or `combine` still surfaces as a normal [`Error`] rather than
/// being resumed as a panic.
const PANIC_STATUS: i32 = i32::MAX;

/// What each seeded accumulator is boxed in. The trailing byte keeps the allocation non-zero-sized,
/// so `seed` hands the library a distinct pointer per partition even when `A` is zero-sized, as the
/// header requires.
#[repr(C)]
struct State<A> {
    value: A,
    _nonzero: u8,
}

/// A fold over CSV records, one instance per partition.
///
/// `combine` takes `&mut Self` rather than `Self` because the library, not the implementation, owns
/// `other`: it frees that state through `free_state` once `combine` returns. Drain `other`, do not
/// consume it.
///
/// `accumulate` runs concurrently on worker threads, each on its own instance, which is why `Send`
/// is required and `Sync` is not. A nonzero code from `accumulate` fails only that partition: if the
/// library later re-reads the partition, the failure is discarded with it and the run continues;
/// otherwise the run ends with that code. A nonzero code from `combine`, or a panic from the seed
/// closure, always ends the run. However it ends, the code is returned to the caller verbatim as
/// [`ErrorKind::Callback`]; use positive codes, since every code the library itself returns is
/// `<= 0`.
/// Sibling workers are never stopped, so both `accumulate` and `combine` must tolerate being called
/// again after returning an error.
pub trait CsvAccumulator: Send {
    /// Folds one record in. The row's cell bytes are only valid for the duration of this call.
    fn accumulate(&mut self, row: RowRef<'_>) -> Result<(), i32>;

    /// Folds `other`'s contents into `self`, leaving `other` for the library to free.
    fn combine(&mut self, other: &mut Self) -> Result<(), i32>;
}

/// What [`aggregate_csv`] reads: a CSV file, or CSV bytes already in memory. A `&str` converts to a
/// path, never to CSV text - pass `text.as_bytes()` for that.
#[derive(Clone, Copy, Debug)]
pub enum CsvInput<'a> {
    Path(&'a Path),
    Bytes(&'a [u8]),
}

impl<'a> From<&'a Path> for CsvInput<'a> {
    fn from(path: &'a Path) -> Self {
        CsvInput::Path(path)
    }
}

impl<'a> From<&'a PathBuf> for CsvInput<'a> {
    fn from(path: &'a PathBuf) -> Self {
        CsvInput::Path(path)
    }
}

impl<'a> From<&'a str> for CsvInput<'a> {
    fn from(path: &'a str) -> Self {
        CsvInput::Path(Path::new(path))
    }
}

impl<'a> From<&'a [u8]> for CsvInput<'a> {
    fn from(bytes: &'a [u8]) -> Self {
        CsvInput::Bytes(bytes)
    }
}

impl<'a, const N: usize> From<&'a [u8; N]> for CsvInput<'a> {
    fn from(bytes: &'a [u8; N]) -> Self {
        CsvInput::Bytes(bytes)
    }
}

/// What every shim reaches through `user_data`. The C ABI shares one `user_data` across all worker
/// threads and requires it to be read-only or internally synchronized, so the seed closure is held
/// by shared reference and the panic slot behind a `Mutex` - the shims never take `&mut Shared`.
struct Shared<'a, A, F> {
    seed: &'a F,
    panic: Mutex<Option<Box<dyn Any + Send>>>,
    _marker: std::marker::PhantomData<fn() -> A>,
}

impl<A, F> Shared<'_, A, F> {
    /// Keeps the FIRST payload: later panics from sibling workers are dropped rather than racing to
    /// overwrite the one that started the abort.
    fn store_panic(&self, payload: Box<dyn Any + Send>) -> i32 {
        let mut slot = self
            .panic
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if slot.is_none() {
            *slot = Some(payload);
        }
        PANIC_STATUS
    }
}

/// # Safety
/// `user_data` is the `Shared` this crate handed to `xl_csv_aggregate_*`, alive for the whole call.
unsafe extern "C" fn seed_shim<A, F>(out_state: *mut *mut c_void, user_data: *mut c_void) -> i32
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    let shared = unsafe { &*(user_data as *const Shared<'_, A, F>) };
    unsafe { *out_state = std::ptr::null_mut() };
    match catch_unwind(AssertUnwindSafe(|| (shared.seed)())) {
        Ok(accumulator) => {
            unsafe {
                *out_state = Box::into_raw(Box::new(State {
                    value: accumulator,
                    _nonzero: 0,
                })) as *mut c_void
            };
            XL_OK
        }
        Err(payload) => shared.store_panic(payload),
    }
}

/// # Safety
/// `state` is a state a previous `seed_shim` boxed, and `row`'s cells are valid for this call only.
unsafe extern "C" fn accumulate_shim<A, F>(
    state: *mut c_void,
    row: *const XlRow,
    user_data: *mut c_void,
) -> i32
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    let shared = unsafe { &*(user_data as *const Shared<'_, A, F>) };
    let accumulator = unsafe { &mut (*(state as *mut State<A>)).value };
    let row = unsafe { &*row };
    match catch_unwind(AssertUnwindSafe(|| {
        accumulator.accumulate(unsafe { RowRef::from_decoded(row.cells, row.cell_count) })
    })) {
        Ok(Ok(())) => XL_OK,
        Ok(Err(code)) => code,
        Err(payload) => shared.store_panic(payload),
    }
}

/// # Safety
/// `acc` and `next` are two distinct states boxed by `seed_shim`; `next` stays the library's to free.
unsafe extern "C" fn combine_shim<A, F>(
    acc: *mut c_void,
    next: *mut c_void,
    user_data: *mut c_void,
) -> i32
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    let shared = unsafe { &*(user_data as *const Shared<'_, A, F>) };
    let accumulator = unsafe { &mut (*(acc as *mut State<A>)).value };
    let other = unsafe { &mut (*(next as *mut State<A>)).value };
    match catch_unwind(AssertUnwindSafe(|| accumulator.combine(other))) {
        Ok(Ok(())) => XL_OK,
        Ok(Err(code)) => code,
        Err(payload) => shared.store_panic(payload),
    }
}

/// # Safety
/// `state` is a state boxed by `seed_shim` that the library is done with. Never called with NULL.
/// `user_data` is the `Shared` this crate handed to `xl_csv_aggregate_*`, alive for the whole call.
unsafe extern "C" fn free_state_shim<A, F>(state: *mut c_void, user_data: *mut c_void)
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    let shared = unsafe { &*(user_data as *const Shared<'_, A, F>) };
    if let Err(payload) = catch_unwind(AssertUnwindSafe(|| {
        drop(unsafe { Box::from_raw(state as *mut State<A>) })
    })) {
        shared.store_panic(payload);
    }
}

fn raw_aggregation<A, F>(shared: &Shared<'_, A, F>) -> XlCsvAggregation
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    XlCsvAggregation {
        struct_size: std::mem::size_of::<XlCsvAggregation>() as i32,
        seed: Some(seed_shim::<A, F>),
        accumulate: Some(accumulate_shim::<A, F>),
        combine: Some(combine_shim::<A, F>),
        free_state: Some(free_state_shim::<A, F>),
        user_data: shared as *const Shared<'_, A, F> as *mut c_void,
    }
}

fn shared<A, F>(seed: &F) -> Shared<'_, A, F> {
    Shared {
        seed,
        panic: Mutex::new(None),
        _marker: std::marker::PhantomData,
    }
}

/// Folds a CSV file or buffer into one `A` across several threads.
///
/// The CSV dialect (`csv_delimiter`, `csv_quote`, `csv_detect_bom`) and `max_cell_bytes` come from
/// `options`; its other settings do not apply. `header_row` is 1-based, `0` meaning no header.
/// `degree_of_parallelism` is `0` for every processor, `1` sequential, `n` up to `n` threads.
/// A buffer is not copied; the borrow keeps it valid for the call.
///
/// `seed` is called at least once per partition, on worker threads, and may be called again for a
/// partition that has to be re-read - a seeded accumulator can therefore be dropped without ever
/// being combined. `seed`, `accumulate` and `combine` never run on the calling thread; each
/// accumulator is dropped on the calling thread once the run has finished.
///
/// The call blocks until the run finishes, so `seed` may borrow locals freely: there is no `'static`
/// bound and none is needed. A panic inside any callback aborts the run and is resumed here, on the
/// calling thread, once the library has freed every state.
pub fn aggregate_csv<'a, A, F>(
    input: impl Into<CsvInput<'a>>,
    options: &OpenOptions,
    header_row: i32,
    degree_of_parallelism: i32,
    seed: F,
) -> Result<A, Error>
where
    A: CsvAccumulator,
    F: Fn() -> A + Sync,
{
    workbook::check_abi_version()?;
    aggregate_unchecked(
        input.into(),
        options,
        header_row,
        degree_of_parallelism,
        seed,
    )
}

/// [`aggregate_csv`] without the ABI check, for the warm-up run in
/// [`workbook::check_abi_version`] itself - calling the public entry point from there would
/// re-enter its `OnceLock` and deadlock.
pub(crate) fn aggregate_unchecked<A, F>(
    input: CsvInput<'_>,
    options: &OpenOptions,
    header_row: i32,
    degree_of_parallelism: i32,
    seed: F,
) -> Result<A, Error>
where
    A: CsvAccumulator,
    F: Fn() -> A + Sync,
{
    let raw_options = options.to_csv_parallel_raw(header_row, degree_of_parallelism);
    let shared = shared::<A, F>(&seed);
    let agg = raw_aggregation(&shared);
    let mut state: *mut c_void = std::ptr::null_mut();

    let status = match input {
        CsvInput::Path(path) => {
            let path = path_str(path)?;
            let path_len = len_i32(path.len(), "path")?;
            unsafe {
                ffi::xl_csv_aggregate_file(path.as_ptr(), path_len, &agg, &raw_options, &mut state)
            }
        }
        CsvInput::Bytes(data) => {
            let data_len = len_i32(data.len(), "buffer")?;
            unsafe {
                ffi::xl_csv_aggregate_memory(
                    data.as_ptr(),
                    data_len,
                    &agg,
                    &raw_options,
                    &mut state,
                )
            }
        }
    };
    finish(shared, status, state)
}

/// Turns the run's outcome into a `Result`, taking ownership of the surviving state.
///
/// A stored panic only wins over the status if the status itself is [`PANIC_STATUS`]: an
/// `accumulate` panic on a chunk whose boundary guess was wrong is discarded exactly like a
/// non-panic error would be once the library re-reads that chunk from the confirmed start, so the
/// panic payload can be stale by the time this runs even though `shared.panic` still holds it. Only
/// resume it when the library's own verdict says the run actually ended on a panic; otherwise a
/// discarded panic from an abandoned first attempt would wrongly kill an otherwise-successful run
/// (or shadow the real error from an unrelated later failure). When the status does say the run
/// ended on a panic, the library has already unwound and freed every state it owns by the time it
/// returns, so resuming here delivers the original payload to the caller without it ever having
/// crossed a foreign frame.
fn finish<A, F>(shared: Shared<'_, A, F>, status: i32, state: *mut c_void) -> Result<A, Error>
where
    A: CsvAccumulator,
    F: Fn() -> A,
{
    let winner = if state.is_null() {
        None
    } else {
        Some(unsafe { Box::from_raw(state as *mut State<A>) })
    };

    if status == PANIC_STATUS {
        if let Some(payload) = shared
            .panic
            .into_inner()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
        {
            drop(winner);
            resume_unwind(payload);
        }
    }
    if status != XL_OK {
        return Err(status_error(status));
    }
    winner.map(|boxed| boxed.value).ok_or_else(|| {
        Error::new(
            ErrorKind::Other,
            "native reported success but wrote no aggregation state",
        )
    })
}

/// Library statuses (all `<= 0`) carry a `xl_last_error_ptr` message; a positive status is a code one
/// of the caller's own callbacks returned, which the library passes through untouched and has no
/// message for.
fn status_error(status: i32) -> Error {
    if status > 0 {
        Error::new(
            ErrorKind::Callback(status),
            format!("an aggregation callback returned status {status}"),
        )
    } else {
        workbook::last_error(status)
    }
}
