//! Integration coverage for `Sheet::aggregate` over `xl_aggregate`.

use excelreader::{
    Accumulator, ErrorKind, Format, OpenOptions, RowRef, Workbook, WriteOptions, WriterHandle,
};
use std::panic::{catch_unwind, AssertUnwindSafe};
use std::sync::atomic::{AtomicI64, AtomicUsize, Ordering};
use std::sync::Arc;

fn generated(format: Format, rows: i64) -> Workbook {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(format);
    let mut writer = WriterHandle::new(&mut bytes, &options).expect("open writer");
    writer.start_sheet("S1").expect("start sheet");
    writer.start_row().expect("start row");
    writer.write_str(Some("id")).expect("write");
    writer.write_str(Some("label")).expect("write");
    writer.end_row().expect("end row");
    for i in 1..=rows {
        writer.start_row().expect("start row");
        writer.write_i64(Some(i)).expect("write");
        writer
            .write_str(Some(&format!("filler-value-{i}")))
            .expect("write");
        writer.end_row().expect("end row");
    }
    writer.end_sheet().expect("end sheet");
    writer.finish().expect("finish");
    OpenOptions::new()
        .format(format)
        .open_bytes(&bytes)
        .expect("open")
}

fn first_number(row: &RowRef<'_>) -> Result<i64, i32> {
    Ok(row.get(0).ok_or(1)?.number().ok_or(1)? as i64)
}

#[derive(Default)]
struct Totals {
    sum: i64,
    rows: i64,
    combines: u64,
}

impl Accumulator for Totals {
    fn accumulate(&mut self, row: RowRef<'_>) -> Result<(), i32> {
        self.sum += first_number(&row)?;
        self.rows += 1;
        Ok(())
    }

    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.sum += other.sum;
        self.rows += other.rows;
        self.combines += other.combines + 1;
        Ok(())
    }
}

#[test]
#[cfg_attr(
    target_os = "macos",
    ignore = "crashes the native runtime on macOS arm64; see \"Parallel aggregation on macOS\" in the crate README"
)]
fn aggregate_sums_every_data_row_once_on_every_format() {
    for format in [Format::Xlsx, Format::Xlsb, Format::Xls, Format::Csv] {
        let workbook = generated(format, 60_000);
        let totals = workbook
            .sheet(0)
            .unwrap()
            .aggregate(1, 4, Totals::default)
            .unwrap();
        assert_eq!(totals.rows, 60_000, "{format:?}");
        assert_eq!(totals.sum, 60_000 * 60_001 / 2, "{format:?}");
        if matches!(format, Format::Xlsx | Format::Xlsb) {
            assert!(totals.combines >= 1, "{format:?} must partition");
        }
    }
}

struct Live {
    live: Arc<AtomicUsize>,
}

impl Drop for Live {
    fn drop(&mut self) {
        self.live.fetch_sub(1, Ordering::SeqCst);
    }
}

impl Accumulator for Live {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        Ok(())
    }

    fn combine(&mut self, _other: &mut Self) -> Result<(), i32> {
        Ok(())
    }
}

#[test]
#[cfg_attr(
    target_os = "macos",
    ignore = "crashes the native runtime on macOS arm64; see \"Parallel aggregation on macOS\" in the crate README"
)]
fn every_seeded_accumulator_is_dropped_once() {
    let workbook = generated(Format::Xlsx, 60_000);
    let live = Arc::new(AtomicUsize::new(0));
    let result = workbook.sheet(0).unwrap().aggregate(1, 4, || {
        live.fetch_add(1, Ordering::SeqCst);
        Live { live: live.clone() }
    });
    drop(result.unwrap());
    assert_eq!(live.load(Ordering::SeqCst), 0);
}

struct Aborting<'a> {
    seen: &'a AtomicI64,
}

impl Accumulator for Aborting<'_> {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        if self.seen.fetch_add(1, Ordering::SeqCst) >= 20_000 {
            return Err(9);
        }
        Ok(())
    }

    fn combine(&mut self, _other: &mut Self) -> Result<(), i32> {
        Ok(())
    }
}

#[test]
#[cfg_attr(
    target_os = "macos",
    ignore = "crashes the native runtime on macOS arm64; see \"Parallel aggregation on macOS\" in the crate README"
)]
fn a_callback_code_aborts_with_callback_kind() {
    let workbook = generated(Format::Xlsx, 60_000);
    let seen = AtomicI64::new(0);
    let result = workbook
        .sheet(0)
        .unwrap()
        .aggregate(1, 4, || Aborting { seen: &seen });
    assert_eq!(result.err().unwrap().kind(), ErrorKind::Callback(9));
}

#[derive(Debug, Default)]
struct Panicking {
    rows: i64,
}

impl Accumulator for Panicking {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        self.rows += 1;
        if self.rows == 100 {
            panic!("boom");
        }
        Ok(())
    }

    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.rows += other.rows;
        Ok(())
    }
}

#[test]
#[cfg_attr(
    target_os = "macos",
    ignore = "crashes the native runtime on macOS arm64; see \"Parallel aggregation on macOS\" in the crate README"
)]
fn a_panic_in_accumulate_is_resumed() {
    let workbook = generated(Format::Xlsb, 60_000);
    let sheet = workbook.sheet(0).unwrap();
    let payload = catch_unwind(AssertUnwindSafe(|| {
        sheet.aggregate(1, 4, Panicking::default)
    }))
    .expect_err("the panic must be resumed");
    assert_eq!(payload.downcast_ref::<&str>(), Some(&"boom"));
}

#[derive(Default)]
struct Count(i64);

impl Accumulator for Count {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        self.0 += 1;
        Ok(())
    }

    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.0 += other.0;
        Ok(())
    }
}

#[test]
#[cfg_attr(
    target_os = "macos",
    ignore = "crashes the native runtime on macOS arm64; see \"Parallel aggregation on macOS\" in the crate README"
)]
fn header_row_zero_aggregates_the_header_too() {
    let workbook = generated(Format::Csv, 10);
    let count = workbook
        .sheet(0)
        .unwrap()
        .aggregate(0, 1, Count::default)
        .unwrap();
    assert_eq!(count.0, 11);
}
