//! Integration coverage for the safe wrapper over `xl_csv_aggregate_file`/`xl_csv_aggregate_memory`.

use excelreader::{aggregate_csv_file, aggregate_csv_memory, CsvAccumulator, CsvParallelOptions, RowRef};
use std::collections::HashMap;

/// Group-by-sum that also counts how many times the library folded two partitions together, so a
/// test can tell a genuinely partitioned run from a silently sequential one.
#[derive(Default)]
struct GroupSums {
    totals: HashMap<Vec<u8>, i64>,
    combines: u64,
}

impl CsvAccumulator for GroupSums {
    fn accumulate(&mut self, row: RowRef<'_>) -> Result<(), i32> {
        let mut cells = row.iter();
        let key = cells.next().ok_or(1)?.as_bytes().to_vec();
        let amount: i64 = std::str::from_utf8(cells.next().ok_or(1)?.as_bytes())
            .map_err(|_| 2)?
            .parse()
            .map_err(|_| 3)?;
        *self.totals.entry(key).or_insert(0) += amount;
        Ok(())
    }

    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        for (key, value) in other.totals.drain() {
            *self.totals.entry(key).or_insert(0) += value;
        }
        self.combines += other.combines + 1;
        Ok(())
    }
}

const GROUPS: [&[u8]; 3] = [b"alpha", b"beta", b"gamma"];

/// Roughly 3.9 MB of CSV: past the library's 256 KB partitioning floor and several times its 1 MB
/// chunk size, so a `degree_of_parallelism` above 1 produces more than one partition.
fn group_fixture() -> (String, HashMap<Vec<u8>, i64>) {
    let mut text = String::from("group,amount\n");
    let mut expected: HashMap<Vec<u8>, i64> = HashMap::new();
    for index in 0..300_000i64 {
        let group = GROUPS[(index % 3) as usize];
        text.push_str(std::str::from_utf8(group).unwrap());
        text.push(',');
        text.push_str(&index.to_string());
        text.push('\n');
        *expected.entry(group.to_vec()).or_insert(0) += index;
    }
    (text, expected)
}

struct TempDir(std::path::PathBuf);

impl TempDir {
    fn new(tag: &str) -> TempDir {
        let path = std::env::temp_dir().join(format!("xlagg-{tag}-{}", std::process::id()));
        std::fs::create_dir_all(&path).unwrap();
        TempDir(path)
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        std::fs::remove_dir_all(&self.0).ok();
    }
}

#[test]
fn group_by_sum_over_a_file_is_partitioned_and_exact() {
    let dir = TempDir::new("file");
    let path = dir.0.join("groups.csv");
    let (text, expected) = group_fixture();
    std::fs::write(&path, &text).unwrap();
    assert!(text.len() > 3 * 1024 * 1024, "fixture must clear the 1 MB chunk floor several times over");

    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    let actual = aggregate_csv_file(&path, GroupSums::default, &options).unwrap();

    assert_eq!(actual.totals, expected);
    assert!(
        actual.combines >= 1,
        "a {}-byte source at degree_of_parallelism=8 must split into more than one partition, so \
         combine must run at least once (observed {})",
        text.len(),
        actual.combines
    );
}

#[test]
fn group_by_sum_over_memory_is_partitioned_and_exact() {
    let (text, expected) = group_fixture();
    let options = CsvParallelOptions {
        degree_of_parallelism: 4,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    let actual = aggregate_csv_memory(text.as_bytes(), GroupSums::default, &options).unwrap();

    assert_eq!(actual.totals, expected);
    assert!(actual.combines >= 1, "observed {} combines, expected a partitioned run", actual.combines);
}

#[test]
fn a_failing_accumulate_returns_its_own_code() {
    #[derive(Debug)]
    struct AlwaysFails;
    impl CsvAccumulator for AlwaysFails {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            Err(42)
        }
        fn combine(&mut self, _other: &mut Self) -> Result<(), i32> {
            Ok(())
        }
    }

    let dir = TempDir::new("fail");
    let path = dir.0.join("fail.csv");
    let (text, _) = group_fixture();
    std::fs::write(&path, &text).unwrap();

    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        ..CsvParallelOptions::default()
    };
    let error = aggregate_csv_file(&path, || AlwaysFails, &options).unwrap_err();

    assert_eq!(error.code(), 42);
}

/// A zero-sized accumulator run for row-exactness: every row must still be counted exactly once
/// across however many partitions and re-reads the library uses. The distinct-pointer rule this
/// case can violate is asserted separately, in
/// `a_zero_sized_accumulator_never_aliases_two_states_in_combine`.
#[test]
fn a_zero_sized_accumulator_gets_a_distinct_state_per_partition() {
    use std::sync::atomic::{AtomicI64, Ordering};
    static ROWS: AtomicI64 = AtomicI64::new(0);

    struct Counter;
    impl CsvAccumulator for Counter {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            ROWS.fetch_add(1, Ordering::Relaxed);
            Ok(())
        }
        fn combine(&mut self, _other: &mut Self) -> Result<(), i32> {
            Ok(())
        }
    }

    let (text, _) = group_fixture();
    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    aggregate_csv_memory(text.as_bytes(), || Counter, &options).unwrap();

    assert_eq!(ROWS.load(Ordering::Relaxed), 300_000);
}

/// The header requires `seed` to return a distinct pointer on every call, and hands `combine` "two
/// distinct states". A zero-sized accumulator is the case that breaks it: `Box::into_raw` of a ZST
/// returns the same dangling address every time, so `acc` and `next` are one and the same object.
#[test]
fn a_zero_sized_accumulator_never_aliases_two_states_in_combine() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static COMBINES: AtomicUsize = AtomicUsize::new(0);
    static ALIASED: AtomicUsize = AtomicUsize::new(0);

    struct Counter;
    impl CsvAccumulator for Counter {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            COMBINES.fetch_add(1, Ordering::Relaxed);
            if std::ptr::eq(self as *const Self, other as *const Self) {
                ALIASED.fetch_add(1, Ordering::Relaxed);
            }
            Ok(())
        }
    }

    let (text, _) = group_fixture();
    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    aggregate_csv_memory(text.as_bytes(), || Counter, &options).unwrap();

    let combines = COMBINES.load(Ordering::Relaxed);
    let aliased = ALIASED.load(Ordering::Relaxed);
    assert!(combines >= 1, "the fixture must partition, so combine must run at least once");
    assert_eq!(aliased, 0, "{combines} combine calls, of which {aliased} folded a state into itself");
}

/// Every accumulator the library seeds must be dropped exactly once.
#[test]
fn every_seeded_accumulator_is_dropped_exactly_once() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static SEEDED: AtomicUsize = AtomicUsize::new(0);
    static DROPPED: AtomicUsize = AtomicUsize::new(0);

    struct Tracked(i64);
    impl Drop for Tracked {
        fn drop(&mut self) {
            DROPPED.fetch_add(1, Ordering::Relaxed);
        }
    }
    impl CsvAccumulator for Tracked {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            self.0 += 1;
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            self.0 += other.0;
            Ok(())
        }
    }

    let (text, _) = group_fixture();
    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    let result = aggregate_csv_memory(
        text.as_bytes(),
        || {
            SEEDED.fetch_add(1, Ordering::Relaxed);
            Tracked(0)
        },
        &options,
    )
    .unwrap();

    assert_eq!(result.0, 300_000);
    drop(result);
    assert_eq!(SEEDED.load(Ordering::Relaxed), DROPPED.load(Ordering::Relaxed));
}

/// Quoted records that span lines are what make a guessed partition boundary wrong, so the run must
/// still count each record exactly once.
#[test]
fn quoted_multi_line_records_straddling_chunks_are_counted_once() {
    struct Rows(i64);
    impl CsvAccumulator for Rows {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            self.0 += 1;
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            self.0 += other.0;
            Ok(())
        }
    }

    let mut text = String::new();
    for _ in 0..8 {
        text.push_str("1,\"");
        for _ in 0..40_000 {
            text.push_str("BOOM,notint,x\n");
        }
        text.push_str("\"\n");
    }

    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        ..CsvParallelOptions::default()
    };
    let result = aggregate_csv_memory(text.as_bytes(), || Rows(0), &options).unwrap();

    assert_eq!(result.0, 8);
}

/// A panic in `accumulate` runs on a library worker thread; it must be caught there, abort the run,
/// and be resumed intact on the calling thread rather than unwinding through the foreign frames.
#[test]
#[should_panic(expected = "accumulate blew up")]
fn a_panicking_accumulate_is_resumed_on_the_calling_thread() {
    struct Explodes;
    impl CsvAccumulator for Explodes {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            panic!("accumulate blew up");
        }
        fn combine(&mut self, _other: &mut Self) -> Result<(), i32> {
            Ok(())
        }
    }

    let _ = aggregate_csv_memory(b"a,b\n1,2\n", || Explodes, &CsvParallelOptions::default());
}

struct RowCounter(i64);

impl CsvAccumulator for RowCounter {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        self.0 += 1;
        Ok(())
    }
    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.0 += other.0;
        Ok(())
    }
}

/// Coverage for a macOS/arm64-only SIGSEGV: NativeAOT's automatic CPU-count detection leaves
/// runtime state that crashes the first native call to follow a concurrent
/// (`degree_of_parallelism` > 1) run. Without `DOTNET_PROCESSOR_COUNT` in the process environment
/// at startup this crashes there (setting it from Rust is too late - the native library is linked,
/// so its runtime initializes at load). CI exports it for the macOS job; see the crate README.
#[test]
fn a_call_after_a_partitioned_run_still_succeeds() {
    let (text, _) = group_fixture();
    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        header_row: 1,
        ..CsvParallelOptions::default()
    };
    let first = aggregate_csv_memory(text.as_bytes(), || RowCounter(0), &options).unwrap();
    assert_eq!(first.0, 300_000);

    let second = aggregate_csv_memory(b"a,b\n1,2\n", || RowCounter(0), &CsvParallelOptions::default()).unwrap();
    assert_eq!(second.0, 2);
}

/// A chunk whose initial boundary guess lands inside a giant quoted field sees garbled interior
/// "rows" there and panics on them. The library re-reads that chunk from the correct boundary once
/// the true start is confirmed, and the re-read succeeds without panicking - the stale panic from
/// the discarded first attempt must not resume and kill an otherwise-successful run.
#[test]
fn a_panic_from_a_discarded_reread_does_not_resume() {
    struct PanicsOnGarbledRows(i64);
    impl CsvAccumulator for PanicsOnGarbledRows {
        fn accumulate(&mut self, row: RowRef<'_>) -> Result<(), i32> {
            let mut cells = row.iter();
            let first = cells.next().ok_or(1)?.as_bytes();
            if first != b"1" {
                panic!("saw a garbled row starting with {:?}", String::from_utf8_lossy(first));
            }
            self.0 += 1;
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            self.0 += other.0;
            Ok(())
        }
    }

    let mut text = String::new();
    for _ in 0..8 {
        text.push_str("1,\"");
        for _ in 0..40_000 {
            text.push_str("BOOM,notint,x\n");
        }
        text.push_str("\"\n");
    }

    let options = CsvParallelOptions {
        degree_of_parallelism: 8,
        ..CsvParallelOptions::default()
    };
    let result = aggregate_csv_memory(text.as_bytes(), || PanicsOnGarbledRows(0), &options).unwrap();

    assert_eq!(result.0, 8);
}

/// The seed closure borrows a local, which only compiles if the wrapper's signature carries no
/// `'static` bound.
#[test]
fn a_seed_closure_may_borrow_a_local() {
    struct CountRows(i64);
    impl CsvAccumulator for CountRows {
        fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
            self.0 += 1;
            Ok(())
        }
        fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
            self.0 += other.0;
            Ok(())
        }
    }

    let start = 100i64;
    let result =
        aggregate_csv_memory(b"a,b\n1,2\n3,4\n", || CountRows(start), &CsvParallelOptions::default())
            .unwrap();
    assert_eq!(result.0, start + 3);
}
