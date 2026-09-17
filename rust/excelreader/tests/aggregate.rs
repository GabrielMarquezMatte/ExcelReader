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
    std::fs::write(&path, b"1,a\n2,b\n").unwrap();

    let error =
        aggregate_csv_file(&path, || AlwaysFails, &CsvParallelOptions::default()).unwrap_err();

    assert_eq!(error.code(), 42);
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
