use excelreader::ReadOptions;
use excelreader::{ColumnData, ErrorKind, ExcelMapper, TableBatches, TypedBatches, Workbook};

#[derive(Default, ExcelMapper)]
struct Record {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna3")]
    coluna3: i64,
}

fn fixture_path() -> String {
    concat!(env!("CARGO_MANIFEST_DIR"), "/../../RealExcel.xlsb").to_string()
}

fn open() -> Workbook {
    Workbook::open(fixture_path()).expect("open must succeed")
}

fn whole_sheet() -> Vec<(String, i64)> {
    let workbook = open();
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Record>(&ReadOptions::new())
        .expect("parse must succeed");
    (0..table.len())
        .map(|i| {
            let row = table.get(i).expect("row must be in range");
            (row.coluna1, row.coluna3)
        })
        .collect()
}

/// The load-bearing property: chunked output equals whole-sheet output, at every batch size. The
/// sizes that are not multiples of 8 are where a validity-bitmap boundary bug would surface.
#[test]
fn batches_equal_the_whole_sheet() {
    let expected = whole_sheet();
    assert_eq!(expected.len(), 100, "RealExcel.xlsb has 100 data rows");

    for batch_size in [1_i64, 7, 8, 9, 1000, 0, 110] {
        let workbook = open();
        let chunks = workbook
            .sheet(0)
            .expect("sheet 0")
            .parse_batches::<Record>(&ReadOptions::new().batch_size(batch_size))
            .expect("parse_batches must open");

        let mut actual = Vec::new();
        let mut batches = 0_i64;
        for batch in chunks {
            let batch = batch.expect("a batch must read without error");
            batches += 1;
            if batch_size != 0 {
                assert!(
                    batch.len() <= batch_size,
                    "batch_size={batch_size} must be respected"
                );
            }
            for i in 0..batch.len() {
                let row = batch.get(i).expect("row must be in range");
                actual.push((row.coluna1, row.coluna3));
            }
        }

        assert_eq!(actual, expected, "batch_size={batch_size}");

        let rows = expected.len() as i64;
        let want = if batch_size == 0 {
            1
        } else {
            (rows + batch_size - 1) / batch_size
        };
        assert_eq!(batches, want, "batch count for batch_size={batch_size}");
    }
}

#[test]
fn a_negative_batch_size_is_rejected() {
    let workbook = open();
    let error = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse_batches::<Record>(&ReadOptions::new().batch_size(-1))
        .expect_err("a negative batch size must be rejected");
    assert_eq!(error.kind(), ErrorKind::InvalidArgument);
}

/// After the sheet is exhausted the iterator must stay exhausted - a second `next()` must not
/// reopen, repeat the last batch, or spin on a latched error.
#[test]
fn the_iterator_stays_exhausted_after_eof() {
    let workbook = open();
    let mut chunks = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse_batches::<Record>(&ReadOptions::new().batch_size(8))
        .expect("parse_batches must open");
    while chunks.next().is_some() {}
    assert!(chunks.next().is_none());
    assert!(chunks.next().is_none());
}

/// Readers are independent: any number may be open on one workbook, and dropping one leaves the
/// others and the workbook usable.
#[test]
fn two_readers_on_one_workbook_each_yield_every_row() {
    let workbook = open();
    let sheet = workbook.sheet(0).expect("sheet 0");
    let one = sheet
        .parse_batches::<Record>(&ReadOptions::new().batch_size(8))
        .expect("first reader must open");
    let two = sheet
        .parse_batches::<Record>(&ReadOptions::new().batch_size(8))
        .expect("a second reader must open alongside it");

    let count = |chunks: TypedBatches<'_, Record>| -> i64 {
        chunks
            .map(|batch| batch.expect("a batch must read").len())
            .sum()
    };
    assert_eq!(count(one), 100);
    assert_eq!(count(two), 100);
}

fn schema_batches(workbook: &Workbook, batch_size: i64) -> TableBatches<'_> {
    let sheet = workbook.sheet(0).expect("sheet 0");
    let schema = sheet
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");
    sheet
        .parse_batches_with_schema(&schema, &ReadOptions::new().batch_size(batch_size))
        .expect("the batches must open")
}

#[test]
fn schema_batches_equal_the_whole_sheet() {
    let expected: Vec<i64> = whole_sheet().into_iter().map(|(_, n)| n).collect();
    for batch_size in [1, 7, 8, 9, 1000] {
        let workbook = open();
        let mut coluna3 = Vec::new();
        for batch in schema_batches(&workbook, batch_size) {
            let batch = batch.expect("a batch must read");
            assert!(batch.len() <= batch_size);
            let column = batch
                .columns()
                .into_iter()
                .find(|c| c.name == Some("Coluna3"))
                .expect("every batch keeps the schema's names");
            let ColumnData::I64(values) = column.data else {
                panic!("Coluna3 must stay an i64 column");
            };
            coluna3.extend_from_slice(values);
        }
        assert_eq!(coluna3, expected, "batch size {batch_size}");
    }
}

#[test]
fn schema_batches_with_batch_size_zero_yield_one_batch() {
    let workbook = open();
    let batches: Vec<_> = schema_batches(&workbook, 0)
        .map(|b| b.expect("a batch must read"))
        .collect();
    assert_eq!(batches.len(), 1);
    assert_eq!(batches[0].len(), 100);
}

#[test]
fn schema_batches_reject_a_negative_batch_size() {
    let workbook = open();
    let sheet = workbook.sheet(0).expect("sheet 0");
    let schema = sheet
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");
    assert!(sheet
        .parse_batches_with_schema(&schema, &ReadOptions::new().batch_size(-1))
        .is_err());
}
