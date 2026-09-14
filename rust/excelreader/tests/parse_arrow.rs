#![cfg(feature = "arrow")]

use arrow::array::{Array, Int64Array, StringArray};
use arrow::record_batch::RecordBatchReader;
use excelreader::arrow::{parse_arrow, parse_arrow_stream};
use excelreader::workbook::{ExcelMapper, Workbook};

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

#[test]
fn parse_arrow_returns_a_record_batch_with_one_column_per_field() {
    let mut workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    let batch = parse_arrow::<Record>(&mut workbook, 1).expect("parse_arrow must succeed");

    assert_eq!(batch.num_columns(), 2); 
    assert_eq!(batch.num_rows(), 100); // RealExcel.xlsb has 100 data rows.
    // Field names come from the column spec's source name (the `#[excel(name = "...")]` value),
    // not the Rust struct field identifier - the native side names each Arrow child field from
    // `spec.Names[0]` (see ExcelReader.Native/NativeApi.Arrow.cs's BuildChildSchema).
    assert_eq!(batch.schema().field(0).name(), "Coluna1");
    assert_eq!(batch.schema().field(1).name(), "Coluna3");

    // Downcasting proves the XL_T_* -> Arrow format-code mapping actually landed, not just that a
    // batch of the right shape came back.
    assert!(batch.column(0).as_any().downcast_ref::<StringArray>().is_some());
    assert!(batch.column(1).as_any().downcast_ref::<Int64Array>().is_some());
}

#[test]
fn parse_arrow_reports_an_error_without_leaving_a_half_built_batch() {
    let mut workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    // header_row is 1-based; a row number past the end of the sheet cannot resolve any column name.
    let result = parse_arrow::<Record>(&mut workbook, 1_000_000);
    assert!(result.is_err());

    // The failed call must not have left the workbook or its shared row cursor in a broken state -
    // a normal, valid parse right after the failure should succeed exactly as if the failed call
    // had never happened.
    let batch = parse_arrow::<Record>(&mut workbook, 1).expect("parse_arrow must succeed");
    assert_eq!(batch.num_rows(), 100); // RealExcel.xlsb has 100 data rows.
}

#[test]
fn arrow_stream_batches_match_the_whole_sheet() {
    let mut workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    let stream = parse_arrow_stream::<Record>(&mut workbook, 1, 8).expect("the stream must open");
    let schema = stream.schema();

    let mut rows = 0;
    let mut batches = 0;
    for batch in stream {
        let batch = batch.expect("a batch must read without error");
        assert!(batch.num_rows() <= 8, "batch size must be respected");
        assert_eq!(batch.schema(), schema, "every batch must share the stream schema");
        rows += batch.num_rows();
        batches += 1;
    }

    assert_eq!(rows, 100, "RealExcel.xlsb has 100 data rows");
    assert_eq!(batches, 13, "100 rows at batch size 8 is 13 batches");
}

#[test]
fn arrow_stream_rejects_a_negative_batch_size() {
    let mut workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    assert!(parse_arrow_stream::<Record>(&mut workbook, 1, -1).is_err());
}

#[test]
fn arrow_stream_with_batch_size_zero_yields_one_batch() {
    let mut workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    let stream = parse_arrow_stream::<Record>(&mut workbook, 1, 0).expect("the stream must open");
    let batches: Vec<_> = stream.map(|b| b.expect("a batch must read")).collect();
    assert_eq!(batches.len(), 1, "batch_size 0 is one whole-sheet batch");
    assert_eq!(batches[0].num_rows(), 100);
}
