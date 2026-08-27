use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};
use excelreader::{Date, OpenOptions, XL_FORMAT_XLSB, XL_T_I64, XL_T_STRING};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna3")]
    coluna3: i64,
}

/// Exercises the widths and temporal types the derive gained beyond String/i64/f64/bool: `Coluna3`
/// is a small integer that must survive the `TryFrom` narrowing, and `Coluna2` is a real date.
#[derive(Default, ExcelMapper)]
struct WideRow {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna2")]
    coluna2: Date,
    #[excel(name = "Coluna3")]
    coluna3: u16,
    #[excel(name = "Coluna16")]
    coluna16: f32,
}

#[derive(Default, ExcelMapper)]
struct AliasRow {
    #[excel(name = "ThisColumnDoesNotExist", alias = "Coluna1")]
    coluna1: String,
}

fn fixture_path() -> String {
    concat!(env!("CARGO_MANIFEST_DIR"), "/../../RealExcel.xlsb").to_string()
}

fn open_fixture() -> Workbook {
    Workbook::open(&fixture_path()).expect("open must succeed")
}

#[test]
fn parses_real_excel_fixture() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);

    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1);

    let all: Vec<Row> = table.iter().collect();
    assert_eq!(all.len(), 100);
}

#[test]
fn get_returns_none_outside_the_row_range() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");

    assert!(
        table.get(table.len() - 1).is_some(),
        "last row is in bounds"
    );
    // Before the bounds check these read past the columnar buffers and returned garbage.
    assert!(table.get(table.len()).is_none(), "one past the end");
    assert!(table.get(i64::MAX).is_none(), "far past the end");
    assert!(table.get(-1).is_none(), "negative row");
}

#[test]
fn iter_yields_exactly_len_rows_and_reports_it_up_front() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");

    let iter = table.iter();
    assert_eq!(iter.len(), 100, "ExactSizeIterator must agree with len()");
    assert_eq!(iter.count(), 100);
}

#[test]
fn parses_integer_widths_floats_and_dates() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<WideRow>(&mut workbook, 1).expect("parse_sheet must succeed");

    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1u16);
    assert!((first.coluna16 - 0.1f32).abs() < f32::EPSILON);
    // 2026-01-01 is 20454 days after 1970-01-01.
    assert_eq!(first.coluna2, Date::new(20_454));
}

#[test]
fn resolves_the_first_alias_present_in_the_header_row() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<AliasRow>(&mut workbook, 1).expect("parse_sheet must succeed via alias");
    assert_eq!(table.len(), 100);
    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
}

#[cfg(feature = "chrono")]
#[test]
fn parses_dates_straight_into_chrono() {
    use chrono::NaiveDate;

    #[derive(Default, ExcelMapper)]
    struct ChronoRow {
        #[excel(name = "Coluna2")]
        coluna2: NaiveDate,
    }

    let mut workbook = open_fixture();
    let table = parse_sheet::<ChronoRow>(&mut workbook, 1).expect("parse_sheet must succeed");
    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna2, NaiveDate::from_ymd_opt(2026, 1, 1).unwrap());
}

#[test]
fn exposes_sheet_navigation() {
    let mut workbook = open_fixture();

    let count = workbook.sheet_count().expect("sheet_count must succeed");
    assert!(count >= 1, "the fixture has at least one sheet");

    let names = workbook.sheet_names().expect("sheet_names must succeed");
    assert_eq!(names.len(), count as usize);
    assert_eq!(
        workbook.sheet_name().expect("sheet_name must succeed"),
        names[0],
        "the first sheet is selected before any move_to_sheet"
    );

    workbook
        .move_to_sheet(0)
        .expect("move_to_sheet must succeed");
    assert_eq!(workbook.sheet_name().unwrap(), names[0]);

    // Reading it is enough - which system the fixture uses is not this test's business.
    workbook.is_date1904().expect("is_date1904 must succeed");
}

#[test]
fn move_to_sheet_rejects_an_index_past_the_end() {
    let mut workbook = open_fixture();
    let count = workbook.sheet_count().unwrap();
    let error = workbook
        .move_to_sheet(count)
        .expect_err("an index past the last sheet must fail");
    assert!(
        !error.message().is_empty(),
        "the failure must carry the native detail, got: {error}"
    );
}

#[test]
fn infers_a_schema_from_the_header_row() {
    let workbook = open_fixture();
    let schema = workbook
        .infer_schema(1, 100)
        .expect("infer_schema must succeed");

    assert!(!schema.is_empty(), "the fixture has columns to infer");
    let first = &schema[0];
    assert_eq!(first.name.as_deref(), Some("Coluna1"));
    assert_eq!(first.column_type, XL_T_STRING);

    let coluna3 = schema
        .iter()
        .find(|c| c.name.as_deref() == Some("Coluna3"))
        .expect("Coluna3 must be inferred");
    assert_eq!(coluna3.column_type, XL_T_I64);
}

#[test]
fn infer_schema_leaves_the_row_cursor_alone() {
    let mut workbook = open_fixture();
    workbook
        .infer_schema(1, 100)
        .expect("infer_schema must succeed");

    // The ABI documents infer_schema as sampling independently of the shared cursor, so a parse
    // afterwards must still see every row.
    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_with_accepts_an_explicit_format_and_options() {
    let options = OpenOptions::new().prefetch_decompression(true);
    let mut workbook = Workbook::open_with(&fixture_path(), XL_FORMAT_XLSB, Some(&options))
        .expect("open_with must succeed");

    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_memory_reads_the_same_bytes() {
    let bytes = std::fs::read(fixture_path()).expect("fixture must be readable");
    let mut workbook =
        Workbook::open_memory(&bytes, XL_FORMAT_XLSB, None).expect("open_memory must succeed");

    let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);
    assert_eq!(table.get(0).unwrap().coluna1, "Valor1");
}

#[test]
fn open_reports_the_native_error_for_a_missing_file() {
    let error = Workbook::open("does-not-exist.xlsx").expect_err("a missing file must fail");
    assert!(
        !error.message().is_empty(),
        "the failure must carry the native detail, got: {error}"
    );
}

#[test]
fn parse_sheet_reports_the_native_error_for_an_unknown_column() {
    #[derive(Default, ExcelMapper)]
    struct Missing {
        #[excel(name = "ThisColumnDoesNotExist")]
        missing: String,
    }

    let mut workbook = open_fixture();
    let error = parse_sheet::<Missing>(&mut workbook, 1).expect_err("an unknown column must fail");
    assert!(
        !error.message().is_empty(),
        "the failure must carry the native detail, got: {error}"
    );
}

#[test]
fn abi_version_matches_the_loaded_library() {
    // Every constructor gates on this; assert it directly so a mismatch is reported as itself
    // rather than as every other test failing at once.
    assert_eq!(
        unsafe { excelreader::xl_abi_version() },
        excelreader::XL_ABI_VERSION,
        "the linked native library speaks a different ABI revision than this crate"
    );
}
