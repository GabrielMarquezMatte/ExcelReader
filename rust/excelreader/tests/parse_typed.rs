use excelreader::workbook::{parse_sheet, parse_sheet_parallel, ExcelMapper, Workbook};
use excelreader::{Date, OpenOptions, XL_FORMAT_CSV, XL_FORMAT_XLSB, XL_T_DATE, XL_T_I64, XL_T_STRING};

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
    let workbook = open_fixture();
    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);

    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1);

    let all: Vec<Row> = table.iter().collect();
    assert_eq!(all.len(), 100);
}

#[test]
fn get_returns_none_outside_the_row_range() {
    let workbook = open_fixture();
    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");

    assert!(
        table.get(table.len() - 1).is_some(),
        "last row is in bounds"
    );
    assert!(table.get(table.len()).is_none(), "one past the end");
    assert!(table.get(i64::MAX).is_none(), "far past the end");
    assert!(table.get(-1).is_none(), "negative row");
}

#[test]
fn iter_yields_exactly_len_rows_and_reports_it_up_front() {
    let workbook = open_fixture();
    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");

    let iter = table.iter();
    assert_eq!(iter.len(), 100, "ExactSizeIterator must agree with len()");
    assert_eq!(iter.count(), 100);
}

#[test]
fn parses_integer_widths_floats_and_dates() {
    let workbook = open_fixture();
    let table = parse_sheet::<WideRow>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");

    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1u16);
    assert!((first.coluna16 - 0.1f32).abs() < f32::EPSILON);
    assert_eq!(first.coluna2, Date::new(20_454));
}

#[test]
fn resolves_the_first_alias_present_in_the_header_row() {
    let workbook = open_fixture();
    let table = parse_sheet::<AliasRow>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed via alias");
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

    let workbook = open_fixture();
    let table = parse_sheet::<ChronoRow>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");
    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna2, NaiveDate::from_ymd_opt(2026, 1, 1).unwrap());
}

#[test]
fn exposes_sheet_navigation() {
    let workbook = open_fixture();

    let count = workbook.sheet_count().expect("sheet_count must succeed");
    assert!(count >= 1, "the fixture has at least one sheet");

    let names = workbook.sheet_names().expect("sheet_names must succeed");
    assert_eq!(names.len(), count as usize);

    workbook.is_date1904().expect("is_date1904 must succeed");
}

#[test]
fn infers_a_schema_from_the_header_row() {
    let workbook = open_fixture();
    let schema = workbook.sheet(0).expect("sheet 0").infer_schema(1, 100)
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
fn infer_schema_leaves_the_sheet_readable() {
    let workbook = open_fixture();
    workbook.sheet(0).expect("sheet 0").infer_schema(1, 100)
        .expect("infer_schema must succeed");

    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_with_accepts_an_explicit_format_and_options() {
    let options = OpenOptions::new().prefetch_decompression(true);
    let workbook = Workbook::open_with(&fixture_path(), XL_FORMAT_XLSB, Some(&options))
        .expect("open_with must succeed");

    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_memory_reads_the_same_bytes() {
    let bytes = std::fs::read(fixture_path()).expect("fixture must be readable");
    let workbook =
        Workbook::open_memory(&bytes, XL_FORMAT_XLSB, None).expect("open_memory must succeed");

    let table = parse_sheet::<Row>(workbook.sheet(0).expect("sheet 0"), 1).expect("parse_sheet must succeed");
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

    let workbook = open_fixture();
    let error = parse_sheet::<Missing>(workbook.sheet(0).expect("sheet 0"), 1).expect_err("an unknown column must fail");
    assert!(
        !error.message().is_empty(),
        "the failure must carry the native detail, got: {error}"
    );
}

#[test]
fn abi_version_matches_the_loaded_library() {
    assert_eq!(
        unsafe { excelreader::xl_abi_version() },
        excelreader::XL_ABI_VERSION,
        "the linked native library speaks a different ABI revision than this crate"
    );
}

fn write_csv(tag: &str, text: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("xlpt-{tag}-{}.csv", std::process::id()));
    std::fs::write(&path, text).expect("the csv must be writable");
    path
}

#[test]
fn infer_schema_parse_text_types_csv_text_fields() {
    let path = write_csv("infer", "id,day
1,2024-01-02
2,2024-03-04
");
    let workbook = Workbook::open_with(path.to_str().unwrap(), XL_FORMAT_CSV, None).expect("open must succeed");
    let sheet = workbook.sheet(0).expect("sheet 0");

    let plain = sheet.infer_schema(1, 100).expect("infer_schema must succeed");
    assert!(plain.iter().all(|c| c.column_type == XL_T_STRING));

    let typed = sheet.infer_schema_parse_text(1, 100).expect("infer_schema_parse_text must succeed");
    assert_eq!(typed[0].column_type, XL_T_I64);
    assert_eq!(typed[1].column_type, XL_T_DATE);
    drop(workbook);
    let _ = std::fs::remove_file(path);
}

#[derive(Default, ExcelMapper)]
struct CsvRow {
    #[excel(name = "id")]
    id: i64,
    #[excel(name = "name")]
    name: String,
}

#[test]
fn parse_sheet_parallel_matches_the_sequential_parse() {
    const ROWS: i64 = 200_000;
    let mut text = String::from("id,name
");
    for i in 0..ROWS {
        text.push_str(&format!("{i},name-{i}
"));
    }
    let path = write_csv("parallel", &text);
    let workbook = Workbook::open_with(path.to_str().unwrap(), XL_FORMAT_CSV, None).expect("open must succeed");

    let sequential = parse_sheet::<CsvRow>(workbook.sheet(0).expect("sheet 0"), 1).expect("sequential parse");
    let parallel = parse_sheet_parallel::<CsvRow>(workbook.sheet(0).expect("sheet 0"), 1, 0).expect("parallel parse");
    assert_eq!(parallel.len(), sequential.len());
    assert_eq!(parallel.len() as i64, ROWS);
    for index in [0, parallel.len() - 1] {
        let (a, b) = (parallel.get(index).unwrap(), sequential.get(index).unwrap());
        assert_eq!((a.id, a.name), (b.id, b.name));
    }

    assert!(parse_sheet_parallel::<CsvRow>(workbook.sheet(0).expect("sheet 0"), 1, -1).is_err());
    drop(workbook);
    let _ = std::fs::remove_file(path);
}
