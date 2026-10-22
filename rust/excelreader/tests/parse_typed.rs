use excelreader::{write_columns_to, ColumnData, ReadOptions, Table, WriteOptions};
use excelreader::{ColumnType, Date, ExcelMapper, Format, OpenOptions, Workbook};

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
    Workbook::open(fixture_path()).expect("open must succeed")
}

#[test]
fn parses_real_excel_fixture() {
    let workbook = open_fixture();
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");
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
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");

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
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");

    let iter = table.iter();
    assert_eq!(iter.len(), 100, "ExactSizeIterator must agree with len()");
    assert_eq!(iter.count(), 100);
}

#[test]
fn parses_integer_widths_floats_and_dates() {
    let workbook = open_fixture();
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<WideRow>(&ReadOptions::new())
        .expect("parse must succeed");

    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1u16);
    assert!((first.coluna16 - 0.1f32).abs() < f32::EPSILON);
    assert_eq!(
        first.coluna2,
        Date {
            days_since_epoch: 20_454
        }
    );
}

#[test]
fn resolves_the_first_alias_present_in_the_header_row() {
    let workbook = open_fixture();
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<AliasRow>(&ReadOptions::new())
        .expect("parse must succeed via alias");
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
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<ChronoRow>(&ReadOptions::new())
        .expect("parse must succeed");
    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna2, NaiveDate::from_ymd_opt(2026, 1, 1).unwrap());
}

#[test]
fn exposes_sheet_navigation() {
    let workbook = open_fixture();

    let sheets = workbook.sheets().expect("sheets must succeed");
    assert!(!sheets.is_empty(), "the fixture has at least one sheet");
    sheets[0].name().expect("name must succeed");

    workbook.is_date1904().expect("is_date1904 must succeed");
}

#[test]
fn infers_a_schema_from_the_header_row() {
    let workbook = open_fixture();
    let schema = workbook
        .sheet(0)
        .expect("sheet 0")
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");

    assert!(!schema.is_empty(), "the fixture has columns to infer");
    let first = &schema[0];
    assert_eq!(first.name.as_deref(), Some("Coluna1"));
    assert_eq!(first.column_type, ColumnType::String);

    let coluna3 = schema
        .iter()
        .find(|c| c.name.as_deref() == Some("Coluna3"))
        .expect("Coluna3 must be inferred");
    assert_eq!(coluna3.column_type, ColumnType::I64);
}

#[test]
fn infer_schema_leaves_the_sheet_readable() {
    let workbook = open_fixture();
    workbook
        .sheet(0)
        .expect("sheet 0")
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");

    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_options_accept_an_explicit_format_and_settings() {
    let workbook = OpenOptions::new()
        .format(Format::Xlsb)
        .prefetch_decompression(true)
        .open(fixture_path())
        .expect("open must succeed");

    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");
    assert_eq!(table.len(), 100);
}

#[test]
fn open_bytes_reads_the_same_bytes() {
    let bytes = std::fs::read(fixture_path()).expect("fixture must be readable");
    let workbook = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("open_bytes must succeed");

    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("parse must succeed");
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
fn parse_reports_the_native_error_for_an_unknown_column() {
    #[derive(Default, ExcelMapper)]
    struct Missing {
        #[excel(name = "ThisColumnDoesNotExist")]
        missing: String,
    }

    let workbook = open_fixture();
    let error = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Missing>(&ReadOptions::new())
        .expect_err("an unknown column must fail");
    assert!(
        !error.message().is_empty(),
        "the failure must carry the native detail, got: {error}"
    );
}

#[test]
fn a_non_utf8_path_is_an_invalid_argument() {
    #[cfg(windows)]
    let path = {
        use std::os::windows::ffi::OsStringExt;
        std::path::PathBuf::from(std::ffi::OsString::from_wide(&[0xD800]))
    };
    #[cfg(unix)]
    let path = {
        use std::os::unix::ffi::OsStrExt;
        std::path::PathBuf::from(std::ffi::OsStr::from_bytes(b"\xFF.xlsx"))
    };
    let error = Workbook::open(path).expect_err("a non-UTF-8 path must fail");
    assert_eq!(error.kind(), excelreader::ErrorKind::InvalidArgument);
}

fn write_csv(tag: &str, text: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("xlpt-{tag}-{}.csv", std::process::id()));
    std::fs::write(&path, text).expect("the csv must be writable");
    path
}

#[test]
fn infer_schema_parse_text_types_csv_text_fields() {
    let path = write_csv(
        "infer",
        "id,day
1,2024-01-02
2,2024-03-04
",
    );
    let workbook = OpenOptions::new()
        .format(Format::Csv)
        .open(&path)
        .expect("open must succeed");
    let sheet = workbook.sheet(0).expect("sheet 0");

    let plain = sheet
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");
    assert!(plain.iter().all(|c| c.column_type == ColumnType::String));

    let typed = sheet
        .infer_schema(1, 100, true)
        .expect("infer_schema with parse_text must succeed");
    assert_eq!(typed[0].column_type, ColumnType::I64);
    assert_eq!(typed[1].column_type, ColumnType::Date);
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
fn parallel_parse_matches_the_sequential_parse() {
    const ROWS: i64 = 200_000;
    let mut text = String::from(
        "id,name
",
    );
    for i in 0..ROWS {
        text.push_str(&format!(
            "{i},name-{i}
"
        ));
    }
    let path = write_csv("parallel", &text);
    let workbook = OpenOptions::new()
        .format(Format::Csv)
        .open(&path)
        .expect("open must succeed");

    let sequential = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<CsvRow>(&ReadOptions::new())
        .expect("sequential parse");
    let parallel = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<CsvRow>(&ReadOptions::new().parallelism(0))
        .expect("parallel parse");
    assert_eq!(parallel.len(), sequential.len());
    assert_eq!(parallel.len(), ROWS);
    for index in [0, parallel.len() - 1] {
        let (a, b) = (parallel.get(index).unwrap(), sequential.get(index).unwrap());
        assert_eq!((a.id, a.name), (b.id, b.name));
    }

    assert!(workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<CsvRow>(&ReadOptions::new().parallelism(-1))
        .is_err());
    drop(workbook);
    let _ = std::fs::remove_file(path);
}

fn column_by_name<'t>(table: &'t Table, name: &str) -> ColumnData<'t> {
    table
        .columns()
        .into_iter()
        .find(|c| c.name == Some(name))
        .unwrap_or_else(|| panic!("{name} must be read"))
        .data
}

#[test]
fn an_inferred_schema_parses_without_a_row_type() {
    let workbook = open_fixture();
    let sheet = workbook.sheet(0).expect("sheet 0");
    let schema = sheet
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");

    let table = sheet
        .parse_with_schema(&schema, &ReadOptions::new())
        .expect("parse_with_schema must succeed");

    assert_eq!(table.len(), 100);
    assert_eq!(table.num_columns(), schema.len());
    assert!(table.column(schema.len()).is_none());
    let ColumnData::Str { offsets, data } = column_by_name(&table, "Coluna1") else {
        panic!("Coluna1 must be a string column");
    };
    assert_eq!(&data[offsets[0] as usize..offsets[1] as usize], b"Valor1");
    let ColumnData::I64(coluna3) = column_by_name(&table, "Coluna3") else {
        panic!("Coluna3 must be an i64 column");
    };
    assert_eq!(coluna3.len(), 100);
    assert_eq!(coluna3[0], 1);
}

#[test]
fn a_schema_without_names_parses_by_index() {
    let workbook = open_fixture();
    let sheet = workbook.sheet(0).expect("sheet 0");
    let schema = sheet
        .infer_schema(0, 100, false)
        .expect("infer_schema must succeed");

    let table = sheet
        .parse_with_schema(&schema, &ReadOptions::new().header_row(0))
        .expect("parse_with_schema must succeed");

    assert_eq!(table.len(), 101, "with no header the header row is data");
    assert!(table.columns().iter().all(|c| c.name.is_none()));
}

#[test]
fn parse_with_schema_rejects_an_empty_schema() {
    let workbook = open_fixture();
    assert!(workbook
        .sheet(0)
        .expect("sheet 0")
        .parse_with_schema(&[], &ReadOptions::new())
        .is_err());
}

#[test]
fn a_parsed_table_writes_back_unchanged() {
    let workbook = open_fixture();
    let sheet = workbook.sheet(0).expect("sheet 0");
    let schema = sheet
        .infer_schema(1, 100, false)
        .expect("infer_schema must succeed");
    let table = sheet
        .parse_with_schema(&schema, &ReadOptions::new())
        .expect("parse_with_schema must succeed");

    let mut bytes = Vec::new();
    write_columns_to(
        &mut bytes,
        &table.columns(),
        &WriteOptions::new().format(Format::Xlsx),
    )
    .expect("a parsed table must write");

    let reopened = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("the written workbook must open");
    let again = reopened
        .sheet(0)
        .expect("sheet 0")
        .parse_with_schema(&schema, &ReadOptions::new())
        .expect("the written workbook must parse with the same schema");
    assert_eq!(again.len(), table.len());
    let (ColumnData::I64(before), ColumnData::I64(after)) = (
        column_by_name(&table, "Coluna3"),
        column_by_name(&again, "Coluna3"),
    ) else {
        panic!("Coluna3 must stay an i64 column");
    };
    assert_eq!(before, after);
}
