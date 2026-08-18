use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna3")]
    coluna3: i64,
}

fn fixture_path() -> String {
    concat!(env!("CARGO_MANIFEST_DIR"), "/../../RealExcel.xlsb").to_string()
}

#[test]
fn parses_real_excel_fixture() {
    let workbook = Workbook::open(&fixture_path()).expect("open must succeed");
    let table = parse_sheet::<Row>(&workbook, 1).expect("parse_sheet must succeed");
    assert_eq!(table.len(), 100);

    let first = table.get(0);
    assert_eq!(first.coluna1, "Valor1");
    assert_eq!(first.coluna3, 1);

    let all: Vec<Row> = table.iter().collect();
    assert_eq!(all.len(), 100);
}
