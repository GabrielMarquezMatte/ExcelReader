use excelreader::workbook::{column_i64, column_str, parse_sheet, ColumnBinding, ExcelMapper, Workbook};
use excelreader::{XL_T_I64, XL_T_STRING};

#[derive(Default)]
struct Row {
    coluna1: String,
    coluna3: i64,
}

impl ExcelMapper for Row {
    fn bindings() -> Vec<ColumnBinding<Self>> {
        vec![
            ColumnBinding {
                name: "Coluna1",
                xl_type: XL_T_STRING,
                assign: |r, col, row| r.coluna1 = column_str(col, row).to_string(),
            },
            ColumnBinding {
                name: "Coluna3",
                xl_type: XL_T_I64,
                assign: |r, col, row| r.coluna3 = column_i64(col, row),
            },
        ]
    }
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
