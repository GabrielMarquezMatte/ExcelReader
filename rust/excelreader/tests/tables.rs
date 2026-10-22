//! Excel tables: written through `WriterHandle` with `TableOptions`, read back through
//! `Workbook::tables`/`table`/`table_by_name`.

use excelreader::{
    Column, ColumnData, ExcelMapper, Format, OpenOptions, ReadOptions, TableInfo, TableOptions,
    Workbook, WriteOptions, WriterHandle,
};

#[derive(Default, Debug, Clone, PartialEq, ExcelMapper)]
struct Item {
    #[excel(name = "name")]
    name: String,
    #[excel(name = "qty")]
    qty: i64,
}

fn item(name: &str, qty: i64) -> Item {
    Item {
        name: name.to_string(),
        qty,
    }
}

fn workbook_with_tables(format: Format, degree: i32) -> Vec<u8> {
    let mut bytes = Vec::new();
    let options = WriteOptions::new()
        .format(format)
        .degree_of_parallelism(degree)
        .batch_rows(1);
    let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");

    handle.start_sheet("Notas").expect("start_sheet must succeed");
    handle.start_row().expect("start_row must succeed");
    handle.write_str(Some("antes")).expect("write_str must succeed");
    handle.end_row().expect("end_row must succeed");
    handle.end_sheet().expect("end_sheet must succeed");

    let sales = [item("widget", 3), item("gadget", 7)];
    handle
        .write_sheet("Vendas", &sales, Some(&TableOptions::new("Sales")))
        .expect("write_sheet as a table must succeed");

    let offsets = [0i32, 1, 2, 3];
    let quantities = [10i64, 20, 30];
    let columns = [
        Column::new(
            Some("name"),
            ColumnData::Str {
                offsets: &offsets,
                data: b"abc",
            },
        ),
        Column::new(Some("qty"), ColumnData::I64(&quantities)),
    ];
    let stock = TableOptions::new("Stock").style(None).row_stripes(false);
    handle
        .write_columns("Estoque", &columns, Some(&stock))
        .expect("write_columns as a table must succeed");
    handle.finish().expect("finish must succeed");
    bytes
}

fn read_items(workbook: &Workbook, sheet: excelreader::Sheet<'_>) -> Vec<Item> {
    let _ = workbook;
    let table = sheet
        .parse::<Item>(&ReadOptions::new())
        .expect("the table must parse");
    (0..table.len())
        .map(|row| {
            let read = table.get(row).expect("the row must exist");
            item(&read.name, read.qty)
        })
        .collect()
}

#[test]
fn tables_written_through_the_handle_read_back() {
    for (format, degree) in [
        (Format::Xlsx, 1),
        (Format::Xlsb, 1),
        (Format::Xlsx, 4),
        (Format::Xlsb, 4),
    ] {
        let bytes = workbook_with_tables(format, degree);
        let workbook = OpenOptions::new()
            .format(format)
            .open_bytes(&bytes)
            .expect("the bytes finish wrote must open");

        let tables = workbook.tables().expect("tables must list");
        let names: Vec<&str> = tables.iter().map(|t| t.name.as_str()).collect();
        assert_eq!(names, ["Sales", "Stock"], "{format:?} at degree {degree}");
        let sales: &TableInfo = &tables[0];
        assert_eq!(
            (
                sales.index,
                sales.sheet,
                sales.first_row,
                sales.last_row,
                sales.first_column,
                sales.last_column,
                sales.header_row_count,
                sales.totals_row_count
            ),
            (0, 1, 0, 2, 0, 1, 1, 0)
        );

        let by_name = workbook
            .table_by_name("sales")
            .expect("lookup must succeed")
            .expect("the table must be found ignoring case");
        assert_eq!(by_name.name().expect("name must read"), "Vendas");
        assert_eq!(
            read_items(&workbook, by_name),
            [item("widget", 3), item("gadget", 7)]
        );
        assert_eq!(
            read_items(&workbook, workbook.table(1).expect("table 1 must exist")),
            [item("a", 10), item("b", 20), item("c", 30)]
        );

        assert!(workbook.table(2).is_err(), "a missing table index must fail");
        assert!(workbook
            .table_by_name("missing")
            .expect("lookup must succeed")
            .is_none());
    }
}

#[test]
fn a_table_is_refused_outside_xlsx_and_xlsb() {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(Format::Xls);
    let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");
    let rows = [item("widget", 3)];
    let error = handle
        .write_sheet("Vendas", &rows, Some(&TableOptions::new("Sales")))
        .expect_err("XLS has no tables");
    assert!(error.to_string().contains("XLSX or XLSB"), "{error}");
    handle
        .write_sheet("Vendas", &rows, None)
        .expect("a plain sheet must still be written");
    handle.finish().expect("finish must succeed");
}
