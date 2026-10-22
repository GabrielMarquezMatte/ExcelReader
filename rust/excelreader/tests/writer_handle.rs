//! Round-trips through `WriterHandle`, the streaming counterpart to `write_columns`: everything
//! written here is read back with the same crate's reader, same reasoning as `write_typed.rs`.

use excelreader::ReadOptions;
use excelreader::{
    Column, ColumnData, Date, ErrorKind, ExcelMapper, Format, OpenOptions, Time, Timestamp,
    Workbook, WriteOptions, WriterHandle,
};
use std::path::PathBuf;
use std::sync::atomic::{AtomicU32, Ordering};

static COUNTER: AtomicU32 = AtomicU32::new(0);

fn temp_path(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!(
        "excelreader-rs-writer-handle-{}-{}-{name}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    path
}

#[derive(Default, Debug, PartialEq, ExcelMapper)]
struct FullRow {
    #[excel(name = "texto")]
    texto: String,
    #[excel(name = "inteiro")]
    inteiro: i64,
    #[excel(name = "numero")]
    numero: f64,
    #[excel(name = "ativo")]
    ativo: bool,
    #[excel(name = "data")]
    data: Date,
    #[excel(name = "hora")]
    hora: Time,
    #[excel(name = "instante")]
    instante: Timestamp,
    #[excel(name = "opcional")]
    opcional: Option<i64>,
}

/// Writes one header row and one data row through every `write_*` method, then a row of blank
/// cells through every `write_*(None)`.
fn write_full_rows(handle: &mut WriterHandle<'_>) {
    handle
        .start_sheet("Planilha1")
        .expect("start_sheet must succeed");

    handle
        .start_row()
        .expect("start_row must succeed for the header row");
    for header in [
        "texto", "inteiro", "numero", "ativo", "data", "hora", "instante", "opcional",
    ] {
        handle
            .write_str(Some(header))
            .expect("writing a header cell must succeed");
    }
    handle
        .end_row()
        .expect("end_row must succeed for the header row");

    handle
        .start_row()
        .expect("start_row must succeed for the data row");
    handle
        .write_str(Some("uma"))
        .expect("write_str must succeed");
    handle.write_i64(Some(1)).expect("write_i64 must succeed");
    handle.write_f64(Some(0.5)).expect("write_f64 must succeed");
    handle
        .write_bool(Some(true))
        .expect("write_bool must succeed");
    handle
        .write_date(Some(Date {
            days_since_epoch: 20454,
        }))
        .expect("write_date must succeed");
    handle
        .write_time(Some(Time {
            micros_since_midnight: 3_600_000_000,
        }))
        .expect("write_time must succeed");
    handle
        .write_timestamp(Some(Timestamp {
            micros_since_epoch: 1_767_225_600_000_000,
        }))
        .expect("write_timestamp must succeed");
    handle
        .write_i64(Some(7))
        .expect("write_i64(Some) must succeed for the optional column");
    handle
        .end_row()
        .expect("end_row must succeed for the data row");

    handle
        .start_row()
        .expect("start_row must succeed for the blank row");
    handle
        .write_str(None)
        .expect("write_str(None) must write a blank cell");
    handle
        .write_i64(None)
        .expect("write_i64(None) must write a blank cell");
    handle
        .write_f64(None)
        .expect("write_f64(None) must write a blank cell");
    handle
        .write_bool(None)
        .expect("write_bool(None) must write a blank cell");
    handle
        .write_date(None)
        .expect("write_date(None) must write a blank cell");
    handle
        .write_time(None)
        .expect("write_time(None) must write a blank cell");
    handle
        .write_timestamp(None)
        .expect("write_timestamp(None) must write a blank cell");
    handle
        .write_i64(None)
        .expect("write_i64(None) must write a blank cell");
    handle
        .end_row()
        .expect("end_row must succeed for the blank row");

    handle.end_sheet().expect("end_sheet must succeed");
}

fn assert_full_rows(workbook: &Workbook) {
    let sheet = workbook.sheet(0).expect("sheet 0");
    assert_eq!(sheet.name().expect("name must read back"), "Planilha1");
    let table = sheet
        .parse::<FullRow>(&ReadOptions::new())
        .expect("the written workbook must parse back");
    assert_eq!(table.len(), 2);

    let first = table.get(0).expect("row 0 must exist");
    assert_eq!(first.texto, "uma");
    assert_eq!(first.inteiro, 1);
    assert_eq!(first.numero, 0.5);
    assert!(first.ativo);
    assert_eq!(
        first.data,
        Date {
            days_since_epoch: 20454
        }
    );
    assert_eq!(
        first.hora,
        Time {
            micros_since_midnight: 3_600_000_000
        }
    );
    assert_eq!(first.opcional, Some(7));

    let second = table.get(1).expect("row 1 must exist");
    assert_eq!(second.texto, "");
    assert_eq!(second.opcional, None);
}

#[test]
fn create_round_trips_through_a_file() {
    let path = temp_path("full.xlsx");
    let mut handle =
        WriterHandle::create(&path, &WriteOptions::new()).expect("create must succeed");
    write_full_rows(&mut handle);
    handle.finish().expect("finish must succeed");

    let workbook = Workbook::open(&path).expect("the file WriterHandle produced must open");
    assert_full_rows(&workbook);
    drop(workbook);
    std::fs::remove_file(&path).ok();
}

#[test]
fn new_writes_the_workbook_to_its_writer_on_finish() {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(Format::Xlsx);
    let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");
    write_full_rows(&mut handle);
    handle.finish().expect("finish must succeed");
    assert!(!bytes.is_empty());

    let workbook = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("the bytes finish wrote must open");
    assert_full_rows(&workbook);
}

#[test]
fn new_needs_an_explicit_format() {
    let error = WriterHandle::new(Vec::new(), &WriteOptions::new())
        .expect_err("a writer has no extension to infer a format from");
    assert_eq!(error.kind(), ErrorKind::InvalidArgument);
}

#[test]
fn dropping_an_unfinished_handle_writes_nothing() {
    let mut bytes = Vec::new();
    {
        let options = WriteOptions::new().format(Format::Xlsx);
        let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");
        handle.start_sheet("S").expect("start_sheet must succeed");
    }
    assert!(bytes.is_empty());
}

#[test]
fn writer_handle_rejects_out_of_order_calls() {
    let path = temp_path("order.xlsx");
    {
        let mut handle =
            WriterHandle::create(&path, &WriteOptions::new()).expect("create must succeed");

        assert!(
            handle.start_row().is_err(),
            "start_row before start_sheet must fail"
        );
        assert!(
            handle.write_i64(Some(1)).is_err(),
            "a cell write before start_row must fail"
        );
        assert!(
            handle.end_row().is_err(),
            "end_row without an open row must fail"
        );
        assert!(
            handle.end_sheet().is_err(),
            "end_sheet without an open sheet must fail"
        );

        handle
            .start_sheet("S")
            .expect("start_sheet must still succeed after the earlier rejected calls");
    }
    std::fs::remove_file(&path).ok();
}

#[derive(Default, Debug, Clone, PartialEq, ExcelMapper)]
struct Item {
    #[excel(name = "name")]
    name: String,
    #[excel(name = "qty")]
    qty: Option<i64>,
}

fn item(name: &str, qty: Option<i64>) -> Item {
    Item {
        name: name.to_string(),
        qty,
    }
}

fn read_items(workbook: &Workbook, index: i32) -> (String, Vec<Item>) {
    let sheet = workbook.sheet(index).expect("the sheet must exist");
    let table = sheet
        .parse::<Item>(&ReadOptions::new())
        .expect("the sheet must parse back");
    let rows = (0..table.len())
        .map(|row| {
            let read = table.get(row).expect("the row must exist");
            item(&read.name, read.qty)
        })
        .collect();
    (sheet.name().expect("name must read back"), rows)
}

#[test]
fn write_sheet_and_write_columns_write_whole_sheets_beside_streamed_ones() {
    for (format, degree) in [
        (Format::Xlsx, 1),
        (Format::Xlsb, 1),
        (Format::Xlsx, 4),
        (Format::Xlsb, 4),
    ] {
        let mut bytes = Vec::new();
        let options = WriteOptions::new()
            .format(format)
            .degree_of_parallelism(degree)
            .batch_rows(2);
        let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");

        let sales = [item("widget", Some(3)), item("gadget", Some(7))];
        handle
            .write_sheet("Vendas", &sales, None)
            .expect("write_sheet must succeed");

        handle
            .start_sheet("Resumo")
            .expect("start_sheet must succeed");
        handle.start_row().expect("start_row must succeed");
        handle
            .write_str(Some("name"))
            .expect("write_str must succeed");
        handle
            .write_str(Some("qty"))
            .expect("write_str must succeed");
        handle.end_row().expect("end_row must succeed");
        handle.start_row().expect("start_row must succeed");
        handle
            .write_str(Some("total"))
            .expect("write_str must succeed");
        handle.write_i64(Some(10)).expect("write_i64 must succeed");
        handle.end_row().expect("end_row must succeed");
        handle.end_sheet().expect("end_sheet must succeed");

        let offsets = [0i32, 1, 2, 3, 4, 5];
        let quantities = [10i64, 20, 0, 40, 50];
        let validity = [0b1_1011u8];
        let columns = [
            Column::new(
                Some("name"),
                ColumnData::Str {
                    offsets: &offsets,
                    data: b"abcde",
                },
            ),
            Column::new(Some("qty"), ColumnData::I64(&quantities)).validity(&validity),
        ];
        handle
            .write_columns("Estoque", &columns, None)
            .expect("write_columns must succeed");
        handle.finish().expect("finish must succeed");

        let workbook = OpenOptions::new()
            .open_bytes(&bytes)
            .expect("the bytes finish wrote must open");
        assert_eq!(
            read_items(&workbook, 0),
            ("Vendas".to_string(), sales.to_vec())
        );
        assert_eq!(
            read_items(&workbook, 1),
            ("Resumo".to_string(), vec![item("total", Some(10))])
        );
        assert_eq!(
            read_items(&workbook, 2),
            (
                "Estoque".to_string(),
                vec![
                    item("a", Some(10)),
                    item("b", Some(20)),
                    item("c", None),
                    item("d", Some(40)),
                    item("e", Some(50)),
                ]
            ),
            "{format:?} at degree {degree}"
        );
    }
}

#[test]
fn write_sheet_fails_while_a_streamed_sheet_is_open_and_the_handle_stays_usable() {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(Format::Xlsx);
    let mut handle = WriterHandle::new(&mut bytes, &options).expect("new must succeed");
    let rows = [item("widget", Some(3))];

    handle
        .start_sheet("Aberta")
        .expect("start_sheet must succeed");
    let error = handle
        .write_sheet("Tabela", &rows, None)
        .expect_err("a whole sheet cannot start while another is open");
    assert!(error.to_string().contains("already open"), "{error}");
    handle.end_sheet().expect("end_sheet must succeed");
    handle
        .write_sheet("Tabela", &rows, None)
        .expect("write_sheet must succeed once the sheet is ended");
    handle.finish().expect("finish must succeed");

    let workbook = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("the bytes finish wrote must open");
    assert_eq!(read_items(&workbook, 1).0, "Tabela");
}
