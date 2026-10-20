use excelreader::{
    CellType, ErrorKind, Format, OpenOptions, RowRef, Workbook, WriteOptions, WriterHandle,
};
use std::panic::{catch_unwind, AssertUnwindSafe};

fn fixture(name: &str) -> String {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .expect("repo root is two levels above the crate");
    root.join(name).to_string_lossy().into_owned()
}

fn generated(format: Format) -> Workbook {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(format);
    let mut writer = WriterHandle::new(&mut bytes, &options).expect("open writer");
    writer.start_sheet("S1").expect("start sheet");
    writer.start_row().expect("start row");
    writer.write_str(Some("id")).expect("write");
    writer.write_str(Some("label")).expect("write");
    writer.end_row().expect("end row");
    for i in 1..=200i64 {
        writer.start_row().expect("start row");
        writer.write_i64(Some(i)).expect("write");
        writer
            .write_str(Some(&format!("label-{i}")))
            .expect("write");
        writer.write_f64(Some(i as f64 * 1e20)).expect("write");
        writer.end_row().expect("end row");
    }
    writer.end_sheet().expect("end sheet");
    writer.finish().expect("finish");
    OpenOptions::new()
        .format(format)
        .open_bytes(&bytes)
        .expect("open")
}

type Cell = (i32, CellType, Option<f64>, String);

fn collect(row: RowRef<'_>) -> Vec<Cell> {
    row.iter()
        .map(|cell| {
            (
                cell.column,
                cell.cell_type,
                cell.number(),
                cell.as_str().unwrap().to_owned(),
            )
        })
        .collect()
}

fn assert_same_as_cursor(workbook: &Workbook) {
    let sheet = workbook.sheet(0).expect("sheet 0");
    let mut pushed: Vec<Vec<Cell>> = Vec::new();
    sheet
        .for_each_row(|row| {
            pushed.push(collect(row));
            Ok(())
        })
        .expect("for_each_row");

    let mut pulled: Vec<Vec<Cell>> = Vec::new();
    let mut cursor = sheet.rows().expect("cursor");
    while let Some(row) = cursor.next_row() {
        pulled.push(collect(row.expect("row")));
    }

    assert_eq!(pushed.len(), pulled.len());
    for (left, right) in pushed.iter().zip(&pulled) {
        assert_eq!(left.len(), right.len());
        for (a, b) in left.iter().zip(right) {
            assert_eq!((a.0, a.1), (b.0, b.1));
            if !a.3.is_empty() && !b.3.is_empty() {
                assert_eq!(a.3, b.3);
            } else {
                assert_eq!(a.2, b.2);
            }
        }
    }
}

#[test]
fn for_each_row_matches_the_cursor_on_every_format() {
    assert_same_as_cursor(&Workbook::open(fixture("RealExcel.xlsx")).expect("open"));
    assert_same_as_cursor(&Workbook::open(fixture("RealExcel.xlsb")).expect("open"));
    assert_same_as_cursor(&generated(Format::Xlsb));
    assert_same_as_cursor(&generated(Format::Xls));
    assert_same_as_cursor(&generated(Format::Csv));
}

#[test]
fn a_positive_callback_code_stops_the_walk() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let mut count = 0;
    let result = workbook.sheet(0).expect("sheet 0").for_each_row(|_| {
        count += 1;
        if count == 3 {
            Err(7)
        } else {
            Ok(())
        }
    });
    assert_eq!(result.unwrap_err().kind(), ErrorKind::Callback(7));
    assert_eq!(count, 3);
}

#[test]
fn a_non_positive_callback_code_is_an_error() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let result = workbook
        .sheet(0)
        .expect("sheet 0")
        .for_each_row(|_| Err(-1));
    assert!(result.is_err());
}

#[test]
fn a_nested_call_inside_the_callback_works() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let sheet = workbook.sheet(0).expect("sheet 0");
    let name = sheet.name().expect("name");
    sheet
        .for_each_row(|_| {
            assert_eq!(sheet.name().expect("nested name"), name);
            Ok(())
        })
        .expect("for_each_row");
}

#[test]
fn a_panic_is_resumed_and_the_workbook_stays_usable() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let sheet = workbook.sheet(0).expect("sheet 0");
    let payload = catch_unwind(AssertUnwindSafe(|| sheet.for_each_row(|_| panic!("boom"))))
        .expect_err("the panic must be resumed");
    assert_eq!(payload.downcast_ref::<&str>(), Some(&"boom"));
    assert!(sheet.for_each_row(|_| Ok(())).is_ok());
}

#[test]
fn number_only_cells_format_their_text() {
    let workbook = Workbook::open(fixture("RealExcel.xlsb")).expect("open");
    let mut seen = 0;
    workbook
        .sheet(0)
        .expect("sheet 0")
        .for_each_row(|row| {
            for cell in row.iter() {
                if cell.as_bytes().is_empty() {
                    if let Some(number) = cell.number() {
                        assert_eq!(cell.text().unwrap(), number.to_string());
                        seen += 1;
                    }
                }
            }
            Ok(())
        })
        .expect("for_each_row");
    assert!(
        seen > 0,
        "the fixture must have a cell that carries only a number"
    );
}
