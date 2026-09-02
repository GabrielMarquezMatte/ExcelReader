use excelreader::workbook::Workbook;
use excelreader::CellType;

fn fixture(name: &str) -> String {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .expect("repo root is two levels above the crate");
    root.join(name).to_string_lossy().into_owned()
}

#[test]
fn cursor_reads_the_header_row() {
    let mut workbook = Workbook::open(&fixture("RealExcel.xlsx")).expect("open");
    let mut cursor = workbook.rows();

    let row = cursor.next_row().expect("a first row").expect("no error");
    let names: Vec<String> = row.iter().map(|cell| cell.as_str().unwrap().to_string()).collect();

    assert!(names.contains(&"Coluna1".to_string()), "got {names:?}");
    assert!(names.contains(&"Coluna3".to_string()), "got {names:?}");
}

#[test]
fn cursor_terminates_cleanly() {
    let mut workbook = Workbook::open(&fixture("RealExcel.xlsx")).expect("open");
    let mut counted = 0usize;
    {
        let mut cursor = workbook.rows();
        while let Some(row) = cursor.next_row() {
            row.expect("no error mid-sheet");
            counted += 1;
        }
    }
    assert!(counted > 1, "fixture should have a header plus data rows");

    // A second cursor on an exhausted sheet yields nothing rather than erroring.
    let mut cursor = workbook.rows();
    assert!(cursor.next_row().is_none());
}

#[test]
fn move_to_sheet_resets_the_cursor() {
    let mut workbook = Workbook::open(&fixture("RealExcel.xlsx")).expect("open");
    let first: Vec<String> = {
        let mut cursor = workbook.rows();
        let row = cursor.next_row().expect("row").expect("ok");
        row.iter().map(|c| c.as_str().unwrap().to_string()).collect()
    };

    workbook.move_to_sheet(0).expect("re-select sheet 0");

    let again: Vec<String> = {
        let mut cursor = workbook.rows();
        let row = cursor.next_row().expect("row").expect("ok");
        row.iter().map(|c| c.as_str().unwrap().to_string()).collect()
    };
    assert_eq!(first, again);
}

#[test]
fn grows_its_buffer_for_an_oversized_row() {
    // One cell far larger than the cursor's initial buffer, so xl_next_row must answer
    // XL_BUFFER_TOO_SMALL at least once and the cursor must retry without losing the row.
    let big = "x".repeat(200_000);
    let dir = std::env::temp_dir().join("excelreader-rust-rows-test");
    std::fs::create_dir_all(&dir).expect("temp dir");
    let path = dir.join("wide.csv");
    std::fs::write(&path, format!("a\n{big}\n")).expect("write fixture");

    let mut workbook =
        Workbook::open_with(&path.to_string_lossy(), excelreader::XL_FORMAT_CSV, None).expect("open csv");
    let mut cursor = workbook.rows();

    let header = cursor.next_row().expect("header").expect("ok");
    assert_eq!(header.get(0).unwrap().as_str().unwrap(), "a");

    let row = cursor.next_row().expect("data row").expect("ok");
    let cell = row.get(0).expect("one cell");
    assert_eq!(cell.cell_type, CellType::String);
    assert_eq!(cell.as_str().unwrap().len(), 200_000);

    drop(cursor);
    drop(workbook);
    std::fs::remove_file(&path).ok();
}
