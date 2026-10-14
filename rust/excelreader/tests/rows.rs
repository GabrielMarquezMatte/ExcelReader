use excelreader::{CellType, Format, OpenOptions, Workbook, WriteOptions, WriterHandle};

fn fixture(name: &str) -> String {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .expect("repo root is two levels above the crate");
    root.join(name).to_string_lossy().into_owned()
}

fn empty_sheet_workbook() -> Workbook {
    let mut bytes = Vec::new();
    let options = WriteOptions::new().format(Format::Xlsx);
    let mut writer = WriterHandle::new(&mut bytes, &options).expect("open writer");
    writer.start_sheet("Empty").expect("start sheet");
    writer.end_sheet().expect("end sheet");
    writer.finish().expect("finish");
    OpenOptions::new().open_bytes(&bytes).expect("open")
}

#[test]
fn cursor_reads_the_header_row() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let mut cursor = workbook.sheet(0).expect("sheet 0").rows().expect("cursor");

    let row = cursor.next_row().expect("a first row").expect("no error");
    let names: Vec<String> = row
        .iter()
        .map(|cell| cell.as_str().unwrap().to_string())
        .collect();

    assert!(names.contains(&"Coluna1".to_string()), "got {names:?}");
    assert!(names.contains(&"Coluna3".to_string()), "got {names:?}");
}

#[test]
fn cursor_terminates_cleanly() {
    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let mut counted = 0usize;
    {
        let mut cursor = workbook.sheet(0).expect("sheet 0").rows().expect("cursor");
        while let Some(row) = cursor.next_row() {
            row.expect("no error mid-sheet");
            counted += 1;
        }
        assert!(cursor.next_row().is_none());
    }
    assert!(counted > 1, "fixture should have a header plus data rows");

    let mut again = workbook.sheet(0).expect("sheet 0").rows().expect("cursor");
    let mut second = 0usize;
    while let Some(row) = again.next_row() {
        row.expect("no error mid-sheet");
        second += 1;
    }
    assert_eq!(second, counted, "a new cursor starts from the top");
}

#[test]
fn grows_its_buffer_for_an_oversized_row() {
    let big = "x".repeat(200_000);
    let dir = std::env::temp_dir().join("excelreader-rust-rows-test");
    std::fs::create_dir_all(&dir).expect("temp dir");
    let path = dir.join("wide.csv");
    std::fs::write(&path, format!("a\n{big}\n")).expect("write fixture");

    let workbook = OpenOptions::new()
        .format(Format::Csv)
        .open(&path)
        .expect("open csv");
    let mut cursor = workbook.sheet(0).expect("sheet 0").rows().expect("cursor");

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

#[test]
fn read_all_agrees_with_the_cursor() {
    let from_cursor: Vec<Vec<String>> = {
        let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
        let mut cursor = workbook.sheet(0).expect("sheet 0").rows().expect("cursor");
        let mut all = Vec::new();
        while let Some(row) = cursor.next_row() {
            let row = row.expect("ok");
            all.push(
                row.iter()
                    .map(|c| c.as_str().unwrap().to_string())
                    .collect(),
            );
        }
        all
    };

    let workbook = Workbook::open(fixture("RealExcel.xlsx")).expect("open");
    let all_rows = workbook
        .sheet(0)
        .expect("sheet 0")
        .read_all()
        .expect("read_all");

    assert_eq!(all_rows.len(), from_cursor.len());
    for (index, expected) in from_cursor.iter().enumerate() {
        let row = all_rows.get(index).expect("row in range");
        let actual: Vec<String> = row
            .iter()
            .map(|c| c.as_str().unwrap().to_string())
            .collect();
        assert_eq!(&actual, expected, "row {index} differs");
    }
}

#[test]
fn read_all_on_an_empty_sheet_is_empty_not_an_error() {
    let workbook = empty_sheet_workbook();
    let all_rows = workbook
        .sheet(0)
        .expect("sheet 0")
        .read_all()
        .expect("an empty sheet is not an error");
    assert_eq!(all_rows.len(), 0);
    assert!(all_rows.is_empty());
    assert_eq!(all_rows.iter().count(), 0);
}

#[test]
fn read_all_grows_its_buffer_past_the_initial_size() {
    let big = "x".repeat(2_000_000);
    let dir = std::env::temp_dir().join("excelreader-rust-all-rows-blob-test");
    std::fs::create_dir_all(&dir).expect("temp dir");
    let path = dir.join("wide.csv");
    std::fs::write(&path, format!("a\n{big}\n")).expect("write fixture");

    let workbook = OpenOptions::new()
        .format(Format::Csv)
        .open(&path)
        .expect("open csv");
    let all_rows = workbook
        .sheet(0)
        .expect("sheet 0")
        .read_all()
        .expect("read_all");

    assert_eq!(all_rows.len(), 2);
    assert_eq!(
        all_rows.get(0).unwrap().get(0).unwrap().as_str().unwrap(),
        "a"
    );
    let data_cell = all_rows.get(1).unwrap().get(0).unwrap();
    assert_eq!(data_cell.cell_type, CellType::String);
    assert_eq!(data_cell.as_str().unwrap().len(), 2_000_000);

    drop(all_rows);
    drop(workbook);
    std::fs::remove_file(&path).ok();
}
