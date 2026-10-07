use excelreader::workbook::{Sheet, SheetVisibility, Workbook};
use excelreader::writer_handle::WriterHandle;
use excelreader::{RowCursor, XL_FORMAT_AUTO, XL_FORMAT_XLSX};

fn two_sheets() -> Workbook {
    let mut writer = WriterHandle::open_memory(XL_FORMAT_XLSX, None).expect("open writer");
    for (name, rows) in [("First", vec!["a1", "a2"]), ("Second", vec!["b1", "b2", "b3"])] {
        writer.start_sheet(name).expect("start sheet");
        for value in rows {
            writer.start_row().expect("start row");
            writer.write_str(Some(value)).expect("write");
            writer.end_row().expect("end row");
        }
        writer.end_sheet().expect("end sheet");
    }
    let bytes = writer.bytes().expect("bytes");
    Workbook::open_memory(&bytes, XL_FORMAT_AUTO, None).expect("open")
}

fn first_column(cursor: &mut excelreader::RowCursor<'_>) -> Vec<String> {
    let mut values = Vec::new();
    while let Some(row) = cursor.next_row() {
        let row = row.expect("no error mid-sheet");
        values.push(row.get(0).expect("a cell").as_str().expect("utf-8").to_string());
    }
    values
}

#[test]
fn a_sheet_reads_its_own_rows() {
    let workbook = two_sheets();
    let second = workbook.sheet(1).expect("sheet 1");
    let first = workbook.sheet(0).expect("sheet 0");

    assert_eq!(first_column(&mut second.rows().expect("cursor")), ["b1", "b2", "b3"]);
    assert_eq!(first_column(&mut first.rows().expect("cursor")), ["a1", "a2"]);
}

#[test]
fn two_cursors_on_one_sheet_each_return_every_row() {
    let workbook = two_sheets();
    let sheet = workbook.sheet(1).expect("sheet 1");
    let mut one = sheet.rows().expect("cursor one");
    let mut two = sheet.rows().expect("cursor two");

    assert!(one.next_row().expect("a row").is_ok());
    assert!(two.next_row().expect("a row").is_ok());

    assert_eq!(first_column(&mut one), ["b2", "b3"]);
    assert_eq!(first_column(&mut two), ["b2", "b3"]);
}

#[test]
fn a_cursor_dropped_midway_leaves_the_workbook_usable() {
    let workbook = two_sheets();
    let sheet = workbook.sheet(1).expect("sheet 1");
    for _ in 0..1000 {
        let mut cursor = sheet.rows().expect("cursor");
        assert!(cursor.next_row().expect("a row").is_ok());
    }

    assert_eq!(first_column(&mut sheet.rows().expect("cursor")), ["b1", "b2", "b3"]);
}

#[test]
fn sheet_rejects_an_index_out_of_range_at_the_call() {
    let workbook = two_sheets();

    assert!(workbook.sheet(2).is_err());
    assert!(workbook.sheet(-1).is_err());
}

#[test]
fn sheets_lists_every_sheet_with_its_name_index_and_visibility() {
    let workbook = two_sheets();
    let sheets = workbook.sheets().expect("sheets");

    assert_eq!(sheets.len(), 2);
    assert_eq!(sheets[1].index(), 1);
    assert_eq!(sheets[1].name().expect("name"), "Second");
    assert_eq!(sheets[0].visibility().expect("visibility"), SheetVisibility::Visible);
}

#[test]
fn sheet_by_name_ignores_case_and_reports_a_miss_as_none() {
    let workbook = two_sheets();

    assert_eq!(workbook.sheet_by_name("Second").expect("ok").expect("found").index(), 1);
    assert_eq!(workbook.sheet_by_name("sEcOnD").expect("ok").expect("found").index(), 1);
    assert!(workbook.sheet_by_name("Missing").expect("ok").is_none());
}

#[test]
fn read_all_reads_the_named_sheet() {
    let workbook = two_sheets();
    let sheet = workbook.sheet(1).expect("sheet 1");

    assert_eq!(sheet.read_all_decoded().expect("decoded").len(), 3);
    assert_eq!(sheet.read_all_blob().expect("blob").len(), 3);
    assert_eq!(workbook.sheet(0).expect("sheet 0").read_all_blob().expect("blob").len(), 2);
}

#[test]
fn infer_schema_samples_the_named_sheet() {
    let workbook = two_sheets();

    let columns = workbook.sheet(1).expect("sheet 1").infer_schema(0, 10).expect("schema");

    assert_eq!(columns.len(), 1);
}

const PARALLEL_SHEETS: usize = 6;
const PARALLEL_ROWS: usize = 500;

fn many_sheets() -> Workbook {
    let mut writer = WriterHandle::open_memory(XL_FORMAT_XLSX, None).expect("open writer");
    for sheet in 0..PARALLEL_SHEETS {
        writer.start_sheet(&format!("sheet{sheet}")).expect("start sheet");
        for row in 0..PARALLEL_ROWS {
            writer.start_row().expect("start row");
            writer.write_str(Some(&format!("s{sheet}-r{row}"))).expect("write");
            writer.end_row().expect("end row");
        }
        writer.end_sheet().expect("end sheet");
    }
    let bytes = writer.bytes().expect("bytes");
    Workbook::open_memory(&bytes, XL_FORMAT_AUTO, None).expect("open")
}

fn read_sheet(sheet: Sheet<'_>) -> Vec<String> {
    first_column(&mut sheet.rows().expect("cursor"))
}

#[test]
fn workbook_and_sheet_cross_threads_and_cursors_move_between_them() {
    fn assert_send_sync<T: Send + Sync>() {}
    fn assert_send<T: Send>() {}

    assert_send_sync::<Workbook>();
    assert_send_sync::<Sheet<'static>>();
    assert_send::<RowCursor<'static>>();
}

#[test]
fn every_sheet_read_on_its_own_thread_matches_a_sequential_read() {
    let workbook = many_sheets();
    let sheets = workbook.sheets().expect("sheets");
    let sequential: Vec<Vec<String>> = sheets.iter().map(|sheet| read_sheet(*sheet)).collect();
    assert_eq!(sequential[3][0], "s3-r0");
    assert_eq!(sequential[3].len(), PARALLEL_ROWS);

    let parallel: Vec<Vec<String>> = std::thread::scope(|scope| {
        let handles: Vec<_> = sheets
            .iter()
            .chain(sheets.iter())
            .map(|sheet| {
                let sheet = *sheet;
                scope.spawn(move || read_sheet(sheet))
            })
            .collect();
        handles.into_iter().map(|handle| handle.join().expect("no panic")).collect()
    });

    for (index, rows) in parallel.iter().enumerate() {
        assert_eq!(rows, &sequential[index % PARALLEL_SHEETS]);
    }
}
