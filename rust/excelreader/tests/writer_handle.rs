//! Round-trips through `WriterHandle`, the streaming counterpart to `xl_write_typed`: everything
//! written here is read back with the same crate's reader, same reasoning as `write_typed.rs`.

use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};
use excelreader::writer_handle::WriterHandle;
use excelreader::{Date, Time, Timestamp, XL_FORMAT_XLSX, XL_T_F64, XL_T_STRING};
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

/// Writes one header row and one data row through every `write_*` method, then a second row using
/// `None`/`write_null` for the blank-cell path.
fn write_full_rows(handle: &mut WriterHandle) {
    handle.start_sheet("Planilha1").expect("start_sheet must succeed");

    handle.start_row().expect("start_row must succeed for the header row");
    for header in [
        "texto",
        "inteiro",
        "numero",
        "ativo",
        "data",
        "hora",
        "instante",
        "opcional",
    ] {
        handle
            .write_str(Some(header))
            .expect("writing a header cell must succeed");
    }
    handle.end_row().expect("end_row must succeed for the header row");

    handle.start_row().expect("start_row must succeed for the data row");
    handle.write_str(Some("uma")).expect("write_str must succeed");
    handle.write_i64(Some(1)).expect("write_i64 must succeed");
    handle.write_f64(Some(0.5)).expect("write_f64 must succeed");
    handle.write_bool(Some(true)).expect("write_bool must succeed");
    handle
        .write_date(Some(Date::new(20454)))
        .expect("write_date must succeed");
    handle
        .write_time(Some(Time::new(3_600_000_000)))
        .expect("write_time must succeed");
    handle
        .write_timestamp(Some(Timestamp::new(1_767_225_600_000_000)))
        .expect("write_timestamp must succeed");
    handle
        .write_i64(Some(7))
        .expect("write_i64(Some) must succeed for the optional column");
    handle.end_row().expect("end_row must succeed for the data row");

    handle
        .start_row()
        .expect("start_row must succeed for the null-cell row");
    handle
        .write_str(None)
        .expect("write_str(None) must write a blank cell");
    handle
        .write_i64(None)
        .expect("write_i64(None) must write a blank cell");
    handle
        .write_null(XL_T_F64)
        .expect("write_null must write a blank cell directly");
    for _ in 0..5 {
        handle
            .write_null(XL_T_STRING)
            .expect("padding the null-cell row out must succeed");
    }
    handle.end_row().expect("end_row must succeed for the null-cell row");

    handle.end_sheet().expect("end_sheet must succeed");
}

#[test]
fn writer_handle_round_trips_through_open_file() {
    let path = temp_path("full.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8").to_string();
    {
        // Scoped so Drop closes and releases the handle - including the exclusive file lock
        // xl_open_write_handle takes - before Workbook::open reopens the same path below.
        let mut handle = WriterHandle::open_with(&target, XL_FORMAT_XLSX, None)
            .expect("WriterHandle::open_with must succeed");
        write_full_rows(&mut handle);
    }

    let mut workbook = Workbook::open(&target).expect("the file WriterHandle produced must open");
    let table = parse_sheet::<FullRow>(&mut workbook, 1).expect("the written file must parse back");
    assert_eq!(table.len(), 2);

    let first = table.get(0).expect("row 0 must exist");
    assert_eq!(first.texto, "uma");
    assert_eq!(first.inteiro, 1);
    assert_eq!(first.numero, 0.5);
    assert!(first.ativo);
    assert_eq!(first.data, Date::new(20454));
    assert_eq!(first.hora, Time::new(3_600_000_000));
    assert_eq!(first.opcional, Some(7));

    let second = table.get(1).expect("row 1 must exist");
    assert_eq!(second.texto, "");
    assert_eq!(second.opcional, None);

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[test]
fn writer_handle_to_memory_round_trips() {
    let mut handle =
        WriterHandle::open_memory(XL_FORMAT_XLSX, None).expect("WriterHandle::open_memory must succeed");
    write_full_rows(&mut handle);

    let bytes = handle.bytes().expect("bytes must succeed");
    assert!(!bytes.is_empty());

    let mut workbook =
        Workbook::open_memory(&bytes, XL_FORMAT_XLSX, None).expect("the bytes bytes() returned must open");
    assert_eq!(
        workbook.sheet_name().expect("name must read back"),
        "Planilha1"
    );
    let table = parse_sheet::<FullRow>(&mut workbook, 1).expect("the returned bytes must parse back");
    assert_eq!(table.len(), 2);
    assert_eq!(table.get(0).expect("row 0").texto, "uma");

    // bytes() must not have released the handle: a second call is still valid and returns the
    // same content.
    let bytes_again = handle.bytes().expect("a second bytes() call must still succeed");
    assert_eq!(bytes_again, bytes);
}

#[test]
fn writer_handle_bytes_rejects_a_file_backed_handle() {
    let path = temp_path("file_backed.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8").to_string();
    {
        let mut handle = WriterHandle::open_with(&target, XL_FORMAT_XLSX, None)
            .expect("WriterHandle::open_with must succeed");
        handle.start_sheet("S").expect("start_sheet must succeed");

        let error = handle
            .bytes()
            .expect_err("bytes() on a file-backed handle must fail");
        assert_eq!(error.code(), excelreader::XL_INVALID_ARGUMENT);
    }
    std::fs::remove_file(&path).ok();
}

#[test]
fn writer_handle_rejects_out_of_order_calls() {
    let path = temp_path("order.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8").to_string();
    {
        let mut handle = WriterHandle::open_with(&target, XL_FORMAT_XLSX, None)
            .expect("WriterHandle::open_with must succeed");

        assert!(
            handle.start_row().is_err(),
            "start_row before start_sheet must fail"
        );
        assert!(
            handle.write_i64(Some(1)).is_err(),
            "a cell write before start_row must fail"
        );
        assert!(handle.end_row().is_err(), "end_row without an open row must fail");
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
