use excelreader::workbook::Workbook;
use excelreader::writer_handle::WriterHandle;
use excelreader::{OpenOptions, RowCursor, Source, XL_FORMAT_AUTO, XL_FORMAT_CSV, XL_FORMAT_XLSX};
use std::io::{self, Read};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;

const SHEETS: usize = 4;
const ROWS: usize = 300;

fn workbook_bytes() -> Vec<u8> {
    let mut writer = WriterHandle::open_memory(XL_FORMAT_XLSX, None).expect("open writer");
    for sheet in 0..SHEETS {
        writer.start_sheet(&format!("sheet{sheet}")).expect("start sheet");
        for row in 0..ROWS {
            writer.start_row().expect("start row");
            writer.write_str(Some(&format!("s{sheet}-r{row}"))).expect("write");
            writer.end_row().expect("end row");
        }
        writer.end_sheet().expect("end sheet");
    }
    writer.bytes().expect("bytes").to_vec()
}

fn first_column(cursor: &mut RowCursor<'_>) -> Vec<String> {
    let mut values = Vec::new();
    while let Some(row) = cursor.next_row() {
        let row = row.expect("a row");
        values.push(row.get(0).expect("a cell").as_str().expect("utf-8").to_string());
    }
    values
}

#[derive(Default, Clone)]
struct Counters {
    reads: Arc<AtomicUsize>,
    drops: Arc<AtomicUsize>,
}

struct MemorySource {
    bytes: Vec<u8>,
    counters: Counters,
}

impl Source for MemorySource {
    fn size(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, buf: &mut [u8]) -> io::Result<usize> {
        self.counters.reads.fetch_add(1, Ordering::SeqCst);
        let start = (offset as usize).min(self.bytes.len());
        let count = buf.len().min(self.bytes.len() - start);
        buf[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }
}

impl Drop for MemorySource {
    fn drop(&mut self) {
        self.counters.drops.fetch_add(1, Ordering::SeqCst);
    }
}

struct FailingSource;

impl Source for FailingSource {
    fn size(&self) -> io::Result<u64> {
        Ok(1024)
    }

    fn read_at(&self, _offset: u64, _buf: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::other("disk on fire"))
    }
}

struct PanickingSource;

impl Source for PanickingSource {
    fn size(&self) -> io::Result<u64> {
        Ok(1024)
    }

    fn read_at(&self, _offset: u64, _buf: &mut [u8]) -> io::Result<usize> {
        panic!("boom");
    }
}

struct OnlyRead(io::Cursor<Vec<u8>>);

impl Read for OnlyRead {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        let limit = buf.len().min(7);
        self.0.read(&mut buf[..limit])
    }
}

#[test]
fn every_sheet_of_a_source_reads_in_parallel_like_memory() {
    let bytes = workbook_bytes();
    let memory = Workbook::open_memory(&bytes, XL_FORMAT_AUTO, None).expect("open memory");
    let expected: Vec<Vec<String>> = memory
        .sheets()
        .expect("sheets")
        .iter()
        .map(|sheet| first_column(&mut sheet.rows().expect("cursor")))
        .collect();

    let counters = Counters::default();
    let workbook = Workbook::open_source(MemorySource { bytes, counters: counters.clone() }, XL_FORMAT_AUTO, None)
        .expect("open source");
    let sheets = workbook.sheets().expect("sheets");
    let parallel: Vec<Vec<String>> = std::thread::scope(|scope| {
        let handles: Vec<_> = sheets
            .iter()
            .chain(sheets.iter())
            .map(|sheet| {
                let sheet = *sheet;
                scope.spawn(move || first_column(&mut sheet.rows().expect("cursor")))
            })
            .collect();
        handles.into_iter().map(|handle| handle.join().expect("no panic")).collect()
    });
    for (index, rows) in parallel.iter().enumerate() {
        assert_eq!(rows, &expected[index % SHEETS]);
    }

    drop(workbook);
    assert_eq!(counters.drops.load(Ordering::SeqCst), 1);
}

#[test]
fn an_io_error_reaches_the_caller_with_its_message() {
    let error = Workbook::open_source(FailingSource, XL_FORMAT_XLSX, None).expect_err("must fail");
    assert!(error.message().contains("disk on fire"), "{}", error.message());
}

#[test]
fn a_panicking_source_becomes_an_error() {
    let error = Workbook::open_source(PanickingSource, XL_FORMAT_XLSX, None).expect_err("must fail");
    assert!(error.message().contains("source panicked"), "{}", error.message());
}

#[test]
fn the_default_cache_fetches_a_small_workbook_once() {
    let bytes = workbook_bytes();
    let cached = Counters::default();
    {
        let workbook = Workbook::open_source(MemorySource { bytes: bytes.clone(), counters: cached.clone() }, XL_FORMAT_XLSX, None)
            .expect("open");
        first_column(&mut workbook.sheet(0).expect("sheet").rows().expect("cursor"));
    }
    assert_eq!(cached.reads.load(Ordering::SeqCst), 1);

    let uncached = Counters::default();
    let options = OpenOptions::new().source_block_size(-1);
    {
        let workbook = Workbook::open_source(MemorySource { bytes, counters: uncached.clone() }, XL_FORMAT_XLSX, Some(&options))
            .expect("open");
        first_column(&mut workbook.sheet(0).expect("sheet").rows().expect("cursor"));
    }
    assert!(uncached.reads.load(Ordering::SeqCst) > 1);
}

#[test]
fn a_reader_serves_a_csv_and_an_xlsx() {
    let csv = OnlyRead(io::Cursor::new(b"name,qty\nwidget,7\ngadget,9\n".to_vec()));
    let workbook = Workbook::open_reader(csv, XL_FORMAT_CSV, None).expect("open csv");
    let names = first_column(&mut workbook.sheet(0).expect("sheet").rows().expect("cursor"));
    assert_eq!(names, vec!["name", "widget", "gadget"]);

    let xlsx = OnlyRead(io::Cursor::new(workbook_bytes()));
    let workbook = Workbook::open_reader(xlsx, XL_FORMAT_AUTO, None).expect("open xlsx");
    let rows = first_column(&mut workbook.sheet(3).expect("sheet").rows().expect("cursor"));
    assert_eq!(rows.len(), ROWS);
    assert_eq!(rows[0], "s3-r0");
}
