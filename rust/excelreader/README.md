![ExcelReader logo](https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png)

# excelreader (Rust)

Read and write Excel/CSV workbooks via ExcelReader's native ABI: opening a workbook (from a path,
memory, your own random-access source or any `Read`), sheet navigation, schema inference,
schema-driven typed parse, schema-driven writing, row-by-row reads, parallel CSV aggregation, and
Arrow export. Every public item lives at the crate root.

## Usage

```toml
[dependencies]
excelreader = "7.0"
```

```rust,no_run
use excelreader::{ExcelMapper, Workbook, ReadOptions};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Name")]
    name: String,
    #[excel(name = "Count")]
    count: u32,
}

let workbook = Workbook::open("book.xlsx")?;
let table = workbook.sheet(0)?.parse::<Row>(&ReadOptions::new())?;
for row in table.iter() { /* ... */ }
# Ok::<(), excelreader::Error>(())
```

`Sheet::parse(header_row, degree_of_parallelism)` converts on the native side and never formats a
cell to text, so it is the fastest way to read a sheet whose columns you know. `header_row` is
1-based (`0` = no header); `degree_of_parallelism` is `0` for every processor, `1` sequential, `n`
up to `n` threads - XLSX, XLSB and CSV are split, and the table is the same whatever it is.

### Field types

| Field type | Column |
|---|---|
| `String` | `ColumnType::String` |
| `i8`..`i64`, `isize`, `u8`..`u64`, `usize` | `ColumnType::I64` |
| `f32`, `f64` | `ColumnType::F64` |
| `bool` | `ColumnType::Bool` |
| `Date` / `chrono::NaiveDate` | `ColumnType::Date` |
| `Time` / `chrono::NaiveTime` | `ColumnType::Time` |
| `Timestamp` / `chrono::NaiveDateTime` | `ColumnType::Timestamp` |

`Option<T>` of any of the above makes the field `None` when the cell is null; a non-`Option` field is
left at its `Default` instead.

Narrower integers convert through `TryFrom` and panic on a value that does not fit, rather than
wrapping silently. `Date`/`Time`/`Timestamp` are dependency-free newtypes over the exact wire
representation; enable the `chrono` feature to use `chrono`'s calendar types directly:

```toml
excelreader = { version = "7.0", features = ["chrono"] }
```

### Opening

`Workbook::open(path)` sniffs the format and uses every library default. Everything else goes
through `OpenOptions`, in the style of `std::fs::OpenOptions`: set what you need, then call `open`,
`open_bytes`, `open_source` or `open_reader`.

```rust,no_run
use excelreader::{Format, OpenOptions};

let options = OpenOptions::new().format(Format::Csv).csv_delimiter(b';' as i32);
let from_path = options.open("data.csv")?;
let from_memory = options.open_bytes(b"a;b\n1;2\n")?;
# Ok::<(), excelreader::Error>(())
```

Format sniffing (`Format::Auto`, the default) does not detect CSV - set `Format::Csv`. The
resource limits (`max_total_decompressed_bytes`, `max_cell_bytes`, `max_shared_string_bytes`,
`max_zip_entries`, `max_buffered_bytes`) protect against hostile files; lower them when opening
untrusted input.

### Sheets and schema inference

A workbook hands out sheets. A `Sheet` is the workbook and an index: it holds no native
resource, it is `Copy`, and every read on it opens its own cursor.

```rust,no_run
# use excelreader::Workbook;
let workbook = Workbook::open("report.xlsx")?;

for sheet in workbook.sheets()? {
    println!("{} {:?}", sheet.name()?, sheet.visibility()?);
}

// Guess a schema from the header row plus a sample of the data, before committing to one.
let sheet = workbook.sheet_by_name("Totals")?.expect("the workbook has a Totals sheet");
for column in sheet.infer_schema(1, 100, false)? {
    println!("{:?} -> {:?}", column.name, column.column_type);
}
# Ok::<(), excelreader::Error>(())
```

`infer_schema`'s last argument, `parse_text`, also types text cells (every CSV field, or numbers
stored as text): integers, decimals, `true`/`false` and ISO dates.

With no row type, `parse_with_schema` reads the sheet through that schema into a `Table`. Its
columns are the `Column`s `write_columns` takes, so the table writes back as it was read:

```rust,no_run
use excelreader::{write_columns, ColumnData, ReadOptions, Workbook, WriteOptions};

let workbook = Workbook::open("book.xlsx")?;
let sheet = workbook.sheet(0)?;
let schema = sheet.infer_schema(1, 100, false)?;
let table = sheet.parse_with_schema(&schema, &ReadOptions::new())?;
for column in table.columns() {
    if let ColumnData::I64(values) = column.data {
        println!("{:?}: {} integers", column.name, values.len());
    }
}
write_columns("copy.xlsx", &table.columns(), &WriteOptions::new())?;

// Or a batch at a time: TableBatches yields one Table per batch.
for batch in sheet.parse_batches_with_schema(&schema, &ReadOptions::new())? {
    println!("{} rows", batch?.len());
}
# Ok::<(), excelreader::Error>(())
```

A column matched by index has no name, and writing needs every column named or none. The schema is
a guess over a sample: a column inferred as `I64` that holds a fraction further down fails the read,
unless the column is `nullable`.

`Workbook` is `Send + Sync`, so sheets can be read in parallel over one open workbook, which
loads the shared-string table once:

```rust,no_run
# use excelreader::Workbook;
# let workbook = Workbook::open("report.xlsx")?;
let tables = std::thread::scope(|scope| {
    let handles: Vec<_> = workbook
        .sheets()?
        .into_iter()
        .map(|sheet| scope.spawn(move || sheet.read_all()))
        .collect();
    handles.into_iter().map(|handle| handle.join().expect("no panic")).collect::<Result<Vec<_>, _>>()
})?;
# Ok::<(), excelreader::Error>(())
```

### Encrypted workbooks

`OpenOptions::password` unlocks a password-protected OOXML workbook (.xlsx/.xlsb/.xlsm):

```rust,no_run
use excelreader::{ErrorKind, OpenOptions};

match OpenOptions::new().password("hunter2").open("protected.xlsx") {
    Ok(workbook) => { /* ... */ }
    Err(e) if e.kind() == ErrorKind::PasswordIncorrect => { /* ask again */ }
    Err(e) if e.kind() == ErrorKind::PasswordRequired => { /* no password was supplied */ }
    Err(other) => { /* unsupported scheme or a corrupt file; not worth retrying */ }
}
```

Every `Error` carries an `ErrorKind` to branch on and the native detail in `message()`. Format
sniffing does not see through the CFB wrapper an encrypted file is stored in, so keep the default
`Format::Auto` rather than setting `Format::Xlsx`, or the file is read as a plain ZIP and fails
before the password is ever checked.

Writing an encrypted workbook is a second step: write the plaintext package with `write_columns`/
`write_sheet`, then wrap it with `encrypt_package`, which produces an agile-encrypted (ECMA-376 4.4)
CFB container - AES-256-CBC, SHA-512, 100,000 spin iterations, with a `dataIntegrity` HMAC:

```rust,no_run
use excelreader::{encrypt_package, write_columns, WriteOptions};
# use excelreader::{Column, ColumnData};
# let ids = [1i64, 2, 3];
# let columns = [Column::new(Some("id"), ColumnData::I64(&ids))];

write_columns("plain.xlsx", &columns, &WriteOptions::new())?;
encrypt_package("plain.xlsx", "secret.xlsx", "hunter2")?;
# Ok::<(), excelreader::Error>(())
```

`package_path` is read twice (it is not disposed or removed), so it must already be a finished file.
Encryption parameters are fixed at Excel's own defaults - there are no options - and only XLSX/XLSB
packages can be encrypted, matching what `OpenOptions::password` can decrypt.

### Reading from your own bytes

Implement [`Source`] to read from anything you can read at an offset. The library calls `read_at`
from several threads, fetching 4 MiB blocks it caches; every sheet can still be read in parallel.

```rust,ignore
struct HttpSource { client: reqwest::blocking::Client, url: String, size: u64 }

impl excelreader::Source for HttpSource {
    fn size(&self) -> std::io::Result<u64> {
        Ok(self.size)
    }

    fn read_at(&self, offset: u64, buf: &mut [u8]) -> std::io::Result<usize> {
        let range = format!("bytes={}-{}", offset, offset + buf.len() as u64 - 1);
        let bytes = self.client.get(&self.url).header("Range", range).send()
            .and_then(|r| r.error_for_status())
            .and_then(|r| r.bytes())
            .map_err(std::io::Error::other)?;
        buf[..bytes.len()].copy_from_slice(&bytes);
        Ok(bytes.len())
    }
}

let workbook = OpenOptions::new().open_source(source)?;
```

Anything that implements `Read` works through `open_reader`. A CSV is read as it arrives; an XLSX,
XLSB or XLS is read whole first.

```rust,no_run
# use excelreader::{Format, OpenOptions};
let workbook = OpenOptions::new().format(Format::Csv).open_reader(std::io::stdin())?;
# Ok::<(), excelreader::Error>(())
```

An `io::Error` from your source, or a panic in it, comes back as the `Error` of the call that
needed the bytes.

### Arrow export (`arrow` feature)

```toml
excelreader = { version = "7.0", features = ["arrow"] }
```

```rust,no_run
# #[cfg(feature = "arrow")]
# fn main() -> Result<(), excelreader::Error> {
use excelreader::{Workbook, ReadOptions};
# #[derive(Default, excelreader::ExcelMapper)]
# struct Row { #[excel(name = "Name")] name: String }

let workbook = Workbook::open("book.xlsx")?;
let batch = workbook.sheet(0)?.to_arrow::<Row>(&ReadOptions::new())?;
println!("{} rows x {} columns", batch.num_rows(), batch.num_columns());
# Ok(())
# }
# #[cfg(not(feature = "arrow"))]
# fn main() {}
```

Off by default — arrow-rs is a large dependency and the typed-parse path needs none of it.

Batched: `Sheet::to_arrow_stream` is `parse_arrow` delivered a batch at a time, returning
`ArrowBatches<'_>`, an `Iterator<Item = Result<RecordBatch, ArrowError>>` that also implements
`RecordBatchReader`:

```rust,no_run
# #[cfg(feature = "arrow")]
# fn main() -> Result<(), excelreader::Error> {
use excelreader::{Workbook, ReadOptions};
# #[derive(Default, excelreader::ExcelMapper)]
# struct Row { #[excel(name = "Name")] name: String }

let workbook = Workbook::open("book.xlsx")?;
let mut chunks = workbook.sheet(0)?.to_arrow_stream::<Row>(&ReadOptions::new())?;
while let Some(batch) = chunks.next() {
    let Ok(batch) = batch else { break }; // ArrowBatches does not fuse, unlike TypedBatches
    println!("{} rows", batch.num_rows());
}
# Ok(())
# }
# #[cfg(not(feature = "arrow"))]
# fn main() {}
```

Same `batch_size` semantics as `parse_batches` below. Unlike `TypedBatches`,
`ArrowBatches` does not fuse after an error — it follows arrow-rs's own semantics — so break on the
first `Err` rather than continuing the loop.

With no row type, `to_arrow_with_schema` and `to_arrow_stream_with_schema` take the columns at run
time instead, such as the ones `infer_schema` returned. A column with a name is matched by it, one
without by its index:

```rust,no_run
# #[cfg(feature = "arrow")]
# fn main() -> Result<(), excelreader::Error> {
use excelreader::{Workbook, ReadOptions};

let workbook = Workbook::open("book.xlsx")?;
let sheet = workbook.sheet(0)?;
let schema = sheet.infer_schema(1, 100, false)?;
let batch = sheet.to_arrow_with_schema(&schema, &ReadOptions::new())?;
# Ok(())
# }
# #[cfg(not(feature = "arrow"))]
# fn main() {}
```

The schema is a guess over a sample: a column inferred as `I64` that holds a fraction further down
fails the read, unless the column is `nullable`.

## Writing

`#[derive(ExcelMapper)]` generates both halves, so the same struct reads and writes - the field
types in the table above apply unchanged:

```rust,no_run
use excelreader::{write_sheet, WriteOptions};
# #[derive(Default, excelreader::ExcelMapper)]
# struct Row { #[excel(name = "Name")] name: String }
# let rows = vec![Row::default()];

write_sheet("out.xlsx", &rows, &WriteOptions::new())?;
# Ok::<(), excelreader::Error>(())
```

`Option<T>` fields become an LSB-first validity bitmap: `None` writes a blank cell rather than a
zero. Only the primary `#[excel(name = "...")]` reaches the header - the `alias` list exists to
resolve a header on the way *in*, and the ABI rejects a write column carrying more than one name.

For buffers that are already columnar, `write_columns` borrows them and copies nothing. The
lifetimes on `Column<'a>` are what turn the ABI's borrow contract into something the compiler
checks:

```rust,no_run
use excelreader::{write_columns, Column, ColumnData, WriteOptions};

let ids = [1i64, 2, 3];
let columns = [Column::new(Some("id"), ColumnData::I64(&ids))];
write_columns("out.xlsx", &columns, &WriteOptions::new())?;
# Ok::<(), excelreader::Error>(())
```

`validity` is checked against the row count before the call: the ABI takes the bitmap without a
length and reads `(rows + 7) / 8` bytes on trust, so a short slice would be a buffer overrun rather
than a wrong answer.

`WriteOptions` sets the format, the sheet name, the CSV dialect, and the XLS/XLSB and XLSX/XLSB
toggles. Its format defaults to `Format::Auto`, inferred from the path's extension; an extension it
does not recognize is an error - a file being created has no signature bytes to sniff.

For XLSX and XLSB, `degree_of_parallelism` formats and compresses batches of `batch_rows` rows on
several threads: `0` uses every processor, `1` writes serially, and leaving it unset is serial.
Unset `batch_rows` keeps the library default of 8,192. With shared strings on, two parallel runs
give the same content in files that differ byte for byte. XLS, CSV and `WriterHandle` ignore both.

```rust,no_run
use excelreader::{write_columns, Column, ColumnData, WriteOptions};
# let ids = [1i64, 2, 3];
# let columns = [Column::new(Some("id"), ColumnData::I64(&ids))];

write_columns("big.xlsx", &columns, &WriteOptions::new().degree_of_parallelism(0))?;
# Ok::<(), excelreader::Error>(())
```

Each write has a path form and an `io::Write` form: `write_columns_to`/`write_sheet_to` build the
workbook in memory and write it to the writer in one `write_all`. A writer has no extension, so
these need an explicit format:

```rust,no_run
use excelreader::{write_sheet_to, Format, WriteOptions};
# #[derive(Default, excelreader::ExcelMapper)]
# struct Row { #[excel(name = "Name")] name: String }
# let rows = vec![Row::default()];

let mut bytes = Vec::new();
write_sheet_to(&mut bytes, &rows, &WriteOptions::new().format(Format::Xlsx))?;
# Ok::<(), excelreader::Error>(())
```

`write_sheet` walks the slice once and appends each field to its own buffer, monomorphized per
field. That transpose is the only copy it makes; `write_columns` pays nothing.

### Streaming writes

`write_sheet`/`write_columns` build the whole table in memory first. `WriterHandle` is the
row-by-row alternative, writing directly as each call arrives instead:

```rust,no_run
use excelreader::{WriteOptions, WriterHandle};

let mut handle = WriterHandle::create("out.xlsx", &WriteOptions::new())?;
handle.start_sheet("Summary")?;
handle.start_row()?;
handle.write_str(Some("Name"))?;
handle.write_i64(Some(42))?;
handle.end_row()?;
handle.end_sheet()?;
handle.finish()?;
# Ok::<(), excelreader::Error>(())
```

Call order mirrors the C ABI's `xl_writer_handle`: `create` (a path) or `new` (any `io::Write`),
then per sheet `start_sheet..end_sheet`, each containing `start_row..end_row` with one `write_*`
call per cell (`None` writes a blank cell), left to right, then `finish`. A call out of order
returns `Err` rather than corrupting output. A handle from `new` writes the whole workbook to its
writer once, in `finish`. Dropping a handle without finishing it closes it on a best-effort basis,
discarding any error, and writes nothing to a writer.

## Reading rows one at a time

`Sheet::rows` returns a cursor over that sheet. Each row borrows a buffer the cursor
reuses, so iterating a sheet allocates nothing per row:

```rust,no_run
use excelreader::Workbook;

let workbook = Workbook::open("book.xlsx")?;
let mut cursor = workbook.sheet(0)?.rows()?;
while let Some(row) = cursor.next_row() {
    let row = row?;
    for cell in row.iter() {
        print!("{}\t", cell.as_str()?);
    }
    println!();
}
# Ok::<(), excelreader::Error>(())
```

Because each row borrows the cursor's buffer, only one row is alive at a time — the borrow checker
enforces it. To hold every row at once, use `read_all`, which reads the whole sheet in one native
call into one buffer and keeps it alive until dropped:

```rust,no_run
# use excelreader::Workbook;
let workbook = Workbook::open("book.xlsx")?;
let rows = workbook.sheet(0)?.read_all()?;
for row in rows.iter() {
    println!("{} cells", row.len());
}
# Ok::<(), excelreader::Error>(())
```

### Chunked reads

For a sheet too large to hold in memory at once, `Sheet::parse_batches` is `Sheet::parse` delivered
a batch at a time:

```rust,no_run
use excelreader::{Workbook, ReadOptions};
# #[derive(Default, excelreader::ExcelMapper)]
# struct Row { #[excel(name = "Name")] name: String }

let workbook = Workbook::open("book.xlsx")?;
let chunks = workbook.sheet(0)?.parse_batches::<Row>(&ReadOptions::new())?;
for batch in chunks {
    let batch = batch?;
    for row in batch.iter() { /* ... */ }
}
# Ok::<(), excelreader::Error>(())
```

`parse_batches` returns `TypedBatches<'_, Row>`, an `Iterator<Item = Result<TableView<Row>, Error>>`
that closes the native reader on drop. `batch_size` is rows per batch — `0` means one unbounded
batch, identical to `parse`; negative is an error. Chunked reads, of either kind (typed or
Arrow — see [Arrow export](#arrow-export-arrow-feature) above), are independent: several can be live
on one workbook, and other reads do not disturb them. `TypedBatches` ends after yielding an `Err`, so
a `for` loop stops there rather than spinning on it.

## Parallel CSV aggregation

`aggregate_csv` folds a CSV file or buffer into one accumulator across several threads, without
building a table. The CSV dialect and cell limit come from `OpenOptions`:

```rust,no_run
use excelreader::{aggregate_csv, CsvAccumulator, OpenOptions, RowRef};

struct Count(u64);
impl CsvAccumulator for Count {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        self.0 += 1;
        Ok(())
    }
    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.0 += other.0;
        Ok(())
    }
}

let options = OpenOptions::new().csv_delimiter(b';' as i32);
let rows = aggregate_csv("big.csv", &options, 1, 0, || Count(0))?;
# Ok::<(), excelreader::Error>(())
```

The input is a path (`&str`, `&Path`, `&PathBuf`) or bytes (`&[u8]`); then `header_row` and
`degree_of_parallelism`, as in `Sheet::parse`. A nonzero code from a callback comes back as
`ErrorKind::Callback(code)`; a panic in one is resumed on the calling thread.

## Callback reads

`Sheet::for_each_row` calls a closure once per row on the calling thread, with cells borrowed from
the library's buffers for that call instead of copied into a cursor:

```rust,no_run
use excelreader::Workbook;

let workbook = Workbook::open("report.xlsx")?;
let mut total = 0.0;
workbook.sheet(0)?.for_each_row(|row| {
    if let Some(amount) = row.get(1).and_then(|cell| cell.number()) {
        total += amount;
    }
    Ok(())
})?;
# Ok::<(), excelreader::Error>(())
```

`Sheet::aggregate` folds any sheet, in any format, into one accumulator across several threads.
New code should use it rather than `aggregate_csv`:

```rust,no_run
use excelreader::{Accumulator, RowRef, Workbook};

struct Count(u64);
impl Accumulator for Count {
    fn accumulate(&mut self, _row: RowRef<'_>) -> Result<(), i32> {
        self.0 += 1;
        Ok(())
    }
    fn combine(&mut self, other: &mut Self) -> Result<(), i32> {
        self.0 += other.0;
        Ok(())
    }
}

let workbook = Workbook::open("big.xlsx")?;
let rows = workbook.sheet(0)?.aggregate(1, 0, || Count(0))?;
# Ok::<(), excelreader::Error>(())
```

A cell in these callbacks carries text, a number, or both. `as_bytes`/`as_str` return only the
stored text, so they are empty for a cell that carries only a number; `text()` formats that number
itself (shortest round-trip form) and `number()` returns it directly, or parses the text when there
is no number. A positive `Err(code)` comes back as `ErrorKind::Callback(code)`, and a panic is
resumed on the calling thread. Calling other methods on the workbook from inside the callback is
allowed.

`Sheet::aggregate` seeds an accumulator per chunk of sheet data, not per thread, and may drop a
chunk's accumulator without combining it. XLS, table views and CSV opened from a `Source` or stream
read sequentially.

## Bounds and panics

`TableView::get` returns `Option<T>` and is `None` outside `0..len()`. The accessors the derive's
generated code reads columns through panic on an out-of-range row, a column type that does not
match the binding, or a string the native library returned as non-UTF-8 - each is a contract
violation rather than recoverable input.

## Build notes

`build.rs` downloads the native `ExcelReader.Native` binary matching your target from the crate's
matching GitHub Release. Set `EXCELREADER_NATIVE_LIB_DIR` to a directory containing a locally-built
copy instead (e.g. from `dotnet publish ../../src/ExcelReader.Native -r win-x64`) to skip that
download - useful when building from source before a release exists yet.

Every constructor first checks the loaded library's ABI revision against the one this crate was
compiled against, and fails with an explanatory error rather than reading native memory through a
layout that may have changed.

## Benchmarks

Criterion suite in `benches/`. Measured on Windows 10 (22H2), 16 logical CPUs @ 3.39 GHz,
rustc 1.97.1 (Release), Criterion 0.8, 100 samples per benchmark (medians shown). `write_bench`
takes 20 samples instead — each of its iterations writes a whole 65,535-row file.

`benches/parse_bench.rs` - `open`/`Sheet::parse`/`infer_schema`, same methodology as the C++ suite:

| Benchmark | RealExcel.xlsb (100 rows) | 65K_Records_Data.xlsb (65,535 rows) |
|---|---:|---:|
| `open` | 85.8 µs | 99.4 µs |
| `parse_sheet` (2 or 6 bound columns) | 54.7 µs | 24.0 ms |
| typed chunks of 1,000 / 10,000 rows | - | 22.8 ms / 22.9 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 83.9 µs | 571.0 µs |

`open` stays nearly flat across the 655x row-count jump (+16%) - XLSB's header carries its
dimensions/index, so opening costs metadata, not row data. `parse_sheet` scales linearly with
rows × columns. Reading the same sheet in typed chunks costs no more than one whole-sheet
`parse_sheet`, while holding only one chunk at a time. `infer_schema` scales with its sample size,
not the file's total row count.

`benches/compare_bench.rs` compares against [calamine](https://github.com/tafia/calamine) reading
`65K_Records_Data.{xlsx,xlsb}` in full (all 14 columns, 65,535 rows). Both sides decode every cell
into an owned value (`String` for text) and fold it into one accumulator, so neither side gets a
zero-copy advantage the other can't take:

| Format | ExcelReader (`Sheet::parse::<FullRow>`) | calamine (`worksheet_range` + `Data` match) |
|---|---:|---:|
| XLSX | 72.3 ms | 271.2 ms |
| XLSB | 46.3 ms | 82.9 ms |

ExcelReader is ~3.8x faster than calamine for XLSX and ~1.8x faster for XLSB on this workload -
calamine is a fast, well-optimized reader in its own right, so the gap is real but not the order
of magnitude seen against slower libraries.

`benches/write_bench.rs` measures the two write layers and
[rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) writing the same 7 columns × 65,535
rows to `.xlsx`, all three starting from the same in-memory `Vec<Row>`:

| Benchmark | Median |
|---|---:|
| `columns` (`write_columns`, pre-transposed) | 33.8 ms |
| `sheet` (`write_sheet`, from `Vec<Row>`) | 40.7 ms |
| `rust_xlsxwriter` (cell-at-a-time) | 322.0 ms |

`sheet` is the matched-work number — it starts from the same shape `rust_xlsxwriter` is handed and
pays the row-to-column transpose itself — and is ~7.9x faster. `columns` is a ceiling no
cell-at-a-time API can reach, since it is handed buffers that are already columnar; read it only
against `sheet`, as the cost of having row-shaped data in the first place. That cost turns out to be
about 21%: not free, but a minority of the cost of producing the file.

Two caveats, both running against the headline number rather than for it. ExcelReader does slightly
*more* work here: it attaches a number format to the date column so Excel shows a date, and writes a
header row, while the `rust_xlsxwriter` case writes that column as a bare serial number and no
header. And `rust_xlsxwriter` carries formatting and formula support this library does not expose at
all, so its number reflects a different feature set, not only a slower path.

Run locally:

```bash
cd rust
EXCELREADER_NATIVE_LIB_DIR=/path/to/native/lib/dir cargo bench -p excelreader
```

`EXCELREADER_NATIVE_LIB_DIR` should point at a directory containing a locally-built
`ExcelReader.Native.{dll,so,dylib}` - see [Build notes](#build-notes) above. Pass
`--bench parse_bench`, `--bench compare_bench` or `--bench write_bench` to run one suite only.
