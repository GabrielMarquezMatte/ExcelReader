# excelreader (Rust)

Read and write Excel/CSV workbooks via ExcelReader's native ABI: opening a workbook (from a path or
memory, with the full open-options surface), sheet navigation, schema inference, schema-driven typed
parse, schema-driven writing, row-by-row decoded reads, and Arrow export.

## Usage

```toml
[dependencies]
excelreader = "2.1"
```

```rust
use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Name")]
    name: String,
    #[excel(name = "Count")]
    count: u32,
}

let workbook = Workbook::open("book.xlsx")?;
let table = parse_sheet::<Row>(workbook.sheet(0)?, 1)?;
for row in table.iter() { /* ... */ }
```

`parse_sheet` takes a `Sheet`, which `Workbook::sheet` hands out (see below).

### Field types

| Field type | Column |
|---|---|
| `String` | `XL_T_STRING` |
| `i8`..`i64`, `isize`, `u8`..`u64`, `usize` | `XL_T_I64` |
| `f32`, `f64` | `XL_T_F64` |
| `bool` | `XL_T_BOOL` |
| `Date` / `chrono::NaiveDate` | `XL_T_DATE` |
| `Time` / `chrono::NaiveTime` | `XL_T_TIME` |
| `Timestamp` / `chrono::NaiveDateTime` | `XL_T_TIMESTAMP` |

`Option<T>` of any of the above makes the field `None` when the cell is null; a non-`Option` field is
left at its `Default` instead.

Narrower integers convert through `TryFrom` and panic on a value that does not fit, rather than
wrapping silently. `Date`/`Time`/`Timestamp` are dependency-free newtypes over the exact wire
representation; enable the `chrono` feature to use `chrono`'s calendar types directly:

```toml
excelreader = { version = "2.1", features = ["chrono"] }
```

### Sheets and schema inference

```rust
let workbook = Workbook::open("book.xlsx")?;
for (index, name) in workbook.sheet_names()?.iter().enumerate() {
    println!("{index}: {name}");
}

// Guess a schema from the header row plus a sample of the data, before committing to one.
for column in workbook.sheet(1)?.infer_schema(1, 100)? {
    println!("{:?} -> type {}", column.name, column.column_type);
}
```

`infer_schema_parse_text(header_row, sample_size)` also types text cells (every CSV field, or
numbers stored as text): integers, decimals, `true`/`false` and ISO dates. `parse_sheet_parallel`
and `arrow::parse_arrow_parallel` take a `degree_of_parallelism` (`0` = every processor, `1` =
sequential); only CSV is split, and the result equals `parse_sheet`'s.

A workbook hands out sheets. A `Sheet` is the workbook and an index: it holds no native
resource, it is `Copy`, and every read on it opens its own cursor.

```rust
let workbook = Workbook::open("report.xlsx")?;

for sheet in workbook.sheets()? {
    println!("{} {:?}", sheet.name()?, sheet.visibility()?);
}

let sheet = workbook.sheet_by_name("Totals")?.expect("the workbook has a Totals sheet");
let mut rows = sheet.rows()?;
while let Some(row) = rows.next_row() {
    let row = row?;
    // ...
}
```

`Workbook` is `Send + Sync`, so sheets can be read in parallel over one open workbook, which
loads the shared-string table once:

```rust
let tables = std::thread::scope(|scope| {
    let handles: Vec<_> = workbook
        .sheets()?
        .into_iter()
        .map(|sheet| scope.spawn(move || sheet.read_all_blob()))
        .collect();
    handles.into_iter().map(|handle| handle.join().expect("no panic")).collect::<Result<Vec<_>, _>>()
})?;
```

`Workbook::open_with` takes an explicit format and `OpenOptions`; `Workbook::open_memory` reads from
a byte slice. Note that format sniffing does not detect CSV - pass `XL_FORMAT_CSV` explicitly.

### Encrypted workbooks

`OpenOptions::password` unlocks a password-protected OOXML workbook (.xlsx/.xlsb/.xlsm) through
`Workbook::open_with`/`open_memory`:

```rust
use excelreader::workbook::Workbook;
use excelreader::{Error, OpenOptions, XL_FORMAT_AUTO};

let options = OpenOptions::new().password("hunter2");
match Workbook::open_with("protected.xlsx", XL_FORMAT_AUTO, Some(&options)) {
    Ok(workbook) => { /* ... */ }
    Err(Error::PasswordIncorrect { .. }) => { /* ask again */ }
    Err(Error::PasswordRequired { .. }) => { /* no password was supplied */ }
    Err(other) => { /* Error::Native - unsupported scheme or a corrupt file; not worth retrying */ }
}
```

`Error::PasswordRequired`/`Error::PasswordIncorrect` are dedicated variants rather than the generic
error carrying `XL_STATUS_PASSWORD_REQUIRED`/`XL_STATUS_PASSWORD_INCORRECT`, so callers can
`match`/`matches!` on "needs a password" without hardcoding the status codes. Format sniffing does not
see through the CFB wrapper an encrypted file is stored in - pass `XL_FORMAT_AUTO` explicitly (as
above) rather than `XL_FORMAT_XLSX`, or the file is read as a plain ZIP and fails before the password
is ever checked.

Writing an encrypted workbook is a second step: write the plaintext package with `write_columns`/
`write_sheet`, then wrap it with `encrypt_package`, which produces an agile-encrypted (ECMA-376 4.4)
CFB container - AES-256-CBC, SHA-512, 100,000 spin iterations, with a `dataIntegrity` HMAC:

```rust
use excelreader::writer::{encrypt_package, write_columns};
use excelreader::XL_FORMAT_XLSX;

write_columns("plain.xlsx", XL_FORMAT_XLSX, &columns, None)?;
encrypt_package("plain.xlsx", "secret.xlsx", "hunter2")?;
```

`package_path` is read twice (it is not disposed or removed), so it must already be a finished file.
Encryption parameters are fixed at Excel's own defaults - there are no options - and only XLSX/XLSB
packages can be encrypted, matching what `Workbook::open_with`/`open_memory` can decrypt.

### Reading from your own bytes

Implement [`Source`] to read from anything you can read at an offset. The library calls `read_at`
from several threads, fetching 4 MiB blocks it caches; every sheet can still be read in parallel.

```rust
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

let workbook = Workbook::open_source(source, XL_FORMAT_AUTO, None)?;
```

Anything that implements `Read` works through `open_reader`. A CSV is read as it arrives; an XLSX,
XLSB or XLS is read whole first.

```rust
let workbook = Workbook::open_reader(std::io::stdin(), XL_FORMAT_CSV, None)?;
```

An `io::Error` from your source, or a panic in it, comes back as the `Error` of the call that
needed the bytes.

### Arrow export (`arrow` feature)

```toml
excelreader = { version = "2.1", features = ["arrow"] }
```

```rust
use excelreader::arrow::parse_arrow;
use excelreader::workbook::Workbook;

let workbook = Workbook::open("book.xlsx")?;
let batch = parse_arrow::<Row>(workbook.sheet(0)?, 1)?;
println!("{} rows x {} columns", batch.num_rows(), batch.num_columns());
```

Off by default — arrow-rs is a large dependency and the typed-parse path needs none of it.

Batched: `parse_arrow_stream` is `parse_arrow` delivered a batch at a time, returning
`ArrowChunks<'_>`, an `Iterator<Item = Result<RecordBatch, ArrowError>>` that also implements
`RecordBatchReader`:

```rust
use excelreader::arrow::parse_arrow_stream;
use excelreader::workbook::Workbook;

let workbook = Workbook::open("book.xlsx")?;
let mut chunks = parse_arrow_stream::<Row>(workbook.sheet(0)?, 1, 10_000)?;
while let Some(batch) = chunks.next() {
    let Ok(batch) = batch else { break }; // ArrowChunks does not fuse, unlike TypedChunks
    println!("{} rows", batch.num_rows());
}
# Ok::<(), excelreader::Error>(())
```

Same `batch_size` semantics as `typed_chunks` below. Unlike `TypedChunks`,
`ArrowChunks` does not fuse after an error — it follows arrow-rs's own semantics — so break on the
first `Err` rather than continuing the loop.

## Writing

`#[derive(ExcelMapper)]` generates both halves, so the same struct reads and writes - the field
types in the table above apply unchanged:

```rust
use excelreader::writer::write_sheet;
use excelreader::XL_FORMAT_XLSX;

write_sheet("out.xlsx", XL_FORMAT_XLSX, &rows, None)?;
```

`Option<T>` fields become an LSB-first validity bitmap: `None` writes a blank cell rather than a
zero. Only the primary `#[excel(name = "...")]` reaches the header - the `alias` list exists to
resolve a header on the way *in*, and the ABI rejects a write column carrying more than one name.

For buffers that are already columnar, `write_columns` borrows them and copies nothing. The
lifetimes on `Column<'a>` are what turn the ABI's borrow contract into something the compiler
checks:

```rust
use excelreader::writer::{write_columns, Column, ColumnData};
use excelreader::XL_FORMAT_XLSX;

let ids = [1i64, 2, 3];
let columns = [Column {
    name: Some("id"),
    data: ColumnData::I64(&ids),
    validity: None,
}];
write_columns("out.xlsx", XL_FORMAT_XLSX, &columns, None)?;
```

`validity` is checked against the row count before the call: the ABI takes the bitmap without a
length and reads `(rows + 7) / 8` bytes on trust, so a short slice would be a buffer overrun rather
than a wrong answer.

`WriteOptions` sets the sheet name, the CSV dialect, and the XLS/XLSB and XLSX/XLSB toggles.
`format_from_path` infers the format from an extension; it returns `XL_FORMAT_AUTO` for anything it
does not recognize, which the write then rejects - a file being created has no signature bytes to
sniff, so there is nothing to fall back on.

`write_sheet` walks the slice once and appends each field to its own buffer, monomorphized per
field. That transpose is the only copy it makes; `write_columns` pays nothing.

### Streaming writes

`write_sheet`/`write_columns` build the whole table in memory first. `writer_handle::WriterHandle` is
the row-by-row alternative, writing directly as each call arrives instead:

```rust
use excelreader::writer_handle::WriterHandle;

let mut handle = WriterHandle::open("out.xlsx", None)?;
handle.start_sheet("Summary")?;
handle.start_row()?;
handle.write_str(Some("Name"))?;
handle.write_i64(Some(42))?;
handle.end_row()?;
handle.end_sheet()?;
```

Call order mirrors the C ABI's `xl_writer_handle`: `open`/`open_with`/`open_memory`, then per sheet
`start_sheet..end_sheet`, each containing `start_row..end_row` with one `write_*` call per cell
(`None` writes a blank cell), left to right. A call out of order returns `Err` rather than
corrupting output. `open_memory` backs the handle with an in-memory buffer instead of a file; read
it out with `bytes()`. Dropping a `WriterHandle` closes and releases it, same as `Workbook` — call
`bytes()` (memory-backed) or reopen the path (file-backed) to observe the result rather than relying
on the drop for that.

## Reading rows one at a time

`Sheet::rows` returns a cursor over that sheet. Each row borrows a buffer the cursor
reuses, so iterating a sheet allocates nothing per row:

```rust
use excelreader::workbook::Workbook;

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
enforces it. To hold every row at once, use `read_all_decoded`, which decodes the whole
sheet in one native call and keeps it alive until dropped:

```rust
let workbook = Workbook::open("book.xlsx")?;
let rows = workbook.sheet(0)?.read_all_decoded()?;
for row in rows.iter() {
    println!("{} cells", row.len());
}
# Ok::<(), excelreader::Error>(())
```

`parse_sheet` remains the fastest way to read a sheet whose columns you know — it converts on the
native side and never formats a cell to text.

### Chunked reads

For a sheet too large to hold in memory at once, `Sheet::typed_chunks` is `parse_sheet` delivered
a batch at a time:

```rust
use excelreader::workbook::Workbook;

let workbook = Workbook::open("book.xlsx")?;
let chunks = workbook.sheet(0)?.typed_chunks::<Row>(1, 10_000)?;
for batch in chunks {
    let batch = batch?;
    for row in batch.iter() { /* ... */ }
}
# Ok::<(), excelreader::Error>(())
```

`typed_chunks` returns `TypedChunks<'_, Row>`, an `Iterator<Item = Result<TableView<Row>, Error>>`
that closes the native reader on drop. `batch_size` is rows per batch — `0` means one unbounded
batch, identical to `parse_sheet`; negative is an error. Chunked reads, of either kind (typed or
Arrow — see [Arrow export](#arrow-export-arrow-feature) above), are independent: several can be live
on one workbook, and other reads do not disturb them. `TypedChunks` ends after yielding an `Err`, so
a `for` loop stops there rather than spinning on it.

## Bounds and panics

`TableView::get` returns `Option<T>` and is `None` outside `0..len()`. The `column_*` accessors used
by generated bindings panic on an out-of-range row, a column type that does not match the binding, or
a string the native library returned as non-UTF-8 - each is a contract violation rather than
recoverable input.

## Parallel aggregation on macOS

`aggregate_csv_file`/`aggregate_csv_memory` over a source large enough to partition (256 KB and up)
can crash the process on macOS arm64. The fault is inside the native runtime's signal handling, not
in this crate: a .NET thread resumes at address 0 after a signal the runtime sent it, which ends the
process with no message. It happens with any `degree_of_parallelism` above 1, and the same runtime
built into a plain C host does not show it, so a Rust host appears to make it far more likely.

Until it is fixed upstream, pass `degree_of_parallelism: 1` on macOS for a sequential run, or keep
sources under the 256 KB partitioning floor, where the library never fans out. Linux and Windows are
unaffected. The crate's own partitioned aggregation tests are `#[ignore]`d on macOS for this reason.

## Build notes

`build.rs` downloads the native `ExcelReader.Native` binary matching your target from the crate's
matching GitHub Release. Set `EXCELREADER_NATIVE_LIB_DIR` to a directory containing a locally-built
copy instead (e.g. from `dotnet publish ../../src/ExcelReader.Native -r win-x64`) to skip that
download - useful when building from source before a release exists yet.

Every constructor first checks the loaded library's `xl_abi_version()` against the `XL_ABI_VERSION`
this crate was compiled against, and fails with a explanatory error rather than reading native memory
through a layout that may have changed.

## Benchmarks

Criterion suite in `benches/`. Measured on Windows 10 (22H2), 16 logical CPUs @ 3.39 GHz,
rustc 1.97.1 (Release), Criterion 0.8, 100 samples per benchmark (medians shown). `write_bench`
takes 20 samples instead — each of its iterations writes a whole 65,535-row file.

`benches/parse_bench.rs` - `open`/`parse_sheet`/`infer_schema`, same methodology as the C++ suite:

| Benchmark | RealExcel.xlsb (100 rows) | 65K_Records_Data.xlsb (65,535 rows) |
|---|---:|---:|
| `open` | 85.8 µs | 98.9 µs |
| `parse_sheet` (2 or 6 bound columns) | 58.7 µs | 24.6 ms |
| typed chunks of 1,000 / 10,000 rows | - | 23.4 ms / 23.5 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 86.8 µs | 577.4 µs |

`open` stays nearly flat across the 655x row-count jump (+15%) - XLSB's header carries its
dimensions/index, so opening costs metadata, not row data. `parse_sheet` scales linearly with
rows × columns. Reading the same sheet in typed chunks costs no more than one whole-sheet
`parse_sheet`, while holding only one chunk at a time. `infer_schema` scales with its sample size,
not the file's total row count.

`benches/compare_bench.rs` compares against [calamine](https://github.com/tafia/calamine) reading
`65K_Records_Data.{xlsx,xlsb}` in full (all 14 columns, 65,535 rows). Both sides decode every cell
into an owned value (`String` for text) and fold it into one accumulator, so neither side gets a
zero-copy advantage the other can't take:

| Format | ExcelReader (`parse_sheet::<FullRow>`) | calamine (`worksheet_range` + `Data` match) |
|---|---:|---:|
| XLSX | 79.1 ms | 279.7 ms |
| XLSB | 46.9 ms | 87.5 ms |

The ExcelReader XLSX run was noisy (Criterion's interval is 76.0–83.0 ms; calamine's is 278.6–281.0 ms), so
read its ratio as 3.4–3.7x.

ExcelReader is ~3.5x faster than calamine for XLSX and ~1.9x faster for XLSB on this workload -
calamine is a fast, well-optimized reader in its own right, so the gap is real but not the order
of magnitude seen against slower libraries.

`benches/write_bench.rs` measures the two write layers and
[rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) writing the same 7 columns × 65,535
rows to `.xlsx`, all three starting from the same in-memory `Vec<Row>`:

| Benchmark | Median |
|---|---:|
| `columns` (`write_columns`, pre-transposed) | 45.7 ms |
| `sheet` (`write_sheet`, from `Vec<Row>`) | 51.9 ms |
| `rust_xlsxwriter` (cell-at-a-time) | 328.3 ms |

`sheet` is the matched-work number — it starts from the same shape `rust_xlsxwriter` is handed and
pays the row-to-column transpose itself — and is ~6.3x faster. `columns` is a ceiling no
cell-at-a-time API can reach, since it is handed buffers that are already columnar; read it only
against `sheet`, as the cost of having row-shaped data in the first place. That cost turns out to be
about 14%: not free, but a minority of the cost of producing the file.

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
