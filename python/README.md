![ExcelReader logo](https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png)

# excelreader (Python)

Read and write XLSX, XLSB, XLS and CSV through ExcelReader's NativeAOT library. No .NET runtime
required — the shared library is self-contained.

## Install

```bash
pip install excelreader-native
```

Wheels ship for Windows x64, macOS (arm64, x64) and Linux (glibc and musl, x64 and arm64). The
package installs as `excelreader-native` and imports as `excelreader`.

## Install (from source)

```bash
python python/scripts/build_native.py   # requires the .NET 10 SDK, once per machine
pip install -e "python[dev]"
```

`build_native.py` publishes `src/ExcelReader.Native` for your platform and copies the resulting
`ExcelReader.Native.{dll,so,dylib}` into `excelreader/_lib/`. To point at a binary you built
elsewhere, set `EXCELREADER_NATIVE_LIB` to its full path.

## Usage

```python
from excelreader import open_workbook

with open_workbook("book.xlsx") as workbook:
    print(len(workbook.sheets), workbook.sheets[0].name)
    for row in workbook.sheets[0].rows():
        for cell in row:
            print(cell.column, cell.type.name, cell.value)
```

### Sheets

A workbook hands out sheets. `workbook.sheets` is a sequence: `len()`, iteration, and lookup by
index (negative allowed), slice or name (names match without regard to case; a miss raises
`KeyError`). A `Sheet` is the workbook and an index; it holds no native resource, and every read on
it opens its own cursor.

```python
import excelreader

with excelreader.open_workbook("report.xlsx") as workbook:
    for sheet in workbook.sheets:
        print(sheet.index, sheet.name, sheet.visibility)

    totals = workbook.sheets["Totals"]
    for row in totals.rows():
        ...
```

One open workbook can be read from several threads, which loads the shared-string table once.
The native calls release the GIL:

```python
from concurrent.futures import ThreadPoolExecutor

with excelreader.open_workbook("report.xlsx") as workbook, ThreadPoolExecutor() as pool:
    tables = list(pool.map(lambda sheet: sheet.to_arrow(sheet.infer_schema()), workbook.sheets))
```

`rows()` and `parse_batches()` are generators: the native cursor opens at the first `next()`. A
generator that has already started keeps working after `workbook.close()`; one that has not started
raises at its first `next()`. `to_arrow_stream()` opens its stream at the call, so it works after
`close()` even if never started.

### Tables

`workbook.tables` lists the workbook's Excel tables ("Format as Table" ranges), with the same
lookups as `sheets`: index, slice or name, ignoring case. A `Table` is a `Sheet` bounded to the
table's range, so every read works on it and starts at the table's header row:

```python
with excelreader.open_workbook("report.xlsx") as workbook:
    sales = workbook.tables["Sales"]
    print(sales.name, sales.sheet.name, sales.first_row, sales.last_row)
    frame = sales.to_pandas(sales.infer_schema())
```

Rows and columns are zero-based and include the header and totals rows; cells keep the column
indexes of the sheet holding the table. `header_row_count` is 0 for a table whose header row is
turned off, and `totals_row_count` is 1 when it shows a totals row. A table is read on one thread
whatever `parallelism` asks for.

### Formats

`open_workbook` sniffs XLS/XLSX/XLSB by file signature. CSV has no signature, so it is chosen by the
`.csv` extension — or explicitly:

```python
open_workbook("data.txt", format="csv")
```

### Dates

`cell.value` is always the raw text as stored, so `CellType.DATE` cells hold Excel serial numbers.
Use `Cell.as_date()` to convert, passing the workbook's epoch flag:

```python
as_date = cell.as_date(workbook.is_date1904)
```

`as_date()` returns `None` for any cell that isn't `CellType.DATE`.

### Reading a sheet

A sheet reads in one of these shapes:

| Method | Returns | Knob | Needs |
|---|---|---|---|
| `rows()` | one list of `Cell` per row, streamed | — | nothing |
| `parse()` | the whole sheet as a `TypedTable` | `parallelism` | nothing |
| `parse_batches()` | one `TypedTable` per batch | `batch_size` | nothing |
| `to_arrow()` | the whole sheet as a `pyarrow.RecordBatch` | `parallelism` | pyarrow |
| `to_arrow_stream()` | an Arrow PyCapsule stream of batches | `batch_size` | nothing |
| `to_pandas()` / `to_polars()` | a DataFrame | `batch_size`, `parallelism` | pandas + pyarrow / polars |

`infer_schema()` guesses the schema every typed method takes.

### Typed columns — the fastest path

`rows()` hands back cell *text*, which means the library formats every value to a string on the way
out and builds one Python object per cell. `parse()` skips that entirely: you give it a schema,
and the conversion happens natively, straight into typed column buffers. On a 65,536 × 14 sheet it is
~30× faster than `rows()`, and faster than `polars.read_excel()` — see [Reading](#reading) below for
the measured numbers.

```python
from excelreader import ColumnSpec, ColumnType

with open_workbook("sales.xlsb") as workbook:
    table = workbook.sheets[0].parse([
        ColumnSpec(ColumnType.STRING, name="Region"),
        ColumnSpec(ColumnType.DATE, name="Order Date"),
        ColumnSpec(ColumnType.F64, name="Total Revenue", nullable=True),
    ])

table.row_count            # rows read
table.names                # ["Region", "Order Date", "Total Revenue"]
region, day, revenue = table.columns
region[0]                  # "Asia" — strings decode on demand, not one str per row up front
day[0]                     # 15477 — days since 1970-01-01
revenue[0]                 # 14862.69
table.validity[2]          # bit-packed nulls, or None when the column has none
```

Each non-string column is a stdlib `array.array`, or a NumPy array if NumPy is installed
(`pip install -e "python[numpy]"`) — NumPy is optional and never required.

Leave `name` out to resolve a column by position instead: `ColumnSpec(ColumnType.I64, index=3)`.
`header_row` defaults to 1 (the first row names the columns); pass `header_row=0` for a sheet with no
header, where every spec must resolve by index.

A column that fails to convert is an error unless its spec sets `nullable=True`, which records the
failure in `table.validity` and keeps reading.

`parse()` always reads the whole sheet from its first row, independent of any `rows()`
generator in progress.

#### Reading a sheet a batch at a time

For a sheet too large to hold in memory at once, `parse_batches()` is `parse()` a batch at a
time — a generator yielding one `TypedTable` per batch:

```python
with open_workbook("sales.xlsb") as workbook:
    schema = [
        ColumnSpec(ColumnType.STRING, name="Region"),
        ColumnSpec(ColumnType.DATE, name="Order Date"),
        ColumnSpec(ColumnType.F64, name="Total Revenue", nullable=True),
    ]
    for batch in workbook.sheets[0].parse_batches(schema, batch_size=10_000):
        region, day, revenue = batch.columns
        ...
```

`batch_size` is rows per batch; `0` means one unbounded batch, identical to `parse()`, and a
negative value is an error. Chunked reads are independent: several can be open on one workbook at
once, typed or Arrow (see [Arrow](#arrow) below), and none disturbs another.

Being a generator, `parse_batches()` opens nothing until the first iteration, so a bad
`batch_size` is only raised there, not at the call.

#### Reading a large sheet on several threads

`parse()`, `to_arrow()`, `to_pandas()` and `to_polars()` take `parallelism`: `1` (the
default) reads on one thread, `0` uses every core, `n` up to n threads. XLSX, XLSB and CSV are split. XLS, or a
sheet under 256 KiB, is read on one thread with the same result. The table is identical
either way. Every partition's columns are held until they are merged, so the peak memory is higher
than a sequential read of the same file. `to_pandas()` and `to_polars()` with `parallelism` other
than 1 parse the whole file instead of streaming it in batches.

`xl_parse_typed` on a Ryzen 7 5700X (8 cores, 16 threads), 14 typed columns, file read from memory:

| File | `parallelism=1` | `parallelism=0` | Speed-up |
|---|---:|---:|---:|
| `65K_Records_Data.csv` (8 MB, 65,535 rows) | 21.4 ms | 5.8 ms | 3.7x |
| the same rows ×20 (160 MB, 1.3M rows) | 425.9 ms | 117.5 ms | 3.6x |

The same call on the real-data workbooks (XLSX 6.7 MB compressed, XLSB 3.7 MB; measured with the
published NativeAOT library): XLSX 54.1 ms at
`parallelism=1` against 20.3 ms at `0` (2.6x), XLSB 28.2 ms against 15.2 ms (1.85x). Inflate runs on
one thread, so under the JIT the sheet stage stops scaling at about three threads.

Peak working set on the 160 MB file, above the interpreter's own, went from 470 MiB to 503 MiB (both
include the file's bytes and the handle's copy of them).

#### Guessing a schema

Writing the `ColumnSpec` list by hand means already knowing every column's name and type. When you
don't, `infer_schema()` samples the sheet and guesses one for you:

```python
with open_workbook("sales.xlsb") as workbook:
    sheet = workbook.sheets[0]
    schema = sheet.infer_schema()   # header_row=1, sample_size=100 by default
    table = sheet.parse(schema)
```

By default each column's type comes from the `CellType` Excel already stored for its sampled cells,
so it costs nothing beyond the sample and is exact for XLSX/XLSB/XLS. A column with a real
mix of kinds, only formula/error results, or nothing sampled falls back to `ColumnType.STRING`;
`nullable` is set when any sampled row left the column empty. CSV cells carry no such type tag, so
by default every CSV column is guessed `ColumnType.STRING`. Pass `parse_text=True` to type text cells
from their exact shape instead: integers, decimals, `true`/`false` and ISO dates or date-times. Codes
with a leading zero (`00123`), scientific notation (`12E4`), padded or comma-decimal numbers and
non-ISO dates stay strings. It is
still a guess over the sample, so a value further down can fail to convert:

```python
with open_workbook("sales.csv") as workbook:
    sheet = workbook.sheets[0]
    schema = sheet.infer_schema(parse_text=True)
    table = sheet.parse(schema, parallelism=0)
```

### Writing

`write_workbook()` writes a `TypedTable` (what `parse()` returns) back out as a single sheet,
through the same `xl_write_typed` native export — one-shot, no writer handle before or after the call:

```python
from excelreader import ColumnType, write_workbook

with open_workbook("sales.xlsb") as workbook:
    sheet = workbook.sheets[0]
    table = sheet.parse(sheet.infer_schema())

types = [ColumnType.STRING, ColumnType.DATE, ColumnType.F64]  # one per table.columns, in order
write_workbook("sales_copy.xlsx", table, types)
```

`types` is required because a `TypedTable` column is a raw buffer (`array`/`StringColumn`/NumPy
array) and nothing about the buffer alone tells I64 from TIME apart — both are 8-byte-per-row
arrays. `format` is inferred from the path's extension (one of xlsx/xlsb/xls/csv) or set explicitly:

```python
write_workbook("report.dat", table, types, format="csv")
```

The destination can also be a binary object with `write()`, such as `io.BytesIO` or an open file. The
workbook is built in memory and written in one `write()` call, and with no extension to infer it from,
`format` is required:

```python
import io

buffer = io.BytesIO()
write_workbook(buffer, table, types, format="xlsx")
payload = buffer.getvalue()
```

`write_workbook()` also takes Arrow-shaped data in place of a `TypedTable`: a `pyarrow.RecordBatch`
or `Table`, a pandas DataFrame (its index is not written), or any table exposing
`__arrow_c_stream__`, such as a polars DataFrame. Each column's type comes from its Arrow type, so
leave `types` out. This path requires `pyarrow`:

```python
write_workbook("report.xlsx", df)          # pandas
write_workbook("report.xlsx", polars_df)   # polars
write_workbook("report.xlsx", record_batch)
```

`WriteOptions` sets the sheet name and CSV dialect, mirroring `xl_write_options` — every field
defaults to `None`, meaning "use the library default":

```python
from excelreader import WriteOptions

write_workbook(
    "report.xlsx", table, types,
    options=WriteOptions(sheet_name="Q3 Results", use_shared_strings=True),
)
```

For xlsx and xlsb, `degree_of_parallelism` formats and compresses batches of `batch_rows` rows on
several threads: `0` uses every processor, `1` writes serially, and `None` is serial. A `batch_rows`
of `None` keeps the library default of 8,192. With shared strings on, two parallel runs give the same
content in files that differ byte for byte. xls and csv ignore both fields.

```python
write_workbook("big.xlsx", table, types, options=WriteOptions(degree_of_parallelism=0))
```

**Phase-1 limits, stated plainly:** `write_workbook` writes a single sheet (for several, see
`write_sheet` below); the whole table must already be in memory (no streaming/chunked writes); no
styling beyond the temporal number formats `xl_write_typed` applies to DATE/TIME/TIMESTAMP columns.
`format="auto"` is not accepted — sniffing reads a file's existing signature bytes, and a file being
created has none.

### Several sheets

`open_writer` returns a `SheetWriter`, whose `write_sheet` writes a whole table as one sheet in a
single native call. It takes the same `data` and `types` as `write_workbook`, so each sheet can be a
`TypedTable` or a pandas, polars or pyarrow frame:

```python
from excelreader import WriteOptions, open_writer

with open_writer("report.xlsx", options=WriteOptions(degree_of_parallelism=0)) as writer:
    writer.write_sheet("Sales", sales_df)
    writer.write_sheet("Stock", table, types)
```

Parallelism and the other options come from `open_writer` and apply to every sheet;
`WriteOptions.sheet_name` does not apply, since each call names its sheet. `write_sheet` raises
`ExcelReaderError` while a sheet started with `start_sheet` is open; between sheets, the two styles
mix freely.

`table=TableOptions(...)` writes the sheet as an Excel table, with filter buttons and a table style
(xlsx and xlsb only). The name must be unique in the workbook; the style defaults to
`TableStyleMedium2`, and `style=None` leaves the table unstyled:

```python
from excelreader import TableOptions

writer.write_sheet("Sales", sales_df, table=TableOptions("Sales", style="TableStyleLight9", column_stripes=True))
```

## Writing a sheet row by row

`open_workbook`'s counterpart for writing one row at a time, instead of building a whole table
first:

```python
import datetime
from excelreader import open_writer

with open_writer("out.xlsx") as writer:
    writer.start_sheet("Data")
    writer.write_row(["name", "qty", "when"])
    writer.write_row(["widget", 7, datetime.date(2026, 1, 31)])
    writer.end_sheet()
```

Only a clean exit from the `with` block, or `close()`, commits the workbook. An exception inside the
block, `abort()`, or a writer collected without being closed discards what was written and leaves an
existing file at the path as it was.

`write_row` picks each cell's type from the Python value: `str`, `bool`, `int` (64-bit), `float`,
`datetime.datetime` (an aware one is converted to UTC), `datetime.date` and `datetime.time`. A `None`
becomes a blank string cell; the tuple `(None, ColumnType.I64)` writes a blank cell of that type:

```python
writer.write_row(["widget", (None, ColumnType.I64), 2.5])
```

To build a workbook without touching the filesystem, open the writer over a binary object with
`write()`, such as `io.BytesIO`. It needs `format`, and receives the finished workbook in one
`write()` call when the writer closes:

```python
import io

from excelreader import open_writer

buffer = io.BytesIO()
with open_writer(buffer, "xlsx") as writer:
    writer.start_sheet("Data")
    writer.write_row(["name", "qty"])
    writer.end_sheet()
payload = buffer.getvalue()
```

### Arrow

With `pyarrow` installed, `to_arrow()` runs the same read and hands the buffers to pyarrow zero-copy
over the Arrow C Data Interface, as one column-named `pyarrow.RecordBatch`:

```python
with open_workbook("sales.xlsb") as workbook:
    batch = workbook.sheets[0].to_arrow(schema)
```

pyarrow owns the buffers from that point on, so the result stays valid after the workbook is closed.

#### Streaming: Arrow streams, pandas, and polars

`to_arrow_stream()` is `to_arrow()` a batch at a time — peak memory is one batch rather than one
sheet. It returns a single-use stream implementing the Arrow PyCapsule interface
(`__arrow_c_stream__`) and needs no third-party package itself; hand it to any consumer, such as
pyarrow for a `RecordBatchReader`:

```python
import pyarrow as pa

with open_workbook("sales.xlsb") as workbook:
    reader = pa.RecordBatchReader.from_stream(workbook.sheets[0].to_arrow_stream(schema, batch_size=10_000))
    for batch in reader:
        frame = batch.to_pandas()   # or polars.from_arrow(batch), for one DataFrame per batch
```

Same `batch_size` rules as `parse_batches()` above, except `to_arrow_stream()` raises
immediately rather than on the first iteration. The stream keeps working after the workbook is closed.

`to_pandas()`/`to_polars()` materialize the whole sheet as a single DataFrame and both take
`batch_size` and `parallelism`, but they spend `batch_size` differently — do not blur them.
`to_polars()` genuinely streams: polars consumes the stream batch by batch, so the whole sheet is
never resident as Arrow buffers at once. `to_pandas()` does **not** bound peak memory the same way —
it concatenates every batch before pandas ever sees them. What `batch_size` buys `to_pandas()` instead
is the conversion: `self_destruct=True` frees each Arrow chunk as pandas consumes it, so the sheet is
never held twice, once as Arrow and once as the DataFrame.

**A faulted stream surfaces differently from every other native error in this library**: pyarrow's
`RecordBatchReader` reports a failed batch as a plain `OSError` carrying the native error message, not
`ExcelReaderError` — the failure crosses the Arrow C Data Interface before this library's own
exception wrapping ever runs.

### From memory, a stream, or your own bytes

`open_workbook()` takes more than a path. `bytes`, `bytearray`, `memoryview` and `io.BytesIO` are
read from memory; the format comes from the content, so pass `format` only for a CSV or to skip the
check.

```python
with open_workbook(payload) as workbook:
    ...
```

An object with a `size` and a `read_at(offset, buffer)` that fills `buffer` and returns the count is
read in blocks on demand: the library asks for 4 MiB blocks and caches them, and every sheet can
still be read on its own thread.

```python
import urllib.request

from excelreader import open_workbook

class HttpSource:
    def __init__(self, url):
        self.url = url
        with urllib.request.urlopen(urllib.request.Request(url, method="HEAD")) as response:
            self.size = int(response.headers["Content-Length"])

    def read_at(self, offset, buffer):
        request = urllib.request.Request(self.url, headers={"Range": f"bytes={offset}-{offset + len(buffer) - 1}"})
        with urllib.request.urlopen(request) as response:
            data = response.read()
        buffer[: len(data)] = data
        return len(data)

with open_workbook(HttpSource("https://example.com/book.xlsx")) as workbook:
    ...
```

Any other object with `readinto()` or `read()`, such as an open file or `sys.stdin.buffer`, is read
as a stream. A CSV is read as it arrives, and its first 16 MiB are kept so that `infer_schema()` followed by a read works;
an XLSX, XLSB or XLS is read whole first.

```python
import sys

from excelreader import open_workbook

with open_workbook(sys.stdin.buffer, "csv") as workbook:
    ...
```

Both callbacks run with the GIL held. For a file on disk, pass its path: that is faster.

### Reader options

`open_workbook()` takes an optional `OpenOptions` for CSV dialect settings and reader
resource limits. Every field defaults to `None`, meaning "use the library default", so you set only
what you want to change.

```python
from excelreader import OpenOptions, open_workbook

# A semicolon-delimited CSV, which the default comma dialect would read as one column per row.
with open_workbook("export.csv", format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
    for row in workbook.sheets[0].rows():
        ...
```

`csv_delimiter` and `csv_quote` are byte values, so pass `ord(";")` rather than `";"`.
`max_cell_bytes` and `intern_strings` apply to every format, a CSV included.

The `max_*` fields are resource limits rather than tuning knobs: they bound what a malformed or
hostile file can make the reader allocate, and exceeding one raises `ExcelReaderError`. Lower them
when parsing untrusted uploads.

```python
options = OpenOptions(
    max_total_decompressed_bytes=64 * 1024 * 1024,  # zip-bomb budget for XLSX/XLSB
    max_cell_bytes=1024 * 1024,
    max_zip_entries=1024,
)
```

`prefetch_decompression=True` overlaps inflating an XLSX/XLSB sheet with parsing it — worth it for
single-file batch work, not for a server already reading many files in parallel. See the root README
for the measured trade.

Each ZIP entry of an XLSX or XLSB is checked against its CRC-32 as it is read, so a damaged file
raises `ExcelReaderError` instead of returning wrong cells. `verify_zip_checksums=False` turns the
check off, for a file whose producer wrote wrong checksums.

### Encrypted workbooks

`open_workbook()` takes a `password` keyword to open a password-protected OOXML
workbook (.xlsx/.xlsb/.xlsm):

```python
from excelreader import PasswordIncorrectError, open_workbook

try:
    with open_workbook("protected.xlsx", password="hunter2") as workbook:
        ...
except PasswordIncorrectError:
    ...  # ask again
```

Omitting `password` for an encrypted file raises `PasswordRequiredError`; a wrong one raises
`PasswordIncorrectError` — both subclass `ExcelReaderError`, so a caller that doesn't care about the
distinction can just catch that. Any other native failure (an unsupported encryption scheme, a
corrupt file) also raises `ExcelReaderError` but is not worth retrying.

An explicit `format="xlsx"`/`format="xlsb"` works for an encrypted file too, the same as leaving
`format` unset — both routes decrypt correctly given the right `password`.

Writing an encrypted workbook is a second step: write the plaintext package with `write_workbook`
or `open_writer`, then wrap it with `encrypt_package`, which produces an
agile-encrypted (ECMA-376 4.4) CFB container — AES-256-CBC, SHA-512, 100,000 spin iterations, with a
`dataIntegrity` HMAC:

```python
from excelreader import encrypt_package, write_workbook

write_workbook("plain.xlsx", table, types)
encrypt_package("plain.xlsx", "secret.xlsx", "hunter2")
```

The package must already be a finished file; it is read, not removed. Encryption parameters are
fixed at Excel's own defaults — there are no options — and only XLSX/XLSB packages can be encrypted,
matching what `open_workbook` can decrypt.

Either side can stay in memory, so the plaintext package never has to be written to disk: the
package can be `bytes` or an `io.BytesIO`, and the destination any binary object with `write()`:

```python
import io

from excelreader import encrypt_package, write_workbook

plain, secret = io.BytesIO(), io.BytesIO()
write_workbook(plain, table, types, format="xlsx")
encrypt_package(plain, secret, "hunter2")
```

## Benchmarks

`benchmarks/bench_read.py` and `benchmarks/bench_write.py` over
`tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb` (65,535 data rows, 14 columns), the same
fixture the .NET, C++ and Rust suites use. Measured on Windows 10 (22H2), AMD Ryzen 7 5700X
(16 logical CPUs @ 3.39 GHz), CPython 3.14.7, 10 runs each (medians shown; `min` is in the scripts' own
output), against a library published from commit 449ca144. Competitors: polars 1.44.1 (with
fastexcel 0.21.0), pandas 3.0.5 (with openpyxl 3.1.5), xlsxwriter 3.2.9.

### Reading

| API | Median | What it produces |
|---|---:|---|
| `to_arrow_stream()`, read by pyarrow | 31.2 ms | Arrow record batches, streamed |
| `to_arrow()` | 30.1 ms | the same parse, handed to pyarrow zero-copy |
| `parse()` | 32.5 ms | typed columnar buffers, converted natively |
| `parse_batches()` | 33.2 ms | the same typed buffers, in batches |
| `to_polars()` | 33.6 ms | typed columnar DataFrame, schema inferred |
| `polars.read_excel()` | 120.6 ms | typed columnar DataFrame, types inferred |
| `rows()` | 975.7 ms | one `Cell` object per cell, streamed per row |

The batched readers cost the same as their whole-sheet counterparts and lower the peak memory:
`to_arrow_stream(batch_size=10000)` peaks at 11.4 MiB against 19.5 MiB for one whole-sheet
record batch on this file.

Only the `to_polars()` / `polars.read_excel()` pair is a like-for-like comparison (~3.6x), and even
that one is loose: both produce a typed columnar DataFrame with inferred types, but the inference rules
are not identical. The rows above it produce different things and are listed to show what each API
costs, not to rank them — `rows()` is ~30x slower than `parse()` because it materializes
917,504 Python objects, which is the price of that shape, not a slow parser.

CSV, from `65K_Records_Data.csv` (8.2 MB), same 14 typed columns:

| API | Median |
|---|---:|
| `parse(parallelism=1)` | 23.3 ms |
| `parse(parallelism=0)` | 11.3 ms |
| `to_polars(parallelism=1)` | 24.2 ms |
| `to_polars(parallelism=0)` | 9.7 ms |
| `polars.read_csv(try_parse_dates=True)` | 6.9 ms |
| `pandas.read_csv`, pyarrow engine | 116.0 ms |
| `pandas.read_csv`, C engine | 132.3 ms |

polars' CSV reader is ahead (multithreaded, and it infers types from the text without our schema
step); ExcelReader is ~10x faster than pandas on either engine. These include opening the file, so
they are above the in-memory `xl_parse_typed` figures in [Reading a large sheet on several
threads](#reading-a-large-sheet-on-several-threads).

### Writing

| API | Median | Output |
|---|---:|---|
| `write_workbook()` → csv | 33.3 ms | 8.2 MB |
| `write_workbook()` → xlsb | 51.7 ms | 4.5 MB |
| `write_workbook()` → xls | 59.2 ms | 13.4 MB |
| `write_workbook()` → xlsx | 68.5 ms | 5.1 MB |
| `write_workbook(polars_df)` → xlsx | 76.6 ms | 5.1 MB |
| `write_workbook(pandas_df)` → xlsx | 82.5 ms | 5.1 MB |
| `polars.DataFrame.write_excel()` | 4,714.9 ms | 5.6 MB |
| `pandas.DataFrame.to_excel()` | 7,371.4 ms | 5.5 MB |

The two DataFrame comparisons are matched work — same DataFrame in, xlsx out both times:
`write_workbook(polars_df)` is ~62x faster than polars' own `write_excel()`, and
`write_workbook(pandas_df)` ~89x faster than `to_excel()`. Both of ours pay a conversion the raw path does not: the DataFrame goes
through Arrow, and each column is cast to the layout the native side reads and copied out of Arrow's
buffers, which is the gap between those rows and the 68.5 ms one. Handing `write_workbook()` buffers
that are already columnar — what `parse()` returns — skips it.

`write_workbook(xlsx)` at 68.5 ms lands within ~4 ms of the C++ binding's
`xl::write_columns` on the same 14 columns (64.2 ms), which is the expected result: both are thin wrappers
over the same `xl_write_typed` call, and neither adds work per cell.

`xls` being the largest file and still faster than xlsx is not a paradox — BIFF8 writes
fixed-width records with no compression, so it trades ~2.6x the bytes for less work per cell. xlsb is faster still, at 51.7 ms.

## Notes

- A `Workbook` can be read from several threads at once; see [Sheets](#sheets).
- Empty cells are skipped, so `cell.column` may skip indices. Do not assume `row[i].column == i`.
- The ABI is documented in `src/ExcelReader.Native/include/excelreader.h`.
