# excelreader (Python)

Read and write XLSX, XLSB, XLS and CSV through ExcelReader's NativeAOT library. No .NET runtime
required — the shared library is self-contained.

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
    print(workbook.sheet_count, workbook.sheet_name)
    for row in workbook.rows():
        for cell in row:
            print(cell.column, cell.type.name, cell.value)
```

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

### Reading everything at once

`rows()` iterates row-by-row; `read_all()` materializes the whole sheet in one call:

```python
all_rows = workbook.read_all()
```

This holds every row in memory at once, so prefer `rows()` for very large sheets.

### Reading everything at once, faster

`read_all()`/`rows()` build one `Cell`/`str` object per cell, which dominates wall-clock time on a
large sheet. `read_all_columnar()` decodes the same data into parallel flat arrays instead — no
per-cell object construction — and is several times faster on large sheets:

```python
sheet = workbook.read_all_columnar()
# sheet.row_offsets[i]:row_offsets[i+1]  -> cell indices for row i
# sheet.columns[j] / sheet.types[j]      -> cell j's column index / CellType
# sheet.value_offsets[j]:[j+1]           -> cell j's byte slice into sheet.values
```

Materialize a single cell on demand instead of decoding every value up front:

```python
from excelreader import decode_cell

first_cell = decode_cell(sheet, 0)
```

Each array is a stdlib `array.array('i')`, or a NumPy `int32` array if NumPy is installed
(`pip install -e "python[numpy]"`) — NumPy is optional and never required.

### Typed columns — the fastest path

Everything above hands back cell *text*, which means the library formats every value to a string on
the way out. `parse_typed()` skips that entirely: you give it a schema, and the conversion happens
natively, straight into typed column buffers. On a 65,536 × 14 sheet it is ~8× faster than
`read_all_columnar()`, ~25× faster than `read_all()`, and faster than `polars.read_excel()` — see
[Reading](#reading) below for the measured numbers.

```python
from excelreader import ColumnSpec, ColumnType

with open_workbook("sales.xlsb") as workbook:
    table = workbook.parse_typed([
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

Leave `name` out to resolve a column by position instead: `ColumnSpec(ColumnType.I64, index=3)`.
`header_row` defaults to 1 (the first row names the columns); pass `header_row=0` for a sheet with no
header, where every spec must resolve by index.

A column that fails to convert is an error unless its spec sets `nullable=True`, which records the
failure in `table.validity` and keeps reading.

Note that `parse_typed()` always reads the whole sheet from its first row, independent of how far
`rows()` has advanced — and it leaves that cursor alone.

#### Reading a sheet a batch at a time

For a sheet too large to hold in memory at once, `iter_parse_typed()` is `parse_typed()` a batch at a
time — a generator yielding one `TypedTable` per batch:

```python
with open_workbook("sales.xlsb") as workbook:
    schema = [
        ColumnSpec(ColumnType.STRING, name="Region"),
        ColumnSpec(ColumnType.DATE, name="Order Date"),
        ColumnSpec(ColumnType.F64, name="Total Revenue", nullable=True),
    ]
    for batch in workbook.iter_parse_typed(schema, batch_size=10_000):
        region, day, revenue = batch.columns
        ...
```

`batch_size` is rows per batch; `0` means one unbounded batch, identical to `parse_typed()`, and a
negative value is an error. A workbook serves one chunked read at a time — of either kind, typed or
Arrow (see [Arrow](#arrow) below) — and any other read on the workbook while one is live invalidates
it: its next call then raises `ExcelReaderError` rather than silently resuming from the moved cursor.
Finish the batches, or call `.close()` on the generator, before starting another read.

Being a generator, `iter_parse_typed()` opens nothing until the first iteration, so a bad
`batch_size` or a rejected second reader is only raised there, not at the call.

#### Reading a large CSV on several threads

`parse_typed()`, `to_arrow()`, `to_record_batch()` and `to_polars()` take `parallelism`: `1` (the
default) reads on one thread, `0` uses every core, `n` up to n threads. Only CSV is split. Any other
format, or a CSV under 256 KiB, is read on one thread with the same result. The table is identical
either way. Every partition's columns are held until they are merged, so the peak memory is higher
than a sequential read of the same file. `to_polars()` with `parallelism` other than 1 parses the whole
file instead of streaming it in batches.

`xl_parse_typed_ex` on a Ryzen 7 5700X (8 cores, 16 threads), 14 typed columns, file read from memory:

| File | `parallelism=1` | `parallelism=0` | Speed-up |
|---|---:|---:|---:|
| `65K_Records_Data.csv` (8 MB, 65,535 rows) | 20.7 ms | 6.9 ms | 3.0x |
| the same rows ×20 (160 MB, 1.3M rows) | 432.2 ms | 125.2 ms | 3.5x |

Peak working set on the 160 MB file went from 916 MB to 955 MB (both include the file's bytes and the
handle's copy of them).

#### Guessing a schema

Writing the `ColumnSpec` list by hand means already knowing every column's name and type. When you
don't, `infer_schema()` samples the sheet and guesses one for you:

```python
with open_workbook("sales.xlsb") as workbook:
    schema = workbook.infer_schema()   # header_row=1, sample_size=100 by default
    table = workbook.parse_typed(schema)
```

By default each column's type comes from the `CellType` Excel already stored for its sampled cells,
so it costs nothing beyond the sample and is exact for XLSX/XLSB/XLS. A column with a real
mix of kinds, only formula/error results, or nothing sampled falls back to `ColumnType.STRING`;
`nullable` is set when any sampled row left the column empty. CSV cells carry no such type tag, so
by default every CSV column is guessed `ColumnType.STRING`. Pass `parse_text=True` to type text cells
from their exact shape instead: integers, decimals, `true`/`false` and ISO dates or date-times. Codes
with a leading zero (`00123`), padded or comma-decimal numbers and non-ISO dates stay strings. It is
still a guess over the sample, so a value further down can fail to convert:

```python
with open_workbook("sales.csv") as workbook:
    schema = workbook.infer_schema(parse_text=True)
    table = workbook.parse_typed(schema, parallelism=0)
```

### Writing

`write_workbook()` writes a `TypedTable` (what `parse_typed()` returns) back out as a single sheet,
through the same `xl_write_typed` native export — one-shot, no writer handle before or after the call:

```python
from excelreader import ColumnType, write_workbook

with open_workbook("sales.xlsb") as workbook:
    table = workbook.parse_typed(workbook.infer_schema())

types = [ColumnType.STRING, ColumnType.DATE, ColumnType.F64]  # one per table.columns, in order
write_workbook("sales_copy.xlsx", table, types)
```

`types` is required because a `TypedTable` column is a raw buffer (`array`/`StringColumn`/NumPy
array) and nothing about the buffer alone tells I64 from TIME apart — both are 8-byte-per-row
arrays. `format` is inferred from the path's extension (one of xlsx/xlsb/xls/csv) or set explicitly:

```python
write_workbook("report.dat", table, types, format="csv")
```

`write_pandas()` and `write_polars()` build the table from a DataFrame instead (both go through
`write_arrow()`, so `pyarrow` must be installed):

```python
from excelreader import write_pandas, write_polars

write_pandas("report.xlsx", df)          # requires pandas + pyarrow
write_polars("report.xlsx", polars_df)   # requires polars + pyarrow
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

**Phase-1 limits, stated plainly:** a single sheet only (no multi-sheet workbooks); the whole table
must already be in memory (no streaming/chunked writes); no styling beyond the temporal number
formats `xl_write_typed` applies to DATE/TIME/TIMESTAMP columns. `format="auto"` is not accepted —
sniffing reads a file's existing signature bytes, and a file being created has none.

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

`write_row` picks each cell's type from the Python value. For control over a column's type — or to
write a typed blank — use the explicit methods: `write_str`, `write_i64`, `write_f64`, `write_bool`,
`write_date`, `write_time`, `write_timestamp`, and `write_null(ColumnType.I64)`. A `None` passed to
`write_row` becomes a blank string cell.

To build a workbook without touching the filesystem, use `open_writer_to_memory` and read the result
out with `bytes()`:

```python
from excelreader import open_writer_to_memory

with open_writer_to_memory("xlsx") as writer:
    writer.start_sheet("Data")
    writer.write_row(["name", "qty"])
    writer.end_sheet()
    payload = writer.bytes()
```

`write_workbook_to_bytes()` is the same idea for the columnar `write_workbook()` path.

### Arrow

With `pyarrow` installed, `to_arrow()` runs the same read and hands the buffers to pyarrow zero-copy
over the Arrow C Data Interface:

```python
import pyarrow as pa

with open_workbook("sales.xlsb") as workbook:
    array = workbook.to_arrow(schema)

batch = pa.RecordBatch.from_struct_array(array)
```

pyarrow owns the buffers from that point on, so the result stays valid after the workbook is closed.

#### Streaming: RecordBatchReader, pandas, and polars

`to_record_batch_reader()` is `to_arrow()` a batch at a time, as a streaming
`pyarrow.RecordBatchReader` — peak memory is one batch rather than one sheet:

```python
with open_workbook("sales.xlsb") as workbook:
    reader = workbook.to_record_batch_reader(schema, batch_size=10_000)
    for batch in reader:
        ...
```

Same `batch_size`/one-chunked-read-at-a-time rules as `iter_parse_typed()` above, except
`to_record_batch_reader()` raises immediately rather than on the first iteration.

`iter_pandas()`/`iter_polars()` build on it, yielding one DataFrame per batch (requires pyarrow, plus
pandas or polars respectively):

```python
with open_workbook("sales.xlsb") as workbook:
    for frame in workbook.iter_polars(schema, batch_size=10_000):
        ...
```

`to_pandas()`/`to_polars()` materialize the whole sheet as a single DataFrame and also take a
`batch_size`, but the two spend it differently — do not blur them. `to_polars()` genuinely streams:
polars consumes the reader batch by batch, so the whole sheet is never resident as Arrow buffers at
once. `to_pandas()` does **not** bound peak memory the same way — `RecordBatchReader.read_all()`
concatenates every batch before pandas ever sees them. What `batch_size` buys `to_pandas()` instead is
the conversion: `self_destruct=True` frees each Arrow chunk as pandas consumes it, so the sheet is
never held twice, once as Arrow and once as the DataFrame.

**A faulted stream surfaces differently from every other native error in this library**: pyarrow's
`RecordBatchReader` reports a failed batch as a plain `OSError` carrying the native error message, not
`ExcelReaderError` — the failure crosses the Arrow C Data Interface before this library's own
exception wrapping ever runs.

### From memory

```python
from excelreader import open_bytes

with open_bytes(payload) as workbook:
    ...
```

### Reader options

`open_workbook()`/`open_bytes()` take an optional `OpenOptions` for CSV dialect settings and reader
resource limits. Every field defaults to `None`, meaning "use the library default", so you set only
what you want to change.

```python
from excelreader import OpenOptions, open_workbook

# A semicolon-delimited CSV, which the default comma dialect would read as one column per row.
with open_workbook("export.csv", format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
    for row in workbook.rows():
        ...
```

`csv_delimiter` and `csv_quote` are byte values, so pass `ord(";")` rather than `";"`.

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

### Encrypted workbooks

`open_workbook()`/`open_bytes()` take a `password` keyword to open a password-protected OOXML
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
(or any of the `write_*` helpers), then wrap it with `encrypt_package`, which produces an
agile-encrypted (ECMA-376 4.4) CFB container — AES-256-CBC, SHA-512, 100,000 spin iterations, with a
`dataIntegrity` HMAC:

```python
from excelreader import encrypt_package, write_workbook

write_workbook("plain.xlsx", table, types)
encrypt_package("plain.xlsx", "secret.xlsx", "hunter2")
```

`package_path` is read twice (it is not disposed or removed), so it must already be a finished file.
Encryption parameters are fixed at Excel's own defaults — there are no options — and only XLSX/XLSB
packages can be encrypted, matching what `open_workbook`/`open_bytes` can decrypt.

## Benchmarks

`benchmarks/bench_read.py` and `benchmarks/bench_write.py` over
`tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb` (65,535 data rows, 14 columns), the same
fixture the .NET, C++ and Rust suites use. Measured on Windows 10 (22H2), 16 logical CPUs
@ 3.39 GHz, CPython 3.14.4, 10 runs each (medians shown; `min` is in the scripts' own output).
Competitors: polars 1.44.1 (with fastexcel 0.21.0), pandas 3.0.5 (with openpyxl 3.1.5), xlsxwriter
3.2.9.

### Reading

| API | Median | What it produces |
|---|---:|---|
| `to_record_batch_reader()` | 35.8 ms | Arrow record batches, streamed |
| `to_arrow()` | 36.1 ms | the same parse, handed to pyarrow zero-copy |
| `parse_typed()` | 39.5 ms | typed columnar buffers, converted natively |
| `iter_parse_typed()` | 38.4 ms | the same typed buffers, in batches |
| `to_polars()` | 38.7 ms | typed columnar DataFrame, schema inferred |
| `polars.read_excel()` | 124.3 ms | typed columnar DataFrame, types inferred |
| `read_all_columnar()` | 471.7 ms | raw columnar cells, no per-cell Python objects |
| `rows()` | 1,086.9 ms | one `Cell` object per cell, streamed per row |
| `read_all()` | 1,656.5 ms | one `Cell` object per cell, all at once |

The batched readers cost the same as their whole-sheet counterparts and lower the peak memory:
`to_record_batch_reader(batch_size=10000)` peaks at 20.2 MiB against 24.8 MiB for one whole-sheet
record batch on this file.

Only the `to_polars()` / `polars.read_excel()` pair is a like-for-like comparison, and even that one
is loose: both produce a typed columnar DataFrame with inferred types, but the inference rules are
not identical. The rows above it produce different things and are listed to show what each API
costs, not to rank them — `read_all()` is ~42x slower than `parse_typed()` because it materializes
917,504 Python objects, which is the price of that shape, not a slow parser.

### Writing

| API | Median | Output |
|---|---:|---|
| `write_workbook()` → csv | 35.7 ms | 8.2 MB |
| `write_workbook()` → xls | 43.9 ms | 17.7 MB |
| `write_workbook()` → xlsb | 71.9 ms | 5.1 MB |
| `write_workbook()` → xlsx | 90.2 ms | 5.1 MB |
| `write_polars()` → xlsx | 470.2 ms | 5.1 MB |
| `write_pandas()` → xlsx | 468.5 ms | 5.1 MB |
| `polars.DataFrame.write_excel()` | 4,707.2 ms | 5.6 MB |
| `pandas.DataFrame.to_excel()` | 7,274.0 ms | 5.5 MB |

The two DataFrame comparisons are matched work — same DataFrame in, xlsx out both times:
`write_polars()` is ~10.0x faster than polars' own `write_excel()`, and `write_pandas()` ~15.5x
faster than `to_excel()`. Both of ours pay a conversion the raw path does not: the DataFrame goes
through Arrow and then a Python list before reaching the native columns, which is most of the gap
between the 470 ms row and the 90 ms one. Handing `write_workbook()` buffers that are already
columnar — what `parse_typed()` returns — skips all of it.

`write_workbook(xlsx)` at 90.2 ms lands within ~5 ms of the C++ binding's
`xl::write_columns` on the same 14 columns (85.1–85.9 ms), which is the expected result: both are thin wrappers
over the same `xl_write_typed` call, and neither adds work per cell.

`xls` being both the largest file and faster than xlsb and xlsx is not a paradox — BIFF8 writes
fixed-width records with no compression, so it trades 3.5x the bytes for less work per cell.

## Notes

- A `Workbook` is **not** thread-safe. Use one per thread.
- Empty cells are skipped, so `cell.column` may skip indices. Do not assume `row[i].column == i`.
- The ABI is documented in `src/ExcelReader.Native/include/excelreader.h`.
