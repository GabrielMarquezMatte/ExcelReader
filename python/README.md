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

#### Guessing a schema

Writing the `ColumnSpec` list by hand means already knowing every column's name and type. When you
don't, `infer_schema()` samples the sheet and guesses one for you:

```python
with open_workbook("sales.xlsb") as workbook:
    schema = workbook.infer_schema()   # header_row=1, sample_size=100 by default
    table = workbook.parse_typed(schema)
```

Each column's type comes from the `CellType` Excel already stored for its sampled cells — not text
sniffing — so it costs nothing beyond the sample and is exact for XLSX/XLSB/XLS. A column with a real
mix of kinds, only formula/error results, or nothing sampled falls back to `ColumnType.STRING`;
`nullable` is set when any sampled row left the column empty. CSV cells carry no such type tag, so
every CSV column is guessed `ColumnType.STRING` — inspect the result (or just try parsing) before
trusting it, especially past the sample.

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

## Benchmarks

`benchmarks/bench_read.py` and `benchmarks/bench_write.py` over
`tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb` (65,535 data rows, 14 columns), the same
fixture the .NET, C++ and Rust suites use. Measured on Windows 10 (22H2), 16 logical CPUs
@ 3.39 GHz, CPython 3.14.5, 10 runs each (medians shown; `min` is in the scripts' own output).

### Reading

| API | Median | What it produces |
|---|---:|---|
| `parse_typed()` | 54.1 ms | typed columnar buffers, converted natively |
| `to_arrow()` | 54.3 ms | the same parse, handed to pyarrow zero-copy |
| `to_polars()` | 56.1 ms | typed columnar DataFrame, schema inferred |
| `polars.read_excel()` | 137.7 ms | typed columnar DataFrame, types inferred |
| `read_all_columnar()` | 508.5 ms | raw columnar cells, no per-cell Python objects |
| `rows()` | 1,054.6 ms | one `Cell` object per cell, streamed per row |
| `read_all()` | 1,616.3 ms | one `Cell` object per cell, all at once |

Only the `to_polars()` / `polars.read_excel()` pair is a like-for-like comparison, and even that one
is loose: both produce a typed columnar DataFrame with inferred types, but the inference rules are
not identical. The rows above it produce different things and are listed to show what each API
costs, not to rank them — `read_all()` is 30x slower than `parse_typed()` because it materializes
917,504 Python objects, which is the price of that shape, not a slow parser.

### Writing

| API | Median | Output |
|---|---:|---|
| `write_workbook()` → xls | 45.6 ms | 17.7 MB |
| `write_workbook()` → csv | 64.2 ms | 8.2 MB |
| `write_workbook()` → xlsb | 74.1 ms | 5.2 MB |
| `write_workbook()` → xlsx | 129.3 ms | 5.1 MB |
| `write_polars()` → xlsx | 507.5 ms | 5.1 MB |
| `write_pandas()` → xlsx | 514.6 ms | 5.1 MB |
| `polars.DataFrame.write_excel()` | 4,487.8 ms | 5.6 MB |
| `pandas.DataFrame.to_excel()` | 6,972.0 ms | 5.5 MB |

The two DataFrame comparisons are matched work — same DataFrame in, xlsx out both times:
`write_polars()` is ~8.8x faster than polars' own `write_excel()`, and `write_pandas()` ~13.5x
faster than `to_excel()`. Both of ours pay a conversion the raw path does not: the DataFrame goes
through Arrow and then a Python list before reaching the native columns, which is most of the gap
between the 507 ms row and the 129 ms one. Handing `write_workbook()` buffers that are already
columnar — what `parse_typed()` returns — skips all of it.

`write_workbook(xlsx)` at 129.3 ms lands within a couple of milliseconds of the C++ binding's
`xl::write_columns` on the same 14 columns, which is the expected result: both are thin wrappers
over the same `xl_write_typed` call, and neither adds work per cell.

`xls` being the fastest and largest is not a paradox — BIFF8 writes fixed-width records with no
compression, so it trades 3.5x the bytes for less work per cell.

## Notes

- A `Workbook` is **not** thread-safe. Use one per thread.
- Empty cells are skipped, so `cell.column` may skip indices. Do not assume `row[i].column == i`.
- The ABI is documented in `src/ExcelReader.Native/include/excelreader.h`.
