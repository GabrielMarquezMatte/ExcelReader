# excelreader (Python)

Read XLSX, XLSB, XLS and CSV through ExcelReader's NativeAOT library. No .NET runtime required —
the shared library is self-contained. Reading only; writing is not exposed yet.

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
[docs/NATIVE_BASELINE.md](../docs/NATIVE_BASELINE.md).

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

## Notes

- A `Workbook` is **not** thread-safe. Use one per thread.
- Empty cells are skipped, so `cell.column` may skip indices. Do not assume `row[i].column == i`.
- The ABI is documented in `src/ExcelReader.Native/include/excelreader.h`.
