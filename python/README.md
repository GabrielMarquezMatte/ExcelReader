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
