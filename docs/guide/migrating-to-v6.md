# Migrating from v5 to v6

What changed in the reading API and how to port v5 code.
See also [reading.md](reading.md) for the full reading guide.

## What changed

A v5 reader owned one "current sheet" that you moved with `MoveToSheet`; in v6 the reader became a workbook that hands out sheets, and each sheet is read on its own. Because a sheet no longer shares a cursor with its workbook, the sheets of one workbook can now be read in parallel.

## Types

| v5 | v6 |
|---|---|
| `XlsxReader`, `XlsbReader`, `XlsReader` | `XlsxWorkbook`, `XlsbWorkbook`, `XlsWorkbook` |
| `XlsxReader.Enumerator` (and the other two) | `XlsxWorkbook.Enumerator` (and `XlsbWorkbook.Enumerator`, `XlsWorkbook.Enumerator`) |
| `IExcelRowReader` | `IExcelWorkbook` and `IExcelSheet` |
| `IExcelRowReader<TEnumerator>` | `IExcelSheet<TEnumerator>` |
| `ExcelSheet`, `reader.Sheets()` | `XlsxSheet` / `XlsbSheet` / `XlsSheet` / `CsvSheet`, `workbook.Sheets` |
| `CsvReader` | `CsvReader` (same name; now a workbook with one sheet) |

An explicit `ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>` declaration becomes `ExcelEnumerable<T, XlsSheet, XlsWorkbook.Enumerator>` (likewise for XLSX and XLSB).

A sheet is a small value (`XlsxSheet` and its siblings are `readonly struct`s) that holds no resources: it is a workbook plus an index. `IExcelSheet` is the format-agnostic view of one, with `Index`, `Name`, `Visibility` and `GetEnumerator()`.

## Members

| v5 | v6 |
|---|---|
| `reader.GetEnumerator()`, `foreach (Row row in reader)` | `workbook.FirstSheet.GetEnumerator()`, `foreach (Row row in workbook.FirstSheet)` |
| `reader.MoveToSheet(i)` then read | read `workbook.Sheets[i]` (`workbook.SheetAt(i)` on an `IExcelWorkbook`) |
| `reader.TryMoveToSheet(name)` | `workbook.TryGetSheet(name, out var sheet)` |
| `reader.SheetName`, `reader.SheetNameAt(i)` | `sheet.Name`, `workbook.Sheets[i].Name` |
| `reader.SheetVisibility`, `reader.SheetVisibilityAt(i)` | `sheet.Visibility` |
| `reader.SheetCount` | `workbook.Sheets.Count` (or `workbook.SheetCount` on `IExcelWorkbook`) |
| `parser.Parse(reader)` | `parser.Parse(workbook.FirstSheet)` |
| `Excel.InferSchema(reader)` | `Excel.InferSchema(sheet)` |
| `new ExcelDataReader(reader)` | `new ExcelDataReader(sheet)` |
| `reader.ToArrowRecordBatch()` | `sheet.ToArrowRecordBatch()` |
| `Excel.Open(...)` returning `IExcelRowReader` | `Excel.Open(...)` returning `IExcelWorkbook`; use `FirstSheet`, `SheetAt(i)` and `TryGetSheet` |

`FirstSheet` is the same sheet as `Sheets[0]` and `SheetAt(0)`.

## Before and after

Reading the first sheet:

```csharp
// v5
using IExcelRowReader reader = Excel.Open("book.xlsx");
foreach (Row row in reader)
{
    Console.WriteLine(row[0].GetString());
}

// v6
using IExcelWorkbook workbook = Excel.Open("book.xlsx");
foreach (Row row in workbook.FirstSheet)
{
    Console.WriteLine(row[0].GetString());
}
```

Walking every sheet:

```csharp
// v5
using IExcelRowReader reader = Excel.Open("book.xlsx");
for (int i = 0; i < reader.SheetCount; i++)
{
    reader.MoveToSheet(i);
    Console.WriteLine(reader.SheetName);
    foreach (Row row in reader)
    {
        Console.WriteLine(row[0].GetString());
    }
}

// v6
using IExcelWorkbook workbook = Excel.Open("book.xlsx");
for (int i = 0; i < workbook.SheetCount; i++)
{
    IExcelSheet sheet = workbook.SheetAt(i);
    Console.WriteLine(sheet.Name);
    foreach (Row row in sheet)
    {
        Console.WriteLine(row[0].GetString());
    }
}
```

When you open a format directly (`Excel.FromXlsxFile` and the like), `workbook.Sheets` is a typed list: `foreach (XlsxSheet sheet in workbook.Sheets)` walks every sheet without boxing.

v5's `ExcelSheet` had a positional `Deconstruct`, so `foreach (var (i, name) in reader.Sheets())` worked. The v6 sheets do not deconstruct; read the properties instead:

```csharp
foreach (XlsxSheet sheet in workbook.Sheets)
{
    Console.WriteLine($"{sheet.Index}: {sheet.Name}");
}
```

Parsing a named sheet:

```csharp
// v5
using IExcelRowReader reader = Excel.Open("changes.xlsx");
if (reader.TryMoveToSheet("Changes"))
{
    foreach (ChangeRow item in ExcelParser.FromAttributes<ChangeRow>().Parse(reader))
    {
        Console.WriteLine(item.File);
    }
}

// v6
using IExcelWorkbook workbook = Excel.Open("changes.xlsx");
if (workbook.TryGetSheet("Changes", out var sheet))
{
    foreach (ChangeRow item in ExcelParser.FromAttributes<ChangeRow>().Parse(sheet))
    {
        Console.WriteLine(item.File);
    }
}
```

A CSV file:

```csharp
// v5
using CsvReader csv = Excel.FromCsvFile("book.csv");
foreach (Row row in csv)
{
    Console.WriteLine(row[0].GetString());
}

// v6
using CsvReader csv = Excel.FromCsvFile("book.csv");
foreach (Row row in csv.FirstSheet)
{
    Console.WriteLine(row[0].GetString());
}
```

## Reading sheets in parallel

Each sheet of a workbook can be read on its own thread:

```csharp
using XlsxWorkbook workbook = Excel.FromXlsxFile(path);

Parallel.For(0, workbook.Sheets.Count, i =>
{
    foreach (Row row in workbook.Sheets[i])
    {
        // consume
    }
});
```

The contract:

- A workbook is safe to share between threads once it is open: reading its metadata, getting sheets and creating enumerators all work concurrently.
- An enumerator belongs to one thread at a time. Using one enumerator from two threads at once is undefined behavior and is not checked.
- You may have several enumerators open at once, including several over the same sheet. The exception is a CSV over a non-seekable stream, which serves one.
- Do not use, reposition or dispose a stream or buffer you passed in until the workbook and every enumerator obtained from it are disposed.

## Behavior changes

- Getting a workbook's sheets (`Sheets`, `FirstSheet`, `SheetAt`, `TryGetSheet`) or opening an enumerator after `Dispose` throws `ObjectDisposedException` (since 5.2.0). `SheetCount` and `IsDate1904` still answer.
- Disposing a workbook does not close its file or stream while an enumerator from it is still undisposed; the resources are released when the last enumerator is disposed (since 5.2.0).
- A failed XLSX shared-string load fails every later enumeration the same way (since 5.2.0).
- A cancelled XLSX shared-string load is retried by the next enumeration, and the bytes it had read still count toward `MaxTotalDecompressedBytes` (since 5.2.0).
- A CSV over a seekable stream is read from the position the stream had when the `CsvReader` was created, by any number of enumerators at once. The stream's position afterwards is unspecified: file streams and memory streams with an exposed buffer are left untouched, any other seekable stream may be moved (since 5.2.0).
- A CSV over a non-seekable stream still allows one enumeration (since 5.2.0).

## Native bindings

The C ABI changes with the library: `XL_ABI_VERSION` is 6. A workbook handle no longer has a current
sheet. Every read takes a zero-based sheet index, and rows are read through a cursor opened per
sheet, so one handle serves several threads.

| v5 | v6 |
|---|---|
| `xl_open_file(path, len, format, &wb)` | `xl_open_file(path, len, format, NULL, &wb)` |
| `xl_open_file_ex`, `xl_open_memory_ex` | `xl_open_file`, `xl_open_memory` with the options pointer |
| `xl_move_to_sheet(wb, i)` | pass `i` to the read function |
| `xl_sheet_name(wb, ...)` | `xl_sheet_name_at(wb, i, ...)` |
| none | `xl_sheet_visibility_at(wb, i, &v)`, `xl_sheet_index(wb, name, len, &i)` |
| `xl_next_row(wb, ...)` | `xl_rows_open(wb, i, &rows)`, `xl_rows_next(rows, ...)`, `xl_rows_close(rows)` |
| `xl_next_row_view(wb, &row)` | `xl_rows_next_view(rows, &row)` |
| `xl_read_all_blob(wb, ...)`, `xl_read_all_decoded(wb, &out)` | `xl_rows_read_all_blob(rows, ...)`, `xl_rows_read_all_decoded(rows, &out)` |
| `xl_rows` (the decoded result struct) | `xl_rows_decoded`; `xl_rows` is now the cursor handle |
| `xl_parse_typed(wb, specs, n, header, &t)` | `xl_parse_typed(wb, i, specs, n, header, 1, &t)` |
| `xl_parse_typed_ex(wb, specs, n, header, dop, &t)` | `xl_parse_typed(wb, i, specs, n, header, dop, &t)` |
| `xl_parse_arrow`, `xl_parse_arrow_ex` | `xl_parse_arrow(wb, i, specs, n, header, dop, &array, &schema)` |
| `xl_parse_arrow_stream(wb, ...)`, `xl_typed_reader_open(wb, ...)` | the same, with `i` after `wb` |
| `xl_infer_schema(wb, header, sample, &s)` | `xl_infer_schema(wb, i, header, sample, 0, &s)` |
| `xl_infer_schema_ex(wb, header, sample, flags, &s)` | `xl_infer_schema(wb, i, header, sample, flags, &s)` |

Behavior changes:

- `xl_parse_typed` and `xl_parse_arrow` used to read sequentially; pass `1` as
  `degree_of_parallelism` to keep that. `0` uses every processor on a CSV.
- Opening a second typed reader or Arrow stream on a workbook used to fail, and reading the workbook
  another way used to invalidate an open one. Both now work: each read is independent.
- A cursor, typed reader or Arrow stream keeps working after `xl_close` on its workbook.
