# ExcelReader

[![CI](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml)
[![CodeQL](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml)
[![Release](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml/badge.svg)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![Downloads](https://img.shields.io/nuget/dt/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![License](https://img.shields.io/github/license/GabrielMarquezMatte/ExcelReader.svg)](LICENSE)
[![Benchmarks](https://img.shields.io/badge/benchmarks-GitHub%20Pages-informational)](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/)

High-performance Excel reading and writing for .NET 10. Reads `.xlsx`, `.xlsb`, `.xls`, and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, handles sparse cells, and includes writers for producing `.xlsx` (Open XML), `.xlsb` (BIFF12), and `.xls` (BIFF8) workbooks.

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.9 (SDK 11.0.100-preview.4). Generated-data benchmarks use 50,000 rows.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | MiniExcel | Sylvan | SpreadCheetah |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 11.652 ms, 12.49 KB | 141.968 ms, 209.00 MB | 35.183 ms, 1.89 MB | - |
| Cell-by-cell read async | 12.827 ms, 14.51 KB | - | - | - |
| Typed row parsing | 13.818 ms, 3.87 MB | 151.842 ms, 196.99 MB | 56.257 ms, 10.47 MB | - |
| Typed row parsing async | 15.135 ms, 3.88 MB | - | 59.240 ms, 10.48 MB | - |
| Workbook writing | 13.910 ms, 4.02 MB | 282.372 ms, 84.89 MB | - | 15.109 ms, 15.84 MB |
| Workbook writing, shared strings | 14.127 ms, 4.06 MB | - | - | - |

ExcelReader is ~12.2x faster than MiniExcel and ~3.0x faster than Sylvan for raw XLSX reads. For typed parsing, it is ~11.0x faster than MiniExcel and ~4.1x faster than Sylvan while allocating much less memory. For XLSX writing, ExcelReader is ~1.1x faster than SpreadCheetah and allocates ~3.9x less memory; it is ~20.3x faster than MiniExcel and allocates ~21x less.

### XLSB (BIFF12)

| Scenario | ExcelReader |
|---|---:|
| Cell-by-cell read | 4.699 ms, 14.23 KB |
| Cell-by-cell read async | 4.873 ms, 16.67 KB |
| Typed row parsing | 6.288 ms, 3.88 MB |
| Typed row parsing async | 6.726 ms, 3.88 MB |
| Workbook writing | 6.892 ms, 4.02 MB |
| Workbook writing, shared strings | 7.418 ms, 4.06 MB |

XLSB is the fastest generated Excel format in these results: raw reads are ~2.5x faster than XLSX reads, typed parsing is ~2.2x faster than XLSX parsing, and writing is ~2.0x faster than XLSX writing. The XLSB writer is also ~2.2x faster than SpreadCheetah on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan |
|---|---:|---:|
| Cell-by-cell read | 4.473 ms, 58.22 KB | 5.268 ms, 1,717.73 KB |
| Cell-by-cell read async | 4.503 ms, 58.29 KB | - |
| Workbook writing | 5.161 ms, 16.03 MB | - |

ExcelReader is ~1.2x faster than Sylvan for generated XLS reads while allocating ~29.5x less memory. The XLS writer is ~2.7x faster than the XLSX writer in this benchmark, but it allocates more because the BIFF8/OLE container is assembled in memory.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv | CsvHelper |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 4.453 ms, 232 B | 7.598 ms, 3.93 KB | 4.759 ms, 1.61 MB | 24.283 ms, 15.52 MB |
| Cell-by-cell read async | 4.837 ms, 352 B | - | - | - |
| Typed row parsing | 5.824 ms, 3.86 MB | 8.280 ms, 3.87 MB | 12.749 ms, 10.95 MB | 22.606 ms, 14.41 MB |
| Typed row parsing async | 6.091 ms, 3.86 MB | - | - | - |
| Row writing | 6.683 ms, 4.00 MB | 6.898 ms, 4.01 MB | 7.163 ms, 4.04 MB | 14.377 ms, 13.79 MB |

For raw CSV reads, ExcelReader is ~1.7x faster than Sep and ~1.1x faster than Sylvan.Data.Csv while allocating ~17x less than Sep and ~7,300x less than Sylvan.Data.Csv; CsvHelper is ~5.5x slower. For typed CSV parsing, ExcelReader is ~1.4x faster than Sep, ~2.2x faster than Sylvan.Data.Csv, and ~3.9x faster than CsvHelper while keeping allocations at ~3.86 MB. For CSV writing, ExcelReader is close to Sep and Sylvan.Data.Csv, and ~2.2x faster than CsvHelper; the extra allocation shown here is primarily the destination `MemoryStream` growth, not per-row writer state.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | Sylvan |
|---|---:|---:|
| XLSX | 70.166 ms, 34.09 KB | 195.927 ms, 644.23 KB |
| XLSM | 69.962 ms, 34.13 KB | 195.598 ms, 644.30 KB |
| XLSB | 23.693 ms, 40.58 KB | 29.870 ms, 338.54 KB |
| XLS | 12.215 ms, 189.85 KB | 18.475 ms, 185.90 KB |
| CSV | 6.065 ms, 232 B | 10.032 ms, 35.74 MB |

On this real-data workload, ExcelReader is ~2.8x faster than Sylvan for XLSX/XLSM, ~1.3x faster for XLSB, ~1.5x faster for XLS, and ~1.7x faster for CSV. Allocations stay under 41 KB for XLSX/XLSM/XLSB and at 232 B for CSV.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

## Install

```bash
dotnet add package ExcelReader.NET
```

## Read rows

```csharp
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

using var reader = Excel.FromFile("report.xlsx");

foreach (var row in reader)
{
    string name = row[0].GetString();

    if (row[1].TryParse(null, out int quantity))
    {
        Console.WriteLine($"{name}: {quantity}");
    }

    if (row[2].Type == CellType.Date && row[2].TryGetDateTime(reader.IsDate1904, out var date))
    {
        Console.WriteLine(date);
    }
}
```

## Open by auto-detecting the format

`Excel.Open` picks the reader from the file signature (XLSX/XLSB are ZIP packages, XLS is an OLE2 document) and returns an `IExcelRowReader`. The interface exposes `GetEnumerator()` directly, so no pattern-match is needed for basic row iteration.

```csharp
using ExcelReader.Core.Reader;

using IExcelRowReader reader = Excel.Open("report.xlsx"); // or report.xlsb / report.xls

foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}
```

Sheet navigation (`SheetCount`, `SheetName`, `MoveToSheet(index)`, `TryMoveToSheet(name)`) is available on `IExcelRowReader` itself, so you can walk every sheet without knowing the format:

```csharp
using IExcelRowReader reader = Excel.Open("report.xlsx");

for (int i = 0; i < reader.SheetCount; i++)
{
    reader.MoveToSheet(i);
    Console.WriteLine(reader.SheetName);
    foreach (var row in reader)
    {
        Console.WriteLine(row[0].GetString());
    }
}
```

CSV is exposed as a single, unnamed sheet (`SheetCount == 1`, `SheetName == ""`). Pattern-match to the concrete type only for reader-specific internals beyond this surface.

`OpenAsync` is the async counterpart. Both require a seekable stream (or a file path) so the signature can be read without consuming the input.

## Read asynchronously

`Row` and `Cell` are `ref struct` types, so async reading uses a manual loop instead of `await foreach`.
For XLSX files, the async reader buffers one row at a time and uses the same row parser as the sync reader, so sync and async reads stay behaviorally aligned while awaits happen only when more bytes are needed.

```csharp
using ExcelReader.Core.Reader;

await using var reader = await Excel.FromFileAsync("report.xlsx", cancellationToken);
await using var rows = await reader.GetAsyncEnumeratorAsync(cancellationToken);

while (await rows.MoveNextAsync())
{
    var row = rows.Current;
    Console.WriteLine(row[0].GetString());
}
```

## Parse typed rows

`ExcelParser<T>` maps worksheet columns to the public settable properties of `T`. Columns match on the property name, or on `[ExcelColumn("header")]` aliases — repeat the attribute to accept several headers. The first row is the header by default.

```csharp
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

public sealed class ChangeRow
{
    [ExcelColumn("file")]
    public string File { get; set; } = "";

    [ExcelColumn("lines_added")]
    public int LinesAdded { get; set; }
}

using var reader = Excel.FromFile("changes.xlsx");
var parser = new ExcelParser<ChangeRow>();

foreach (var item in parser.Parse(reader))
{
    Console.WriteLine($"{item.File}: +{item.LinesAdded}");
}
```

Built-in property types: `string`, `bool`, `DateTime`, `DateOnly`, `Guid`, every integral and floating type plus `decimal`, and `enum`s (matched by member name or numeric value). Each also works as a `Nullable<T>`. Empty cells leave the property at its default; an unparseable cell is skipped (keeps the default) unless the column is required. `T` needs no parameterless-constructor constraint, so models with `required` members are supported.

`Parse` and `ParseAsync` also accept the `IExcelRowReader` from `Excel.Open`, so you can parse without knowing the concrete format:

```csharp
using IExcelRowReader reader = Excel.Open("changes.xlsx"); // or .xlsb / .xls
foreach (var item in new ExcelParser<ChangeRow>().Parse(reader)) { /* ... */ }
```

## Parser configuration

Pass an `ExcelParserConfig` to control header handling and culture:

```csharp
using System.Globalization;
using ExcelReader.Core.Parser;

var config = new ExcelParserConfig
{
    HeaderRow = 1,                                   // 1-based row holding the headers
    ColumnNameComparer = StringComparer.OrdinalIgnoreCase,
    HeaderNormalization = HeaderNormalization.Trim | HeaderNormalization.CollapseSpaces,
    Culture = CultureInfo.GetCultureInfo("pt-BR"),   // parse "1.234,56" as 1234.56m
};

var parser = new ExcelParser<ChangeRow>(config);
```

`Culture` applies when parsing text-backed numeric/`Guid` cells (XLSX inline and shared strings); binary numeric cells (XLS/XLSB) carry a raw value and ignore it. `HeaderNormalization` flags (`Trim`, `CollapseSpaces`, `RemoveDiacritics`) are applied to both the sheet headers and the property names before matching.

## Required columns

Mark a property `[ExcelRequired]` to assert its column exists and carries a value:

```csharp
public sealed class Order
{
    [ExcelRequired]
    public int Id { get; set; }

    [ExcelRequired(AllowEmpty = true)]   // column must exist; blank cells allowed
    public string? Note { get; set; }
}
```

- A missing required header throws when the header row is read, listing every missing column.
- By default each data row must have a non-empty cell; the first blank throws, naming the column and row number. `AllowEmpty = true` relaxes this to column presence only.
- The check covers presence, not parseability — a present-but-malformed value does not throw here.

## Custom converters

For types the built-in parsers do not handle — money strings, custom formats, domain value objects — implement `IExcelCellConverter<T>` and attach it with `[ExcelConverter]`. `T` must be the property's exact type. One instance is created and reused across all rows, so converters must be stateless.

```csharp
using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.ValueObjects;

public sealed class BrlMoneyConverter : IExcelCellConverter<decimal>
{
    public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out decimal value)
    {
        string text = cell.GetString().Replace("R$", "", StringComparison.Ordinal).Trim();
        return decimal.TryParse(text, NumberStyles.Currency, CultureInfo.GetCultureInfo("pt-BR"), out value);
    }
}

public sealed class Invoice
{
    [ExcelConverter(typeof(BrlMoneyConverter))]
    public decimal Total { get; set; }
}
```

Return `false` to signal a parse failure (the property keeps its default). Empty cells are skipped before the converter runs.

## Write XLSX workbooks

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xlsx");
await using var workbook = await WorkbookWriter.CreateAsync(stream);

await workbook.StartAsync();
await using (var sheet = workbook.AddSheet("Summary"))
{
    await sheet.StartAsync();

    await using (var row = await sheet.StartRowAsync())
    {
        row.Write("Name");
        row.Write("Total");
        row.Write("Created");
    }

    await using (var row = await sheet.StartRowAsync())
    {
        row.Write("Q1");
        row.Write(42);
        row.Write(DateTime.UtcNow);
    }
}

await workbook.EndAsync();
```

By default, the XLSX writer emits inline strings to keep memory usage flat while rows stream out.
If your workbook repeats many strings and smaller files matter more than the extra lookup table, opt in to shared strings:

```csharp
await using var workbook = await WorkbookWriter.CreateAsync(stream, useSharedStrings: true);
```

## Read and write XLSB workbooks (BIFF12)

Use `Excel.FromXlsbFile`, `Excel.FromXlsb`, `Excel.FromXlsbFileAsync`, or `Excel.FromXlsbAsync` to open XLSB directly. For writing, use `XlsbWorkbookWriter`, `XlsbSheetWriter`, and `XlsbRowWriter`.

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xlsb");
await using var workbook = await XlsbWorkbookWriter.CreateAsync(stream);

await workbook.StartAsync();
await using (XlsbSheetWriter sheet = workbook.AddSheet("Summary"))
{
    await sheet.StartAsync();

    await using (XlsbRowWriter row = await sheet.StartRowAsync())
    {
        row.Write("Name");
        row.Write("Total");
        row.Write("Created");
    }
}

await workbook.EndAsync();
```

The XLSB writer also defaults to inline string cells. Pass `useSharedStrings: true` to deduplicate repeated text into `sharedStrings.bin`.

## Write XLS workbooks (BIFF8)

`XlsWorkbookWriter` emits a binary BIFF8 `.xls` file. The sheet and row APIs are synchronous; only the final `EndAsync` (which assembles and flushes the OLE container) is async. BIFF8 is capped at 65,536 rows × 256 columns per sheet.

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xls");
await using var workbook = XlsWorkbookWriter.Create(stream);

workbook.Start();
using (var sheet = workbook.AddSheet("Summary"))
{
    sheet.Start();

    using (var row = sheet.StartRow())
    {
        row.Write("Name");
        row.Write("Total");
        row.Write("Created");
    }

    using (var row = sheet.StartRow())
    {
        row.Write("Q1");
        row.Write(42);
        row.Write(DateTime.UtcNow);
    }
}

await workbook.EndAsync();
```

## Write typed records

The low-level writers above give you cell-by-cell control. When you just want to dump a collection of objects to a sheet, `WorkbookRecordWriter` writes a header row followed by one row per record, mapping each public readable property to a column. It is generic over the low-level interfaces, so the same API targets XLSX, XLSB, XLS, and CSV — pick the format with a `RecordWriter.Create*` factory.

```csharp
using ExcelReader.Core.Writer;

public sealed class Sale
{
    public string? Region { get; set; }
    public int Units { get; set; }
    public decimal Revenue { get; set; }
    public DateOnly Date { get; set; }
}

var sales = new[]
{
    new Sale { Region = "North", Units = 42, Revenue = 1234.50m, Date = new DateOnly(2026, 1, 2) },
    new Sale { Region = "South", Units = 17, Revenue = 512.00m,  Date = new DateOnly(2026, 1, 3) },
};

await using var stream = File.Create("sales.xlsx");
await using var writer = await RecordWriter.CreateXlsxAsync(stream);   // or CreateXlsbAsync / CreateXlsAsync / CreateCsv
await writer.WriteSheetAsync("Sales", sales);
```

Each `WriteSheetAsync` call targets a new sheet (a duplicate name throws), so one workbook can hold sheets of different record types. `RecordWriter.CreateCsv` is the exception: a CSV file is a single sheet, so a second `WriteSheetAsync` throws (the sheet name is ignored). An `IAsyncEnumerable<T>` overload streams records that are produced asynchronously. The written file round-trips straight back through `ExcelParser<T>` because the headers are the property names.

Column behavior mirrors the parser attributes:

- **`[ExcelColumn("Header")]`** — use a custom header instead of the property name (the first alias wins).
- **`[ExcelIgnore]`** — exclude a property from both writing and parsing (for computed/transient members).
- **`[ExcelConverter(typeof(MyConverter))]`** — if the converter also implements `IExcelCellWriter<T>`, it controls how the value is written, so a custom type round-trips through the same converter it reads with.

`DateTime` and `DateOnly` are written as Excel date serials; `TimeOnly` as a time-of-day fraction. Numeric properties become number cells; any other type is written as its `ToString()` text. (`CreateCsv` follows the CSV rules instead — see [Write CSV](#write-csv) — writing `DateTime`/`DateOnly` as ISO text and `TimeOnly` as a time-of-day fraction, all still round-tripping through `ExcelParser<T>`.)

## Read CSV

`CsvReader` streams RFC 4180 CSV (quoted fields, embedded delimiters/newlines, `""`-escaped quotes) through the same `Row`/`Cell` model as the Excel readers, so `ExcelParser<T>` works on it unchanged.

```csharp
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

using var reader = Excel.FromCsvFile("report.csv");

foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}

// Typed parsing works exactly like the Excel readers:
foreach (var item in new ExcelParser<ChangeRow>().Parse(reader))
{
    Console.WriteLine($"{item.File}: +{item.LinesAdded}");
}
```

`Excel.FromCsv`/`FromCsvFile`/`FromCsvAsync`/`FromCsvFileAsync` mirror the other formats' factory shape. Pass `CsvReaderOptions` to change the delimiter/quote character, supply a non-UTF-8 `Encoding` (transcoded to UTF-8 internally), or turn off BOM detection:

```csharp
var options = new CsvReaderOptions { Delimiter = (byte)';' };
using var reader = Excel.FromCsvFile("relatorio.csv", options);
```

Every CSV cell is text (`CellType.ExcelString`, or `CellType.Empty` for a blank field); at the reader level there is no binary numeric or date representation, so `Cell.TryGetDateTime`/`IsDate1904` (always `false` for CSV) do not apply. The typed parser, however, is CSV-specialized: `ExcelParser<T>.Parse(CsvReader)` parses `DateTime`/`DateOnly` columns directly from the cell text (ISO or culture format, honoring `Culture` — e.g. pt-BR `02/07/2026`), so no `[ExcelConverter]` is needed for dates. All the usual attributes work unchanged (`[ExcelColumn]` aliases, `[ExcelRequired]`, `[ExcelConverter]`), and a converter still takes precedence over the built-in date parsing. (Holding the reader as `IExcelRowReader` instead routes through the generic Excel pipeline, where dates use serial-number semantics — prefer the concrete `Parse(CsvReader)` overload for CSV.)

`Excel.Open`/`OpenAsync` do **not** auto-detect CSV — plain text has no magic-byte signature to sniff, so open CSV explicitly via `Excel.FromCsv*`.

## Write CSV

`CsvWriter` emits RFC 4180 CSV: no sheets, styles, or shared strings, so rows stream straight to the output.

```csharp
using ExcelReader.Core.Writer;

using var stream = File.Create("out.csv");
using var writer = CsvWriter.Create(stream);

using (CsvRowWriter row = writer.StartRow())
{
    row.Write("Name");
    row.Write("Total");
    row.Write("Created");
}

using (CsvRowWriter row = writer.StartRow())
{
    row.Write("Q1");
    row.Write(42);
    row.Write(DateTime.UtcNow);
}
```

Fields are quoted only when they contain the delimiter, quote character, `\r`, or `\n`; embedded quotes are doubled. `bool` writes as lowercase `true`/`false` and `DateTime`/`DateOnly` as round-trip ISO 8601 (`"O"`); `TimeOnly` as a time-of-day fraction — all matching what `ExcelParser<T>.Parse(CsvReader)` expects, so a file written by `CsvWriter` parses back without configuration. `Skip(count)` writes empty fields to keep column positions aligned (CSV has no sparse-cell concept). Pass `CsvWriterOptions` to change the delimiter/quote byte, mirroring `CsvReaderOptions`.

To dump a collection of typed records instead of writing cells by hand, use `RecordWriter.CreateCsv(stream)` — the same [record-writing API](#write-typed-records) as the Excel formats, restricted to a single sheet.

## Notes

- Reads `.xlsx`, `.xlsb` (BIFF12), `.xls` (BIFF8), and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.
- Reads one sheet at a time (XLSX/XLSB/XLS); use `MoveToSheet(index)` or `TryMoveToSheet(name)` to switch sheets. CSV has no sheets.
- Missing cells in sparse rows are exposed as empty cells.
- String conversion allocates only when you call `GetString()`.
- The XLSX scanner accepts the SpreadsheetML shapes commonly emitted by non-Excel producers, including single-quoted attributes, comments in `sheetData`, and CDATA text runs.
- Readers bound untrusted input by default: 512 MB total decompressed ZIP data, 32 MB per cell/row value buffer, and 128 MB for shared strings. Pass `ExcelReaderOptions` to the `Excel.From*`/`Excel.Open*` factories to tune these limits; set a limit to `0` to opt out and restore unlimited behavior for that limit. `CsvReader` has its own `CsvReaderOptions.MaxCellBytes` (default 32 MB) for the same purpose.
- The XLSX writer emits a compact workbook with strings, numbers, booleans, dates, and blank cells; shared strings are opt-in.
- The XLSB writer emits BIFF12 workbook parts inside the standard XLSB ZIP package; shared strings are opt-in.
- The XLS writer buffers records in memory and assembles the OLE container at `EndAsync`; choose it when write throughput matters more than peak allocation.

## Build

```bash
dotnet restore ExcelReader.slnx
dotnet build ExcelReader.slnx --configuration Release
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj --configuration Release
```

## License

ExcelReader is licensed under the MIT License. See [LICENSE](LICENSE).
