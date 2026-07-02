# ExcelReader

[![CI](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml)
[![CodeQL](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml)
[![Release](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml/badge.svg)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![Downloads](https://img.shields.io/nuget/dt/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![License](https://img.shields.io/github/license/GabrielMarquezMatte/ExcelReader.svg)](LICENSE)
[![Benchmarks](https://img.shields.io/badge/benchmarks-GitHub%20Pages-informational)](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/)

High-performance Excel reading and writing for .NET 10. Reads `.xlsx`, `.xlsb`, `.xls`, and `.csv`; writes `.xlsx`, `.xlsb`, and `.xls`.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, handles sparse cells, and includes writers for producing `.xlsx` (Open XML), `.xlsb` (BIFF12), and `.xls` (BIFF8) workbooks.

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10, AMD Ryzen 7 5700X, .NET 10.0.9 (SDK 11.0.100-preview.4). Each benchmark uses 50,000 rows.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | MiniExcel | Sylvan | SpreadCheetah |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 19.06 ms, 14.07 KB | 145.63 ms, 210.54 MB | 38.30 ms, 1.89 MB | — |
| Cell-by-cell read async | 19.94 ms, 16.09 KB | — | — | — |
| Typed row parsing | 20.02 ms, 3.88 MB | 161.86 ms, 199.31 MB | 59.35 ms, 10.47 MB | — |
| Typed row parsing async | 22.48 ms, 3.88 MB | — | 61.04 ms, 10.48 MB | — |
| Workbook writing | 16.26 ms, 2.11 MB | 282.20 ms, 85.14 MB | — | 14.58 ms, 12.32 MB |
| Workbook writing, shared strings | 16.36 ms, 2.12 MB | — | — | — |

ExcelReader is ~7.6x faster than MiniExcel for reads and ~17.4x faster for writes. Compared with Sylvan, it is ~2.0x faster for raw reads and ~3.0x faster for typed parsing, with substantially lower allocations. Shared strings are opt-in and are effectively even with inline strings in this repeated-text benchmark. For writing, SpreadCheetah is ~1.1x faster but allocates ~5.8x more memory; the XLSB writer beats both on speed and allocation (see below).

### XLSB (BIFF12)

| Scenario | ExcelReader |
|---|---:|
| Cell-by-cell read | 4.37 ms, 15.83 KB |
| Cell-by-cell read async | 5.04 ms, 18.27 KB |
| Typed row parsing | 9.37 ms, 3.88 MB |
| Typed row parsing async | 10.68 ms, 3.88 MB |
| Workbook writing | 10.55 ms, 2.04 MB |
| Workbook writing, shared strings | 9.90 ms, 2.05 MB |

XLSB is the fastest path in these results: raw reads are ~4.4x faster than XLSX reads, typed parsing is ~2.1x faster than XLSX parsing, and writing is ~1.5x faster than XLSX writing. The XLSB writer is also ~1.4x faster than SpreadCheetah while allocating ~83% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan |
|---|---:|---:|
| Cell-by-cell read | 4.66 ms, 59.21 KB | 5.25 ms, 1,717.73 KB |
| Cell-by-cell read async | 4.79 ms, 59.28 KB | — |
| Workbook writing | 6.70 ms, 12.37 MB | — |

XLS reading allocates ~29x less than Sylvan at comparable speed. The XLS writer is ~2.6x faster than the XLSX writer (6.70 ms vs 17.16 ms), but allocates more in this benchmark.

### CSV

Run separately on Windows 11, Intel Core i7-1355U, .NET 10.0.9 (SDK 10.0.301) — a different machine from the XLSX/XLSB/XLS results above, so compare ratios within this table, not absolute times across formats.

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv | CsvHelper |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 5.87 ms, 705 B | 6.11 ms, 4.83 KB | 3.87 ms, 1.69 MB | 23.85 ms, 16.27 MB |
| Cell-by-cell read async | 8.48 ms, 1.38 KB | — | — | — |
| Typed row parsing | 10.87 ms, 3.86 MB | 6.85 ms, 3.87 MB | 9.63 ms, 10.95 MB | 19.76 ms, 14.41 MB |
| Typed row parsing async | 14.34 ms, 3.86 MB | — | — | — |

For raw cell-by-cell reads, ExcelReader is competitive with Sep and allocates ~7x less; Sylvan.Data.Csv is faster but allocates ~2,400x more, and CsvHelper is ~4.1x slower and allocates ~23,000x more. For typed row parsing, `ExcelParser<T>.Parse(CsvReader)` uses a CSV-specialized projection (dense field binding, single-pass, native text date parsing) rather than the generic Excel pipeline: it now allocates the least of any library here (~3.86 MB, edging out Sep) and is faster than CsvHelper (~1.8x). Sep remains ~1.6x faster and Sylvan ~1.1x faster on raw parse time, but ExcelReader gives up far less than before while keeping the lowest allocation.

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

Pattern-match to the concrete type (`XlsxReader`, `XlsbReader`, `XlsReader`) only when you need format-specific methods such as `MoveToSheet`.

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

## Notes

- Reads `.xlsx`, `.xlsb` (BIFF12), `.xls` (BIFF8), and `.csv`; writes `.xlsx`, `.xlsb`, and `.xls`.
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
