# ExcelReader

[![CI](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml)
[![CodeQL](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml)
[![Release](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml/badge.svg)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![Downloads](https://img.shields.io/nuget/dt/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![License](https://img.shields.io/github/license/GabrielMarquezMatte/ExcelReader.svg)](LICENSE)
[![Benchmarks](https://img.shields.io/badge/benchmarks-GitHub%20Pages-informational)](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/)

High-performance Excel reading and writing for .NET 10. Reads `.xlsx` and `.xls`; writes both formats.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, handles sparse cells, and includes writers for producing `.xlsx` (Open XML) and `.xls` (BIFF8) workbooks.

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10, AMD Ryzen 7 5700X, .NET 10.0.9. Each benchmark uses 50,000 rows.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | MiniExcel | Sylvan |
|---|---:|---:|---:|
| Cell-by-cell read | 24.88 ms, 13.41 KB | 198.03 ms, 375.74 MB | 42.84 ms, 347.42 KB |
| Typed row parsing | 29.34 ms, 3.87 MB | 228.63 ms, 400.05 MB | 70.67 ms, 10.49 MB |
| Workbook writing | 35.02 ms, 6.34 MB | 288.16 ms, 85.14 MB | — |

ExcelReader is ~8x faster than MiniExcel for reads and writes. Compared with Sylvan, ~1.7x faster for raw reads and ~2.4x faster for typed parsing, with substantially lower allocations.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan |
|---|---:|---:|
| Cell-by-cell read | 4.775 ms, 61.34 KB | 5.469 ms, 1,717.98 KB |
| Workbook writing | 7.004 ms, 1.92 MB | — |

XLS reading allocates ~28x less than Sylvan at comparable speed. The XLS writer is **5x faster** than the XLSX writer (7.00 ms vs 34.73 ms) and allocates **3.3x less** (1.92 MB vs 6.34 MB).

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

`Excel.Open` picks the reader from the file signature (XLSX is a ZIP, XLS is an OLE2 document) and returns an `IExcelReader`. Pattern-match to the concrete reader to enumerate rows.

```csharp
using ExcelReader.Core.Reader;

using IExcelReader reader = Excel.Open("report.xlsx"); // or report.xls

if (reader is XlsxReader xlsx)
{
    using var rows = xlsx.GetEnumerator();
    while (rows.MoveNext())
    {
        Console.WriteLine(rows.Current[0].GetString());
    }
}
else if (reader is XlsReader xls)
{
    using var rows = xls.GetEnumerator();
    while (rows.MoveNext())
    {
        Console.WriteLine(rows.Current[0].GetString());
    }
}
```

`OpenAsync` is the async counterpart. Both require a seekable stream (or a file path) so the signature can be read without consuming the input.

## Read asynchronously

`Row` and `Cell` are `ref struct` types, so async reading uses a manual loop instead of `await foreach`.

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

Use `[ExcelColumn]` when the spreadsheet header does not match the property name.

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

## Notes

- Reads `.xlsx` and `.xls` (BIFF8) files; writes both formats.
- Reads one sheet at a time; use `MoveToSheet(index)` or `TryMoveToSheet(name)` to switch sheets.
- Missing cells in sparse rows are exposed as empty cells.
- String conversion allocates only when you call `GetString()`.
- The XLSX writer emits a compact workbook with strings, numbers, booleans, dates, and blank cells.
- The XLS writer buffers records in memory and assembles the OLE container at `EndAsync`; choose it when write throughput matters more than peak allocation.

## Build

```bash
dotnet restore ExcelReader.slnx
dotnet build ExcelReader.slnx --configuration Release
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj --configuration Release
```

## License

ExcelReader is licensed under the MIT License. See [LICENSE](LICENSE).
