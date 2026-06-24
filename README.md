# ExcelReader

High-performance XLSX reading, typed row parsing, and minimal workbook writing for .NET 10.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, handles sparse cells, and includes a small writer for producing simple `.xlsx` workbooks.

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10, AMD Ryzen 7 5700X, .NET 10.0.9. Each benchmark uses 50,000 rows and compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | MiniExcel | Sylvan |
|---|---:|---:|---:|
| Cell-by-cell read | 24.88 ms, 13.41 KB | 198.03 ms, 375.74 MB | 42.84 ms, 347.42 KB |
| Typed row parsing | 29.34 ms, 3.87 MB | 228.63 ms, 400.05 MB | 70.67 ms, 10.49 MB |
| Workbook writing | 35.43 ms, 6.34 MB | 288.16 ms, 85.14 MB | - |

In these runs, ExcelReader was about 8x faster than MiniExcel for raw reads, 7.8x faster for typed parsing, and 8.1x faster for writing. Compared with Sylvan, ExcelReader was about 1.7x faster for raw reads and 2.4x faster for typed parsing, while allocating substantially less memory.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

## Install

```bash
dotnet add package ExcelReader.Core
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

## Write workbooks

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

## Notes

- Supports `.xlsx` files, not legacy `.xls`.
- Reads one sheet at a time; use `MoveToSheet(index)` or `TryMoveToSheet(name)` to switch sheets.
- Missing cells in sparse rows are exposed as empty cells.
- String conversion allocates only when you call `GetString()`.
- The writer emits a compact workbook with strings, numbers, booleans, dates, and blank cells.

## Build

```bash
dotnet restore ExcelReader.slnx
dotnet build ExcelReader.slnx --configuration Release
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj --configuration Release
```

## License

ExcelReader is licensed under the MIT License. See [LICENSE](LICENSE).
