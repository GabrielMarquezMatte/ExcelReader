# ExcelReader

[![CI](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml)
[![CodeQL](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml)
[![Release](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml/badge.svg)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![Downloads](https://img.shields.io/nuget/dt/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![License](https://img.shields.io/github/license/GabrielMarquezMatte/ExcelReader.svg)](LICENSE)
[![Benchmarks](https://img.shields.io/badge/benchmarks-GitHub%20Pages-informational)](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/)

High-performance Excel reading and writing for .NET 10. Reads `.xlsx`, `.xlsb`, `.xls`, and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, handles sparse cells, and includes writers for producing `.xlsx` (Open XML), `.xlsb` (BIFF12), and `.xls` (BIFF8) workbooks. The library also supports opening workbook data directly from in-memory buffers without requiring a stream, which makes it convenient for API and network-based scenarios. `Excel.FromCsv(ReadOnlyMemory<byte>)` and `Excel.FromXls(ReadOnlyMemory<byte>)` now accept caller-owned buffers directly, and `Excel.Open(ReadOnlyMemory<byte>)` routes XLS workbooks through the same true-memory path instead of wrapping the bytes in `MemoryStream`.

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.9 (SDK 11.0.100-preview.4). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [`tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/results`](tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/results).

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | MiniExcel | Sylvan | SpreadCheetah |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 11.114 ms, 12.59 KB | 145.614 ms, 209.00 MB | 36.697 ms, 1.89 MB | - |
| Cell-by-cell read async | 11.457 ms, 14.72 KB | - | - | - |
| Typed row parsing | 15.691 ms, 3.88 MB | 156.414 ms, 197.78 MB | 57.585 ms, 10.47 MB | - |
| Typed row parsing async | 15.250 ms, 3.88 MB | - | 60.788 ms, 10.48 MB | - |
| Workbook writing | 14.611 ms, 4.02 MB | 286.522 ms, 84.89 MB | - | 15.763 ms, 15.84 MB |
| Workbook writing, shared strings | 15.002 ms, 4.06 MB | - | - | - |

ExcelReader is ~13.1x faster than MiniExcel and ~3.3x faster than Sylvan for raw XLSX reads, allocating ~17,000x and ~154x less respectively. For typed parsing, it is ~10.0x faster than MiniExcel and ~3.7x faster than Sylvan. For XLSX writing, ExcelReader is ~1.1x faster than SpreadCheetah and allocates ~3.9x less memory; it is ~19.6x faster than MiniExcel and allocates ~21x less.

### XLSB (BIFF12)

| Scenario | ExcelReader |
|---|---:|
| Cell-by-cell read | 4.659 ms, 14.32 KB |
| Cell-by-cell read async | 4.719 ms, 16.88 KB |
| Typed row parsing | 6.909 ms, 3.88 MB |
| Typed row parsing async | 7.036 ms, 3.88 MB |
| Workbook writing | 7.381 ms, 4.02 MB |
| Workbook writing, shared strings | 7.142 ms, 4.06 MB |

XLSB is the fastest generated Excel format in these results: raw reads are ~2.4x faster than XLSX reads, typed parsing is ~2.3x faster than XLSX parsing, and writing is ~2.0x faster than XLSX writing. The XLSB writer is also ~2.1x faster than SpreadCheetah on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan |
|---|---:|---:|
| Cell-by-cell read | 4.362 ms, 2.89 KB | 5.372 ms, 1,717.73 KB |
| Cell-by-cell read async | 4.353 ms, 2.96 KB | - |
| Workbook writing | 5.654 ms, 16.03 MB | - |

ExcelReader is ~1.2x faster than Sylvan for generated XLS reads while allocating ~594x less memory. The XLS writer is ~2.6x faster than the XLSX writer in this benchmark, but it allocates more because the BIFF8/OLE container is assembled in memory.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv | CsvHelper |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 5.414 ms, 232 B | 8.106 ms, 3.93 KB | 4.714 ms, 1.61 MB | 24.883 ms, 15.52 MB |
| Cell-by-cell read async | 4.913 ms, 352 B | - | - | - |
| Typed row parsing | 6.277 ms, 3.86 MB | 9.744 ms, 3.87 MB | 12.717 ms, 10.95 MB | 24.501 ms, 14.41 MB |
| Typed row parsing async | 6.435 ms, 3.86 MB | - | - | - |
| Row writing | 6.813 ms, 4.00 MB | 7.181 ms, 4.01 MB | 7.128 ms, 4.04 MB | 14.765 ms, 13.79 MB |

For raw CSV reads, ExcelReader is ~1.5x faster than Sep while allocating ~17x less; Sylvan.Data.Csv is marginally faster here (~13%) but allocates ~7,280x more (1.61 MB vs 232 B) — worth it only if raw wall-clock time matters more than memory pressure; CsvHelper is ~4.6x slower and allocates ~70,000x more. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.6x faster than Sep, ~2.0x faster than Sylvan.Data.Csv, and ~3.9x faster than CsvHelper, with the lowest allocation of the group. For CSV writing, ExcelReader, Sep, and Sylvan.Data.Csv are all within ~5% of each other, and ~2.2x faster than CsvHelper; the ~4 MB shown across the first three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 64.866 ms, 34.52 KB | 44.077 ms, 124.29 KB | 201.885 ms, 644.23 KB |
| XLSM | 66.762 ms, 34.56 KB | 42.765 ms, 123.56 KB | 196.790 ms, 644.30 KB |
| XLSB | 23.300 ms, 40.66 KB | 14.111 ms, 60.30 KB | 30.478 ms, 338.54 KB |
| XLS | 12.535 ms, 9.97 KB | n/a | 18.868 ms, 185.90 KB |
| CSV | 6.172 ms, 232 B | n/a | 10.750 ms, 35.75 MB |

On this real-data workload, ExcelReader is ~3.1x faster than Sylvan for XLSX, ~2.9x faster for XLSM, ~1.3x faster for XLSB, ~1.5x faster for XLS, and ~1.7x faster for CSV — allocating ~18.7x less for XLSX/XLSM, ~8.3x less for XLSB, ~18.6x less for XLS, and ~161,000x less for CSV (232 B vs 35.75 MB). The prefetch column is the opt-in [`PrefetchDecompression`](#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads (latest)

The latest real-data benchmark also measures the in-memory path for workbook content loaded directly into memory.

| Method | Mean | Error | StdDev | Allocated |
|---|---:|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 65.86 ms | 0.446 ms | 0.396 ms | 4.98 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 42.88 ms | 0.198 ms | 0.166 ms | 71.33 KB |
| Xlsm_ExcelReader_Memory | 66.48 ms | 0.597 ms | 0.558 ms | 4.98 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 43.59 ms | 0.574 ms | 0.509 ms | 72.99 KB |
| Xlsb_ExcelReader_Memory | 24.98 ms | 0.247 ms | 0.231 ms | 14.43 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 19.40 ms | 0.256 ms | 0.200 ms | 33.06 KB |
| Xls_ExcelReader_Memory | 12.40 ms | 0.132 ms | 0.117 ms | 1.99 KB |
| Csv_ExcelReader_Memory | 6.03 ms | 0.084 ms | 0.074 ms | 232 B |

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 57.48 ms, 756.72 KB | 37.88 ms, 826.43 KB | 184.77 ms, 17,825.41 KB |
| XLSB | 37.23 ms, 756.84 KB | 26.80 ms, 822.35 KB | 62.79 ms, 17,799.79 KB |

Both formats now handle this well: ~3.2x and ~2.3x faster than Sylvan for XLSX and XLSB respectively, at roughly ~24x and ~23x less memory, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~36x on this workload.

### Typed record writing

`WorkbookRecordWriter`/`RecordWriter` (the header-plus-one-row-per-object API — see [Write typed records](#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 15.399 ms | 4.02 MB |
| XLSB | 8.284 ms | 4.02 MB |
| XLS | 5.556 ms | 4.03 MB |
| CSV | 7.468 ms | 4.00 MB |

Relative ordering matches the lower-level writers above (XLS fastest, then XLSB, then CSV, then XLSX) — the record-mapping layer adds negligible overhead over hand-written cell-by-cell writes.

### Ref struct typed parsing (zero-copy)

`RefParser.ParseNamed<T>` (see [Parse into a ref struct](#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 15.69 ms | 3.88 MB |
| `struct` (`ExcelParser<T>`) | 14.51 ms | 1.59 MB |
| `ref struct` + span binding (`RefParser.ParseNamed<T>`) | 13.11 ms | 13.07 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.7% less than the `class` baseline — and is ~16% faster, since there's no per-row model allocation and no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### Cold start

First use of `ExcelParser<T>`/`RecordWriter` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 40.98 ms | 29.56 KB |
| First typed record write | 20.94 ms | 79.41 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

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

Every reader supports `await foreach`. For XLSX files, the async reader buffers one row at a time and uses the same row parser as the sync reader, so sync and async reads stay behaviorally aligned while awaits happen only when more bytes are needed.

```csharp
using ExcelReader.Core.Reader;

await using var reader = await Excel.FromFileAsync("report.xlsx", cancellationToken);

await foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}
```

`await foreach` binds to the reader's `GetAsyncEnumerator()` by pattern — the sheet is opened synchronously and only each row advance is awaited. Because `Row` and `Cell` are `ref struct` types, the current row cannot be held across an `await` inside the loop body: read its cells (or copy the values out) before awaiting anything else.

When you need the sheet *opened* asynchronously too (e.g. the first read touches the network), or you need to `await` while a row is in scope, drive the enumerator manually via `GetAsyncEnumeratorAsync`, which awaits the open and threads the cancellation token:

```csharp
await using var reader = await Excel.FromFileAsync("report.xlsx", cancellationToken);
await using var rows = await reader.GetAsyncEnumeratorAsync(cancellationToken);

while (await rows.MoveNextAsync())
{
    var row = rows.Current;
    Console.WriteLine(row[0].GetString());
}
```

`await foreach` does not accept `.WithCancellation(ct)`: `Row` being a `ref struct` rules out `IAsyncEnumerable<Row>`, so the loop binds to the pattern rather than the interface. Pass the token at open time (as above), or use the manual `GetAsyncEnumeratorAsync(ct)` loop.

## Prefetch decompression (XLSX/XLSB)

XLSX and XLSB are ZIP-backed, and inflating a sheet's compressed bytes competes for
wall-clock time with parsing it. `ExcelReaderOptions.PrefetchDecompression` overlaps the
two: a background thread inflates ahead while the calling thread parses. It is **opt-in,
defaults to `false`**, and only affects XLSX/XLSB — XLS and CSV have nothing to
decompress, so the option is silently ignored for them.

```csharp
var options = new ExcelReaderOptions { PrefetchDecompression = true };
using var reader = Excel.FromFile("report.xlsx", options);

foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}
```

Measured across both read benchmarks (see [Real data reads](#real-data-reads) and
[String-heavy reads](#string-heavy-reads)), on 65K-row workbooks:

| Workload | Default | `PrefetchDecompression = true` | Gain |
|---|---:|---:|---:|
| XLSX, real data | 64.9 ms | 44.1 ms | 32% |
| XLSM, real data | 66.8 ms | 42.8 ms | 36% |
| XLSB, real data | 23.3 ms | 14.1 ms | 39% |
| XLSX, string-heavy | 57.5 ms | 37.9 ms | 34% |
| XLSB, string-heavy | 37.5 ms | 30.5 ms | 19% |

The gain tracks how much of a read is decompression rather than parsing, so it is largest
on XLSB with numeric data (where inflate dominates) and smallest on string-heavy XLSB
(where string materialization does). Allocations rise from the producer task and the
pooled decompression buffers — on the real-data corpus, roughly 35 KB to 124 KB for XLSX
and 41 KB to 60 KB for XLSB — and neither path triggers a garbage collection.

Do **not** enable it for concurrent server workloads: a caller already reading many files
in parallel is CPU-saturated, and an extra background thread per read only doubles thread
demand for no gain. It's meant for single-file batch processing.

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

## Parse into a ref struct (zero-copy)

`RefParser.ParseNamed<T>` (.NET 9+) targets a `ref struct` model instead of a class/struct — same attribute-driven column matching as `ExcelParser<T>` (`[ExcelColumn]`, `[ExcelRequired]`, `[ExcelConverter]`), but a `ReadOnlySpan<byte>` property binds directly to the cell's raw bytes instead of allocating a `string`:

```csharp
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

public readonly ref struct ChangeRowRef
{
    public ReadOnlySpan<byte> File { get; init; }   // zero-copy — aliases the reader's row buffer
    public int LinesAdded { get; init; }
}

using var reader = Excel.FromFile("changes.xlsx");

foreach (ChangeRowRef item in RefParser.ParseNamed<ChangeRowRef>(reader))
{
    Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
}
```

The sequence also supports `await foreach`, so a `ref struct` model can be parsed asynchronously — the rows are streamed via `MoveNextAsync` while the model stays a zero-copy `ref struct`:

```csharp
await using var reader = await Excel.FromFileAsync("changes.xlsx");

await foreach (ChangeRowRef item in RefParser.ParseNamed<ChangeRowRef>(reader))
{
    Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
}
```

A few differences from `ExcelParser<T>`:

- **Span fields alias the reader's row buffer** — valid only until the next row. Copy them out (e.g. `Encoding.UTF8.GetString(span)`) if you need to keep the value past the loop body. Under `await foreach`, the same rule means the model can't be held across an `await` in the loop body.
- **`foreach` / `await foreach` only.** Consumption is pattern-based — the sequence cannot be surfaced through `IEnumerable<T>`, `IAsyncEnumerable<T>`, or LINQ, because a `ref struct` element can't be boxed through those interfaces (`IAsyncEnumerable<T>` in particular forbids a `ref struct` element type — CS9267). Iterate it directly.
- **Not AOT/trim-safe**, same tradeoff as `ExcelParser<T>` (both reflect over `T`'s properties and compile setters at runtime).
- A regular `struct`/`class` model works with `ParseNamed` too — only a genuine `ref struct` model gets the extra zero-copy span-property binding.

## Write XLSX workbooks

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xlsx");
await using var workbook = await XlsxWorkbookWriter.CreateAsync(stream);

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
await using var workbook = await XlsxWorkbookWriter.CreateAsync(stream, useSharedStrings: true);
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
