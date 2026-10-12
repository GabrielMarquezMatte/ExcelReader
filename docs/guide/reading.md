# Reading rows

Opening a workbook, enumerating rows, and the async and prefetch paths.
See also [parsing.md](parsing.md) for binding rows to types, and [csv.md](csv.md) for delimited text.

## Read rows

```csharp
using ExcelReader.Core.Reader;

using var workbook = Excel.FromXlsxFile("report.xlsx");

foreach (var row in workbook.FirstSheet)
{
    string name = row[0].GetString();

    if (row[1].TryParse(null, out int quantity))
    {
        Console.WriteLine($"{name}: {quantity}");
    }

    if (row[2].Type == CellType.Date && row[2].TryGetDateTime(workbook.IsDate1904, out var date))
    {
        Console.WriteLine(date);
    }
}
```

An opened workbook ([`XlsxWorkbook`](xref:ExcelReader.Core.Reader.Xlsx.XlsxWorkbook), [`XlsbWorkbook`](xref:ExcelReader.Core.Reader.Xlsb.XlsbWorkbook), [`XlsWorkbook`](xref:ExcelReader.Core.Reader.Xls.XlsWorkbook), [`CsvReader`](xref:ExcelReader.Core.Reader.Csv.CsvReader)) hands out its sheets: `FirstSheet` is the first one, `Sheets[i]` any of them. A sheet holds no resources of its own; iterating it starts a new, independent read. Dispose the workbook, not the sheets.

### Parallel reading (opt-in)

[`SheetParallel`](xref:ExcelReader.Core.Reader.SheetParallel) reads one sheet on several threads. It takes any sheet, so it works the same on XLSX, XLSB and CSV:

```csharp
using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

await foreach (var row in SheetParallel.ParseAsync(workbook.FirstSheet, ExcelParser.FromAttributes<SalesRow>(), new SheetParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += row.Revenue;
}
```

You own the workbook: keep it open until the task completes or the enumeration ends. For CSV, give the dialect when you open it (`Excel.FromCsvFile(path, options)`), not in the options. `DegreeOfParallelism` of `0` (the default) uses every processor; `1` reads sequentially.

There are three families, each over a sheet:

| Family | Result | Order |
|---|---|---|
| `ParseAsync<T>` | an `IAsyncEnumerable<T>` | the sheet's own order, identical to `parser.Parse(sheet)` |
| `ForEachAsync` | calls your callback on the worker threads | none |
| `AggregateAsync` | folds one accumulator per partition into a result | the same result as a single-threaded run |

Parsing and type conversion run across threads; whatever your loop body does per row does not, so if that dominates, raising the degree will not help.

Both of the next two examples read a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and the per-row model allocation disappears; a class model still costs one instance per row.

A fold over an XLSX sheet, with one accumulator per partition merged in sheet order:

```csharp
public ref struct SaleRow
{
    [ExcelColumn("Region")] public ReadOnlySpan<byte> Region { get; set; }
    [ExcelColumn("Units")]  public int Units { get; set; }
}

public sealed class UnitTotal : IRowAccumulator<UnitTotal, SaleRow>
{
    public long Units;

    public void Add(SaleRow model) => Units += model.Units;

    public void Merge(UnitTotal following) => Units += following.Units;
}

using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

UnitTotal total = await SheetParallel.AggregateAsync<UnitTotal, SaleRow>(
    workbook.FirstSheet,
    ExcelParser.FromAttributes<SaleRow>(),
    new SheetParallelOptions { HeaderRow = 1 });
```

The same call takes a [`RowAggregation<TState>`](xref:ExcelReader.Core.Reader.RowAggregation%601) when you would rather pass delegates than write a type, and an `IRowRecord<TRecord>` type in place of the parser.

A per-row callback:

```csharp
using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

long total = 0;
await SheetParallel.ForEachAsync(
    workbook.FirstSheet,
    ExcelParser.FromAttributes<SaleRow>(),
    row => Interlocked.Add(ref total, row.Units),
    new SheetParallelOptions { DegreeOfParallelism = 8, HeaderRow = 1 });
```

Unlike `ParseAsync`, the callback runs on the worker threads, so records arrive in no particular order and the callback must be safe to call from several threads at once. Spans are valid only for the duration of the call.

`HeaderRow` follows the same rule as before: `ParseAsync` takes the header from the parser's [`ExcelParserOptions.HeaderRow`](xref:ExcelReader.Core.Parser.ExcelParserOptions.HeaderRow) and rejects a non-zero [`SheetParallelOptions.HeaderRow`](xref:ExcelReader.Core.Reader.SheetParallelOptions.HeaderRow), while the parser-based `ForEachAsync` and `AggregateAsync` use `SheetParallelOptions.HeaderRow`.

#### When it reads sequentially

These cases read on one thread with the same results and no error:

- `DegreeOfParallelism` of 1.
- A sheet under 256 KB uncompressed.
- XLS.
- A table view (`ExcelTable.AsSheet()`).
- A CSV that cannot be partitioned: a stream that is not a file, or a non-UTF-8 encoding. Only in-memory data (`ReadOnlyMemory<byte>`, `byte[]`) and a `FileStream`-backed source partition; a CSV over a `MemoryStream` does not, so pass its bytes as `ReadOnlyMemory<byte>`.

`PrefetchDecompression` does not apply on this path: XLSX and XLSB go through a pipeline that already overlaps inflate and parse.

#### Delivery guarantees

- **`ParseAsync` and `AggregateAsync`** are exact on every format. `ParseAsync` yields the same items in the same order as the sequential parse; `AggregateAsync` equals its own single-threaded run.
- **XLSB** delivers every row exactly once, in `ForEachAsync` too.
- **XLSX** delivers exactly once for a well-formed sheet. The pipeline cuts the inflated sheet at what looks like a row start; on a sheet where those bytes sit inside cell content (a CDATA run, a comment) or the file is truncated, a worker detects it, and the read recovers by reading the rest sequentially. `ParseAsync` and `AggregateAsync` stay exact. In `ForEachAsync`, workers may already have called the callback for chunks that are then dropped, so on such a sheet rows can be delivered twice, or be rows the sequential reader would not produce.
- **CSV** is at least once in `ForEachAsync`, unchanged from v6: the callback may be invoked more than once for the same record. A partition whose start offset was guessed wrongly is read again from the confirmed offset, and the discarded pass may already have delivered a whole partition's worth of records, 1 MiB to 64 MiB of them, not just a few near the seam. On a source with quoted fields those records may also be misparsed, carrying values that appear nowhere in the file, and an exception the callback threw during a discarded pass is discarded with it. A start is only guessed wrongly inside a quoted field, so a source in which the quote character never appears delivers every record exactly once. For that guarantee on any source, use `AggregateAsync`, where the discarded partition's accumulator is thrown away, and apply side effects once it returns.

A callback that writes to a database or sends messages must be idempotent on CSV, and on XLSX if the sheet may be malformed, and must also tolerate records that do not exist: deduplicating on a key does not filter those out.

[`ExcelParseException.Row`](xref:ExcelReader.Core.Parser.ExcelParseException.Row) on `ParseAsync` counts the rows emitted so far. On `ForEachAsync` and `AggregateAsync` it is the exact row when the read is sequential and `0` when it was partitioned.

See the [parallel sheet benchmarks](../performance/benchmarks.md#parallel-sheet-reading) and the [parallel CSV benchmarks](../performance/benchmarks.md#parallel-csv) for what this buys. On a real-data XLSX the sheet stage drops from 29 ms (prefetched, sequential) to about 18 ms, and on XLSB from 16 ms to about 13 ms; the floor is inflate, which runs on one thread. CSV has no such floor: its `AggregateAsync` at dop 16 is ~2.8x faster than the typed path while allocating ~410x less, with zero garbage collections.

Properties need setters. A get-only property is skipped during binding and silently receives nothing.

### Apache Arrow conversion (opt-in)

```bash
dotnet add package ExcelReader.Arrow
```

```csharp
using ExcelReader.Arrow;
using ExcelReader.Core.Reader;

using var workbook = Excel.FromXlsxFile("report.xlsx");
Apache.Arrow.RecordBatch batch = workbook.FirstSheet.ToArrowRecordBatch();
```

The schema defaults to [`Excel.InferSchema`](xref:ExcelReader.Core.Reader.Excel.InferSchema*)'s guess; pass one in `new ArrowReadOptions { Schema = ... }` to skip inference, and set `HeaderRow` there (default 1, 0 for none). The whole sheet is materialized into one `RecordBatch` — there is no chunked/streaming variant yet. Inference costs about the same time and ~1.8x the allocation of an explicit schema on the same file. See [Arrow conversion](../performance/benchmarks.md#arrow-conversion) in the benchmarks.

`WriteRecordBatch`/`WriteRecordBatchAsync` are the write-side mirror — one call from a `RecordBatch` to a sheet, for XLSX, XLSB, XLS, and CSV:

```csharp
using ExcelReader.Arrow;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create("report.xlsx"));
workbook.WriteRecordBatch(batch); // adds the sheet, writes the header + every row, ends the workbook
```

Supports string, large string, string view, every integer width (signed and unsigned), half float, float, double, decimal128, decimal256, boolean, date32, date64, time32, time64, timestamp and null columns; any other type (binary, nested, dictionary) throws `NotSupportedException`. Pass `new ArrowWriteOptions { WriteHeader = false }` to skip the header row, or set `SheetName` (XLSX/XLSB/XLS only — CSV has no sheet name) to rename it from the `"Sheet1"` default.

## Open by auto-detecting the format

[`Excel.Open`](xref:ExcelReader.Core.Reader.Excel.Open*) picks the workbook type from the file signature (XLSX/XLSB are ZIP packages, XLS is an OLE2 document) and returns an [`IExcelWorkbook`](xref:ExcelReader.Core.Reader.IExcelWorkbook). `FirstSheet` gives the first sheet, and a sheet iterates its rows directly, so no pattern-match is needed for basic row iteration.

```csharp
using ExcelReader.Core.Reader;

using IExcelWorkbook workbook = Excel.Open("report.xlsx"); // or report.xlsb / report.xls

foreach (var row in workbook.FirstSheet)
{
    Console.WriteLine(row[0].GetString());
}
```

`SheetCount`, `SheetAt(index)` and `TryGetSheet(name)` are on `IExcelWorkbook` itself, so you can walk every sheet without knowing the format. Each [`IExcelSheet`](xref:ExcelReader.Core.Reader.IExcelSheet) carries its own `Index`, `Name` and `Visibility`:

```csharp
using IExcelWorkbook workbook = Excel.Open("report.xlsx");

for (int i = 0; i < workbook.SheetCount; i++)
{
    IExcelSheet sheet = workbook.SheetAt(i);
    Console.WriteLine(sheet.Name);
    foreach (var row in sheet)
    {
        Console.WriteLine(row[0].GetString());
    }
}
```

`TryGetSheet` matches a name ignoring case and opens nothing:

```csharp
if (workbook.TryGetSheet("Totals", out var totals))
{
    foreach (var row in totals)
    {
        Console.WriteLine(row[0].GetString());
    }
}
```

When you open a format directly ([`Excel.FromXlsxFile`](xref:ExcelReader.Core.Reader.Excel.FromXlsxFile*) and the like) the workbook's `Sheets` is a typed, indexable list (`Sheets.Count`, `Sheets[i]`, `foreach`) of [`XlsxSheet`](xref:ExcelReader.Core.Reader.Xlsx.XlsxSheet), [`XlsbSheet`](xref:ExcelReader.Core.Reader.Xlsb.XlsbSheet) or [`XlsSheet`](xref:ExcelReader.Core.Reader.Xls.XlsSheet) values, so walking it boxes nothing. Reaching a sheet only looks up a name and an index; the file is not read until you iterate the sheet.

`Visibility` reports whether a sheet is shown in the workbook's tab bar, read from the format's own encoding of it (XLSX's `state` attribute, XLSB's `BrtBundleSh.hsState`, XLS's `BoundSheet8.hsState`). It is reported, never enforced — a hidden sheet enumerates its rows like any other, so converters and exporters filter on it themselves:

```csharp
using ExcelReader.Core; // ExcelSheetVisibility, shared by the readers and writers

for (int i = 0; i < workbook.SheetCount; i++)
{
    IExcelSheet sheet = workbook.SheetAt(i);
    if (sheet.Visibility != ExcelSheetVisibility.Visible)
    {
        continue; // skip hidden and veryHidden sheets
    }
    Console.WriteLine(sheet.Name);
}
```

[`ExcelSheetVisibility.VeryHidden`](xref:ExcelReader.Core.ExcelSheetVisibility.VeryHidden) is the state Excel's own unhide dialog does not offer; both it and `Hidden` come back here. A sheet whose format says nothing about its state — or says something no producer agrees on — reads as `Visible`.

CSV is exposed as a single, unnamed sheet (`SheetCount == 1`, `Name == ""`, always `Visible`). Pattern-match to the concrete type only for workbook-specific internals beyond this surface.

`OpenAsync` is the async counterpart. Both require a seekable stream (or a file path) so the signature can be read without consuming the input.

Detection covers the signed formats only. To open a source whose format you already know — including CSV, which has no signature to detect — pass an [`ExcelFileFormat`](xref:ExcelReader.Core.Reader.ExcelFileFormat) and let [`ExcelReaderOptions.Csv`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.Csv) carry the dialect; see [CSV](csv.md#read-csv).

```csharp
using IExcelWorkbook workbook = Excel.Open("report.csv", ExcelFileFormat.Csv);
```

## Read an Excel table

A workbook's tables (ranges formatted with **Format as Table**) are listed on `Tables`, ordered by sheet
and creation, and found by name with `TryGetTable`, ignoring case. XLSX/XLSM and XLSB report them; XLS
and CSV always report none. `AsSheet()` reads only the table: its header row first, then its data rows
without the totals row, each cut to the table's columns, so it plugs into anything that takes a sheet.

```csharp
using IExcelWorkbook workbook = Excel.Open("sales.xlsx");
if (workbook.TryGetTable("Sales", out ExcelTable? sales))
{
    foreach (SaleRow row in ExcelParser.FromAttributes<SaleRow>().Parse(sales.AsSheet()))
    {
        Console.WriteLine($"{row.Product}: {row.Qty}");
    }
}
```

- `FirstRow`, `FirstColumn`, `LastRow` and `LastColumn` are zero-based and include the header and totals
  rows; `Ref` is the same range in A1 notation.
- Column indexes inside `AsSheet()` rows stay those of the sheet: a table starting at column C reports
  its first cell at index 2. `Excel.InferSchema` keeps those indexes in [`ExcelColumnSchema.Index`](xref:ExcelReader.Core.Reader.Schema.ExcelColumnSchema.Index) but
  lists only the table's columns. [`SheetDataReader`](xref:ExcelReader.Core.Reader.SheetDataReader) ordinals are relative to the table: ordinal 0 is
  its first column.
- A table whose header row is turned off (`HeaderRowCount == 0`) still enumerates, but `Parse` throws
  `InvalidOperationException` for every map, and `SheetDataReader` and `InferSchema` throw it when
  `HeaderRow` is greater than 0. Read it with `new SheetDataReader(table.AsSheet(), new SheetDataReaderOptions { HeaderRow = 0 })`,
  `Excel.InferSchema(table.AsSheet(), new SchemaInferenceOptions { HeaderRow = 0 })`, or by column index on the rows.
- `StyleName` is the table style's name, such as `TableStyleMedium2`, or `null` when the table has none.

## Reading sheets in parallel

This is one thread per sheet. To read a single sheet on several threads, see [Parallel reading](#parallel-reading-opt-in). Each sheet of a workbook can be read on its own thread:

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
- You may have several enumerators open at once, including several over the same sheet.
- Do not touch a stream you passed in while the workbook is open.

Disposing the workbook closes it to new sheets and enumerators; its file or stream is released when the last enumerator already obtained from it is disposed.

## Read asynchronously

Every sheet supports `await foreach`. For XLSX files, the async enumerator buffers one row at a time and uses the same row parser as the synchronous enumerator, so sync and async reads stay behaviorally aligned while awaits happen only when more bytes are needed.

```csharp
using ExcelReader.Core.Reader;

await using var workbook = await Excel.FromXlsxFileAsync("report.xlsx", ct: cancellationToken);

await foreach (var row in workbook.FirstSheet)
{
    Console.WriteLine(row[0].GetString());
}
```

`await foreach` binds to the sheet's `GetAsyncEnumerator(CancellationToken ct = default)` by pattern. The call itself does no I/O: opening the sheet part (and, for XLSX, loading shared strings) happens asynchronously on the first `MoveNextAsync`. Because [`Row`](xref:ExcelReader.Core.Reader.Row) and [`Cell`](xref:ExcelReader.Core.Reader.Cell) are `ref struct` types, the current row cannot be held across an `await` inside the loop body: read its cells (or copy the values out) before awaiting anything else.

`await foreach` does not accept `.WithCancellation(ct)`: `Row` being a `ref struct` rules out `IAsyncEnumerable<Row>`, so the loop binds to the pattern rather than the interface. To pass a token, or to `await` while a row is in scope, drive the enumerator manually:

```csharp
await using var workbook = await Excel.FromXlsxFileAsync("report.xlsx", ct: cancellationToken);
await using var rows = workbook.FirstSheet.GetAsyncEnumerator(cancellationToken);

while (await rows.MoveNextAsync())
{
    var row = rows.Current;
    Console.WriteLine(row[0].GetString());
}
```

Errors from opening the sheet, such as cancellation or a corrupt part, surface from that first `MoveNextAsync`, not from `GetAsyncEnumerator`.

## Prefetch decompression (XLSX/XLSB)

XLSX and XLSB are ZIP-backed, and inflating a sheet's compressed bytes competes for
wall-clock time with parsing it. [`ExcelReaderOptions.PrefetchDecompression`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.PrefetchDecompression) overlaps the
two: a background thread inflates ahead while the calling thread parses. It is **opt-in,
defaults to `false`**, and only affects XLSX/XLSB — XLS and CSV have nothing to
decompress, so the option is silently ignored for them.

```csharp
var options = new ExcelReaderOptions { PrefetchDecompression = true };
using var workbook = Excel.FromXlsxFile("report.xlsx", options);

foreach (var row in workbook.FirstSheet)
{
    Console.WriteLine(row[0].GetString());
}
```

Measured across both read benchmarks (see [Real data reads](../performance/benchmarks.md#real-data-reads) and
[String-heavy reads](../performance/benchmarks.md#string-heavy-reads)), on 65K-row workbooks:

| Workload | Default | `PrefetchDecompression = true` | Gain |
|---|---:|---:|---:|
| XLSX, real data | 49.7 ms | 27.8 ms | 44% |
| XLSM, real data | 50.3 ms | 28.1 ms | 44% |
| XLSB, real data | 27.7 ms | 16.9 ms | 39% |
| XLSX, string-heavy | 41.2 ms | 27.4 ms | 34% |
| XLSB, string-heavy | 39.5 ms | 26.0 ms | 34% |

The gain tracks how much of a read is decompression rather than parsing, so it is largest
on XLSX and XLSM with numeric data (where inflate dominates). Allocations rise from the producer task and the
pooled decompression buffers — on the real-data corpus, roughly 8 KB to 25 KB for XLSX
and 7 KB to 16 KB for XLSB — and neither path triggers a garbage collection.

Do **not** enable it for concurrent server workloads: a caller already reading many files
in parallel is CPU-saturated, and an extra background thread per read only doubles thread
demand for no gain. It's meant for single-file batch processing.

Writing has the same option, under a different name: see
[Prefetch compression](writing.md#prefetch-compression-xlsxxlsb-writing).

## Bridge to ADO.NET (`IDataReader`)

`SheetDataReader` adapts one `IExcelSheet` to `System.Data.IDataReader`, so it drops straight into `SqlBulkCopy`, `DataTable.Load`, Dapper, or any other ADO.NET consumer:

```csharp
using System.Data;
using ExcelReader.Core.Reader;

using IExcelWorkbook workbook = Excel.Open("report.xlsx");
using IDataReader data = new SheetDataReader(workbook.FirstSheet); // header on row 1 by default

var table = new DataTable();
table.Load(data);
```

The header row (1-based, default 1) fixes the column shape; pass `new SheetDataReaderOptions { HeaderRow = 0 }` for a header-less sheet, whose columns come back named `Column0`, `Column1`, ... sized from the first data row. There is no schema-inference pass — `GetFieldType`/`GetValue`/the typed getters read the *current* row's own cell type, so a consumer building its schema from the first row (like `DataTable.Load`) locks in that row's types for the whole load. `NextResult()` always returns `false`; build one `SheetDataReader` per sheet instead of chaining result sets.

The bridge costs ~1.3–1.4x over iterating rows directly, and `GetBytes` stays within ~10% of the direct read. See [ADO.NET bridge](../performance/benchmarks.md#adonet-bridge-idatareader) in the benchmarks.

