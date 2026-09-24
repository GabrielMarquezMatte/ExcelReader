# Reading rows

Opening a workbook, enumerating rows, and the async and prefetch paths.
See also [parsing.md](parsing.md) for binding rows to types, and [csv.md](csv.md) for delimited text.

## Read rows

```csharp
using ExcelReader.Core.Reader;

using var reader = Excel.FromXlsxFile("report.xlsx");

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

### Parallel CSV parsing (opt-in)

```csharp
await foreach (var row in CsvParallel.ParseAsync("big.csv", ExcelParser.FromAttributes<SalesRow>(), new CsvParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += row.Revenue;
}
```

Rows arrive in file order, identical to the sequential parser. Parsing and type conversion run across
threads; whatever your loop body does per row does not — if that dominates, raising the degree will
not help. Sources that cannot be partitioned (non-seekable streams, non-UTF-8 encodings, small files)
fall back to sequential parsing with the same results.

For a fold or a per-row side effect, `CsvParallel.ForEachAsync` accepts a `ref struct` row type, so text
columns stay as `ReadOnlySpan<byte>` and the per-row model allocation disappears entirely — a class
model still costs one instance per row:

```csharp
public ref struct SaleRow
{
    [ExcelColumn("Region")] public ReadOnlySpan<byte> Region { get; set; }
    [ExcelColumn("Units")]  public int Units { get; set; }
}

long total = 0;
await CsvParallel.ForEachAsync(
    "big.csv",
    ExcelParser.FromAttributes<SaleRow>(),
    row => Interlocked.Add(ref total, row.Units),
    new CsvParallelOptions { DegreeOfParallelism = 8, HeaderRow = 1 });
```

Unlike `CsvParallel.ParseAsync`, the callback runs on the worker threads, so records arrive in no
particular order and the callback must be safe to call from several threads at once. Spans are valid
only for the duration of the call.

Delivery is at least once, not exactly once: the callback may be invoked more than once for the same
record. A partition whose start offset was guessed wrongly is read again from the confirmed offset, and
the discarded pass may already have delivered a whole partition's worth of records — 1 MiB to 64 MiB of them, not just a few near the seam.
On a source with quoted fields those records may also be misparsed, carrying values that appear nowhere
in the file, and an exception the callback threw during a discarded pass is discarded with it. A start
is only guessed wrongly inside a quoted field, so a source in which the quote character never appears
delivers every record exactly once; for that guarantee on any source, use `CsvParallel.AggregateAsync`,
where the discarded partition's accumulator is thrown away, and apply side effects once it returns.
A callback that writes to a database or sends messages must be idempotent and must also tolerate
records that do not exist — deduplicating on a key does not filter those out.

See the [parallel CSV benchmarks](../performance/benchmarks.md#parallel-csv) for what this buys. Those figures are
`CsvParallel.AggregateAsync`'s, measured over the same projection path this shares: at dop 16 it is
~2.7x faster than the typed path while allocating ~394x less, with zero garbage collections.

Properties need setters. A get-only property is skipped during binding and silently receives nothing.

### Apache Arrow conversion (opt-in)

```bash
dotnet add package ExcelReader.Arrow
```

```csharp
using ExcelReader.Arrow;
using ExcelReader.Core.Reader;

using var reader = Excel.FromXlsxFile("report.xlsx");
Apache.Arrow.RecordBatch batch = reader.ToArrowRecordBatch();
```

`schema` defaults to `Excel.InferSchema`'s guess; pass an explicit `ExcelColumnSchema[]` to skip inference. The whole sheet is materialized into one `RecordBatch` — there is no chunked/streaming variant yet.

`WriteRecordBatch`/`WriteRecordBatchAsync` are the write-side mirror — one call from a `RecordBatch` to a sheet, for XLSX, XLSB, XLS, and CSV:

```csharp
using ExcelReader.Arrow;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create("report.xlsx"));
workbook.WriteRecordBatch(batch); // adds the sheet, writes the header + every row, ends the workbook
```

Supports the same seven Arrow types `ToArrowRecordBatch` produces (string, int64, double, boolean, date32, time64, timestamp); any other type throws `NotSupportedException`. Pass `writeHeader: false` to skip the header row, or a `sheetName` (XLSX/XLSB/XLS only — CSV has no sheet name) to rename it from the `"Sheet1"` default.

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

`Sheets()` is the same walk as an extension method, yielding an `ExcelSheet { Index, Name }` per sheet after selecting it as current:

```csharp
using IExcelRowReader reader = Excel.Open("report.xlsx");

foreach (var sheet in reader.Sheets())
{
    Console.WriteLine(sheet.Name);
    foreach (var row in reader) // reads the sheet Sheets() just selected
    {
        Console.WriteLine(row[0].GetString());
    }
}
```

`SheetVisibility` and `SheetVisibilityAt(index)` report whether a sheet is shown in the workbook's tab bar, read from the format's own encoding of it (XLSX's `state` attribute, XLSB's `BrtBundleSh.hsState`, XLS's `BoundSheet8.hsState`). It is reported, never enforced — a hidden sheet enumerates its rows like any other, so converters and exporters filter on it themselves:

```csharp
using ExcelReader.Core; // ExcelSheetVisibility, shared by the readers and writers

foreach (var sheet in reader.Sheets())
{
    if (sheet.Visibility != ExcelSheetVisibility.Visible)
    {
        continue; // skip hidden and veryHidden sheets
    }
    Console.WriteLine(sheet.Name);
}
```

`ExcelSheetVisibility.VeryHidden` is the state Excel's own unhide dialog does not offer; both it and `Hidden` come back here. A sheet whose format says nothing about its state — or says something no producer agrees on — reads as `Visible`.

CSV is exposed as a single, unnamed sheet (`SheetCount == 1`, `SheetName == ""`, always `Visible`). Pattern-match to the concrete type only for reader-specific internals beyond this surface.

`OpenAsync` is the async counterpart. Both require a seekable stream (or a file path) so the signature can be read without consuming the input.

Detection covers the signed formats only. To open a source whose format you already know — including CSV, which has no signature to detect — pass an `ExcelFileFormat` and let `ExcelReaderOptions.Csv` carry the dialect; see [CSV](csv.md#read-csv).

```csharp
using IExcelRowReader reader = Excel.Open("report.csv", ExcelFileFormat.Csv);
```

## Read asynchronously

Every reader supports `await foreach`. For XLSX files, the async reader buffers one row at a time and uses the same row parser as the sync reader, so sync and async reads stay behaviorally aligned while awaits happen only when more bytes are needed.

```csharp
using ExcelReader.Core.Reader;

await using var reader = await Excel.FromXlsxFileAsync("report.xlsx", cancellationToken);

await foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}
```

`await foreach` binds to the reader's `GetAsyncEnumerator(CancellationToken ct = default)` by pattern. The call itself does no I/O: opening the sheet part (and, for XLSX, loading shared strings) happens asynchronously on the first `MoveNextAsync`. Because `Row` and `Cell` are `ref struct` types, the current row cannot be held across an `await` inside the loop body: read its cells (or copy the values out) before awaiting anything else.

`await foreach` does not accept `.WithCancellation(ct)`: `Row` being a `ref struct` rules out `IAsyncEnumerable<Row>`, so the loop binds to the pattern rather than the interface. To pass a token, or to `await` while a row is in scope, drive the enumerator manually:

```csharp
await using var reader = await Excel.FromXlsxFileAsync("report.xlsx", cancellationToken);
await using var rows = reader.GetAsyncEnumerator(cancellationToken);

while (await rows.MoveNextAsync())
{
    var row = rows.Current;
    Console.WriteLine(row[0].GetString());
}
```

Errors from opening the sheet, such as cancellation or a corrupt part, surface from that first `MoveNextAsync`, not from `GetAsyncEnumerator`.

## Prefetch decompression (XLSX/XLSB)

XLSX and XLSB are ZIP-backed, and inflating a sheet's compressed bytes competes for
wall-clock time with parsing it. `ExcelReaderOptions.PrefetchDecompression` overlaps the
two: a background thread inflates ahead while the calling thread parses. It is **opt-in,
defaults to `false`**, and only affects XLSX/XLSB — XLS and CSV have nothing to
decompress, so the option is silently ignored for them.

```csharp
var options = new ExcelReaderOptions { PrefetchDecompression = true };
using var reader = Excel.FromXlsxFile("report.xlsx", options);

foreach (var row in reader)
{
    Console.WriteLine(row[0].GetString());
}
```

Measured across both read benchmarks (see [Real data reads](../performance/benchmarks.md#real-data-reads) and
[String-heavy reads](../performance/benchmarks.md#string-heavy-reads)), on 65K-row workbooks:

| Workload | Default | `PrefetchDecompression = true` | Gain |
|---|---:|---:|---:|
| XLSX, real data | 64.3 ms | 42.9 ms | 33% |
| XLSM, real data | 65.0 ms | 42.6 ms | 35% |
| XLSB, real data | 28.9 ms | 15.4 ms | 47% |
| XLSX, string-heavy | 57.3 ms | 34.8 ms | 39% |
| XLSB, string-heavy | 39.5 ms | 25.8 ms | 35% |

The gain tracks how much of a read is decompression rather than parsing, so it is largest
on XLSB with numeric data (where inflate dominates). Allocations rise from the producer task and the
pooled decompression buffers — on the real-data corpus, roughly 18 KB to 38 KB for XLSX
and 19 KB to 30 KB for XLSB — and neither path triggers a garbage collection.

Do **not** enable it for concurrent server workloads: a caller already reading many files
in parallel is CPU-saturated, and an extra background thread per read only doubles thread
demand for no gain. It's meant for single-file batch processing.

Writing has the same option, under a different name: see
[Prefetch compression](writing.md#prefetch-compression-xlsxxlsb-writing).

## Bridge to ADO.NET (`IDataReader`)

`ExcelDataReader` adapts an `IExcelRowReader`'s current sheet to `System.Data.IDataReader`, so it drops straight into `SqlBulkCopy`, `DataTable.Load`, Dapper, or any other ADO.NET consumer:

```csharp
using System.Data;
using ExcelReader.Core.Reader;

using IExcelRowReader reader = Excel.Open("report.xlsx");
using IDataReader data = new ExcelDataReader(reader); // headerRow: 1 by default

var table = new DataTable();
table.Load(data);
```

The header row (1-based, default 1) fixes the column shape; pass `headerRow: 0` for a header-less sheet, whose columns come back named `Column0`, `Column1`, ... sized from the first data row. There is no schema-inference pass — `GetFieldType`/`GetValue`/the typed getters read the *current* row's own cell type, so a consumer building its schema from the first row (like `DataTable.Load`) locks in that row's types for the whole load. `NextResult()` always returns `false`; combine with `Sheets()` to build one `ExcelDataReader` per sheet instead of chaining result sets.

