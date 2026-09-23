# Writing workbooks

Writers for XLSX (Open XML), XLSB (BIFF12) and XLS (BIFF8), cell styles,
typed records, and prefetch compression. See [csv.md](csv.md) for CSV output.

## Write XLSX workbooks

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xlsx");
await using var workbook = XlsxWorkbookWriter.Create(stream);

await using (var sheet = workbook.AddSheet("Summary"))
{

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
await using var workbook = XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { UseSharedStrings = true });
```

## Hidden sheets

Every `IWorkbookWriter<TSheet>` takes an optional `ExcelSheetVisibility` alongside the sheet name, written to whatever the format uses for it — XLSX's `state` attribute, XLSB's `BrtBundleSh.hsState`, XLS's `BoundSheet8.hsState` — so it reads back through [`SheetVisibility`](reading.md#open-by-auto-detecting-the-format):

```csharp
await using var data = workbook.AddSheet("Data");                                 // visible
await using var lookups = workbook.AddSheet("Lookups", ExcelSheetVisibility.Hidden);
await using var audit = workbook.AddSheet("Audit", ExcelSheetVisibility.VeryHidden); // not in Excel's unhide dialog
```

A workbook whose sheets are *all* hidden is rejected when it is ended (`InvalidOperationException`): the formats allow it, but Excel reports such a file as damaged, so it is caught at write time rather than shipped to a user. CSV accepts the argument and ignores it — delimited text has no tab bar to hide from.

## Cell styles on write

Every `IWorkbookWriter<TSheet>` supports column- and row-level styling: a number format (currency, date, percentage), bold, and italic. Register a `CellStyle` once with `AddStyle` and apply its returned index to a column (before the first row) or to a whole row (when starting it):

```csharp
using ExcelReader.Core.Writer;

await using var workbook = XlsxWorkbookWriter.Create(stream);

int currency = workbook.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
int header = workbook.AddStyle(new CellStyle { Bold = true });

await using var sheet = workbook.AddSheet("Summary");
sheet.SetColumnStyle(columnIndex: 1, currency); // before the first row
sheet.SetColumnWidth(columnIndex: 0, width: 20);

await using (var row = await sheet.StartRowAsync(header))
{
    row.Write("Product");
    row.Write("Total");
}
await using (var row = await sheet.StartRowAsync())
{
    row.Write("Widget");
    row.Write(1234.5);
}

await workbook.EndAsync();
```

`AddStyle` deduplicates by value: registering the same `CellStyle` twice returns the same index, and index 0 is always the general/default style. `SetColumnStyle`/`SetColumnWidth` must be called before `StartAsync` — the column layout (XLSX `<cols>`, XLSB `BrtColInfo`, XLS `COLINFO`) has to be written ahead of the row data. A row's style (from `StartRowAsync(int, CancellationToken)`) takes precedence over its column's style for any cell in that row. CSV has no cell concept of style: every style member is a documented no-op there.

Cell-level styling (one specific cell rather than a whole column or row) is out of scope. Bold/italic render only in XLSX today; XLSB and XLS apply the number format but keep the default font, since their font records are opaque binary blobs this library isn't confident hand-editing without a verified field map.

## Read and write XLSB workbooks (BIFF12)

Use `Excel.FromXlsbFile`, `Excel.FromXlsb`, `Excel.FromXlsbFileAsync`, or `Excel.FromXlsbAsync` to open XLSB directly. For writing, use `XlsbWorkbookWriter`, `XlsbSheetWriter`, and `XlsbRowWriter`.

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xlsb");
await using var workbook = XlsbWorkbookWriter.Create(stream);

await using (XlsbSheetWriter sheet = workbook.AddSheet("Summary"))
{

    await using (XlsbRowWriter row = await sheet.StartRowAsync())
    {
        row.Write("Name");
        row.Write("Total");
        row.Write("Created");
    }
}

await workbook.EndAsync();
```

The XLSB writer also defaults to inline string cells. Set `XlsbWriterOptions.UseSharedStrings` to deduplicate repeated text into `sharedStrings.bin`.

## Write XLS workbooks (BIFF8)

`XlsWorkbookWriter` emits a binary BIFF8 `.xls` file. The sheet and row APIs are synchronous; only the final `EndAsync` (which assembles and flushes the OLE container) is async. BIFF8 is capped at 65,536 rows × 256 columns per sheet.

```csharp
using ExcelReader.Core.Writer;

await using var stream = File.Create("out.xls");
await using var workbook = XlsWorkbookWriter.Create(stream);

using (var sheet = workbook.AddSheet("Summary"))
{

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
await using var writer = RecordWriter.CreateXlsx(stream);   // or CreateXlsb / CreateXls / CreateCsv
await writer.WriteSheetAsync("Sales", sales);
```

Each `WriteSheetAsync` call targets a new sheet (a duplicate name throws), so one workbook can hold sheets of different record types. `RecordWriter.CreateCsv` is the exception: a CSV file is a single sheet, so a second `WriteSheetAsync` throws (the sheet name is ignored). An `IAsyncEnumerable<T>` overload streams records that are produced asynchronously. The written file round-trips straight back through `ExcelParser<T>` because the headers are the property names.

Column behavior mirrors the parser attributes:

- **`[ExcelColumn("Header")]`** — use a custom header instead of the property name (the first alias wins).
- **`[ExcelIgnore]`** — exclude a property from both writing and parsing (for computed/transient members).
- **`[ExcelConverter(typeof(MyConverter))]`** — if the converter also implements `IExcelCellWriter<T>`, it controls how the value is written, so a custom type round-trips through the same converter it reads with.

`DateTime` and `DateOnly` are written as Excel date serials; `TimeOnly` as a time-of-day fraction. Numeric properties become number cells; any other type is written as its `ToString()` text. (`CreateCsv` follows the CSV rules instead — see [Write CSV](csv.md#write-csv) — writing `DateTime`/`DateOnly` as ISO text and `TimeOnly` as a time-of-day fraction, all still round-tripping through `ExcelParser<T>`.)

For a model marked `[ExcelSerializable]`, use `MappedRecordWriter.Create*` instead — same behavior, but driven by the source-generated map instead of reflection, so it stays Native AOT/trim-safe. See [Generate typed maps at compile time](parsing.md#generate-typed-maps-at-compile-time-native-aot--trimming).

## Prefetch compression (XLSX/XLSB writing)

The write-side mirror of [Prefetch decompression](reading.md#prefetch-decompression-xlsxxlsb). XLSX and XLSB
are ZIP-backed, so every row a writer serializes has to be deflated before it reaches the stream,
and by default that happens on the calling thread. Set `PrefetchWrite` on the writer options to move the deflate
onto a background thread, so the caller keeps building the next batch of rows while the previous one
compresses:

```csharp
await using var wb = XlsxWorkbookWriter.Create(stream, leaveOpen: true, options: new XlsxWriterOptions { PrefetchWrite = true });
XlsxSheetWriter sheet = wb.AddSheet("S1");

foreach (var record in records)
{
    using XlsxRowWriter row = sheet.StartRow();
    row.Write(record.Name);
    row.Write(record.Value);
}

await sheet.EndAsync();
```

`XlsbWriterOptions` has the same property, and `RecordWriter.CreateXlsx`/`CreateXlsb` and their
`MappedRecordWriter` counterparts take the same options. XLS and CSV are uncompressed, so
they have nothing to overlap and do not offer it.

Writing a 50,000-row workbook (`WriteBenchmark`, both figures from one run):

| Workload | Default | `PrefetchWrite = true` | Gain |
|---|---:|---:|---:|
| XLSX | 12.261 ms | 7.731 ms | 37% |
| XLSB | 7.515 ms | 5.668 ms | 25% |

Allocations are unchanged (4.02 MB vs. 4.03 MB) — the background writer hands over buffers the row
writer already owns rather than copying them.

The same caveat as the read side applies: **do not enable it for concurrent server workloads**. A
caller already writing many files in parallel is CPU-saturated, and an extra background thread per
writer only doubles thread demand for no gain. It's meant for single-file batch work.

