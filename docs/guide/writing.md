# Writing workbooks

Writers for XLSX (Open XML), XLSB (BIFF12) and XLS (BIFF8), cell styles,
typed records, and prefetch compression. See [csv.md](csv.md) for CSV output.

## Write XLSX workbooks

<!-- snippet: writing-xlsx -->
```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

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

<!-- snippet: writing-xlsx-shared-strings -->
```csharp
await using var workbook = XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { UseSharedStrings = true });
```

## Hidden sheets

Every [`IWorkbookWriter<TSheet>`](xref:ExcelReader.Core.Writer.IWorkbookWriter%601) takes an optional [`ExcelSheetVisibility`](xref:ExcelReader.Core.ExcelSheetVisibility) alongside the sheet name, written to whatever the format uses for it — XLSX's `state` attribute, XLSB's `BrtBundleSh.hsState`, XLS's `BoundSheet8.hsState` — so it reads back through [`IExcelSheet.Visibility`](reading.md#open-by-auto-detecting-the-format):

<!-- snippet: writing-hidden-sheets -->
```csharp
using ExcelReader.Core; // ExcelSheetVisibility, shared by the readers and writers

await using (var data = workbook.AddSheet("Data")) { /* rows */ } // visible
await using (var lookups = workbook.AddSheet("Lookups", ExcelSheetVisibility.Hidden)) { /* rows */ }
await using (var audit = workbook.AddSheet("Audit", ExcelSheetVisibility.VeryHidden)) { /* rows */ } // not in Excel's unhide dialog
```

A workbook whose sheets are *all* hidden is rejected when it is ended (`InvalidOperationException`): the formats allow it, but Excel reports such a file as damaged, so it is caught at write time rather than shipped to a user. CSV accepts the argument and ignores it — delimited text has no tab bar to hide from.

## Cell styles on write

Every `IWorkbookWriter<TSheet>` supports column-, row- and cell-level styling: number format, font (bold, italic, underline, strikethrough, name, size, color), solid fill, per-edge borders, alignment, wrap, indent and rotation. Register a [`CellStyle`](xref:ExcelReader.Core.Writer.CellStyle) once with `AddStyle` and apply its returned index to a column (before the first row) or to a whole row (when starting it):

<!-- snippet: writing-styles -->
```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

await using var workbook = XlsxWorkbookWriter.Create(stream);

int currency = workbook.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
int header = workbook.AddStyle(new CellStyle { Bold = true });

await using var sheet = workbook.AddSheet("Summary");
sheet.SetColumnStyle(columnIndex: 1, currency); // before the first row
sheet.SetColumnWidth(columnIndex: 0, width: 20);
sheet.SetShowGridLines(false);

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

`AddStyle` deduplicates by value: registering the same `CellStyle` twice returns the same index, and index 0 is always the general/default style. `SetColumnStyle`/`SetColumnWidth` must be called before the sheet's first `StartRow`/`StartRowAsync` (calling them later throws `InvalidOperationException`) — the column layout (XLSX `<cols>`, XLSB `BrtColInfo`, XLS `COLINFO`) has to be written ahead of the row data. `SetShowGridLines` follows the same rule (it sets XLSX `<sheetView showGridLines>`, the XLSB `BrtBeginWsView` flag and the XLS `WINDOW2` flag); sheets show gridlines unless it is called with `false`. A row's style (from `StartRowAsync(int, CancellationToken)`) takes precedence over its column's style for any cell in that row. CSV has no cell concept of style: every style member is a documented no-op there.

### Style one cell, and set row heights

`row.Style(styleId)` styles the next cell written on that row, and only that cell; it replaces the row and column style for it, and `Style(0)` forces the general style. A `null` write under a pending style writes a styled blank cell. `Skip` or ending the row while a style is pending throws `InvalidOperationException`. `StartRow(styleId, height)`/`StartRowAsync(styleId, height)` set the row height in points (greater than 0, at most 409):

<!-- snippet: writing-cell-styles -->
```csharp
await using var workbook = XlsxWorkbookWriter.Create(stream);

int title = workbook.AddStyle(new CellStyle
{
    Bold = true,
    FontSize = 14,
    FontColor = ExcelColor.FromRgb(0xFFFFFF),
    FillColor = ExcelColor.FromRgb(0x1F4E78),
    HorizontalAlignment = ExcelHorizontalAlignment.Center,
    VerticalAlignment = ExcelVerticalAlignment.Center,
});
var thin = new CellBorder(ExcelBorderStyle.Thin, ExcelColor.FromRgb(0x808080));
int boxed = workbook.AddStyle(new CellStyle { Left = thin, Right = thin, Top = thin, Bottom = thin });
int warning = workbook.AddStyle(new CellStyle { FontColor = ExcelColor.FromRgb(0xC00000), Italic = true });

await using var sheet = workbook.AddSheet("Report");
await using (var row = await sheet.StartRowAsync(title, height: 30)) // row style and height in points
{
    row.Write("Quarterly report");
}
await using (var row = await sheet.StartRowAsync())
{
    row.Style(boxed).Write("Revenue"); // Style applies to the next cell only
    row.Style(boxed).Write(1234.5);
    row.Style(warning).Write("estimate");
}

await workbook.EndAsync();
```

### Styled record layouts

`[ExcelStyle]` on a property styles that column's data cells (never its header) when a layout from `ExcelRecordLayout.FromAttributes<T>()` or the `[ExcelSerializable]` generator writes it; it wins over `SetColumnStyle`, except that a column number format set with `SetColumnStyle` is kept when the attribute names none. Colors are `0xRRGGBB` ints, `-1` for unset, and `Border` sets all four edges. `WithHeaderStyle` styles every header cell, including a table's header row. For a style that depends on the value, write the column yourself in an `IExcelRecordMap<T>` and call `row.Style` before the write:

<!-- snippet: writing-styled-records -->
```csharp
public sealed class Payment
{
    public string Customer { get; init; } = "";

    [ExcelStyle(NumberFormat = "#,##0.00", Bold = true)]
    public decimal Amount { get; init; }
}

public sealed class Balance : IExcelRecordMap<Balance>
{
    public static int Negative; // a style id registered on the workbook before writing

    public string Account { get; init; } = "";

    public double Value { get; init; }

    public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<Balance, TRow> builder)
        where TRow : IRowWriter
    {
        builder
            .Column("Account", static (row, b) => row.Write(b.Account))
            .Column("Value", static (row, b) =>
            {
                if (b.Value < 0)
                {
                    row.Style(Negative);
                }
                row.Write(b.Value);
            });
    }
}

await using var workbook = XlsxWorkbookWriter.Create(stream);
Balance.Negative = workbook.AddStyle(new CellStyle { FontColor = ExcelColor.FromRgb(0xFF0000) });

ExcelRecordLayout<Payment> layout = ExcelRecordLayout.FromAttributes<Payment>()
    .WithHeaderStyle(new CellStyle { Bold = true, FillColor = ExcelColor.FromRgb(0xD9E1F2) });

await using (var sheet = workbook.AddSheet("Payments"))
{
    await sheet.WriteRecordsAsync(payments, layout);
}
await using (var sheet = workbook.AddSheet("Balances"))
{
    await sheet.WriteRecordsAsync(balances, ExcelRecordLayout.Generated<Balance>());
}

await workbook.EndAsync();
```

Styled layouts work in serial, table and parallel writes. An invalid `[ExcelStyle]` value throws `ArgumentException` from `FromAttributes`; with the generator it is compile error `EXR011`. CSV accepts styled layouts and writes the same text as without them. A custom `ISheetWriter<TRow>` given a styled layout throws `NotSupportedException`.

Limits are checked when a style is added: at most 512 fonts in every format, 65,490 cell styles in XLSX/XLSB and 4,000 in XLS. XLS also has a 56-color palette: the 57th distinct color (font, fill or border) throws `ExcelLimitExceededException`.

### Styles from the native library and the bindings

The native writer handle (ABI 7) takes the same styles: `xl_add_style` registers one, and `xl_set_column_style`, `xl_start_row_styled` and `xl_style_next_cell` apply it; `xl_sheet_options` carries column styles, widths and a header style for a typed `xl_write_sheet`. Column styles reach a typed sheet's data cells, never its header. A column or overlay style without a number format keeps the date/time format of a date, time or timestamp column; an explicit number format wins. Each README has a "Cell styles" section:

- [C++](../../cpp/README.md#cell-styles)
- [Python](../../python/README.md#cell-styles)
- [Rust](../../rust/excelreader/README.md#cell-styles)
- [.NET Framework](https://github.com/GabrielMarquezMatte/ExcelReader/blob/master/src/ExcelReader.NetFramework/README.md#cell-styles)
- [C ABI](https://github.com/GabrielMarquezMatte/ExcelReader/blob/master/src/ExcelReader.Native/README.md#cell-styles)

## Read and write XLSB workbooks (BIFF12)

Use [`Excel.FromXlsbFile`](xref:ExcelReader.Core.Reader.Excel.FromXlsbFile*), [`Excel.FromXlsb`](xref:ExcelReader.Core.Reader.Excel.FromXlsb*), [`Excel.FromXlsbFileAsync`](xref:ExcelReader.Core.Reader.Excel.FromXlsbFileAsync*), or [`Excel.FromXlsbAsync`](xref:ExcelReader.Core.Reader.Excel.FromXlsbAsync*) to open XLSB directly. For writing, use [`XlsbWorkbookWriter`](xref:ExcelReader.Core.Writer.Xlsb.XlsbWorkbookWriter), [`XlsbSheetWriter`](xref:ExcelReader.Core.Writer.Xlsb.XlsbSheetWriter), and [`XlsbRowWriter`](xref:ExcelReader.Core.Writer.Xlsb.XlsbRowWriter).

<!-- snippet: writing-xlsb -->
```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;

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

The XLSB writer also defaults to inline string cells. Set [`XlsbWriterOptions.UseSharedStrings`](xref:ExcelReader.Core.Writer.Xlsb.XlsbWriterOptions.UseSharedStrings) to deduplicate repeated text into `sharedStrings.bin`.

## Write XLS workbooks (BIFF8)

[`XlsWorkbookWriter`](xref:ExcelReader.Core.Writer.Xls.XlsWorkbookWriter) emits a binary BIFF8 `.xls` file. The sheet and row APIs are synchronous; only the final `EndAsync` (which assembles and flushes the OLE container) is async. BIFF8 is capped at 65,536 rows × 256 columns per sheet.

<!-- snippet: writing-xls -->
```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;

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

The low-level writers above give you cell-by-cell control. When you just want to dump a collection of objects to a sheet, `sheet.WriteRecordsAsync(records, layout)` writes a header row followed by one row per record, with the columns coming from an [`ExcelRecordLayout<T>`](xref:ExcelReader.Core.Writer.ExcelRecordLayout%601). `ExcelRecordLayout.FromAttributes<T>()` maps each public readable property to a column by reflection. It is an extension on [`ISheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ISheetWriter%601), so the same call works on a sheet from any of the writers above — XLSX, XLSB, XLS, and CSV.

<!-- snippet: writing-records -->
```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

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
await using var workbook = XlsxWorkbookWriter.Create(stream);   // or XlsbWorkbookWriter / XlsWorkbookWriter / CsvWorkbookWriter
await using (var sheet = workbook.AddSheet("Sales"))
{
    await sheet.WriteRecordsAsync(sales, ExcelRecordLayout.FromAttributes<Sale>());
}
```

Disposing the sheet ends it; add the next sheet after that (a duplicate name throws, ignoring case), so one workbook can hold sheets of different record types. The sheet stays open after `WriteRecordsAsync`, so you can append more rows before disposing it. [`CsvWorkbookWriter`](xref:ExcelReader.Core.Writer.Csv.CsvWorkbookWriter) is the exception: a CSV file is a single sheet, so a second `AddSheet` throws (the sheet name is ignored). An `IAsyncEnumerable<T>` overload streams records that are produced asynchronously. The written file round-trips straight back through [`ExcelParser<T>`](xref:ExcelReader.Core.Parser.ExcelParser%601) because the headers are the property names.

Column behavior mirrors the parser attributes:

- **`[ExcelColumn("Header")]`** — use a custom header instead of the property name (the first alias wins).
- **`[ExcelIgnore]`** — exclude a property from both writing and parsing (for computed/transient members).
- **`[ExcelConverter(typeof(MyConverter))]`** — if the converter also implements [`IExcelCellWriter<T>`](xref:ExcelReader.Core.Writer.IExcelCellWriter%601), it controls how the value is written, so a custom type round-trips through the same converter it reads with.

`DateTime` and `DateOnly` are written as Excel date serials; `TimeOnly` as a time-of-day fraction. Numeric properties (including `Half`) become number cells; any other type is written as text formatted with the invariant culture — `DateTimeOffset` as ISO 8601 (`"O"`), enums by name — so the file parses back the same whatever the writing machine's culture. `Int128`/`UInt128` stay text because a number cell is a double and would drop their low digits. (`CsvWorkbookWriter` follows the CSV rules instead — see [Write CSV](csv.md#write-csv) — writing `DateTime`/`DateOnly` as ISO text and `TimeOnly` as a time-of-day fraction, all still round-tripping through `ExcelParser<T>`.)

For a model marked `[ExcelSerializable]`, pass `ExcelRecordLayout.Generated<T>()` instead — same behavior, but driven by the source-generated map instead of reflection, so it stays Native AOT/trim-safe. See [Generate typed maps at compile time](parsing.md#generate-typed-maps-at-compile-time-native-aot--trimming).

## Write an Excel table

The XLSX and XLSB sheet writers implement [`ITableSheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ITableSheetWriter%601). `BeginTable` writes the header row
and opens a table there; every row written until `EndTable` (or the end of the sheet) belongs to it. The
file opens in Excel already formatted as a table, with filter buttons and the chosen built-in style, and
reads back through [`IExcelWorkbook.Tables`](xref:ExcelReader.Core.Reader.IExcelWorkbook.Tables).

<!-- snippet: writing-table -->
```csharp
using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create("report.xlsx"));
XlsxSheetWriter sheet = workbook.AddSheet("Report");
sheet.BeginTable("Sales", ["Region", "Units", "Revenue"]);
foreach (Sale s in sales)
{
    using XlsxRowWriter row = sheet.StartRow();
    row.Write(s.Region);
    row.Write(s.Units);
    row.Write(s.Revenue);
}
sheet.EndTable();
```

For records, `WriteTableAsync` does all three steps with the same layouts as `WriteRecordsAsync`:

<!-- snippet: writing-table-records -->
```csharp
await sheet.WriteTableAsync(sales, "Sales", ExcelRecordLayout.FromAttributes<Sale>());
```

- [`ExcelTableOptions`](xref:ExcelReader.Core.Writer.ExcelTableOptions) sets `FirstColumn`, `StyleName` (a built-in `TableStyleLight1`–`21`,
  `TableStyleMedium1`–`28` or `TableStyleDark1`–`11`, or `null` for none; default `TableStyleMedium2`) and
  the banding flags `ShowRowStripes`, `ShowColumnStripes`, `ShowFirstColumn`, `ShowLastColumn`.
- With `FirstColumn > 0`, rows you write yourself must `Skip(FirstColumn)` to reach the table;
  `WriteTableAsync` does it for you.
- One table can be open per sheet at a time; write as many as you like one after another. Table names
  must be valid Excel names, must not read as cell references (`A1`, `R1C1`), and must be unique in the
  workbook, ignoring case. A table with no data rows gets one empty row, as Excel requires.
- Totals rows, custom table styles, and tables in XLS or CSV are not supported.

## Prefetch compression (XLSX/XLSB writing)

The write-side mirror of [Prefetch decompression](reading.md#prefetch-decompression-xlsxxlsb). XLSX and XLSB
are ZIP-backed, so every row a writer serializes has to be deflated before it reaches the stream,
and by default that happens on the calling thread. Set `PrefetchWrite` on the writer options to move the deflate
onto a background thread, so the caller keeps building the next batch of rows while the previous one
compresses:

<!-- snippet: writing-prefetch -->
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

[`XlsbWriterOptions`](xref:ExcelReader.Core.Writer.Xlsb.XlsbWriterOptions) has the same property. XLS and CSV are uncompressed, so
they have nothing to overlap and do not offer it.

Writing a 50,000-row workbook (`WriteBenchmark`):

| Workload | Default | `PrefetchWrite = true` | Gain |
|---|---:|---:|---:|
| XLSX | 8.213 ms | 5.577 ms | 32% |
| XLSB | 5.388 ms | 3.271 ms | 39% |

Allocations are unchanged (4.02–4.03 MB either way) — the background writer hands over buffers the row
writer already owns rather than copying them.

The same caveat as the read side applies: **do not enable it for concurrent server workloads**. A
caller already writing many files in parallel is CPU-saturated, and an extra background thread per
writer only doubles thread demand for no gain. It's meant for single-file batch work.

## Write records in parallel (XLSX/XLSB)

`PrefetchWrite` moves compression off the calling thread, but the rows are still built on one thread,
and building them is most of a write. `WriteRecordsParallelAsync` and `WriteTableParallelAsync` use
several cores for both: the records are cut into batches, each batch is formatted and compressed on
a pool thread, and the batches reach the file in order. They are extensions on
[`XlsxSheetWriter`](xref:ExcelReader.Core.Writer.Xlsx.XlsxSheetWriter) and
[`XlsbSheetWriter`](xref:ExcelReader.Core.Writer.Xlsb.XlsbSheetWriter), and take an `IEnumerable<T>`
or an `IAsyncEnumerable<T>`. The cell-by-cell API stays single-threaded, and XLS and CSV have no
parallel form.

<!-- snippet: writing-parallel -->
```csharp
await using var wb = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
XlsxSheetWriter sheet = wb.AddSheet("S1");

// Header row from the layout, then one row per record.
await sheet.WriteRecordsParallelAsync(records, ExcelRecordLayout.FromAttributes<Metric>());

// Or format each row yourself; this delegate runs on several threads at once.
await sheet.WriteRecordsParallelAsync(records, static (XlsxRowWriter row, Metric record) =>
{
    row.Write(record.Name);
    row.Write(record.Value);
}, new ParallelWriteOptions { DegreeOfParallelism = 4 });

await sheet.EndAsync();
```

The rules:

- **Your delegate runs on several threads at once**, though never twice at once on the same row
  writer. Whatever it captures must be thread-safe. The library's own layouts are.
- **Order is kept.** Rows come out in the order the source yields them, and the source itself is
  enumerated by one flow at a time, never concurrently.
- **Rows may be written before and after the call, not during it.** Row numbering carries on, and
  the call can be made more than once on a sheet.
- **A failure or a cancellation ends the sheet and the workbook.** A serial `WriteRecordsAsync`
  cancelled between rows leaves the sheet usable; a parallel call does not. The original exception is
  rethrown once every batch in flight has stopped.
- **With `UseSharedStrings`, string indexes depend on thread timing**, so two runs give files that
  differ byte for byte. Their content is the same.
- More than 1,048,576 rows throws `ExcelLimitExceededException`.

[`ParallelWriteOptions`](xref:ExcelReader.Core.Writer.ParallelWriteOptions) has two knobs.
`DegreeOfParallelism` is the most batches in flight at once (0, the default, means one per core).
`BatchRows` is the rows per batch, 8,192 by default. Memory grows with their product: at the
defaults each batch in flight holds about 3 MB (its formatted rows, the compressor's window and the
compressed result), whatever the size of the sheet.

Writing 1,000,000 rows of four columns (text, integer, date, number) to a `MemoryStream` at
`Fastest`, on a Ryzen 7 5700X (`ParallelRecordWriteBenchmark`):

| Write | XLSX | XLSB |
|---|---:|---:|
| Serial | 160.09 ms | 110.06 ms |
| Serial, `PrefetchWrite` | 99.24 ms | 61.54 ms |
| Parallel, 1 batch in flight | 174.04 ms | 107.89 ms |
| Parallel, 2 | 87.64 ms | 53.74 ms |
| Parallel, 4 | 45.61 ms | 28.72 ms |
| Parallel, 8 | 28.93 ms | 18.66 ms |
| Parallel, 16 | 22.07 ms | 15.11 ms |

The file comes out 0.18% (XLSX) or 0.04% (XLSB) larger, because each batch is compressed on its own.
As with `PrefetchWrite`, this is for single-file batch work: a server already writing many files at
once has no idle cores to give it.

### From the native library and the bindings

`xl_open_write_handle` and `xl_open_write_handle_to_memory` take the same two knobs in `xl_write_options`
(as do the deprecated `xl_write_typed` and `xl_write_typed_to_memory`):
`degree_of_parallelism` and `batch_rows`, with the same meaning (`0` is every processor and the
8,192-row default). A zero-initialized struct therefore writes XLSX and XLSB in parallel; a struct
with the 7.0 size (32 bytes) or a NULL options pointer writes serially, as before. The C++, Rust and
Python wrappers default to a serial write and turn it on only when asked:

```cpp
auto writer = xl::WriterHandle::open("big.xlsx", {.degree_of_parallelism = 0}).value();
writer.write_columns("Sheet1", columns);
writer.finish();
```

```rust
let mut handle = WriterHandle::create("big.xlsx", &WriteOptions::new().degree_of_parallelism(0))?;
handle.write_columns("Sheet1", &columns, None)?;
handle.finish()?;
```

```python
with open_writer("big.xlsx", options=WriteOptions(degree_of_parallelism=0)) as writer:
    writer.write_sheet("Sheet1", table, types)
```

The table's buffers are read from several threads, so they must not change while the call runs.
XLS, CSV and the `xl_writer_handle` API ignore both fields.
