# CSV

Reading, dialect sniffing and writing delimited text.
Parallel CSV parsing is covered in [parsing.md](parsing.md).

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
foreach (var item in ExcelParser.FromAttributes<ChangeRow>().Parse(reader))
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

`Excel.Open`/`OpenAsync` do **not** auto-detect CSV — plain text has no magic-byte signature to sniff. Name the format instead: every `Open`/`OpenAsync` overload takes an optional `ExcelFileFormat`, and `ExcelReaderOptions.Csv` carries the dialect, so a caller that decides the format from an extension or a content-type header configures both families through one options object:

```csharp
ExcelFileFormat format = Path.GetExtension(path) is ".csv" ? ExcelFileFormat.Csv : ExcelFileFormat.Unknown;

using IExcelRowReader reader = Excel.Open(path, format, new ExcelReaderOptions
{
    Password = secret,                                                  // used by the Excel formats
    Csv = CsvReaderOptions.Default with { SniffDialect = true },        // used by the CSV format
});
```

`ExcelFileFormat.Unknown` keeps the signature-based detection the no-format overloads do; `EncryptedOoxml` is a detection result rather than something to open, so passing it throws `ArgumentOutOfRangeException`.

## Sniff a CSV dialect

`CsvSniffer.Detect` infers the delimiter, quote character, and encoding (from a leading byte-order mark) from a sample of bytes, so a `;`-separated pt-BR export or a TSV can be read without the caller knowing the dialect up front. It is deterministic (ties break by candidate order) and never throws on arbitrary input — an indecisive sample returns `CsvDialect.Default` (comma, `"`, UTF-8).

```csharp
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

CsvDialect dialect = CsvSniffer.DetectFile("export.csv");
using var reader = Excel.FromCsvFile("export.csv", CsvReaderOptions.Default.WithDialect(dialect));
```

`CsvSniffer.Detect` takes a span of bytes or a seekable `Stream`, `DetectFile` takes a path, and `DetectAsync`/`DetectFileAsync` are the async siblings; only the first 64 KiB are examined. The `Stream` overloads require a seekable source — they read a bounded sample and restore the stream's position — so a non-seekable stream throws `ArgumentException`; buffer it first, or pass the bytes as a span instead. Pass `CsvSnifferOptions` to change the candidate delimiters/quotes (and their priority order) or the number of sample lines considered.

`CsvReaderOptions.SniffDialect` does the sample-then-apply above for you, at the entry point that owns the source — the `Excel.FromCsv*` factories, `Excel.Open` with `ExcelFileFormat.Csv`, and the parallel CSV entry points:

```csharp
using var reader = Excel.FromCsvFile("export.csv", CsvReaderOptions.Default with { SniffDialect = true });
```

An explicitly set `Encoding` survives sniffing unless the source carries a byte-order mark naming a different one — a sniffed encoding only ever comes from a BOM, so it never silently overrides a caller who knows their file is Windows-1252.

## Write CSV

`CsvWorkbookWriter` emits RFC 4180 CSV: no styles or shared strings, so rows stream straight to the output. A CSV file is a single sheet, so `AddSheet` is called once (the name is ignored).

```csharp
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;

using var stream = File.Create("out.csv");
using var workbook = CsvWorkbookWriter.Create(stream);
using var writer = workbook.AddSheet("Sheet1");

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

Fields are quoted only when they contain the delimiter, quote character, `\r`, or `\n`; embedded quotes are doubled. `bool` writes as lowercase `true`/`false` and `DateTime`/`DateOnly` as round-trip ISO 8601 (`"O"`); `TimeOnly` as a time-of-day fraction — all matching what `ExcelParser<T>.Parse(CsvReader)` expects, so a file written by `CsvWorkbookWriter` parses back without configuration. `Skip(count)` writes empty fields to keep column positions aligned (CSV has no sparse-cell concept).

`CsvWriterOptions` mirrors `CsvReaderOptions` property for property, so a file written with one set of settings reads back with the matching set:

```csharp
using var workbook = CsvWorkbookWriter.Create(stream, leaveOpen: false, new CsvWriterOptions
{
    Delimiter = (byte)';',
    Encoding = Encoding.Latin1,          // fields are formatted as UTF-8, then transcoded on the way out
    WriteByteOrderMark = true,           // what makes Excel open a UTF-8 file as UTF-8
    NewLine = CsvNewLine.LineFeed,       // default is CRLF, as RFC 4180 specifies
});
```

Fields containing `\r` or `\n` are quoted whichever `NewLine` is chosen, so the terminator never changes how the file reads back.

To dump a collection of typed records instead of writing cells by hand, call `WriteRecordsAsync` on the sheet from `CsvWorkbookWriter.Create(stream)` — the same [record-writing API](writing.md#write-typed-records) as the Excel formats, restricted to a single sheet.

