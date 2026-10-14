# Migrating from v6 to v7

What changed in parallel reading and in the rest of the public API, and how to port v6 code.
See also [reading.md](reading.md#parallel-reading-opt-in) for the full guide.

## What changed

Parallel reading now takes a sheet, and it works on XLSX and XLSB as well as CSV. `CsvParallel` is gone: its eighteen path, memory and stream methods became six [`SheetParallel`](xref:ExcelReader.Core.Reader.SheetParallel) methods that take any [`IExcelSheet`](xref:ExcelReader.Core.Reader.IExcelSheet). You open the workbook yourself, pass its sheet, and the library parallelises as far as the format allows. The contract types lost their `Csv` prefix because they now serve every format.

## Types

| v6 | v7 |
|---|---|
| `CsvParallel` | `SheetParallel` |
| `CsvParallelOptions` | [`SheetParallelOptions`](xref:ExcelReader.Core.Reader.SheetParallelOptions) (no `Reader` property) |
| `ICsvRecord<TSelf>` | [`IRowRecord<TSelf>`](xref:ExcelReader.Core.Reader.IRowRecord%601) |
| `ICsvAccumulator<TSelf, TModel>` | [`IRowAccumulator<TSelf, TModel>`](xref:ExcelReader.Core.Reader.IRowAccumulator%602) |
| `CsvAggregation<TState>` | [`RowAggregation<TState>`](xref:ExcelReader.Core.Reader.RowAggregation%601) |
| `CsvRowAction<TState>` | [`RowAction<TState>`](xref:ExcelReader.Core.Reader.RowAction%601) |

The old names are removed, with no `[Obsolete]` forwarders. Members, contracts and results are unchanged.

## Before and after

A CSV file:

```csharp
// v6
await foreach (Sale s in CsvParallel.ParseAsync(path, parser, new CsvParallelOptions { DegreeOfParallelism = 8, Reader = options }))
{
    Total += s.Revenue;
}

// v7
using CsvReader csv = Excel.FromCsvFile(path, new ExcelReaderOptions { Csv = options });
await foreach (Sale s in SheetParallel.ParseAsync(csv.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += s.Revenue;
}
```

CSV bytes in memory:

```csharp
// v6
await foreach (Sale s in CsvParallel.ParseAsync(data, parser, new CsvParallelOptions { DegreeOfParallelism = 8, Reader = options }))
{
    Total += s.Revenue;
}

// v7
using CsvReader csv = Excel.FromCsv(data, new ExcelReaderOptions { Csv = options });
await foreach (Sale s in SheetParallel.ParseAsync(csv.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += s.Revenue;
}
```

A stream:

```csharp
// v6
await foreach (Sale s in CsvParallel.ParseAsync(stream, parser, new CsvParallelOptions { DegreeOfParallelism = 8, Reader = options }))
{
    Total += s.Revenue;
}

// v7
using CsvReader csv = Excel.FromCsv(stream, leaveOpen: true, new ExcelReaderOptions { Csv = options });
await foreach (Sale s in SheetParallel.ParseAsync(csv.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += s.Revenue;
}
```

The other five methods port the same way: open a workbook, pass `csv.FirstSheet` (or any sheet) in place of the path, bytes or stream, and drop `Reader` from the options. XLSX and XLSB need no other change:

```csharp
using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");
await foreach (Sale s in SheetParallel.ParseAsync(workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 8 }))
{
    Total += s.Revenue;
}
```

## Behavior changes

- The caller owns the workbook and keeps it open until the task completes or the enumeration ends. v6 opened and closed its own source.
- The CSV dialect is given when the workbook is opened ([`ExcelReaderOptions.Csv`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.Csv)), not in the parallel options.
- A CSV in a `MemoryStream` is no longer partitioned: the stream overload of [`Excel.FromCsv`](xref:ExcelReader.Core.Reader.Excel.FromCsv*) reads it sequentially. Pass the bytes as `ReadOnlyMemory<byte>` to keep partitioning.
- `HeaderRow` rules are the same as v6. `ParseAsync` takes the header from [`ExcelParserOptions.HeaderRow`](xref:ExcelReader.Core.Parser.ExcelParserOptions.HeaderRow) and rejects a non-zero [`SheetParallelOptions.HeaderRow`](xref:ExcelReader.Core.Reader.SheetParallelOptions.HeaderRow); the parser-based `ForEachAsync` and `AggregateAsync` use `SheetParallelOptions.HeaderRow`.
- XLSX and XLSB sheets are now read on several threads when `DegreeOfParallelism` allows it. XLS, a table view (`ExcelTable.AsSheet()`), a sheet under 256 KB uncompressed and a CSV that cannot be partitioned still read sequentially with the same results.
- `PrefetchDecompression` does not apply to a parallel XLSX or XLSB read; the pipeline already overlaps inflate and parse.

## Other breaking changes

These changes replace positional parameters and loose arguments with types that can grow without another break.

| v6 | v7 |
|---|---|
| `ExcelParserConfig` | [`ExcelParserOptions`](xref:ExcelReader.Core.Parser.ExcelParserOptions); the [`ExcelParser`](xref:ExcelReader.Core.Parser.ExcelParser) factories' parameter is `options` |
| `IExcelCellConverter<T>.TryConvert(in Cell, bool isDate1904, IFormatProvider provider, out T)` | `TryConvert(in Cell, in ExcelRowContext context, out T)` |
| `ExcelCellReader<T>`, `ExcelRowParser<T>`, `ExcelCellReaders.*`: `(…, bool isDate1904, IFormatProvider provider, …)` | `(…, in ExcelRowContext context, …)` |
| `Excel.InferSchema(sheet, headerRow, sampleSize, parseText)` returning `ExcelColumnSchema[]` | `Excel.InferSchema(sheet, new SchemaInferenceOptions { HeaderRow = …, SampleSize = …, ParseText = … })` returning `IReadOnlyList<ExcelColumnSchema>` |
| `new ExcelDataReader(sheet, headerRow)` | `new SheetDataReader(sheet, new SheetDataReaderOptions { HeaderRow = … })` |
| `sheet.ToArrowRecordBatch(schema, headerRow)` | `sheet.ToArrowRecordBatch(new ArrowReadOptions { Schema = schema, HeaderRow = … })` |
| `WriteRecordBatch(batch, sheetName, writeHeader)` and the async form | `WriteRecordBatch(batch, new ArrowWriteOptions { SheetName = …, WriteHeader = … })` |
| `XlsWorkbookWriter.Create(stream, leaveOpen, date1904)` | `XlsWorkbookWriter.Create(stream, leaveOpen, new XlsWriterOptions { Date1904 = … })` |
| `Excel.FromCsv*(…, CsvReaderOptions options)` | `Excel.FromCsv*(…, new ExcelReaderOptions { Csv = options })` |
| `CsvReaderOptions.MaxCellBytes`, `CsvReaderOptions.InternStrings` | [`ExcelReaderOptions.MaxCellBytes`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.MaxCellBytes), [`ExcelReaderOptions.InternStrings`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.InternStrings) |
| `CellType.Formula` | Removed. A formula cell reports its cached result's type: `Number`, `Date`, `ExcelString`, `Boolean` or `Error` |
| `ExcelDataReader` (its name collided with Sylvan.Data.Excel and the ExcelDataReader package) | [`SheetDataReader`](xref:ExcelReader.Core.Reader.SheetDataReader) |
| `Cell.TryGetDateTime(out DateTime)`, which silently assumed the 1900 date system | `Cell.TryGetDateTime(workbook.IsDate1904, out DateTime)` |

A converter reads the date system and culture from the context:

```csharp
// v6
public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out decimal value)
    => decimal.TryParse(cell.GetString(), NumberStyles.Currency, provider, out value);

// v7
public bool TryConvert(in Cell cell, in ExcelRowContext context, out decimal value)
    => decimal.TryParse(cell.GetString(), NumberStyles.Currency, context.FormatProvider, out value);
```

To unit-test a converter, build the context with an initializer: `new ExcelRowContext { IsDate1904 = false, FormatProvider = CultureInfo.InvariantCulture }`.

Two behavior changes come with them:

- [`Excel.Open`](xref:ExcelReader.Core.Reader.Excel.Open*) with [`ExcelFileFormat.Csv`](xref:ExcelReader.Core.Reader.ExcelFileFormat.Csv) now applies `ExcelReaderOptions.MaxCellBytes` and `InternStrings`. v6 ignored both and used the CSV defaults.
- Formula cells are typed by their cached result in every format. v6 reported an XLSX text result and an XLS numeric result as `Formula`, so an XLS formula with a date style lost its `Date` type, and XLSB never used `Formula`.
- The CLI `convert` command no longer shifts dates by four years when the source workbook uses the 1904 date system.
- [`XlsxWriterOptions`](xref:ExcelReader.Core.Writer.Xlsx.XlsxWriterOptions) gains `Date1904`, which v6 lacked. It defaults to `false`, so existing output is unchanged.

## Stricter reading and writing

v7 refuses several inputs that v6 accepted and then got wrong.

- **Damaged packages.** Each ZIP entry of an XLSX or XLSB package is checked against its CRC-32. A mismatch throws `InvalidDataException` when the entry's last byte is read, so rows before the damage may already have been returned. v6 checked only the entry's size. Set [`ExcelReaderOptions.VerifyZipChecksums`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.VerifyZipChecksums) to `false` to read a package whose producer wrote wrong checksums; the bindings have the same switch as `verify_zip_checksums`. The check costs 2% to 4% of a read, and under 1% with `PrefetchDecompression`, where it runs on the background thread.
- **Line breaks in XLSX text.** A cell with a line break now reads with `
`, as XLSB and XLS already did. Excel stores the break as a CR LF pair in the XML, which v6 returned as `
`. A carriage return stored as data (`_x000D_` or `&#13;`) is still returned. Code that split XLSX text on `
` should split on `
`.
- **Sheet names.** `AddSheet` throws `ArgumentException` for a name that begins or ends with an apostrophe, or contains a control character (tab and line feed included), U+FFFE or U+FFFF. Excel refuses a workbook with such a name.
- **Column settings.** `SetColumnWidth` and `SetColumnStyle` throw `ArgumentOutOfRangeException` for a column past the format's last one (16,384 columns; 256 for XLS) and for a width above 255, Excel's maximum. v6 wrote a file Excel refused for the first, and an XLS width of 256 came out as a hidden column.
- **Number formats.** `AddStyle` throws `ExcelLimitExceededException` for the 201st distinct number format of a workbook, and `ArgumentException` for a format longer than 255 characters. Excel does not open a workbook past either.
- **Rows in XLSB.** A sheet stops at 1,048,576 rows. v6 accepted one more, which Excel does not read.
- **Values with no numeric form.** `Write<T>` with a type that does not convert to a number, such as `char`, throws `ArgumentException` in every writer. XLSB and XLS threw `InvalidCastException`.
- **Text in XLSX.** A carriage return, U+FFFE and U+FFFF in cell text are now written escaped. v6 wrote them raw: Excel read the carriage return as a line feed and refused a file with either of the other two.
- **`ExcelLimitExceededException.Message`** no longer says "bytes" for limits counted in rows, columns or entries. `LimitName`, `Limit` and `Actual` are unchanged.

## Native bindings

The C ABI is unchanged: no signature moved, and `XL_ABI_VERSION` stays 6. Behavior changed. `degree_of_parallelism` of `0` on `xl_parse_typed` and the Arrow exports now parallelises XLSX and XLSB as well as CSV, using every processor. Pass `1` for the old sequential behavior. The Rust, Python and C++ bindings take the same value (`degree_of_parallelism`, `parallelism`) and need no code change; a `parallelism=0` call on an XLSX or XLSB sheet is now faster, and holds every partition's columns until they merge, so its peak memory is higher.

`XL_CELL_FORMULA` (5) is no longer produced; formula cells come back with their result's cell type. The constant stays defined in the header and in the bindings.

`xl_write_options.date1904` now applies to XLSX as well as XLSB and XLS; v6 ignored it for XLSX.

`xl_csv_aggregate` keeps its signature. Inside, it opens a [`CsvReader`](xref:ExcelReader.Core.Reader.Csv.CsvReader) with the dialect from the native struct and calls `SheetParallel`.

### Python: one way to open a workbook

`open_bytes()` is removed. `open_workbook()` takes the bytes instead, and also accepts a `bytearray`, a `memoryview` or an `io.BytesIO`. It also takes a `Source` or a readable stream, the inputs that 7.0's release candidates opened with `open_source()` and `open_stream()`, which are gone as well. Arguments and keywords are unchanged.

| v6 | v7 |
|---|---|
| `open_bytes(data, "xlsx", password=p)` | `open_workbook(data, "xlsx", password=p)` |
| `open_bytes(buffer.getvalue())` | `open_workbook(buffer)` |

### Python: one way to write a workbook

`write_workbook_to_bytes()`, `write_arrow_to_bytes()`, `write_pandas_to_bytes()` and `write_polars_to_bytes()` are removed. `write_workbook()`, `write_arrow()`, `write_pandas()` and `write_polars()` take the destination first, a path or a binary object with `write()`, such as `io.BytesIO`. A destination without a path needs `format`. The frame writers now take `format` and `options` as keywords instead of passing `**kwargs` through.

| v6 | v7 |
|---|---|
| `payload = write_workbook_to_bytes(table, types, format="xlsx")` | `buffer = io.BytesIO()`<br>`write_workbook(buffer, table, types, format="xlsx")`<br>`payload = buffer.getvalue()` |
| `write_pandas_to_bytes(df, format="xlsx")` | `write_pandas(buffer, df, format="xlsx")` |

The streaming writer and the encryption helper follow the same pattern. `open_writer_to_memory()` and `SheetWriter.bytes()` are removed: `open_writer()` takes a binary object with `write()` and hands it the finished workbook when the writer closes. `encrypt_package_bytes()` is removed: `encrypt_package()` takes the package as a path, `bytes` or `io.BytesIO`, and the destination as a path or an object with `write()`.

| v6 | v7 |
|---|---|
| `with open_writer_to_memory("xlsx") as w:`<br>`    ...`<br>`    payload = w.bytes()` | `buffer = io.BytesIO()`<br>`with open_writer(buffer, "xlsx") as w:`<br>`    ...`<br>`payload = buffer.getvalue()` |
| `secret = encrypt_package_bytes(payload, password)` | `encrypt_package(payload, out, password)` with `out = io.BytesIO()` |

### Python: one shape per read

A sheet reads as rows (`rows()`), a typed table (`parse_typed()`, or `iter_parse_typed()` in batches), Arrow (`to_arrow()`, or `to_arrow_stream()` in batches) or a DataFrame (`to_pandas()`, `to_polars()`). `write_workbook()` takes a `TypedTable` or any Arrow-compatible table, including pandas and polars DataFrames. `SheetWriter` writes whole rows only.

| v6 | v7 |
|---|---|
| `workbook.sheet_count` | `len(workbook.sheets)` |
| `workbook.sheet_names` | `[s.name for s in workbook.sheets]` |
| `workbook.sheet_name_at(i)` | `workbook.sheets[i].name` |
| `sheet.read_all()` | `list(sheet.rows())`, or `parse_typed()` for speed |
| `sheet.read_all_columnar()`, `ColumnarSheet`, `decode_cell()` | `sheet.parse_typed(sheet.infer_schema())` |
| `sheet.to_arrow()` returning a `StructArray` | `sheet.to_arrow()` returns a `pyarrow.RecordBatch` |
| `sheet.to_record_batch(...)` | `sheet.to_arrow(...)` |
| `sheet.to_record_batch_reader(...)` | `pyarrow.RecordBatchReader.from_stream(sheet.to_arrow_stream(...))` |
| `sheet.iter_pandas(...)`, `sheet.iter_polars(...)` | iterate the reader above and convert each batch |
| `write_arrow(dest, batch)`, `write_pandas(dest, df)`, `write_polars(dest, df)` | `write_workbook(dest, batch_or_df)` |
| `writer.start_row()`, `write_str(v)` … `write_timestamp(v)`, `end_row()` | `writer.write_row([...])` |
| `writer.write_null(ColumnType.X)` | `(None, ColumnType.X)` as a value in `write_row` |
| `OpenOptions(csv_max_cell_bytes=n)`, `OpenOptions(csv_intern_strings=b)` | `OpenOptions(max_cell_bytes=n)`, `OpenOptions(intern_strings=b)`, which now apply to CSV too |
| `from excelreader import ArrowStream` | `from excelreader.arrow_stream import ArrowStream` |

### C++: one overload set per operation

`Workbook::open` takes a path, a byte span, a `Source` or an `InputStream`. `write_columns`, `write_sheet` and `WriterHandle::open` take a path or a `std::vector<std::uint8_t>&`. Options are passed as `const Options& = {}`, and the format moved into `OpenOptions::format` and `WriteOptions::format` (`XL_FORMAT_AUTO` infers it from a path's extension).

| v6 | v7 |
|---|---|
| `Workbook::open(path, format, &opts)` | `Workbook::open(path, {.format = format, ...})` |
| `Workbook::open_memory(span, ...)`, `open_source(...)`, `open_stream(...)` | `Workbook::open(span | source | stream, opts)` |
| `sheet_count()`, `sheet_name_at(i)`, `sheet_names()` | `sheets()`, `sheet(i).name()` |
| `Sheet::read_all_decoded()`, `RowCursor::read_all_decoded()` | `Sheet::read_all()` |
| `typed_reader` / `arrow_stream` with `batch_size = 0` | `parse_sheet` / `parse_arrow` |
| `TableView::to_vector()` | `std::vector(t.begin(), t.end())` |
| `handle()`, `from_raw`, constructors from C handles, `to_c()` | removed or private |
| `XlType<T>`, `FieldBinding` | `xl::detail` |
| `OpenOptions::password(sv)` | `OpenOptions::password` field |
| `OpenOptions::csv_max_cell_bytes`, `csv_intern_strings` | `max_cell_bytes`, `intern_strings`, which now apply to CSV too |
| `write_columns_to_memory(format, cols, &opts)` | `write_columns(out_vector, cols, {.format = format})` |
| `write_sheet_to_memory(format, rows, &opts)` | `write_sheet(out_vector, rows, {.format = format})` |
| `format_from_path(path)` | internal |
| `WriterHandle::open_memory(format, &opts)` + `bytes()` | `WriterHandle::open(out_vector, {.format = format})` + `finish()` |
| `aggregate_csv_file` / `aggregate_csv_memory` with `CsvParallelOptions` | `aggregate_csv(path | span, seed, opts, header_row, dop)` |
| `ArrowSchemaGuard`, `ArrowArrayGuard` | `ArrowTable` |

### Rust: one path per item, no raw FFI

Every public item lives at the crate root. The raw C interface (`Xl*`, `xl_*`, `XL_*`) is private, and integer codes became the enums `Format`, `ColumnType` and `ErrorKind`. `OpenOptions` and `WriteOptions` are builder-only and carry the format.

| v6 | v7 |
|---|---|
| `excelreader::workbook::*`, `writer::*`, `rows::*`, … | `excelreader::*` |
| `XL_FORMAT_*`, `XL_T_*`, `Error::code()` | `Format`, `ColumnType`, `Error::kind()` |
| `Workbook::open_with(path, fmt, Some(&o))` | `o.format(fmt).open(path)` |
| `Workbook::open_memory` / `open_source` / `open_reader` | `OpenOptions::open_bytes` / `open_source` / `open_reader` |
| `sheet_count` / `sheet_names` / `sheet_name_at` | `sheets()`, `Sheet::name()` |
| `parse_sheet` / `parse_sheet_parallel` | `sheet.parse::<T>(header_row, parallelism)` |
| `parse_arrow` / `parse_arrow_parallel` / `parse_arrow_stream` | `sheet.parse_arrow::<T>(…)` / `sheet.parse_arrow_stream::<T>(…)` |
| `read_all_decoded` / `read_all_blob` | `read_all()` |
| `infer_schema_parse_text(h, n)` | `infer_schema(h, n, true)` |
| `write_columns(path, fmt, cols, opts)` | `write_columns(path, cols, &opts)`, format in `WriteOptions` |
| `write_columns_to_memory` / `write_sheet_to_memory` | `write_columns_to(writer, …)` / `write_sheet_to(writer, …)` |
| `WriterHandle::open` / `open_with` | `WriterHandle::create(path, &opts)` |
| `WriterHandle::open_memory` + `bytes()` | `WriterHandle::new(writer, &opts)` + `finish()` |
| `WriterHandle::write_null(t)` | `write_*(None)` |
| `aggregate_csv_file` / `aggregate_csv_memory`, `CsvParallelOptions` | `aggregate_csv(input, &opts, header_row, parallelism, seed)` |
| `OpenOptions::csv_max_cell_bytes` / `csv_intern_strings` | `max_cell_bytes` / `intern_strings`, which now apply to CSV too |

### One vocabulary across the bindings

Python, C++ and Rust read a sheet with the same four names: `parse` (typed, whole sheet), `parse_batches` (typed, in batches), `to_arrow` (Arrow, whole sheet) and `to_arrow_stream` (Arrow, in batches). The header row, parallelism and batch size travel together: as keywords in Python, a `ReadOptions` aggregate in C++ and a `ReadOptions` builder in Rust.

| Binding | v7 release candidates | v7 |
|---|---|---|
| Python | `sheet.parse_typed(schema, 1, 0)`, `sheet.iter_parse_typed(schema, 1, 500)` | `sheet.parse(schema, parallelism=0)`, `sheet.parse_batches(schema, batch_size=500)`; every read option is keyword-only |
| Python | `cell.as_date(workbook.is_date1904)` | `cell.as_date()`; the cell carries `date1904` |
| C++ | `parse_sheet<T>(s, h, n)`, `typed_reader<T>(s, h, b)` | `parse<T>(s, {.header_row = h, .parallelism = n})`, `parse_batches<T>(s, {.batch_size = b})` |
| C++ | `parse_arrow<T>(...)`, `arrow_stream<T>(...)` | `to_arrow<T>(s, opts)`, `to_arrow_stream<T>(s, opts)` |
| Rust | `sheet.parse::<T>(h, n)`, `sheet.typed_chunks::<T>(h, b)` | `sheet.parse::<T>(&ReadOptions::new().header_row(h).parallelism(n))`, `sheet.parse_batches::<T>(&opts)` |
| Rust | `sheet.parse_arrow::<T>(...)`, `sheet.parse_arrow_stream::<T>(...)`, `TypedChunks`, `ArrowChunks` | `sheet.to_arrow::<T>(&opts)`, `sheet.to_arrow_stream::<T>(&opts)`, `TypedBatches`, `ArrowBatches` |
| Rust | `Date::new(d)`, `Time::new(t)`, `Timestamp::new(ts)` | `Date { days_since_epoch: d }`, `Time { micros_since_midnight: t }`, `Timestamp { micros_since_epoch: ts }` |
