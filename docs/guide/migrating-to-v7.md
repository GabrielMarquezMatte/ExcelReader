# Migrating from v6 to v7

What changed in parallel reading and how to port v6 code.
See also [reading.md](reading.md#parallel-reading-opt-in) for the full guide.

## What changed

Parallel reading now takes a sheet, and it works on XLSX and XLSB as well as CSV. `CsvParallel` is gone: its eighteen path, memory and stream methods became six `SheetParallel` methods that take any `IExcelSheet`. You open the workbook yourself, pass its sheet, and the library parallelises as far as the format allows. The contract types lost their `Csv` prefix because they now serve every format.

## Types

| v6 | v7 |
|---|---|
| `CsvParallel` | `SheetParallel` |
| `CsvParallelOptions` | `SheetParallelOptions` (no `Reader` property) |
| `ICsvRecord<TSelf>` | `IRowRecord<TSelf>` |
| `ICsvAccumulator<TSelf, TModel>` | `IRowAccumulator<TSelf, TModel>` |
| `CsvAggregation<TState>` | `RowAggregation<TState>` |
| `CsvRowAction<TState>` | `RowAction<TState>` |

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
using CsvReader csv = Excel.FromCsvFile(path, options);
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
using CsvReader csv = Excel.FromCsv(data, options);
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
using CsvReader csv = Excel.FromCsv(stream, leaveOpen: true, options);
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
- The CSV dialect is given when the workbook is opened (`CsvReaderOptions`), not in the parallel options.
- A CSV in a `MemoryStream` is no longer partitioned: the stream overload of `Excel.FromCsv` reads it sequentially. Pass the bytes as `ReadOnlyMemory<byte>` to keep partitioning.
- `HeaderRow` rules are the same as v6. `ParseAsync` takes the header from `ExcelParserOptions.HeaderRow` and rejects a non-zero `SheetParallelOptions.HeaderRow`; the parser-based `ForEachAsync` and `AggregateAsync` use `SheetParallelOptions.HeaderRow`.
- XLSX and XLSB sheets are now read on several threads when `DegreeOfParallelism` allows it. XLS, a table view (`ExcelTable.AsSheet()`), a sheet under 256 KB uncompressed and a CSV that cannot be partitioned still read sequentially with the same results.
- `PrefetchDecompression` does not apply to a parallel XLSX or XLSB read; the pipeline already overlaps inflate and parse.

## Native bindings

The C ABI is unchanged: no signature moved, and `XL_ABI_VERSION` stays 6. Behavior changed. `degree_of_parallelism` of `0` on `xl_parse_typed` and the Arrow exports now parallelises XLSX and XLSB as well as CSV, using every processor. Pass `1` for the old sequential behavior. The Rust, Python and C++ bindings take the same value (`degree_of_parallelism`, `parallelism`) and need no code change; a `parallelism=0` call on an XLSX or XLSB sheet is now faster, and holds every partition's columns until they merge, so its peak memory is higher.

`xl_csv_aggregate` keeps its signature. Inside, it opens a `CsvReader` with the dialect from the native struct and calls `SheetParallel`.
