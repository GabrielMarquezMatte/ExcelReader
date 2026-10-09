# Benchmarks

Throwaway numbers go stale; these are regenerated from the benchmark suite in
`tests/ExcelReader.Benchmarks`. Live results: https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.11 (SDK 10.0.400). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [GitHub Pages benchmark dashboard](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/) — `tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/` is `.gitignore`-excluded and never actually reaches GitHub, so a link to it 404s for every reader.

**Benchmark methodology.** In the Excel-format tables below, the **"Cell-by-cell read"** rows (including the per-format rows in "Real data reads" / "String-heavy reads") read ExcelReader's `cell.Value` — a zero-copy `ReadOnlySpan<byte>`, no decode or allocation — against each competitor's own idiomatic read call. For Sylvan.Data.Excel, that's the ADO.NET-style `GetString(i)`, which is forced to materialize a UTF-16 `string`; its API has no zero-copy accessor, so it cannot avoid that cost. These rows are therefore not matched work: part of the reported gap is "we parse faster" and part is "we skipped an allocation you were never offered a way to skip." Each affected benchmark class also has a `*_Materialized` sibling (calling `cell.GetString()`, the same UTF-16 materialization Sylvan pays) so the matched-work number is measurable — see `tests/ExcelReader.Benchmarks/Shared/BenchmarkAccumulators.cs`. **Those results are published under [Matched-work reads](#matched-work-reads-_materialized), and reading them is the honest way to judge the comparison:** treat the Excel cell-by-cell ratios in the tables below as an upper bound, not a like-for-like number. The CSV cell-by-cell rows are matched work: Sep and Sylvan.Data.Csv both expose span accessors (`row[i].Span`, `GetFieldSpan(i)`), and the benchmarks use them. The **"Typed row parsing"** / **"Typed record writing"** rows are unaffected — both sides already materialize real objects/strings there, so those comparisons are matched work as published.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | Sylvan | OfficeIMO | SpreadCheetah |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 7.207 ms, 3.82 KB | 35.203 ms, 1.89 MB | 104.726 ms, 6.84 MB | - |
| Cell-by-cell read async | 7.953 ms, 3.89 KB | - | - | - |
| Typed row parsing | 10.935 ms, 3.87 MB | 56.855 ms, 10.47 MB | 104.530 ms, 5.69 MB | - |
| Typed row parsing async | 11.241 ms, 3.87 MB | 60.501 ms, 10.48 MB | - | - |
| Typed row parsing, shared strings | 10.849 ms, 2.29 MB | - | - | - |
| Workbook writing | 10.913 ms, 4.02 MB | - | 18.166 ms, 4.02 MB | 15.128 ms, 15.84 MB |
| Workbook writing, shared strings | 10.848 ms, 4.06 MB | - | - | - |

ExcelReader is ~4.9x faster than Sylvan for raw XLSX reads, allocating ~508x less, and ~14.5x faster than OfficeIMO. For typed parsing it is ~5.2x faster than Sylvan (~5.4x async against async). For XLSX writing, ExcelReader is ~1.4x faster than SpreadCheetah while allocating ~3.9x less memory, and ~1.7x faster than OfficeIMO.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~1% faster than the inline-string sheet above (10.849 ms vs. 10.935 ms) and allocates ~41% less (2.29 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | OfficeIMO |
|---|---:|---:|
| Cell-by-cell read | 5.178 ms, 4.13 KB | 10.406 ms, 6.70 MB |
| Cell-by-cell read async | 5.249 ms, 4.20 KB | - |
| Typed row parsing | 8.211 ms, 3.87 MB | - |
| Typed row parsing async | 8.261 ms, 3.87 MB | - |
| Workbook writing | 7.478 ms, 4.02 MB | - |
| Workbook writing, shared strings | 7.370 ms, 4.06 MB | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than XLSX reads, typed parsing is ~1.3x faster than XLSX parsing, and writing is ~1.5x faster than XLSX writing. Against OfficeIMO's XLSB reader, ExcelReader is ~2.0x faster while allocating ~1,660x less. The XLSB writer is also ~2.0x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 3.916 ms, 2.84 KB | 5.173 ms, 1,717.73 KB | 6.856 ms, 10,499.25 KB |
| Cell-by-cell read async | 3.236 ms, 2.91 KB | - | - |
| Workbook writing | 5.296 ms, 16.03 MB | - | - |

ExcelReader is ~1.3x faster than Sylvan for generated XLS reads while allocating ~605x less memory, and ~1.8x faster than OfficeIMO. The XLS writer is ~2.1x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 5.296 ms vs 10.964 ms in the same run). Most of its 16.03 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.03 MB like the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 3.053 ms, 848 B | 8.012 ms, 3.93 KB | 4.359 ms, 37.78 KB |
| Cell-by-cell read async | 2.941 ms, 920 B | - | - |
| Typed row parsing | 4.631 ms, 3.86 MB | 8.454 ms, 3.87 MB | 11.832 ms, 10.95 MB |
| Typed row parsing async | 4.987 ms, 3.86 MB | - | - |
| Row writing | 4.749 ms, 4.00 MB | 7.030 ms, 4.01 MB | 7.276 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~2.6x faster than Sep while allocating ~5x less, and ~1.4x faster than Sylvan.Data.Csv while allocating ~46x less. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.8x faster than Sep and ~2.6x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`CsvParallel.ParseAsync<T>` and `CsvParallel.AggregateAsync` on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|
| 1 | 904.0 ms, 671.42 MB | 518.6 ms, 245.76 MB | 784.1 ms, 1.20 MB | 478.1 ms, 1.20 MB |
| 2 | 623.7 ms, 680.13 MB | 382.9 ms, 254.68 MB | 405.4 ms, 1.26 MB | 262.9 ms, 1.28 MB |
| 4 | 343.0 ms, 680.24 MB | 217.2 ms, 254.72 MB | 218.3 ms, 1.37 MB | 137.6 ms, 1.44 MB |
| 8 | 285.2 ms, 680.15 MB | 204.7 ms, 254.63 MB | 158.3 ms, 1.54 MB | 94.6 ms, 1.62 MB |
| 16 | 324.9 ms, 680.32 MB | 189.0 ms, 255.15 MB | 114.4 ms, 1.66 MB | 73.0 ms, 1.69 MB |

The typed path scales ~2.8x on the conversion-heavy corpus (bimodal at dop 16: 86.7 ms standard deviation, median 275.8 ms) and ~2.7x on narrow ints, where it flattens out after dop 4 (217.2 ms vs. 189.0 ms at dop 16). Its ~680 MB is one materialized row object per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~6.9x (conversion-heavy) and ~6.5x (narrow) at dop 16 — because they are not fighting the allocator: at dop 16 the conversion-heavy aggregate is ~2.8x faster than typed while allocating ~410x less (1.66 MB against 680.32 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 784.1 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 490.6 ms, 244.15 MB | 187.6 ms, 255.14 MB | 521.7 ms, 244.15 MB | 429.3 ms, 247.77 MB |
| Conversion-heavy | 626.6 ms, 467.31 MB | 184.9 ms, 474.82 MB | 817.5 ms, 467.31 MB | 605.5 ms, 468.84 MB |

Sequentially, ExcelReader's typed parser is ~1.1x faster than Sep on narrow ints (490.6 ms vs. 521.7 ms), and ~1.3x faster on the conversion-heavy corpus. In parallel ExcelReader is ~2.3x (narrow) and ~3.3x (conversion-heavy) faster than Sep's, while yielding rows in file order.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 43.792 ms, 7.23 KB | 29.544 ms, 23.39 KB | 201.558 ms, 644.23 KB |
| XLSM | 44.486 ms, 7.23 KB | 29.169 ms, 23.42 KB | 207.288 ms, 644.30 KB |
| XLSB | 25.704 ms, 6.79 KB | 17.403 ms, 17.01 KB | 30.864 ms, 338.54 KB |
| XLS | 11.335 ms, 12.02 KB | n/a | 19.146 ms, 185.90 KB |
| CSV | 5.284 ms, 840 B | n/a | 4.630 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~4.6x faster than Sylvan for XLSX and ~4.7x for XLSM, ~1.2x faster for XLSB, and ~1.7x faster for XLS — allocating ~89x less for XLSX and XLSM, ~51x less for XLSB and ~15x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~14% faster, while ExcelReader allocates ~49x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 44.249 ms | 0.341 ms | 6.98 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 30.115 ms | 0.803 ms | 23.19 KB |
| Xlsm_ExcelReader_Memory | 45.242 ms | 0.357 ms | 6.98 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 29.307 ms | 0.718 ms | 23.05 KB |
| Xlsb_ExcelReader_Memory | 25.606 ms | 0.255 ms | 8.49 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 16.688 ms | 0.857 ms | 16.64 KB |
| Xls_ExcelReader_Memory | 11.330 ms | 0.105 ms | 11.88 KB |
| Csv_ExcelReader_Memory | 4.919 ms | 0.043 ms | 568 B |

`Csv_ExcelReader_Memory` is both faster (4.919 ms vs. 5.284 ms for `Csv_ExcelReader`) and allocates less (568 B vs. 840 B) than its stream twin.

`Xls_ExcelReader_Memory` matches its stream twin (11.330 ms vs. 11.335 ms for `Xls_ExcelReader`), with near-identical allocation (11.88 KB vs. 12.02 KB). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 33.16 ms, 2.18 MB | 23.35 ms, 2.21 MB | 187.18 ms, 17.41 MB |
| XLSB | 35.92 ms, 2.18 MB | 21.86 ms, 2.21 MB | 66.81 ms, 17.38 MB |

Both formats handle this well: ~5.6x and ~1.9x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 33.79 ms | 2.18 MB |
| Whole read, prefetch | 23.61 ms | 2.21 MB |
| Shared-string table only (open + first row) | 9.72 ms | 747 KB |
| Whole read, stored | 20.10 ms | 2.18 MB |
| Shared-string table only, stored | 3.24 ms | 746 KB |

About 40% of the default read is inflate: without it the whole read drops from 33.79 ms to 20.10 ms, and the shared-string table from 9.72 ms to 3.24 ms. What remains, ~17 ms, is the worksheet scan.

### Inflate

`InflateBenchmark` decompresses one ZIP entry per method with `System.IO.Compression.DeflateStream` and with the library's managed `InflateStream`, which every ZIP-based reader now uses:

| Entry | `DeflateStream` | `InflateStream` | Ratio |
|---|---:|---:|---:|
| XLSX `sheet1.xml` (6.7 MB → 34.4 MB) | 23.433 ms, 280 B | 16.501 ms, 200 B | 0.70 |
| XLSB `sheet1.bin` (3.7 MB → 15.9 MB) | 13.230 ms, 280 B | 10.259 ms, 200 B | 0.78 |
| String-heavy `sharedStrings.xml` | 10.522 ms, 281 B | 6.659 ms, 200 B | 0.63 |

Without prefetch a read is inflate plus parse, so the faster inflate converts directly; with prefetch the two overlap and the read is bounded by the slower stage, which is why the prefetch columns above moved little.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 46.53 ms | 0.19 ms | 16.70 KB |
| XLSM | 46.16 ms | 0.48 ms | 16.70 KB |
| XLSB | 28.46 ms | 0.58 ms | 16.26 KB |
| XLS | 13.36 ms | 0.27 ms | 21.48 KB |
| CSV | 19.38 ms | 0.10 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 8.41 ms | 0.02 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 57.19 ms | 2.01 ms | 1,111.1 | 1,000.0 | 333.3 | 15.25 MB |
| XLSB | 57.84 ms | 1.98 ms | 1,100.0 | 1,000.0 | 300.0 | 15.25 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 840 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 16.26 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.25 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan, with 333 Gen2 collections, and is still ~3.3x faster (57.19 ms vs. 187.18 ms). The same allocation holds for XLSB (15.25 MB vs. Sylvan's 17.38 MB, also ~12% less, 300 Gen2 collections), where matched-work time is ~14% ahead of Sylvan (57.84 ms vs. 66.81 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 12.099 ms | 4.02 MB |
| XLSB | 7.597 ms | 4.02 MB |
| XLS | 5.021 ms | 4.03 MB |
| CSV | 4.762 ms | 4.00 MB |

Relative ordering: CSV and XLS close at the front, then XLSB, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 7.597 ms vs. 7.478 ms, XLSX 12.099 ms vs. 10.913 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 20.487 ms | 18.02 KB |
| Native string columns → CSV | 5.098 ms | 416 B |
| Arrow `StringArray` → XLSX | 26.644 ms | 18.06 KB |
| Styled rows → XLSX | 7.068 ms | 18.20 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~18 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 10.94 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 9.93 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 9.30 ms | 4.36 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.9% less than the `class` baseline — and is ~6% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`ExcelDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 6.878 ms | 1.00 | 3.87 KB |
| `GetBytes` (text column only) | 7.222 ms | 1.05 | 4.55 KB |
| `GetValue` (every column) | 9.660 ms | 1.40 | 5.01 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 9.059 ms | 1.32 | 1.58 MB |
| `DataTable.Load` | 63.615 ms | 9.25 | 22.84 MB |

The bridge costs ~1.3–1.4x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it stays within ~5% of the raw read. `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 9.660 ms to read through `GetValue` alone.

### Arrow conversion

`ArrowConversionBenchmark`: `ToArrowRecordBatch` from [ExcelReader.Arrow](../guide/reading.md#apache-arrow-conversion-opt-in), 100,000 rows into one `RecordBatch`. The typed rows use the generated records (a text name, an int64 id, a timestamp, a double) as CSV and as XLSB. The all-string row is a separate header-less CSV with 8 text columns.

| Scenario | Mean | Allocated |
|---|---:|---:|
| CSV, typed, explicit schema | 10.76 ms | 8.13 MB |
| CSV, typed, inferred schema | 10.77 ms | 14.89 MB |
| XLSB, typed, explicit schema | 15.85 ms | 8.14 MB |
| CSV, 8 text columns, explicit schema | 17.22 ms | 16.27 MB |

On the same file, letting `Excel.InferSchema` guess the schema costs about the same time (10.77 ms vs. 10.76 ms) and ~1.8x the allocation of passing an explicit `ExcelColumnSchema[]`. The allocation grows with the sheet, because the batch holds every row at once and there is no streaming variant.

### Encrypted workbooks

`EncryptedWorkbookBenchmark`: agile encryption (AES-256, SHA-512, 100,000 spin iterations — see [Encrypted workbooks](../guide/encryption.md)). The small rows use a 15 KB fixture; the large rows use a generated 300,000-row workbook, encrypted with `Excel.EncryptPackage`.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Small, plain | 0.198 ms | 9.15 KB |
| Small, encrypted, stream | 26.129 ms | 38.39 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 25.641 ms | 40.54 KB |
| Small, encrypted, in memory | 26.001 ms | 53.53 KB |
| Small, encrypted, open only | 25.780 ms | 32.72 KB |
| 300,000 rows, plain | 39.180 ms | 3.87 KB |
| 300,000 rows, encrypted | 66.278 ms | 29.79 KB |
| 300,000 rows, encrypted, open only | 25.744 ms | 29.10 KB |

Nearly all of the cost is a fixed ~26 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 25.78 ms for 15 KB and 25.74 ms for 300,000 rows). After that, decrypting while reading costs ~3% over a plain read: 66.28 ms − 25.74 ms = 40.53 ms, against 39.18 ms unencrypted. On the small fixture, integrity verification and the in-memory path cost nothing measurable. The fixture is too small to show the extra pass that verification makes over a large stream.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 40.69 ms | 21.09 KB |
| First typed record write | 18.57 ms | 84.25 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 40.69 ms | 21.09 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 36.11 ms | 24.03 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 46.15 ms | 26.16 KB |

`Build` is ~11% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slower than reflection (~13%, with a 5.4 ms standard deviation), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

### Native C ABI reads

`NativeRowReadBenchmark` and `NativeTypedParseBenchmark` call the C ABI entry points from .NET, reading the real-data XLSB workbook (65,535 rows × 14 columns) from memory. This isolates what the ABI itself costs, without any binding's marshalling on top.

| Entry point | Shape handed back | Mean | Managed allocated |
|---|---|---:|---:|
| `xl_next_row` | one row per call, serialized into a caller buffer | 40.35 ms | 3.71 MB |
| `xl_read_all_blob` | whole sheet in one buffer | 43.59 ms | 21.06 MB |
| `xl_read_all_decoded` | whole sheet as native row structs | 54.56 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 40.79 ms | 11.61 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 39.51 ms | 11.62 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Row by row, `xl_next_row` costs ~1.6x the managed in-memory read of the same file (40.35 ms vs. 25.606 ms for `Xlsb_ExcelReader_Memory` under [In-memory real-data reads](#in-memory-real-data-reads)). The managed figure is from a later run, after `InflateStream`; this native table has not been re-run since, so read the ratio as an upper bound. Every cell crosses the ABI as text, so each binary XLSB number is formatted to UTF-8 on the way out. `xl_read_all_blob` adds ~8% over it, because the whole sheet is buffered before it is copied out. `xl_read_all_decoded` adds ~35%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` cost about the same as a row-by-row read but return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 8.73 ms | 1.29 MB |
| Batches of 10,000 rows | 8.49 ms | 1.59 MB |
| Batches of 1,000 rows | 8.47 ms | 1.51 MB |

Batching costs nothing in time: both batch sizes land within ~3% of one whole-sheet table. A caller can bound its native memory to one batch at a time without paying for it.

### C++ and Rust bindings

The C++ and Rust wrappers around the native library each have their own benchmark suite (Google
Benchmark and Criterion respectively), measured on the same machine as the .NET results above.
Full tables and how to run them locally: [`cpp/README.md#benchmarks`](../../cpp/README.md#benchmarks) and
[`rust/excelreader/README.md#benchmarks`](../../rust/excelreader/README.md#benchmarks).

Headline numbers, reading a real 65,535-row workbook in full (all 14 columns) against each
language's own competing library, both sides decoding into owned values so neither gets a
zero-copy advantage the other can't take:

| Comparison | Format | ExcelReader | Competitor | Ratio |
|---|---|---:|---:|---:|
| vs. [calamine](https://github.com/tafia/calamine) (Rust) | XLSX | 79.1 ms | 279.7 ms | ~3.5x faster |
| vs. calamine (Rust) | XLSB | 46.9 ms | 87.5 ms | ~1.9x faster |
| vs. [DuckDB](https://github.com/duckdb/duckdb) `read_xlsx` (C++) | XLSX | 62.9 ms | 410.8 ms | ~6.5x faster |
| vs. [xlsxio](https://github.com/brechtsanders/xlsxio) (C) | XLSX | 62.9 ms | 512.9 ms | ~8.2x faster |
| vs. [xlnt](https://github.com/tfussell/xlnt) (C++) | XLSX | 62.9 ms | 2,394.1 ms | ~38x faster |
| vs. [zsv](https://github.com/liquidaty/zsv) `fast` engine (C) | CSV | 27.2 ms | 20.4 ms | ~1.3x **slower** |

Writing the same 14 columns × 65,535 rows from row-shaped data, the C++ suite measures
`write_sheet` at 95.5–96.5 ms against DuckDB's 829–833 ms (~8.6–8.7x), libxlsxwriter's 1,194 ms (~12x),
xlsxio's 2,433 ms (~25x) and xlnt's 5,436 ms (~57x). The Rust suite measures 51.9 ms against
rust_xlsxwriter's 328.3 ms (~6.3x) over 7 columns. Caveats for both are in the binding READMEs.

DuckDB, xlsxio and xlnt do not read `.xlsb`, so those comparisons are XLSX-only. calamine is a fast,
well-optimized reader in its own right — the gap there is real but not the order of magnitude seen
against the C/C++ competitors.

zsv is the one competitor ahead. The CSV row is the typed workload, where both sides build the same
14-column rows with owned strings. Just counting cell bytes, zsv is ~3.3x faster, and that gap is
ExcelReader's C ABI, which serializes every row into a buffer, not its CSV parser. That suite is
built with GCC (zsv does not compile under MSVC). Its caveats are in the C++ README.

