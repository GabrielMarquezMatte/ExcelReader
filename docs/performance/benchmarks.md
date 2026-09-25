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
| Cell-by-cell read | 7.778 ms, 11.18 KB | 36.633 ms, 1.89 MB | 105.146 ms, 6.84 MB | - |
| Cell-by-cell read async | 8.322 ms, 13.31 KB | - | - | - |
| Typed row parsing | 11.304 ms, 3.87 MB | 56.964 ms, 10.47 MB | 106.258 ms, 5.69 MB | - |
| Typed row parsing async | 11.696 ms, 3.88 MB | 60.071 ms, 10.48 MB | - | - |
| Typed row parsing, shared strings | 11.038 ms, 2.30 MB | - | - | - |
| Workbook writing | 10.910 ms, 4.02 MB | - | 18.042 ms, 4.02 MB | 15.136 ms, 15.84 MB |
| Workbook writing, shared strings | 10.508 ms, 4.06 MB | - | - | - |

ExcelReader is ~4.7x faster than Sylvan for raw XLSX reads, allocating ~173x less, and ~13.5x faster than OfficeIMO. For typed parsing it is ~5.0x faster than Sylvan (~5.1x async against async). For XLSX writing, ExcelReader is ~1.4x faster than SpreadCheetah while allocating ~3.9x less memory, and ~1.7x faster than OfficeIMO.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~2% faster than the inline-string sheet above (11.038 ms vs. 11.304 ms) and allocates ~41% less (2.30 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | OfficeIMO |
|---|---:|---:|
| Cell-by-cell read | 5.409 ms, 13.28 KB | 11.495 ms, 6.70 MB |
| Cell-by-cell read async | 5.756 ms, 15.84 KB | - |
| Typed row parsing | 8.247 ms, 3.88 MB | - |
| Typed row parsing async | 8.319 ms, 3.88 MB | - |
| Workbook writing | 7.555 ms, 4.02 MB | - |
| Workbook writing, shared strings | 7.268 ms, 4.06 MB | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than XLSX reads, typed parsing is ~1.4x faster than XLSX parsing, and writing is ~1.4x faster than XLSX writing. Against OfficeIMO's XLSB reader, ExcelReader is ~2.1x faster while allocating ~516x less. The XLSB writer is also ~2.0x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 3.315 ms, 2.45 KB | 5.308 ms, 1,717.73 KB | 6.876 ms, 10,499.25 KB |
| Cell-by-cell read async | 3.252 ms, 2.52 KB | - | - |
| Workbook writing | 5.150 ms, 16.03 MB | - | - |

ExcelReader is ~1.6x faster than Sylvan for generated XLS reads while allocating ~700x less memory, and ~2.1x faster than OfficeIMO. The XLS writer is ~2.1x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 5.150 ms vs 10.802 ms in the same run). Most of its 16.03 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.03 MB like the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 3.609 ms, 440 B | 7.709 ms, 3.93 KB | 4.495 ms, 37.78 KB |
| Cell-by-cell read async | 2.911 ms, 512 B | - | - |
| Typed row parsing | 5.065 ms, 3.86 MB | 8.880 ms, 3.87 MB | 12.424 ms, 10.95 MB |
| Typed row parsing async | 5.129 ms, 3.86 MB | - | - |
| Row writing | 4.624 ms, 4.00 MB | 6.919 ms, 4.01 MB | 6.940 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~2.1x faster than Sep while allocating ~9x less, and ~1.2x faster than Sylvan.Data.Csv while allocating ~88x less. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.8x faster than Sep and ~2.5x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`CsvParallel.ParseAsync<T>` and `CsvParallel.AggregateAsync` on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|
| 1 | 892.6 ms, 670.98 MB | 511.0 ms, 245.32 MB | 802.7 ms, 847.48 KB | 486.8 ms, 842.34 KB |
| 2 | 619.8 ms, 680.07 MB | 390.6 ms, 254.64 MB | 420.9 ms, 1.24 MB | 269.7 ms, 1.29 MB |
| 4 | 360.6 ms, 680.24 MB | 219.3 ms, 254.69 MB | 218.8 ms, 1.35 MB | 137.5 ms, 1.43 MB |
| 8 | 282.3 ms, 680.11 MB | 190.2 ms, 254.59 MB | 148.7 ms, 1.55 MB | 95.0 ms, 1.63 MB |
| 16 | 257.0 ms, 680.31 MB | 188.1 ms, 255.11 MB | 109.8 ms, 1.66 MB | 72.4 ms, 1.71 MB |

The typed path scales ~3.5x on the conversion-heavy corpus and ~2.7x on narrow ints, where it stops improving after dop 8 (190.2 ms vs. 188.1 ms at dop 16). Its ~680 MB is one materialized row object per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~7.3x (conversion-heavy) and ~6.7x (narrow) at dop 16 — because they are not fighting the allocator: at dop 16 the conversion-heavy aggregate is ~2.3x faster than typed while allocating ~410x less (1.66 MB against 680.31 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 802.7 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 504.3 ms, 244.15 MB | 198.3 ms, 255.11 MB | 508.9 ms, 244.15 MB | 424.6 ms, 247.08 MB |
| Conversion-heavy | 620.6 ms, 467.31 MB | 230.7 ms, 474.77 MB | 823.3 ms, 467.31 MB | 594.1 ms, 469.95 MB |

Sequentially, ExcelReader's typed parser and Sep are tied on narrow ints (504.3 ms vs. 508.9 ms), and ExcelReader is ~1.3x faster than Sep on the conversion-heavy corpus. In parallel ExcelReader is ~2.1x (narrow) and ~2.6x (conversion-heavy) faster than Sep's, while yielding rows in file order. The conversion-heavy parallel row is bimodal in this run (66.8 ms standard deviation, median 188.9 ms), so its ratio is the least stable number in the table.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 54.064 ms, 17.86 KB | 31.842 ms, 37.46 KB | 201.175 ms, 644.23 KB |
| XLSM | 53.988 ms, 17.86 KB | 31.343 ms, 35.83 KB | 207.479 ms, 644.30 KB |
| XLSB | 28.944 ms, 18.80 KB | 16.532 ms, 29.34 KB | 30.223 ms, 338.54 KB |
| XLS | 10.584 ms, 11.30 KB | n/a | 18.765 ms, 185.90 KB |
| CSV | 5.082 ms, 376 B | n/a | 4.611 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~3.7x faster than Sylvan for XLSX and ~3.8x for XLSM, essentially tied for XLSB (~1.04x), and ~1.8x faster for XLS — allocating ~36x less for XLSX and XLSM, ~18x less for XLSB and ~16x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~10% faster, while ExcelReader allocates ~110x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 53.556 ms | 0.628 ms | 7.01 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 31.833 ms | 0.356 ms | 26.23 KB |
| Xlsm_ExcelReader_Memory | 54.460 ms | 0.879 ms | 7.01 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 31.199 ms | 0.209 ms | 25.61 KB |
| Xlsb_ExcelReader_Memory | 28.853 ms | 0.435 ms | 8.69 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 17.263 ms | 0.311 ms | 17.14 KB |
| Xls_ExcelReader_Memory | 10.143 ms | 0.059 ms | 11.30 KB |
| Csv_ExcelReader_Memory | 4.884 ms | 0.055 ms | 312 B |

`Csv_ExcelReader_Memory` is both faster (4.884 ms vs. 5.082 ms for `Csv_ExcelReader`) and allocates less (312 B vs. 376 B) than its stream twin.

`Xls_ExcelReader_Memory` is ~4% faster than its stream twin (10.143 ms vs. 10.584 ms for `Xls_ExcelReader`), with identical allocation (11.30 KB both). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 42.19 ms, 2.19 MB | 28.22 ms, 2.27 MB | 187.79 ms, 17.41 MB |
| XLSB | 39.43 ms, 2.19 MB | 25.56 ms, 2.26 MB | 61.75 ms, 17.38 MB |

Both formats handle this well: ~4.5x and ~1.6x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 41.59 ms | 2.19 MB |
| Whole read, prefetch | 28.37 ms | 2.27 MB |
| Shared-string table only (open + first row) | 13.70 ms | 755 KB |
| Whole read, stored | 20.44 ms | 2.19 MB |
| Shared-string table only, stored | 3.14 ms | 754 KB |

About half of the default read is inflate: without it the whole read drops from 41.59 ms to 20.44 ms, and the shared-string table from 13.70 ms to 3.14 ms. What remains, ~17 ms, is the worksheet scan.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 54.55 ms | 0.37 ms | 27.33 KB |
| XLSM | 54.00 ms | 0.16 ms | 27.33 KB |
| XLSB | 30.38 ms | 0.56 ms | 28.27 KB |
| XLS | 12.10 ms | 0.07 ms | 20.77 KB |
| CSV | 19.56 ms | 0.49 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 9.08 ms | 0.30 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 63.39 ms | 1.21 ms | 1,000.0 | 875.0 | 250.0 | 15.26 MB |
| XLSB | 59.95 ms | 0.96 ms | 1,111.1 | 1,000.0 | 333.3 | 15.26 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 376 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 28.27 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.26 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan, with 250 Gen2 collections, and is still ~3.0x faster (63.39 ms vs. 187.79 ms). The same allocation holds for XLSB (15.26 MB vs. Sylvan's 17.38 MB, also ~12% less, 333 Gen2 collections), where matched-work time is ~3% ahead of Sylvan (59.95 ms vs. 61.75 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 11.684 ms | 4.02 MB |
| XLSB | 7.552 ms | 4.02 MB |
| XLS | 4.835 ms | 4.03 MB |
| CSV | 4.701 ms | 4.00 MB |

Relative ordering: CSV and XLS close at the front, then XLSB, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 7.552 ms vs. 7.555 ms, XLSX 11.684 ms vs. 10.910 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 20.487 ms | 18.02 KB |
| Native string columns → CSV | 5.053 ms | 416 B |
| Arrow `StringArray` → XLSX | 26.360 ms | 18.06 KB |
| Styled rows → XLSX | 7.136 ms | 18.20 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~18 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 11.30 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 10.97 ms | 1.59 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 9.99 ms | 11.72 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.7% less than the `class` baseline — and is ~9% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`ExcelDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 7.485 ms | 1.00 | 11.18 KB |
| `GetBytes` (text column only) | 7.628 ms | 1.02 | 11.87 KB |
| `GetValue` (every column) | 10.550 ms | 1.41 | 5.02 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 10.850 ms | 1.45 | 1.59 MB |
| `DataTable.Load` | 61.443 ms | 8.21 | 22.85 MB |

The bridge costs ~1.4–1.5x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it stays within ~2% of the raw read. `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 10.550 ms to read through `GetValue` alone.

### Arrow conversion

`ArrowConversionBenchmark`: `ToArrowRecordBatch` from [ExcelReader.Arrow](../guide/reading.md#apache-arrow-conversion-opt-in), 100,000 rows into one `RecordBatch`. The typed rows use the generated records (a text name, an int64 id, a timestamp, a double) as CSV and as XLSB. The all-string row is a separate header-less CSV with 8 text columns.

| Scenario | Mean | Allocated |
|---|---:|---:|
| CSV, typed, explicit schema | 11.01 ms | 8.13 MB |
| CSV, typed, inferred schema | 12.44 ms | 14.89 MB |
| XLSB, typed, explicit schema | 16.47 ms | 8.15 MB |
| CSV, 8 text columns, explicit schema | 19.92 ms | 16.27 MB |

On the same file, letting `Excel.InferSchema` guess the schema costs ~13% more time and ~1.8x the allocation of passing an explicit `ExcelColumnSchema[]`. The allocation grows with the sheet, because the batch holds every row at once and there is no streaming variant.

### Encrypted workbooks

`EncryptedWorkbookBenchmark`: agile encryption (AES-256, SHA-512, 100,000 spin iterations — see [Encrypted workbooks](../guide/encryption.md)). The small rows use a 15 KB fixture; the large rows use a generated 300,000-row workbook, encrypted with `Excel.EncryptPackage`.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Small, plain | 0.184 ms | 19.42 KB |
| Small, encrypted, stream | 26.272 ms | 48.52 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 25.905 ms | 50.73 KB |
| Small, encrypted, in memory | 26.139 ms | 53.13 KB |
| Small, encrypted, open only | 25.982 ms | 43.06 KB |
| 300,000 rows, plain | 45.711 ms | 11.18 KB |
| 300,000 rows, encrypted | 76.240 ms | 37.10 KB |
| 300,000 rows, encrypted, open only | 25.998 ms | 36.51 KB |

Nearly all of the cost is a fixed ~26 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 25.98 ms for 15 KB and 26.00 ms for 300,000 rows). After that, decrypting while reading costs ~10% over a plain read: 76.24 ms − 26.00 ms = 50.24 ms, against 45.71 ms unencrypted. On the small fixture, integrity verification and the in-memory path cost nothing measurable. The fixture is too small to show the extra pass that verification makes over a large stream.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 33.17 ms | 28.43 KB |
| First typed record write | 18.49 ms | 84.25 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 33.17 ms | 28.43 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 28.47 ms | 31.37 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 35.96 ms | 33.49 KB |

`Build` is ~14% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slightly slower than reflection (~8%), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

### Native C ABI reads

`NativeRowReadBenchmark` and `NativeTypedParseBenchmark` call the C ABI entry points from .NET, reading the real-data XLSB workbook (65,535 rows × 14 columns) from memory. This isolates what the ABI itself costs, without any binding's marshalling on top.

| Entry point | Shape handed back | Mean | Managed allocated |
|---|---|---:|---:|
| `xl_next_row` | one row per call, serialized into a caller buffer | 41.62 ms | 3.71 MB |
| `xl_read_all_blob` | whole sheet in one buffer | 44.18 ms | 21.06 MB |
| `xl_read_all_decoded` | whole sheet as native row structs | 56.87 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 42.20 ms | 11.61 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 41.51 ms | 11.62 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Row by row, `xl_next_row` costs ~1.4x the managed in-memory read of the same file (41.62 ms vs. 28.853 ms for `Xlsb_ExcelReader_Memory` under [In-memory real-data reads](#in-memory-real-data-reads); different benchmark class, same run). Every cell crosses the ABI as text, so each binary XLSB number is formatted to UTF-8 on the way out. `xl_read_all_blob` adds ~6% over it, because the whole sheet is buffered before it is copied out. `xl_read_all_decoded` adds ~37%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` cost about the same as a row-by-row read but return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 9.79 ms | 1.31 MB |
| Batches of 10,000 rows | 9.57 ms | 1.60 MB |
| Batches of 1,000 rows | 9.68 ms | 1.51 MB |

Batching costs nothing in time: both batch sizes land within ~2% of one whole-sheet table. A caller can bound its native memory to one batch at a time without paying for it.

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
| vs. [calamine](https://github.com/tafia/calamine) (Rust) | XLSX | 85.8 ms | 272.0 ms | ~3.2x faster |
| vs. calamine (Rust) | XLSB | 53.0 ms | 86.8 ms | ~1.6x faster |
| vs. [DuckDB](https://github.com/duckdb/duckdb) `read_xlsx` (C++) | XLSX | 73.7 ms | 433.6 ms | ~5.9x faster |
| vs. [xlsxio](https://github.com/brechtsanders/xlsxio) (C) | XLSX | 73.7 ms | 505.0 ms | ~6.9x faster |
| vs. [xlnt](https://github.com/tfussell/xlnt) (C++) | XLSX | 73.7 ms | 2,494.2 ms | ~34x faster |
| vs. [zsv](https://github.com/liquidaty/zsv) `fast` engine (C) | CSV | 30.4 ms | 22.7 ms | ~1.3x **slower** |

Writing the same 14 columns × 65,535 rows from row-shaped data, the C++ suite measures
`write_sheet` at 97.8–100.1 ms against DuckDB's 833–870 ms (~8.5–8.7x), libxlsxwriter's 1,184 ms (~12x),
xlsxio's 2,454 ms (~25x) and xlnt's 5,641 ms (~56x). The Rust suite measures 53.3 ms against
rust_xlsxwriter's 326.9 ms (~6.1x) over 7 columns. Caveats for both are in the binding READMEs.

DuckDB, xlsxio and xlnt do not read `.xlsb`, so those comparisons are XLSX-only. calamine is a fast,
well-optimized reader in its own right — the gap there is real but not the order of magnitude seen
against the C/C++ competitors.

zsv is the one competitor ahead. The CSV row is the typed workload, where both sides build the same
14-column rows with owned strings. Just counting cell bytes, zsv is ~4x faster, and that gap is
ExcelReader's C ABI, which serializes every row into a buffer, not its CSV parser. That suite is
built with GCC (zsv does not compile under MSVC). Its caveats are in the C++ README.

