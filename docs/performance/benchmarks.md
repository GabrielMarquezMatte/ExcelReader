# Benchmarks

Throwaway numbers go stale; these are regenerated from the benchmark suite in
`tests/ExcelReader.Benchmarks`. Live results: https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.16.0-preview.1 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.13 (built with SDK 11.0.100-rc.2, targeting `net10.0`). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [GitHub Pages benchmark dashboard](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/) — `tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/` is `.gitignore`-excluded and never actually reaches GitHub, so a link to it 404s for every reader.

Every BenchmarkDotNet table is a `--job Short` run (1 launch, 3 warmup and 3 measured iterations) on an otherwise idle machine. Methods whose 99.9% confidence interval came out wider than 10% of the mean on any row were run again with `--job Medium` (2 launches, 15 iterations) and carry those figures instead: the parallel CSV tables (except Sep's sequential rows), the large-sheet tables, the parallel sheet read, the serial, `PrefetchWrite`, 2-batch XLSB and 16-batch XLSX parallel writes, the XLS and XLSX rows of the XLS writer comparison, the XLS typed record write, the XLSB `PrefetchWrite` write, the XLSX real-data matched-work read, the string-heavy XLSX matched-work read, the interned XLSB matched-work read and the Sylvan XLSX string-heavy read, the stored whole read, the inferred-schema Arrow conversion, the decoded native read-all and the batched native typed reads. The read and typed-parse rows of the XLSX and XLSB tables, the generated-workbook matched-work read, the `ref struct` parsing table and the ADO.NET bridge table come from a later `--job Short --launchCount 5` run instead: Tiered PGO settles each process into a different code shape, which moved these rows by up to ±10% between single-launch runs. The cold-start rows set their own launch count, and the file-size and working-set figures under the parallel record writes, the managed deflate and CRC-32 figures, the Python/C++/Rust figures and the native-library comparison come from other harnesses that were not re-run.

**Benchmark methodology.** In the Excel-format tables below, the **"Cell-by-cell read"** rows (including the per-format rows in "Real data reads" / "String-heavy reads") read ExcelReader's `cell.Value` — a zero-copy `ReadOnlySpan<byte>`, no decode or allocation — against each competitor's own idiomatic read call. For Sylvan.Data.Excel, that's the ADO.NET-style `GetString(i)`, which is forced to materialize a UTF-16 `string`; its API has no zero-copy accessor, so it cannot avoid that cost. These rows are therefore not matched work: part of the reported gap is "we parse faster" and part is "we skipped an allocation you were never offered a way to skip." Each affected benchmark class also has a `*_Materialized` sibling (calling `cell.GetString()`, the same UTF-16 materialization Sylvan pays) so the matched-work number is measurable — see `tests/ExcelReader.Benchmarks/Shared/BenchmarkAccumulators.cs`. **Those results are published under [Matched-work reads](#matched-work-reads-_materialized), and reading them is the honest way to judge the comparison:** treat the Excel cell-by-cell ratios in the tables below as an upper bound, not a like-for-like number. The CSV cell-by-cell rows are matched work: Sep and Sylvan.Data.Csv both expose span accessors (`row[i].Span`, `GetFieldSpan(i)`), and the benchmarks use them. The **"Typed row parsing"** / **"Typed record writing"** rows are unaffected — both sides already materialize real objects/strings there, so those comparisons are matched work as published.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | Sylvan | ExcelDataReader | OfficeIMO | SpreadCheetah |
|---|---:|---:|---:|---:|---:|
| Cell-by-cell read | 7.098 ms, 3.44 KB | 37.056 ms, 1.89 MB | 87.075 ms, 32.88 MB | 13.892 ms, 3.51 MB | - |
| Cell-by-cell read async | 7.428 ms, 3.51 KB | - | - | - | - |
| Typed row parsing | 10.913 ms, 3.87 MB | 60.050 ms, 10.48 MB | - | 11.790 ms, 2.36 MB | - |
| Typed row parsing async | 10.223 ms, 3.87 MB | 58.884 ms, 10.48 MB | - | - | - |
| Typed row parsing, shared strings | 7.737 ms, 2.29 MB | - | - | - | - |
| Typed row parsing, `InternStrings` | 10.633 ms, 2.30 MB | - | - | - | - |
| Workbook writing | 8.191 ms, 4.02 MB | - | - | 18.162 ms, 4.04 MB | 14.828 ms, 15.84 MB |
| Workbook writing, shared strings | 7.981 ms, 4.03 MB | - | - | - | - |

ExcelReader is ~5.2x faster than Sylvan for raw XLSX reads, allocating ~564x less, ~12.3x faster than ExcelDataReader, and ~2.0x faster than OfficeIMO. For typed parsing it is ~5.5x faster than Sylvan (~5.8x async against async) and ~1.1x faster than OfficeIMO, which allocates less (2.36 MB against 3.87 MB) because it always interns strings. `ExcelReaderOptions.InternStrings` is off by default; turned on, ExcelReader allocates 2.30 MB at 10.633 ms, ~10% faster than OfficeIMO in the same run. For XLSX writing, ExcelReader is ~1.8x faster than SpreadCheetah while allocating ~3.9x less memory, and ~2.2x faster than OfficeIMO. The OfficeIMO rows are OfficeIMO.Excel 3.5.1, and the ExcelDataReader rows ExcelDataReader 3.9.0.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~29% faster than the inline-string sheet above (7.737 ms vs. 10.913 ms) and allocates ~41% less (2.29 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | ExcelDataReader | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 4.942 ms, 3.67 KB | 28.106 ms, 18.06 MB | 9.841 ms, 5.12 MB |
| Cell-by-cell read async | 5.071 ms, 3.74 KB | - | - |
| Typed row parsing | 7.375 ms, 3.87 MB | - | - |
| Typed row parsing async | 7.276 ms, 3.87 MB | - | - |
| Workbook writing | 5.462 ms, 4.02 MB | - | - |
| Workbook writing, shared strings | 5.521 ms, 4.03 MB | - | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than their XLSX counterparts, and typed parsing and writing ~1.5x. Against OfficeIMO's XLSB reader, ExcelReader is ~2.0x faster while allocating ~1,430x less, and against ExcelDataReader's ~5.7x faster. The XLSB writer is also ~2.7x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | ExcelDataReader | OfficeIMO |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 3.529 ms, 2.84 KB | 5.199 ms, 1,717.73 KB | 38.718 ms, 43,522.77 KB | 7.629 ms, 10,499.44 KB |
| Cell-by-cell read async | 3.279 ms, 2.91 KB | - | - | - |
| Workbook writing | 4.736 ms, 16.31 MB | - | - | - |

ExcelReader is ~1.5x faster than Sylvan for generated XLS reads while allocating ~605x less memory, ~2.2x faster than OfficeIMO, and ~11x faster than ExcelDataReader. The XLS writer is ~1.7x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 4.736 ms vs 7.941 ms in the same run). Most of its 16.31 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.31 MB, close to the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 2.986 ms, 888 B | 8.005 ms, 3.93 KB | 4.034 ms, 37.78 KB |
| Cell-by-cell read async | 2.907 ms, 960 B | - | - |
| Typed row parsing | 4.618 ms, 3.86 MB | 8.438 ms, 3.87 MB | 11.703 ms, 10.95 MB |
| Typed row parsing async | 4.876 ms, 3.86 MB | - | - |
| Row writing | 4.634 ms, 4.00 MB | 6.861 ms, 4.01 MB | 6.925 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~2.7x faster than Sep while allocating ~4.5x less, and ~1.4x faster than Sylvan.Data.Csv while allocating ~44x less. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.8x faster than Sep and ~2.5x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`SheetParallel.ParseAsync<T>` and `SheetParallel.AggregateAsync` over a `CsvSheet`, on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Conversion-heavy, struct + `InternStrings` | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|---:|
| 1 | 612.4 ms, 671.42 MB | 561.2 ms, 1.64 MB | 512.1 ms, 245.76 MB | 453.5 ms, 1.21 MB | 410.6 ms, 1.20 MB |
| 2 | 472.5 ms, 680.04 MB | 447.5 ms, 10.86 MB | 399.0 ms, 254.48 MB | 263.3 ms, 0.91 MB | 245.3 ms, 0.92 MB |
| 4 | 268.9 ms, 680.10 MB | 254.9 ms, 11.33 MB | 224.3 ms, 254.56 MB | 136.1 ms, 1.07 MB | 130.5 ms, 1.09 MB |
| 8 | 219.9 ms, 680.10 MB | 197.2 ms, 12.21 MB | 193.7 ms, 254.52 MB | 92.2 ms, 1.27 MB | 93.3 ms, 1.26 MB |
| 16 | 419.6 ms, 680.12 MB | 171.3 ms, 14.17 MB | 255.5 ms, 255.02 MB | 68.1 ms, 1.38 MB | 70.5 ms, 1.34 MB |

The typed path scales ~2.3x by dop 4 on both corpora (612.4 to 268.9 ms conversion-heavy, 512.1 to 224.3 ms narrow), then flattens out: narrow ints reach 193.7 ms at dop 8 and slow to 255.5 ms at dop 16, and the conversion-heavy corpus gets *slower* past dop 8 (419.6 ms at dop 16 against 219.9 ms at dop 8, standard deviation 105.4 ms). Its ~680 MB is one materialized row object and two strings per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The struct + `InternStrings` column parses the same rows into a `struct` with `ExcelReaderOptions.InternStrings` on, so there is no row object and each repeated region or country name is one cached string. It never slows down: 197.2 ms at dop 8, 171.3 ms at dop 16, 2.4x faster than the class at dop 16 while allocating ~48x less (14.17 MB, under one Gen0 collection per run). What remains is per-worker caches and the ordered merge's buffers. The interning win depends on the data: the corpus has 4 regions and 6 countries, so every string is a cache hit. Text columns with mostly distinct values (names, ids, free text) allocate a string per cell either way, and on those the struct alone removes only the row object.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~4.9x (conversion-heavy) and ~4.4x (narrow) at dop 8, ~6.7x and ~5.8x at dop 16 — because they are not fighting the allocator: at dop 4 the conversion-heavy aggregate is ~2.0x faster than typed while allocating ~630x less (1.07 MB against 680.10 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

The typed path's slowdown past dop 8 is the garbage collector. Every degree allocates the same ~680 MB and runs 42 gen0 collections, but at dop 16 each of them also collects gen1 (41 gen1 collections against 26 at dop 8), and the pauses add up to 100-200 ms against 39 ms. In a console harness on the 4.3 million row corpus, with the default workstation GC, dop 16 took 365 ms; with `DOTNET_gcServer=1` it took 182 ms, faster than dop 8. Not allocating per row, as the struct + `InternStrings` and aggregate columns do, avoids it without changing the GC mode.

Type conversion is the floor under all of them. A single-threaded pass that allocates nothing still costs 453.5 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus, row type | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints, class | 474.9 ms, 244.15 MB | 269.1 ms, 255.02 MB | 507.5 ms, 244.15 MB | 433.0 ms, 245.46 MB |
| Narrow ints, struct | 392.0 ms, 13.93 KB | 162.0 ms, 12.15 MB | 476.2 ms, 9.34 KB | 251.7 ms, 2.81 MB |
| Conversion-heavy, class | 407.6 ms, 467.31 MB | 353.3 ms, 474.68 MB | 835.3 ms, 467.31 MB | 596.5 ms, 468.41 MB |
| Conversion-heavy, struct + pooled strings | 377.3 ms, 14.55 KB | 119.6 ms, 11.28 MB | 857.6 ms, 10.77 KB | 392.7 ms, 9.33 MB |

The struct rows give both libraries the same help: the row is a `struct`, and repeated text is pooled, through `ExcelReaderOptions.InternStrings` on ExcelReader's side and `SepToString.PoolPerCol()` (sequential) or `PoolPerColThreadSafe()` (parallel) on Sep's. Every row except Sep's sequential ones is from a Medium job, since the Short run spread too much; Sep's conversion-heavy parallel struct row still has a 35.1 ms standard deviation, and ExcelReader's conversion-heavy parallel class row 29.4 ms.

Sequentially, ExcelReader's typed parser is ~1.1x faster than Sep on narrow ints (474.9 ms vs. 507.5 ms), and ~2.0x faster on the conversion-heavy corpus, where decimal parsing dominates. In parallel ExcelReader is ~1.6x (narrow) and ~1.7x (conversion-heavy) faster than Sep's, while yielding rows in file order. With struct rows and pooled strings both libraries speed up and stop collecting, and ExcelReader stays ahead: ~1.2x and ~2.3x sequentially, ~1.6x and ~3.3x in parallel. The class rows gain less from parallelism (353.3 ms against 407.6 ms sequential on the conversion-heavy corpus) because they are bound by the allocator, like the typed dop 16 row above: 28 Gen1 collections per run; the struct rows run 3.2x faster in parallel on the same corpus. As in the table above, pooling pays off here because the conversion-heavy corpus repeats 4 regions and 6 countries.

### Parallel sheet reading

`SheetParallel.AggregateAsync<TAccumulator, TModel>` over the first sheet of the real-data workbook (65,535 rows, 14 columns; XLSX 6.7 MB compressed, 34.4 MB sheet XML; XLSB 3.7 MB, 15.9 MB sheet records), folding every row of a typed model into a per-partition sum (`SheetParallelBenchmark`). One thread inflates the sheet and cuts it into row-aligned chunks; `Dop - 1` workers parse them.

| Dop | XLSX | XLSB |
|---:|---:|---:|
| 1 | 46.07 ms, 8.02 MB | 29.86 ms, 8.02 MB |
| 2 | 29.36 ms, 8.13 MB | 22.49 ms, 8.07 MB |
| 4 | 16.69 ms, 8.12 MB | 13.28 ms, 8.07 MB |
| 8 | 16.79 ms, 8.17 MB | 13.18 ms, 8.09 MB |

At dop 4 XLSX reads ~2.8x faster than sequentially and XLSB ~2.2x. Dop 2 is a single parse worker running alongside the inflating thread. Past dop 4 nothing improves: the one inflating thread is the floor (~14 ms for XLSX, ~9 ms for XLSB, see [Inflate](#inflate), plus the XLSB record walk that finds row boundaries). Allocation barely moves with dop (8.02 to 8.17 MB): it is the typed model's strings.

### Large sheets

`LargeSheetParallelBenchmark` reads sheets generated in memory at the largest sizes the formats allow. XLSX and XLSB hold 1,000,000 rows by four columns (text, integer, date, number), the limit being 1,048,576 rows per sheet; CSV holds 4,000,000 conversion-heavy rows (two strings, a date, two decimals and an int), and has no such limit. "Aggregate" folds every row into a per-partition sum; "typed parse" is the ordered `ParseAsync` into a class (`ref struct` for the CSV aggregate).

| XLSX, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 223.16 ms, 77.25 MB | 249.45 ms, 77.25 MB |
| 2 | 163.11 ms, 77.62 MB | 190.50 ms, 106.81 MB |
| 4 | 59.25 ms, 77.57 MB | 73.08 ms, 106.75 MB |
| 8 | 45.08 ms, 77.64 MB | 55.79 ms, 106.78 MB |
| 16 | 44.48 ms, 77.95 MB | 55.15 ms, 107.11 MB |

| XLSB, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 157.36 ms, 77.25 MB | 170.55 ms, 77.25 MB |
| 2 | 108.82 ms, 77.53 MB | 118.56 ms, 99.69 MB |
| 4 | 47.71 ms, 77.48 MB | 56.25 ms, 99.62 MB |
| 8 | 49.58 ms, 77.61 MB | 54.14 ms, 99.76 MB |
| 16 | 49.46 ms, 77.87 MB | 53.82 ms, 100.03 MB |

| CSV, Dop | Aggregate | Typed parse | Typed parse, struct + `InternStrings` |
|---:|---:|---:|---:|
| 1 | 418.65 ms, 1.99 KB | 509.94 ms, 623.07 MB | 506.93 ms, 14.05 KB |
| 2 | 204.59 ms, 11.58 KB | 325.09 ms, 626.87 MB | 323.26 ms, 4.46 MB |
| 4 | 115.50 ms, 10.94 KB | 200.87 ms, 626.68 MB | 178.60 ms, 4.70 MB |
| 8 | 84.80 ms, 18.88 KB | 165.18 ms, 626.40 MB | 139.31 ms, 5.33 MB |
| 16 | 61.33 ms, 35.14 KB | 216.44 ms, 626.55 MB | 122.48 ms, 7.40 MB |

The XLSX and XLSB aggregates and typed parses each materialize a `Record` and its name string per row, which is where their ~77-107 MB comes from; the CSV aggregate binds a `ref struct` and allocates kilobytes, while the CSV typed parse allocates ~626 MB at every degree, one row object and two strings per record. The struct + `InternStrings` parse drops both (the corpus repeats 4 regions and 6 countries, see [Parallel CSV](#parallel-csv)) and keeps a few MB of per-worker caches and merge buffers.

XLSX and XLSB reach ~3.0-3.8x at dop 4. The XLSB aggregate stops there (3.3x), as in the real-data table above, while the XLSX aggregate improves to 5.0x at dop 8 and then stops. The CSV aggregate keeps scaling because nothing but the parser runs on the workers: 4.9x at dop 8 and 6.8x at dop 16. The typed CSV parse reaches 3.1x at dop 8 and then loses 31% at dop 16 to the collector, while the struct + `InternStrings` parse keeps going to 4.1x at dop 16, 1.8x faster than the class. The aggregate is still the fastest, ~2.0x ahead of the struct parse at dop 16, because it never builds the row at all.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 37.237 ms, 7.20 KB | 23.663 ms, 23.34 KB | 195.758 ms, 644.27 KB |
| XLSM | 37.999 ms, 7.20 KB | 23.061 ms, 23.35 KB | 199.923 ms, 644.35 KB |
| XLSB | 24.746 ms, 6.72 KB | 16.642 ms, 15.09 KB | 29.421 ms, 338.55 KB |
| XLS | 11.196 ms, 12.02 KB | n/a | 18.221 ms, 185.90 KB |
| CSV | 4.949 ms, 888 B | n/a | 4.502 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~5.3x faster than Sylvan for XLSX and XLSM, ~1.2x faster for XLSB, and ~1.6x faster for XLS — allocating ~90x less for XLSX and XLSM, ~50x less for XLSB and ~15x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~10% faster, while ExcelReader allocates ~46x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 37.534 ms | 0.007 ms | 6.71 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 23.609 ms | 0.043 ms | 22.83 KB |
| Xlsm_ExcelReader_Memory | 37.225 ms | 0.045 ms | 6.71 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 23.467 ms | 0.052 ms | 22.78 KB |
| Xlsb_ExcelReader_Memory | 24.418 ms | 0.012 ms | 8.31 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 15.382 ms | 0.047 ms | 16.58 KB |
| Xls_ExcelReader_Memory | 11.031 ms | 0.051 ms | 11.87 KB |
| Csv_ExcelReader_Memory | 4.744 ms | 0.004 ms | 616 B |

`Csv_ExcelReader_Memory` is both faster (4.744 ms vs. 4.949 ms for `Csv_ExcelReader`) and allocates less (616 B vs. 888 B) than its stream twin.

`Xls_ExcelReader_Memory` is ~1% faster than its stream twin (11.031 ms vs. 11.196 ms for `Xls_ExcelReader`), with near-identical allocation (11.87 KB vs. 12.02 KB). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since mostly caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 30.00 ms, 2.18 MB | 20.01 ms, 2.21 MB | 176.47 ms, 17.41 MB |
| XLSB | 33.56 ms, 2.18 MB | 19.52 ms, 2.20 MB | 58.80 ms, 17.38 MB |

Both formats handle this well: ~5.9x and ~1.8x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 30.01 ms | 2.18 MB |
| Whole read, prefetch | 19.56 ms | 2.21 MB |
| Shared-string table only (open + first row) | 9.22 ms | 747 KB |
| Whole read, stored | 17.65 ms | 2.18 MB |
| Shared-string table only, stored | 3.34 ms | 746 KB |

About 41% of the default read is inflate: without it the whole read drops from 30.01 ms to 17.65 ms, and the shared-string table from 9.22 ms to 3.34 ms. What remains, ~14 ms, is the worksheet scan.

### Inflate

`InflateBenchmark` decompresses one ZIP entry per method with `System.IO.Compression.DeflateStream` and with the library's managed `InflateStream`, which every ZIP-based reader now uses:

| Entry | `DeflateStream` | `InflateStream` | Ratio |
|---|---:|---:|---:|
| XLSX `sheet1.xml` (6.7 MB → 34.4 MB) | 22.244 ms, 280 B | 13.463 ms, 200 B | 0.61 |
| XLSB `sheet1.bin` (3.7 MB → 15.9 MB) | 12.600 ms, 280 B | 9.040 ms, 200 B | 0.72 |
| String-heavy `sharedStrings.xml` | 10.360 ms, 280 B | 5.763 ms, 200 B | 0.56 |

Without prefetch a read is inflate plus parse, so the faster inflate converts directly; with prefetch the two overlap and the read is bounded by the slower stage, which is why the prefetch columns above moved little.

### Write path: managed deflate and CRC-32

XLSX and XLSB packages are written by the library's own ZIP writer. `CompressionLevel.Fastest`, the default, compresses with a managed fixed-Huffman deflate encoder and checksums with a managed CRC-32 that folds with PCLMULQDQ. Each piece alone on the same Ryzen 7 5700X: the deflate rows against zlib-ng at `Fastest`, which `ZipArchive` used, and the CRC row against `System.IO.Hashing.Crc32`, the fastest managed alternative (`ZipArchive` itself calls zlib's native CRC, which was not measured):

| | zlib-ng `Fastest` / System.IO.Hashing | Managed | Ratio |
|---|---:|---:|---:|
| Deflate, writer XLSX sheet (6.0 MB) | 4.75 ms | 2.98 ms | 1.59x, same size |
| Deflate, writer XLSB sheet (4.6 MB) | 4.60 ms | 2.52 ms | 1.82x, same size |
| Deflate, 65K real XLSX sheet (34.4 MB) | 60.3 ms | 38.4 ms | 1.57x, same size |
| CRC-32, 64 KB chunks in cache | 16.4 GB/s | 35.8 GB/s | 2.19x |

Whole writes of 50,000 rows by four columns (text, integer, date, number) into a `MemoryStream`, before and after the switch. Each figure is the median of three interleaved runs of 41 writes:

| Write | `ZipArchive` | Managed ZIP writer | Change | File size |
|---|---:|---:|---:|---:|
| XLSX, `Fastest` | 10.43 ms | 9.11 ms | -12.7% | 773,487 B → 774,344 B |
| XLSX, `Fastest` + `PrefetchWrite` | 6.15 ms | 5.84 ms | -5.0% | same as above |
| XLSX, `Optimal` | 30.90 ms | 31.38 ms | within noise | 483,613 B → 483,607 B |
| XLSB, `Fastest` | 7.66 ms | 5.42 ms | -29.2% | 861,060 B → 861,568 B |
| XLSB, `Fastest` + `PrefetchWrite` | 5.37 ms | 3.08 ms | -42.6% | same as above |
| XLSB, `Optimal` | 38.17 ms | 37.98 ms | within noise | 466,917 B → 466,917 B |

XLSB gains more because deflate was a larger share of its write; with `PrefetchWrite` an XLSX write is bounded by building the XML, so the faster deflate moves it little.

The encoder's limits are deliberate:

- It emits fixed-Huffman blocks only, so `Optimal` and `SmallestSize` still go through zlib, inside the same container.
- A block covers at most 65,535 input bytes, so any block that would grow is written stored instead. Incompressible input grows by at most 5 bytes per block.

### Parallel record writes

`WriteRecordsParallelAsync` formats and compresses batches of 8,192 records on pool threads and appends them to the sheet in order. `ParallelRecordWriteBenchmark` writes 1,000,000 rows by four columns (text, integer, date, number) into a pre-sized `MemoryStream` at `Fastest`, on the same Ryzen 7 5700X, with a number of batches in flight from 1 to 16:

| Write | XLSX | vs serial | XLSB | vs serial |
|---|---:|---:|---:|---:|
| Serial | 165.08 ms, 16.09 KB | 1.0x | 113.61 ms, 13.42 KB | 1.0x |
| Serial, `PrefetchWrite` | 102.18 ms, 18.35 KB | 1.6x | 62.39 ms, 15.64 KB | 1.8x |
| Parallel, 1 batch in flight | 164.23 ms, 16.15 KB | 1.0x | 114.03 ms, 13.48 KB | 1.0x |
| Parallel, 2 | 84.48 ms, 82.67 KB | 2.0x | 58.53 ms, 81.12 KB | 1.9x |
| Parallel, 4 | 44.27 ms, 79.71 KB | 3.7x | 30.04 ms, 78.45 KB | 3.8x |
| Parallel, 8 | 27.54 ms, 74.09 KB | 6.0x | 19.48 ms, 70.47 KB | 5.8x |
| Parallel, 16 | 21.64 ms, 79.04 KB | 7.6x | 15.37 ms, 67.00 KB | 7.4x |

It scales because a write is bound by building rows, not by deflate: `PrefetchWrite` already takes deflate off the calling thread and stops at 1.6x to 1.8x, so compressing chunks in parallel would add nothing on top of it, while formatting the rows in parallel does. Past dop 8 the gain shrinks (6.0x to 7.6x for XLSX, 5.8x to 7.4x for XLSB), because the machine has 8 physical cores. None of the rows triggers a garbage collection, and every write allocates under 85 KB for 1,000,000 rows.

Each batch is compressed with no reference to the one before it, which costs 0.18% of file size for XLSX (15,481,372 B to 15,508,760 B) and 0.04% for XLSB (16,863,131 B to 16,869,306 B). The process's peak working set rose by 12 MB over the serial write with 4 batches in flight and by 23 MB with 8, about 3 MB per batch, and does not depend on the number of rows. The file sizes and working sets come from a console harness, because BenchmarkDotNet does not report them.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 39.13 ms | 0.08 ms | 16.66 KB |
| XLSM | 39.48 ms | 0.04 ms | 16.66 KB |
| XLSB | 27.93 ms | 0.01 ms | 16.19 KB |
| XLS | 12.73 ms | 0.02 ms | 21.48 KB |
| CSV | 18.95 ms | 0.09 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 8.42 ms | 0.24 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 50.77 ms | 0.30 ms | 1,100.0 | 1,000.0 | 300.0 | 15.25 MB |
| XLSB | 54.01 ms | 0.24 ms | 1,100.0 | 1,000.0 | 300.0 | 15.25 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 888 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 16.19 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.25 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan and is still ~3.5x faster (50.77 ms vs. 176.47 ms). The same allocation holds for XLSB (15.25 MB vs. Sylvan's 17.38 MB, also ~12% less), where matched-work time is ~8% ahead of Sylvan (54.01 ms vs. 58.80 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 8.381 ms | 4.02 MB |
| XLSB | 5.838 ms | 4.02 MB |
| XLS | 6.182 ms | 4.31 MB |
| CSV | 4.814 ms | 4.00 MB |

Relative ordering: CSV at the front, XLSB and XLS close behind it, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 5.838 ms vs. 5.462 ms, XLSX 8.381 ms vs. 8.191 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 14.275 ms | 15.93 KB |
| Native string columns → CSV | 5.525 ms | 416 B |
| Arrow `StringArray` → XLSX | 20.674 ms | 15.94 KB |
| Styled rows → XLSX | 5.817 ms | 16.10 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~16 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 10.91 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 10.68 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 9.02 ms | 3.98 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.9% less than the `class` baseline — and is ~16% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`SheetDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 6.542 ms | 1.00 | 3.48 KB |
| `GetBytes` (text column only) | 6.983 ms | 1.07 | 4.18 KB |
| `GetValue` (every column) | 9.568 ms | 1.46 | 5.01 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 9.357 ms | 1.43 | 1.58 MB |
| `DataTable.Load` | 62.94 ms | 9.62 | 22.84 MB |

The bridge costs ~1.4x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it allocates nothing per row and costs ~7% over the raw read. `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 9.568 ms to read through `GetValue` alone.

### Arrow conversion

`ArrowConversionBenchmark`: `ToArrowRecordBatch` from [ExcelReader.Arrow](../guide/reading.md#apache-arrow-conversion-opt-in), 100,000 rows into one `RecordBatch`. The typed rows use the generated records (a text name, an int64 id, a timestamp, a double) as CSV and as XLSB. The all-string row is a separate header-less CSV with 8 text columns.

| Scenario | Mean | Allocated |
|---|---:|---:|
| CSV, typed, explicit schema | 12.26 ms | 8.13 MB |
| CSV, typed, inferred schema | 12.48 ms | 14.89 MB |
| XLSB, typed, explicit schema | 16.88 ms | 8.14 MB |
| CSV, 8 text columns, explicit schema | 22.31 ms | 16.27 MB |

On the same file, letting `Excel.InferSchema` guess the schema costs ~2% more time (12.48 ms vs. 12.26 ms) and ~1.8x the allocation of passing an explicit `ExcelColumnSchema[]`. The allocation grows with the sheet, because the batch holds every row at once and there is no streaming variant.

### Encrypted workbooks

`EncryptedWorkbookBenchmark`: agile encryption (AES-256, SHA-512, 100,000 spin iterations — see [Encrypted workbooks](../guide/encryption.md)). The small rows use a 15 KB fixture; the large rows use a generated 300,000-row workbook, encrypted with `Excel.EncryptPackage`.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Small, plain | 0.210 ms | 8.72 KB |
| Small, encrypted, stream | 26.966 ms | 37.93 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 26.745 ms | 40.05 KB |
| Small, encrypted, in memory | 26.166 ms | 52.99 KB |
| Small, encrypted, open only | 27.432 ms | 32.35 KB |
| 300,000 rows, plain | 36.460 ms | 3.48 KB |
| 300,000 rows, encrypted | 62.931 ms | 29.41 KB |
| 300,000 rows, encrypted, open only | 24.645 ms | 28.72 KB |

Nearly all of the cost is a fixed ~25 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 27.43 ms for 15 KB and 24.64 ms for 300,000 rows). After that, decrypting while reading costs ~5% over a plain read: 62.93 ms − 24.64 ms = 38.29 ms, against 36.46 ms unencrypted. On the small fixture, integrity verification and the in-memory path stay within ~3% of the plain encrypted read, inside the spread of the key derivation itself. The fixture is too small to show the extra pass that verification makes over a large stream.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 46.05 ms | 20.72 KB |
| First typed record write | 16.66 ms | 80.81 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 46.05 ms | 20.72 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 41.57 ms | 23.66 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 49.10 ms | 25.78 KB |

`Build` is ~10% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slower than reflection (~7%, with a 0.5 ms standard deviation), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

### Native C ABI reads

`NativeRowReadBenchmark` and `NativeTypedParseBenchmark` call the C ABI entry points from .NET, reading the real-data XLSB workbook (65,535 rows × 14 columns) from memory. This isolates what the ABI itself costs, without any binding's marshalling on top.

| Entry point | Shape handed back | Mean | Managed allocated |
|---|---|---:|---:|
| `xl_rows_next` | one row per call, serialized into a caller buffer | 38.42 ms | 3.71 MB |
| `xl_rows_read_all_blob` | whole sheet in one buffer | 41.82 ms | 21.06 MB |
| `xl_rows_read_all_decoded` | whole sheet as native row structs | 52.05 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 34.84 ms | 3.71 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 36.02 ms | 3.72 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_rows_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Row by row, `xl_rows_next` costs ~1.6x the managed in-memory read of the same file (38.42 ms vs. 24.418 ms for `Xlsb_ExcelReader_Memory` under [In-memory real-data reads](#in-memory-real-data-reads)). Every cell crosses the ABI as text, so each binary XLSB number is formatted to UTF-8 on the way out. `xl_rows_read_all_blob` adds ~9% over it, because the whole sheet is buffered before it is copied out. `xl_rows_read_all_decoded` adds ~35%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` are slightly faster than a row-by-row read (~6–9%) and return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 8.76 ms | 9.87 KB |
| Batches of 10,000 rows | 8.83 ms | 16.82 KB |
| Batches of 1,000 rows | 8.14 ms | 81.15 KB |

Batching costs nothing in time: batches of 10,000 rows land within 1% of one whole-sheet table, and batches of 1,000 rows are ~7% faster. A caller can bound its native memory to one batch at a time without paying for it.

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
| vs. [calamine](https://github.com/tafia/calamine) (Rust) | XLSX | 72.3 ms | 271.2 ms | ~3.8x faster |
| vs. calamine (Rust) | XLSB | 46.3 ms | 82.9 ms | ~1.8x faster |
| vs. [DuckDB](https://github.com/duckdb/duckdb) `read_xlsx` (C++) | XLSX | 61.3 ms | 413.8 ms | ~6.7x faster |
| vs. [xlsxio](https://github.com/brechtsanders/xlsxio) (C) | XLSX | 61.3 ms | 497.4 ms | ~8.1x faster |
| vs. [xlnt](https://github.com/tfussell/xlnt) (C++) | XLSX | 61.3 ms | 2,335.0 ms | ~38x faster |
| vs. [zsv](https://github.com/liquidaty/zsv) `fast` engine (C) | CSV | 27.4 ms | 20.0 ms | ~1.4x **slower** |

Writing the same 14 columns × 65,535 rows from row-shaped data, the C++ suite measures
`write_sheet` at 74.8–76.0 ms against DuckDB's 831 ms (~11x), libxlsxwriter's 1,207 ms (~16x),
xlsxio's 2,405 ms (~32x) and xlnt's 5,525 ms (~74x). The Rust suite measures 40.7 ms against
rust_xlsxwriter's 322.0 ms (~7.9x) over 7 columns. Caveats for both are in the binding READMEs.

DuckDB, xlsxio and xlnt do not read `.xlsb`, so those comparisons are XLSX-only. calamine is a fast,
well-optimized reader in its own right — the gap there is real but not the order of magnitude seen
against the C/C++ competitors.

zsv is the one competitor ahead. The CSV row is the typed workload, where both sides build the same
14-column rows with owned strings. Just counting cell bytes, zsv is ~3.3x faster, and that gap is
ExcelReader's C ABI, which serializes every row into a buffer, not its CSV parser. That suite is
built with GCC (zsv does not compile under MSVC). Its caveats are in the C++ README.
