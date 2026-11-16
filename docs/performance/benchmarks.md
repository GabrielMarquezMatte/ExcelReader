# Benchmarks

Throwaway numbers go stale; these are regenerated from the benchmark suite in
`tests/ExcelReader.Benchmarks`. Live results: https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.16.0-preview.2 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.14 (built with SDK 10.0.403, targeting `net10.0`). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [GitHub Pages benchmark dashboard](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/) — `tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/` is `.gitignore`-excluded and never actually reaches GitHub, so a link to it 404s for every reader.

Every BenchmarkDotNet table is one `--job Medium` run of the whole suite (2 launches, 10 warmup and 15 measured iterations) on an otherwise idle machine. Tiered PGO settles each process into a different code shape, so a row can still move by up to ±10% between runs; compare rows within a table, not across releases. The cold-start rows set their own launch count, and the file-size and working-set figures under the parallel record writes, the managed deflate and CRC-32 figures, the Python/C++/Rust figures and the native-library comparison come from other harnesses that were not re-run.

**Benchmark methodology.** In the Excel-format tables below, the **"Cell-by-cell read"** rows (including the per-format rows in "Real data reads" / "String-heavy reads") read ExcelReader's `cell.Value` — a zero-copy `ReadOnlySpan<byte>`, no decode or allocation — against each competitor's own idiomatic read call. For Sylvan.Data.Excel, that's the ADO.NET-style `GetString(i)`, which is forced to materialize a UTF-16 `string`; its API has no zero-copy accessor, so it cannot avoid that cost. These rows are therefore not matched work: part of the reported gap is "we parse faster" and part is "we skipped an allocation you were never offered a way to skip." Each affected benchmark class also has a `*_Materialized` sibling (calling `cell.GetString()`, the same UTF-16 materialization Sylvan pays) so the matched-work number is measurable — see `tests/ExcelReader.Benchmarks/Shared/BenchmarkAccumulators.cs`. **Those results are published under [Matched-work reads](#matched-work-reads-_materialized), and reading them is the honest way to judge the comparison:** treat the Excel cell-by-cell ratios in the tables below as an upper bound, not a like-for-like number. The CSV cell-by-cell rows are matched work: Sep and Sylvan.Data.Csv both expose span accessors (`row[i].Span`, `GetFieldSpan(i)`), and the benchmarks use them. The **"Typed row parsing"** / **"Typed record writing"** rows are unaffected — both sides already materialize real objects/strings there, so those comparisons are matched work as published.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | Sylvan | ExcelDataReader | OfficeIMO | SpreadCheetah |
|---|---:|---:|---:|---:|---:|
| Cell-by-cell read | 6.363 ms, 3.48 KB | 37.532 ms, 1.89 MB | 90.746 ms, 32.88 MB | 12.554 ms, 3.51 MB | - |
| Cell-by-cell read async | 6.613 ms, 3.55 KB | - | - | - | - |
| Typed row parsing | 8.875 ms, 3.87 MB | 57.041 ms, 10.48 MB | - | 10.327 ms, 2.36 MB | - |
| Typed row parsing async | 10.120 ms, 3.87 MB | 60.316 ms, 10.48 MB | - | - | - |
| Typed row parsing, shared strings | 7.629 ms, 2.29 MB | - | - | - | - |
| Typed row parsing, `InternStrings` | 9.103 ms, 2.34 MB | - | - | - | - |
| Workbook writing | 8.213 ms, 4.02 MB | - | - | 18.105 ms, 4.04 MB | 15.310 ms, 15.84 MB |
| Workbook writing, shared strings | 7.965 ms, 4.03 MB | - | - | - | - |

ExcelReader is ~5.9x faster than Sylvan for raw XLSX reads, allocating ~558x less, ~14.3x faster than ExcelDataReader, and ~2.0x faster than OfficeIMO. For typed parsing it is ~6.4x faster than Sylvan (~6.0x async against async) and ~1.2x faster than OfficeIMO, which allocates less (2.36 MB against 3.87 MB) because it always interns strings. `ExcelReaderOptions.InternStrings` is off by default; turned on, ExcelReader allocates 2.34 MB at 9.103 ms, ~12% faster than OfficeIMO in the same run. For XLSX writing, ExcelReader is ~1.9x faster than SpreadCheetah while allocating ~3.9x less memory, and ~2.2x faster than OfficeIMO. The OfficeIMO rows are OfficeIMO.Excel 3.5.3, and the ExcelDataReader rows ExcelDataReader 3.9.0.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~14% faster than the inline-string sheet above (7.629 ms vs. 8.875 ms) and allocates ~41% less (2.29 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | ExcelDataReader | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 4.427 ms, 3.71 KB | 28.009 ms, 18.06 MB | 10.117 ms, 5.12 MB |
| Cell-by-cell read async | 4.432 ms, 3.78 KB | - | - |
| Typed row parsing | 7.070 ms, 3.87 MB | - | - |
| Typed row parsing async | 7.139 ms, 3.87 MB | - | - |
| Workbook writing | 5.388 ms, 4.02 MB | - | - |
| Workbook writing, shared strings | 5.377 ms, 4.03 MB | - | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than their XLSX counterparts, typed parsing ~1.3x and writing ~1.5x. Against OfficeIMO's XLSB reader, ExcelReader is ~2.3x faster while allocating ~1,410x less, and against ExcelDataReader's ~6.3x faster. The XLSB writer is also ~2.8x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | ExcelDataReader | OfficeIMO |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 2.956 ms, 3.14 KB | 5.369 ms, 1,717.73 KB | 38.608 ms, 43,522.77 KB | 7.941 ms, 10,499.44 KB |
| Cell-by-cell read async | 3.498 ms, 3.21 KB | - | - | - |
| Workbook writing | 4.557 ms, 16.32 MB | - | - | - |

ExcelReader is ~1.8x faster than Sylvan for generated XLS reads while allocating ~547x less memory, ~2.7x faster than OfficeIMO, and ~13x faster than ExcelDataReader. The XLS writer is ~1.8x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 4.557 ms vs 8.243 ms in the same run). Most of its 16.32 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.32 MB, close to the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 3.080 ms, 896 B | 7.814 ms, 3.93 KB | 4.109 ms, 37.78 KB |
| Cell-by-cell read async | 3.359 ms, 968 B | - | - |
| Typed row parsing | 4.474 ms, 3.86 MB | 8.588 ms, 3.87 MB | 12.068 ms, 10.95 MB |
| Typed row parsing async | 4.840 ms, 3.86 MB | - | - |
| Row writing | 4.660 ms, 4.00 MB | 6.979 ms, 4.01 MB | 7.149 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~2.5x faster than Sep while allocating ~4.5x less, and ~1.3x faster than Sylvan.Data.Csv while allocating ~43x less. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.9x faster than Sep and ~2.7x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`SheetParallel.ParseAsync<T>` and `SheetParallel.AggregateAsync` over a `CsvSheet`, on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Conversion-heavy, struct + `InternStrings` | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|---:|
| 1 | 574.1 ms, 671.41 MB | 547.7 ms, 1.68 MB | 464.7 ms, 245.75 MB | 453.4 ms, 1.21 MB | 432.2 ms, 1.21 MB |
| 2 | 452.2 ms, 679.93 MB | 452.6 ms, 10.86 MB | 405.1 ms, 254.36 MB | 253.6 ms, 0.85 MB | 247.7 ms, 0.86 MB |
| 4 | 257.8 ms, 680.00 MB | 248.7 ms, 11.43 MB | 230.5 ms, 254.43 MB | 137.2 ms, 1.02 MB | 130.1 ms, 1.02 MB |
| 8 | 208.1 ms, 679.99 MB | 189.7 ms, 12.28 MB | 192.1 ms, 254.41 MB | 90.7 ms, 1.20 MB | 90.1 ms, 1.19 MB |
| 16 | 486.8 ms, 680.03 MB | 167.5 ms, 14.31 MB | 268.3 ms, 254.91 MB | 67.2 ms, 1.32 MB | 70.1 ms, 1.28 MB |

The typed path scales ~2.2x by dop 4 on the conversion-heavy corpus and ~2.0x on the narrow one (574.1 to 257.8 ms, 464.7 to 230.5 ms), then flattens out: narrow ints reach 192.1 ms at dop 8 and slow to 268.3 ms at dop 16, and the conversion-heavy corpus gets *slower* past dop 8 (486.8 ms at dop 16 against 208.1 ms at dop 8, standard deviation 24.1 ms). Its ~680 MB is one materialized row object and two strings per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The struct + `InternStrings` column parses the same rows into a `struct` with `ExcelReaderOptions.InternStrings` on, so there is no row object and each repeated region or country name is one cached string. It never slows down: 189.7 ms at dop 8, 167.5 ms at dop 16, 2.9x faster than the class at dop 16 while allocating ~48x less (14.31 MB, under one Gen0 collection per run). What remains is per-worker caches and the ordered merge's buffers. The interning win depends on the data: the corpus has 4 regions and 6 countries, so every string is a cache hit. Text columns with mostly distinct values (names, ids, free text) allocate a string per cell either way, and on those the struct alone removes only the row object.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~5.0x (conversion-heavy) and ~4.8x (narrow) at dop 8, ~6.7x and ~6.2x at dop 16 — because they are not fighting the allocator: at dop 4 the conversion-heavy aggregate is ~1.9x faster than typed while allocating ~670x less (1.02 MB against 680.00 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

The typed path's slowdown past dop 8 is the garbage collector. Every degree allocates the same ~680 MB and runs 42 gen0 collections, but at dop 16 each of them also collects gen1 (41 gen1 collections against 24 at dop 8), and the pauses add up to 100-200 ms against 39 ms. In a console harness on the 4.3 million row corpus, with the default workstation GC, dop 16 took 365 ms; with `DOTNET_gcServer=1` it took 182 ms, faster than dop 8. Not allocating per row, as the struct + `InternStrings` and aggregate columns do, avoids it without changing the GC mode.

Type conversion is the floor under all of them. A single-threaded pass that allocates nothing still costs 453.4 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus, row type | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints, class | 451.0 ms, 244.14 MB | 271.5 ms, 255.06 MB | 502.0 ms, 244.15 MB | 446.5 ms, 247.41 MB |
| Narrow ints, struct | 444.6 ms, 38.24 KB | 153.0 ms, 11.90 MB | 480.0 ms, 9.34 KB | 257.3 ms, 4.22 MB |
| Conversion-heavy, class | 375.6 ms, 467.30 MB | 320.6 ms, 474.61 MB | 830.7 ms, 467.31 MB | 602.3 ms, 469.10 MB |
| Conversion-heavy, struct + pooled strings | 389.5 ms, 75.22 KB | 114.7 ms, 11.41 MB | 918.5 ms, 10.77 KB | 380.5 ms, 5.01 MB |

The struct rows give both libraries the same help: the row is a `struct`, and repeated text is pooled, through `ExcelReaderOptions.InternStrings` on ExcelReader's side and `SepToString.PoolPerCol()` (sequential) or `PoolPerColThreadSafe()` (parallel) on Sep's. The parallel rows spread the most: ExcelReader's narrow and conversion-heavy parallel class rows have a 33.3 ms and 28.0 ms standard deviation, and Sep's conversion-heavy parallel struct row 27.6 ms.

Sequentially, ExcelReader's typed parser is ~1.1x faster than Sep on narrow ints (451.0 ms vs. 502.0 ms), and ~2.2x faster on the conversion-heavy corpus, where decimal parsing dominates. In parallel ExcelReader is ~1.6x (narrow) and ~1.9x (conversion-heavy) faster than Sep's, while yielding rows in file order. With struct rows and pooled strings both libraries stop collecting, and ExcelReader stays ahead: ~1.1x and ~2.4x sequentially, ~1.7x and ~3.3x in parallel. The class rows gain less from parallelism (320.6 ms against 375.6 ms sequential on the conversion-heavy corpus) because they are bound by the allocator, like the typed dop 16 row above: ~28 Gen1 collections per run; the struct rows run 3.4x faster in parallel on the same corpus. As in the table above, pooling pays off here because the conversion-heavy corpus repeats 4 regions and 6 countries.

### Parallel sheet reading

`SheetParallel.AggregateAsync<TAccumulator, TModel>` over the first sheet of the real-data workbook (65,535 rows, 14 columns; XLSX 6.7 MB compressed, 34.4 MB sheet XML; XLSB 3.7 MB, 15.9 MB sheet records), folding every row of a typed model into a per-partition sum (`SheetParallelBenchmark`). One thread inflates the sheet and cuts it into row-aligned chunks; `Dop - 1` workers parse them.

| Dop | XLSX | XLSB |
|---:|---:|---:|
| 1 | 44.39 ms, 8.02 MB | 24.34 ms, 8.02 MB |
| 2 | 28.55 ms, 8.14 MB | 15.01 ms, 8.07 MB |
| 4 | 16.13 ms, 8.11 MB | 12.91 ms, 8.07 MB |
| 8 | 16.25 ms, 8.17 MB | 12.96 ms, 8.09 MB |

At dop 4 XLSX reads ~2.8x faster than sequentially and XLSB ~1.9x. Dop 2 is a single parse worker running alongside the inflating thread. Past dop 4 nothing improves: the one inflating thread is the floor (~14 ms for XLSX, ~9 ms for XLSB, see [Inflate](#inflate), plus the XLSB record walk that finds row boundaries). Allocation barely moves with dop (8.02 to 8.17 MB): it is the typed model's strings.

### Large sheets

`LargeSheetParallelBenchmark` reads sheets generated in memory at the largest sizes the formats allow. XLSX and XLSB hold 1,000,000 rows by four columns (text, integer, date, number), the limit being 1,048,576 rows per sheet; CSV holds 4,000,000 conversion-heavy rows (two strings, a date, two decimals and an int), and has no such limit. "Aggregate" folds every row into a per-partition sum; "typed parse" is the ordered `ParseAsync` into a class (`ref struct` for the CSV aggregate).

| XLSX, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 179.48 ms, 77.25 MB | 205.45 ms, 77.25 MB |
| 2 | 134.76 ms, 77.68 MB | 145.41 ms, 106.87 MB |
| 4 | 48.73 ms, 77.58 MB | 60.19 ms, 106.78 MB |
| 8 | 39.57 ms, 77.66 MB | 48.65 ms, 106.84 MB |
| 16 | 39.85 ms, 78.00 MB | 50.17 ms, 107.17 MB |

| XLSB, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 130.43 ms, 77.25 MB | 155.74 ms, 77.25 MB |
| 2 | 95.28 ms, 77.57 MB | 104.01 ms, 99.73 MB |
| 4 | 45.25 ms, 77.48 MB | 51.91 ms, 99.62 MB |
| 8 | 45.65 ms, 77.62 MB | 52.67 ms, 99.76 MB |
| 16 | 45.75 ms, 77.88 MB | 53.06 ms, 100.02 MB |

| CSV, Dop | Aggregate | Typed parse | Typed parse, struct + `InternStrings` |
|---:|---:|---:|---:|
| 1 | 500.39 ms, 1.99 KB | 497.73 ms, 623.07 MB | 508.79 ms, 74.70 KB |
| 2 | 206.19 ms, 7.05 KB | 326.13 ms, 626.88 MB | 321.23 ms, 4.50 MB |
| 4 | 106.38 ms, 10.94 KB | 180.35 ms, 626.72 MB | 169.94 ms, 4.80 MB |
| 8 | 76.57 ms, 25.73 KB | 154.98 ms, 626.40 MB | 127.37 ms, 5.48 MB |
| 16 | 58.07 ms, 35.14 KB | 201.60 ms, 626.54 MB | 120.13 ms, 7.65 MB |

The XLSX and XLSB aggregates and typed parses each materialize a `Record` and its name string per row, which is where their ~77-107 MB comes from; the CSV aggregate binds a `ref struct` and allocates kilobytes, while the CSV typed parse allocates ~626 MB at every degree, one row object and two strings per record. The struct + `InternStrings` parse drops both (the corpus repeats 4 regions and 6 countries, see [Parallel CSV](#parallel-csv)) and keeps a few MB of per-worker caches and merge buffers.

XLSX and XLSB reach ~2.9-3.7x at dop 4. The XLSB aggregate stops there (2.9x), as in the real-data table above, while the XLSX aggregate improves to 4.5x at dop 8 and then stops. The CSV aggregate keeps scaling because nothing but the parser runs on the workers: 206.19 ms at dop 2, 76.57 ms at dop 8 and 58.07 ms at dop 16. Its dop 1 row (500.39 ms) is no faster than the typed parse at dop 1 in this run, so ratios against it would overstate the scaling. The typed CSV parse reaches 3.2x at dop 8 and then loses 30% at dop 16 to the collector, while the struct + `InternStrings` parse keeps going to 4.2x at dop 16, 1.7x faster than the class. The aggregate is still the fastest, ~2.1x ahead of the struct parse at dop 16, because it never builds the row at all.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 38.378 ms, 5.46 KB | 24.991 ms, 21.60 KB | 194.303 ms, 644.27 KB |
| XLSM | 37.822 ms, 5.46 KB | 24.501 ms, 21.52 KB | 192.195 ms, 644.33 KB |
| XLSB | 18.939 ms, 4.98 KB | 10.395 ms, 11.35 KB | 30.055 ms, 338.55 KB |
| XLS | 9.383 ms, 10.55 KB | n/a | 18.110 ms, 185.90 KB |
| CSV | 4.958 ms, 896 B | n/a | 4.595 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~5.1x faster than Sylvan for XLSX and XLSM, ~1.6x faster for XLSB, and ~1.9x faster for XLS — allocating ~118x less for XLSX and XLSM, ~68x less for XLSB and ~18x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~7% faster, while ExcelReader allocates ~46x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 38.792 ms | 0.197 ms | 5.09 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 24.980 ms | 0.390 ms | 21.17 KB |
| Xlsm_ExcelReader_Memory | 38.033 ms | 0.208 ms | 5.09 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 24.836 ms | 0.080 ms | 21.17 KB |
| Xlsb_ExcelReader_Memory | 18.590 ms | 0.083 ms | 4.70 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 10.433 ms | 0.049 ms | 9.20 KB |
| Xls_ExcelReader_Memory | 9.080 ms | 0.098 ms | 10.39 KB |
| Csv_ExcelReader_Memory | 4.761 ms | 0.004 ms | 616 B |

`Csv_ExcelReader_Memory` is both faster (4.761 ms vs. 4.958 ms for `Csv_ExcelReader`) and allocates less (616 B vs. 896 B) than its stream twin.

`Xls_ExcelReader_Memory` is ~3% faster than its stream twin (9.080 ms vs. 9.383 ms for `Xls_ExcelReader`), with near-identical allocation (10.39 KB vs. 10.55 KB). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since mostly caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 30.50 ms, 746.87 KB | 20.47 ms, 783.11 KB | 184.25 ms, 17.41 MB |
| XLSB | 26.40 ms, 746.34 KB | 16.81 ms, 791.60 KB | 60.08 ms, 17.38 MB |

Both formats handle this well: ~6.0x and ~2.3x faster than Sylvan for XLSX and XLSB respectively, at ~24x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 32.20 ms | 747 KB |
| Whole read, prefetch | 19.94 ms | 782 KB |
| Shared-string table only (open + first row) | 9.57 ms | 747 KB |
| Whole read, stored | 18.80 ms | 746 KB |
| Shared-string table only, stored | 3.56 ms | 746 KB |

About 42% of the default read is inflate: without it the whole read drops from 32.20 ms to 18.80 ms, and the shared-string table from 9.57 ms to 3.56 ms. What remains, ~15 ms, is the worksheet scan.

### Inflate

`InflateBenchmark` decompresses one ZIP entry per method with `System.IO.Compression.DeflateStream` and with the library's managed `InflateStream`, which every ZIP-based reader now uses:

| Entry | `DeflateStream` | `InflateStream` | Ratio |
|---|---:|---:|---:|
| XLSX `sheet1.xml` (6.7 MB → 34.4 MB) | 22.908 ms, 280 B | 13.537 ms, 200 B | 0.59 |
| XLSB `sheet1.bin` (3.7 MB → 15.9 MB) | 12.977 ms, 280 B | 9.139 ms, 200 B | 0.70 |
| String-heavy `sharedStrings.xml` | 10.721 ms, 280 B | 5.842 ms, 200 B | 0.54 |

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
| Serial | 160.09 ms, 19.44 KB | 1.0x | 110.06 ms, 16.98 KB | 1.0x |
| Serial, `PrefetchWrite` | 99.24 ms, 21.72 KB | 1.6x | 61.54 ms, 19.22 KB | 1.8x |
| Parallel, 1 batch in flight | 174.04 ms, 19.52 KB | 0.9x | 107.89 ms, 17.04 KB | 1.0x |
| Parallel, 2 | 87.64 ms, 86.28 KB | 1.8x | 53.74 ms, 86.70 KB | 2.0x |
| Parallel, 4 | 45.61 ms, 83.28 KB | 3.5x | 28.72 ms, 83.38 KB | 3.8x |
| Parallel, 8 | 28.93 ms, 77.82 KB | 5.5x | 18.66 ms, 74.32 KB | 5.9x |
| Parallel, 16 | 22.07 ms, 81.85 KB | 7.3x | 15.11 ms, 70.65 KB | 7.3x |

It scales because a write is bound by building rows, not by deflate: `PrefetchWrite` already takes deflate off the calling thread and stops at 1.6x to 1.8x, so compressing chunks in parallel would add nothing on top of it, while formatting the rows in parallel does. Past dop 8 the gain shrinks (5.5x to 7.3x for XLSX, 5.9x to 7.3x for XLSB), because the machine has 8 physical cores. None of the rows triggers a garbage collection, and every write allocates under 90 KB for 1,000,000 rows.

Each batch is compressed with no reference to the one before it, which costs 0.18% of file size for XLSX (15,481,372 B to 15,508,760 B) and 0.04% for XLSB (16,863,131 B to 16,869,306 B). The process's peak working set rose by 12 MB over the serial write with 4 batches in flight and by 23 MB with 8, about 3 MB per batch, and does not depend on the number of rows. The file sizes and working sets come from a console harness, because BenchmarkDotNet does not report them.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 39.44 ms | 0.78 ms | 16.61 KB |
| XLSM | 39.51 ms | 0.22 ms | 16.61 KB |
| XLSB | 19.77 ms | 0.03 ms | 16.13 KB |
| XLS | 10.29 ms | 0.07 ms | 21.70 KB |
| CSV | 12.85 ms | 0.06 ms | 34.21 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 7.03 ms | 0.16 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 48.52 ms | 0.72 ms | 1,181.8 | 1,090.9 | 363.6 | 15.25 MB |
| XLSB | 46.11 ms | 0.34 ms | 1,181.8 | 1,090.9 | 363.6 | 15.25 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 34.21 MB materialized against 896 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 16.13 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.25 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan and is still ~3.8x faster (48.52 ms vs. 184.25 ms). The same allocation holds for XLSB (15.25 MB vs. Sylvan's 17.38 MB, also ~12% less), where matched-work time is ~23% ahead of Sylvan (46.11 ms vs. 60.08 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 8.194 ms | 4.02 MB |
| XLSB | 6.053 ms | 4.02 MB |
| XLS | 4.938 ms | 4.32 MB |
| CSV | 4.961 ms | 4.00 MB |

Relative ordering: XLS and CSV tied at the front, then XLSB, then XLSX. Against hand-written cell-by-cell writes the record-mapping layer adds nothing for XLSX (8.194 ms vs. 8.213 ms) and ~12% for XLSB (6.053 ms vs. 5.388 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 19.732 ms | 19.27 KB |
| Native string columns → CSV | 4.460 ms | 424 B |
| Arrow `StringArray` → XLSX | 20.812 ms | 19.32 KB |
| Styled rows → XLSX | 5.260 ms | 19.47 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~19 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 8.88 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 8.98 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 8.15 ms | 4.02 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.9% less than the `class` baseline — and is ~9% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`SheetDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 6.139 ms | 1.00 | 3.52 KB |
| `GetBytes` (text column only) | 6.037 ms | 0.99 | 4.68 KB |
| `GetValue` (every column) | 8.473 ms | 1.39 | 5.01 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 8.565 ms | 1.40 | 1.58 MB |
| `DataTable.Load` | 35.01 ms | 5.73 | 16.60 MB |

The bridge costs ~1.4x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it allocates nothing per row and costs nothing measurable over the raw read (6.037 ms vs. 6.139 ms, against a 0.46 ms standard deviation on the baseline). `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 8.473 ms to read through `GetValue` alone.

### Arrow conversion

`ArrowConversionBenchmark`: `ToArrowRecordBatch` from [ExcelReader.Arrow](../guide/reading.md#apache-arrow-conversion-opt-in), 100,000 rows into one `RecordBatch`. The typed rows use the generated records (a text name, an int64 id, a timestamp, a double) as CSV and as XLSB. The all-string row is a separate header-less CSV with 8 text columns.

| Scenario | Mean | Allocated |
|---|---:|---:|
| CSV, typed, explicit schema | 8.81 ms | 3.13 MB |
| CSV, typed, inferred schema | 7.28 ms | 5.62 MB |
| XLSB, typed, explicit schema | 12.01 ms | 3.13 MB |
| CSV, 8 text columns, explicit schema | 10.24 ms | 6.69 MB |

`ToArrowRecordBatch` reads through the same column builder as the native `xl_parse_arrow`, then copies each finished buffer once into the managed array Apache.Arrow takes. Before that (per-cell appenders into Apache.Arrow builders) the same rows took 12.26 ms / 8.13 MB (CSV typed), 16.88 ms / 8.14 MB (XLSB) and 22.31 ms / 16.27 MB (8 text columns). The inferred schema is faster than the explicit one because `Excel.InferSchema` types every CSV column as a string, which skips the int64/timestamp/double conversion; it allocates ~1.8x more. The batch holds every row at once and there is no streaming variant.

On the 65K-row real-data fixtures (14 columns, 5 of them text, one thread), `ExcelColumnType.StringDictionaryColumn` on the text columns cut XLSB from 27.2 to 24.0 ms and XLSX from 45.5 to 43.4 ms; CSV, which has no shared strings to look up by index, stayed the same (22.0 vs 22.4 ms). Those come from a quick console harness, not BenchmarkDotNet.

### Encrypted workbooks

`EncryptedWorkbookBenchmark`: agile encryption (AES-256, SHA-512, 100,000 spin iterations — see [Encrypted workbooks](../guide/encryption.md)). The small rows use a 15 KB fixture; the large rows use a generated 300,000-row workbook, encrypted with `Excel.EncryptPackage`.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Small, plain | 0.215 ms | 8.72 KB |
| Small, encrypted, stream | 25.094 ms | 38.67 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 24.696 ms | 40.81 KB |
| Small, encrypted, in memory | 24.761 ms | 53.75 KB |
| Small, encrypted, open only | 25.578 ms | 33.05 KB |
| 300,000 rows, plain | 33.942 ms | 3.52 KB |
| 300,000 rows, encrypted | 59.446 ms | 29.87 KB |
| 300,000 rows, encrypted, open only | 24.877 ms | 29.16 KB |

Nearly all of the cost is a fixed ~25 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 25.58 ms for 15 KB and 24.88 ms for 300,000 rows). After that, decrypting while reading costs ~2% over a plain read: 59.45 ms − 24.88 ms = 34.57 ms, against 33.94 ms unencrypted. On the small fixture, integrity verification and the in-memory path stay within ~2% of the plain encrypted read, inside the spread of the key derivation itself. The fixture is too small to show the extra pass that verification makes over a large stream.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 56.25 ms | 20.54 KB |
| First typed record write | 17.90 ms | 84.20 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 56.25 ms | 20.54 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 44.98 ms | 23.70 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 58.96 ms | 25.60 KB |

`Build` is ~20% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slower than reflection (~5%, with a 0.3 ms standard deviation), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

### Native C ABI reads

`NativeRowReadBenchmark` and `NativeTypedParseBenchmark` call the C ABI entry points from .NET, reading the real-data XLSB workbook (65,535 rows × 14 columns) from memory. This isolates what the ABI itself costs, without any binding's marshalling on top.

| Entry point | Shape handed back | Mean | Managed allocated |
|---|---|---:|---:|
| `xl_rows_next` | one row per call, serialized into a caller buffer | 31.65 ms | 3.71 MB |
| `xl_rows_next_cells` | one row per call, cursor-owned cell views, numbers binary | 19.49 ms | 3.71 MB |
| `xl_rows_read_all_blob` | whole sheet in one buffer | 35.24 ms | 21.05 MB |
| `xl_rows_read_all_decoded` | whole sheet as native row structs | 46.44 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 25.75 ms | 3.71 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 26.01 ms | 3.72 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_rows_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Every cell that crosses the ABI as text costs time: each binary XLSB number is formatted to UTF-8 on the way out. `xl_rows_next_cells` passes numbers as doubles and copies only the text, so it reads the sheet ~38% faster than `xl_rows_next`, and faster than `xl_parse_typed`, which also converts every cell to its column type. `xl_rows_read_all_blob` adds ~11% over `xl_rows_next`, because the whole sheet is buffered before it is copied out. `xl_rows_read_all_decoded` adds ~47%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads; `xl_rows_next_cells` is the one for reading cells as they are.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 6.97 ms | 10.63 KB |
| Batches of 10,000 rows | 6.68 ms | 18.97 KB |
| Batches of 1,000 rows | 7.00 ms | 97.37 KB |

Batching costs nothing in time: batches of 10,000 rows are ~4% faster than one whole-sheet table, and batches of 1,000 rows land within 1% of it. A caller can bound its native memory to one batch at a time without paying for it.

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
| vs. [calamine](https://github.com/tafia/calamine) (Rust) | XLSX | 69.5 ms | 276.4 ms | ~4.0x faster |
| vs. calamine (Rust) | XLSB | 47.8 ms | 87.0 ms | ~1.8x faster |
| vs. [DuckDB](https://github.com/duckdb/duckdb) `read_xlsx` (C++) | XLSX | 58.0 ms | 417.8 ms | ~7.2x faster |
| vs. [xlsxio](https://github.com/brechtsanders/xlsxio) (C) | XLSX | 58.0 ms | 515.5 ms | ~8.9x faster |
| vs. [xlnt](https://github.com/tfussell/xlnt) (C++) | XLSX | 58.0 ms | 2,495.0 ms | ~43x faster |
| vs. [zsv](https://github.com/liquidaty/zsv) `fast` engine (C) | CSV | 28.1 ms | 21.2 ms | ~1.3x **slower** |

Writing the same 14 columns × 65,535 rows from row-shaped data, the C++ suite measures
`write_sheet` at 76.2–76.5 ms against DuckDB's 835 ms (~11x), libxlsxwriter's 1,200 ms (~16x),
xlsxio's 2,441 ms (~32x) and xlnt's 5,554 ms (~73x). The Rust suite measures 40.8 ms against
rust_xlsxwriter's 338.5 ms (~8.3x) over 7 columns. Caveats for both are in the binding READMEs.

DuckDB, xlsxio and xlnt do not read `.xlsb`, so those comparisons are XLSX-only. calamine is a fast,
well-optimized reader in its own right — the gap there is real but not the order of magnitude seen
against the C/C++ competitors.

zsv is the one competitor ahead. The CSV row is the typed workload, where both sides build the same
14-column rows with owned strings. Just counting cell bytes, zsv is ~3.4x faster, and that gap is
ExcelReader's C ABI, which copies every row's cells out to the caller, not its CSV parser. That suite is
built with GCC (zsv does not compile under MSVC). Its caveats are in the C++ README.
