# Benchmarks

Throwaway numbers go stale; these are regenerated from the benchmark suite in
`tests/ExcelReader.Benchmarks`. Live results: https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.12 (SDK 10.0.401). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [GitHub Pages benchmark dashboard](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/) — `tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/` is `.gitignore`-excluded and never actually reaches GitHub, so a link to it 404s for every reader.

Every BenchmarkDotNet table is a `--job Short` run (1 launch, 3 warmup and 3 measured iterations) on an otherwise idle machine. Rows whose 99.9% confidence interval came out wider than 25% of the mean were run again with `--job Medium` (2 launches, 15 iterations) and carry those figures instead: the parallel CSV tables, the CSV cell-by-cell read, the Sylvan XLS read, the Sylvan string-heavy reads, the string-heavy XLSX matched-work read, the generated-workbook matched-work read, `DataTable.Load`, the large encrypted read and the `InternStrings` typed parse. The cold-start rows set their own launch count, and the file-size and working-set figures under the parallel record writes, the Python/C++/Rust figures and the native-library comparison come from other harnesses that were not re-run. The large-sheet and parallel-write tables are `--job Medium` runs.

**Benchmark methodology.** In the Excel-format tables below, the **"Cell-by-cell read"** rows (including the per-format rows in "Real data reads" / "String-heavy reads") read ExcelReader's `cell.Value` — a zero-copy `ReadOnlySpan<byte>`, no decode or allocation — against each competitor's own idiomatic read call. For Sylvan.Data.Excel, that's the ADO.NET-style `GetString(i)`, which is forced to materialize a UTF-16 `string`; its API has no zero-copy accessor, so it cannot avoid that cost. These rows are therefore not matched work: part of the reported gap is "we parse faster" and part is "we skipped an allocation you were never offered a way to skip." Each affected benchmark class also has a `*_Materialized` sibling (calling `cell.GetString()`, the same UTF-16 materialization Sylvan pays) so the matched-work number is measurable — see `tests/ExcelReader.Benchmarks/Shared/BenchmarkAccumulators.cs`. **Those results are published under [Matched-work reads](#matched-work-reads-_materialized), and reading them is the honest way to judge the comparison:** treat the Excel cell-by-cell ratios in the tables below as an upper bound, not a like-for-like number. The CSV cell-by-cell rows are matched work: Sep and Sylvan.Data.Csv both expose span accessors (`row[i].Span`, `GetFieldSpan(i)`), and the benchmarks use them. The **"Typed row parsing"** / **"Typed record writing"** rows are unaffected — both sides already materialize real objects/strings there, so those comparisons are matched work as published.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | Sylvan | OfficeIMO | SpreadCheetah |
|---|---:|---:|---:|---:|
| Cell-by-cell read | 6.780 ms, 3.63 KB | 35.084 ms, 1.89 MB | 14.447 ms, 3.51 MB | - |
| Cell-by-cell read async | 6.953 ms, 3.70 KB | - | - | - |
| Typed row parsing | 10.979 ms, 3.87 MB | 56.677 ms, 10.48 MB | 11.802 ms, 2.36 MB | - |
| Typed row parsing async | 10.958 ms, 3.87 MB | 58.753 ms, 10.48 MB | - | - |
| Typed row parsing, shared strings | 9.896 ms, 2.29 MB | - | - | - |
| Typed row parsing, `InternStrings` | 11.58 ms, 2.30 MB | - | - | - |
| Workbook writing | 8.527 ms, 4.02 MB | - | 18.100 ms, 4.04 MB | 15.393 ms, 15.84 MB |
| Workbook writing, shared strings | 8.416 ms, 4.06 MB | - | - | - |

ExcelReader is ~5.2x faster than Sylvan for raw XLSX reads, allocating ~534x less, and ~2.1x faster than OfficeIMO. For typed parsing it is ~5.2x faster than Sylvan (~5.4x async against async) and ~1.1x faster than OfficeIMO, which allocates less (2.36 MB against 3.87 MB) because it always interns strings. `ExcelReaderOptions.InternStrings` is off by default; turned on, ExcelReader allocates 2.30 MB at 11.58 ms, against OfficeIMO's 11.93 ms in the same `--job Medium` run. For XLSX writing, ExcelReader is ~1.8x faster than SpreadCheetah while allocating ~3.9x less memory, and ~2.1x faster than OfficeIMO. The OfficeIMO rows are OfficeIMO.Excel 3.4.5: that release made its XLSX read and typed parse several times faster than the 3.4.4 figures published before (106.861 ms and 107.805 ms).

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~10% faster than the inline-string sheet above (9.896 ms vs. 10.979 ms) and allocates ~41% less (2.29 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | OfficeIMO |
|---|---:|---:|
| Cell-by-cell read | 4.881 ms, 3.89 KB | 10.928 ms, 5.12 MB |
| Cell-by-cell read async | 5.220 ms, 3.96 KB | - |
| Typed row parsing | 7.851 ms, 3.87 MB | - |
| Typed row parsing async | 8.014 ms, 3.87 MB | - |
| Workbook writing | 5.786 ms, 4.02 MB | - |
| Workbook writing, shared strings | 5.734 ms, 4.05 MB | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than XLSX reads, typed parsing is ~1.4x faster than XLSX parsing, and writing is ~1.5x faster than XLSX writing. Against OfficeIMO's XLSB reader, ExcelReader is ~2.2x faster while allocating ~1,350x less. The XLSB writer is also ~2.7x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 3.961 ms, 2.84 KB | 5.177 ms, 1,717.73 KB | 7.788 ms, 10,499.26 KB |
| Cell-by-cell read async | 3.253 ms, 2.91 KB | - | - |
| Workbook writing | 4.753 ms, 16.31 MB | - | - |

ExcelReader is ~1.3x faster than Sylvan for generated XLS reads while allocating ~605x less memory, and ~2.0x faster than OfficeIMO. The XLS writer is ~1.8x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 4.753 ms vs 8.600 ms in the same run). Most of its 16.31 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.31 MB, close to the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 2.499 ms, 888 B | 7.827 ms, 3.93 KB | 4.044 ms, 37.78 KB |
| Cell-by-cell read async | 2.935 ms, 960 B | - | - |
| Typed row parsing | 4.889 ms, 3.86 MB | 8.736 ms, 3.87 MB | 12.161 ms, 10.95 MB |
| Typed row parsing async | 4.915 ms, 3.86 MB | - | - |
| Row writing | 4.551 ms, 4.00 MB | 6.880 ms, 4.01 MB | 6.983 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~3.1x faster than Sep while allocating ~4.5x less, and ~1.6x faster than Sylvan.Data.Csv while allocating ~44x less. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.8x faster than Sep and ~2.5x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`SheetParallel.ParseAsync<T>` and `SheetParallel.AggregateAsync` over a `CsvSheet`, on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|
| 1 | 587.8 ms, 671.42 MB | 516.9 ms, 245.76 MB | 472.0 ms, 1.20 MB | 463.9 ms, 1.20 MB |
| 2 | 461.6 ms, 680.28 MB | 385.8 ms, 254.80 MB | 267.3 ms, 1.35 MB | 265.6 ms, 1.34 MB |
| 4 | 256.9 ms, 680.39 MB | 217.8 ms, 254.81 MB | 145.2 ms, 1.50 MB | 140.1 ms, 1.51 MB |
| 8 | 210.5 ms, 680.31 MB | 187.8 ms, 254.77 MB | 96.3 ms, 1.69 MB | 93.1 ms, 1.68 MB |
| 16 | 471.2 ms, 680.33 MB | 187.0 ms, 255.23 MB | 74.0 ms, 1.79 MB | 73.3 ms, 1.76 MB |

The typed path scales ~2.3x by dop 4 on the conversion-heavy corpus (587.8 to 256.9 ms) and ~2.4x on narrow ints (516.9 to 217.8 ms), then flattens out: narrow ints reach only 187.0 ms at dop 16, and the conversion-heavy corpus gets *slower* past dop 8 (471.2 ms at dop 16 against 210.5 ms at dop 8, standard deviation 23.2 ms). Its ~680 MB is one materialized row object per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~4.9x (conversion-heavy) and ~5.0x (narrow) at dop 8, ~6.4x and ~6.3x at dop 16 — because they are not fighting the allocator: at dop 4 the conversion-heavy aggregate is ~1.8x faster than typed while allocating ~450x less (1.50 MB against 680.39 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

The typed path's slowdown past dop 8 is the garbage collector. Every degree allocates the same ~680 MB and runs 42 gen0 collections, but at dop 16 each of them also collects gen1 (41 gen1 collections against 24 at dop 8), and the pauses add up to 100-200 ms against 39 ms. In a console harness on the 4.3 million row corpus, with the default workstation GC, dop 16 took 365 ms; with `DOTNET_gcServer=1` it took 182 ms, faster than dop 8. The aggregate path allocates no row objects and is unaffected.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 472.0 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 479.4 ms, 244.15 MB | 189.8 ms, 255.24 MB | 515.1 ms, 244.15 MB | 434.8 ms, 245.60 MB |
| Conversion-heavy | 397.9 ms, 467.31 MB | 330.4 ms, 474.83 MB | 808.7 ms, 467.31 MB | 619.0 ms, 469.93 MB |

Sequentially, ExcelReader's typed parser is ~1.1x faster than Sep on narrow ints (479.4 ms vs. 515.1 ms), and ~2.0x faster on the conversion-heavy corpus, where decimal parsing dominates. In parallel ExcelReader is ~2.3x (narrow) and ~1.9x (conversion-heavy) faster than Sep's, while yielding rows in file order. The conversion-heavy parallel run gains less than the narrow one (330.4 ms against 397.9 ms sequential) because it is bound by the allocator, like the typed dop 16 row above: 28 Gen1 collections per run.

### Parallel sheet reading

`SheetParallel.AggregateAsync<TAccumulator, TModel>` over the first sheet of the real-data workbook (65,535 rows, 14 columns; XLSX 6.7 MB compressed, 34.4 MB sheet XML; XLSB 3.7 MB, 15.9 MB sheet records), folding every row of a typed model into a per-partition sum (`SheetParallelBenchmark`). One thread inflates the sheet and cuts it into row-aligned chunks; `Dop - 1` workers parse them.

| Dop | XLSX | XLSB |
|---:|---:|---:|
| 1 | 52.76 ms, 8.02 MB | 33.82 ms, 8.02 MB |
| 2 | 36.63 ms, 8.13 MB | 25.35 ms, 8.07 MB |
| 4 | 16.64 ms, 8.11 MB | 13.25 ms, 8.07 MB |
| 8 | 16.73 ms, 8.16 MB | 13.25 ms, 8.09 MB |

At dop 4 XLSX reads ~3.2x faster than sequentially and XLSB ~2.6x. Dop 2 is a single parse worker running alongside the inflating thread. Past dop 4 nothing improves: the one inflating thread is the floor (~14 ms for XLSX, ~9 ms for XLSB, see [Inflate](#inflate), plus the XLSB record walk that finds row boundaries). Allocation barely moves with dop (8.02 to 8.16 MB): it is the typed model's strings.

### Large sheets

`LargeSheetParallelBenchmark` reads sheets generated in memory at the largest sizes the formats allow. XLSX and XLSB hold 1,000,000 rows by four columns (text, integer, date, number), the limit being 1,048,576 rows per sheet; CSV holds 4,000,000 conversion-heavy rows (two strings, a date, two decimals and an int), and has no such limit. "Aggregate" folds every row into a per-partition sum; "typed parse" is the ordered `ParseAsync` into a class (`ref struct` for the CSV aggregate).

| XLSX, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 215.91 ms, 77.25 MB | 243.27 ms, 77.25 MB |
| 2 | 169.84 ms, 77.62 MB | 185.00 ms, 106.81 MB |
| 4 | 64.81 ms, 77.58 MB | 74.01 ms, 106.75 MB |
| 8 | 51.11 ms, 77.52 MB | 55.44 ms, 106.78 MB |
| 16 | 44.51 ms, 77.94 MB | 57.32 ms, 107.11 MB |

| XLSB, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 160.91 ms, 77.25 MB | 189.08 ms, 77.25 MB |
| 2 | 119.50 ms, 77.53 MB | 127.86 ms, 99.69 MB |
| 4 | 48.76 ms, 77.46 MB | 58.04 ms, 99.59 MB |
| 8 | 46.92 ms, 77.61 MB | 53.51 ms, 99.74 MB |
| 16 | 49.19 ms, 77.88 MB | 54.30 ms, 100.00 MB |

| CSV, Dop | Aggregate | Typed parse |
|---:|---:|---:|
| 1 | 426.07 ms, 1.99 KB | 543.94 ms, 623.08 MB |
| 2 | 226.22 ms, 6.77 KB | 338.08 ms, 626.87 MB |
| 4 | 118.30 ms, 10.60 KB | 191.11 ms, 626.71 MB |
| 8 | 90.67 ms, 18.32 KB | 189.98 ms, 626.43 MB |
| 16 | 64.16 ms, 34.29 KB | 207.63 ms, 626.57 MB |

The XLSX and XLSB aggregates and typed parses each materialize a `Record` and its name string per row, which is where their ~77-107 MB comes from; the CSV aggregate binds a `ref struct` and allocates kilobytes, while the CSV typed parse allocates ~626 MB at every degree, one row object and two strings per record.

XLSX and XLSB reach ~3.3x at dop 4 and then barely improve, as in the real-data table above (the XLSX aggregate gets to 4.9x at dop 16, the XLSB aggregate to 3.4x at dop 8). The CSV aggregate keeps scaling because nothing but the parser runs on the workers: 4.7x at dop 8 and 6.6x at dop 16. The typed CSV parse reaches 2.9x at dop 8 and then loses 9% at dop 16 to the collector (see [Parallel CSV](#parallel-csv)), so for a 4-million-row file the aggregate at dop 16 is ~3.2x faster than the typed parse and allocates no row objects.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 41.206 ms, 7.38 KB | 27.823 ms, 23.54 KB | 191.883 ms, 644.25 KB |
| XLSM | 43.502 ms, 7.38 KB | 28.364 ms, 23.55 KB | 197.799 ms, 644.33 KB |
| XLSB | 25.775 ms, 6.94 KB | 16.722 ms, 15.28 KB | 29.536 ms, 338.54 KB |
| XLS | 10.750 ms, 12.02 KB | n/a | 18.124 ms, 185.90 KB |
| CSV | 4.992 ms, 888 B | n/a | 4.482 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~4.7x faster than Sylvan for XLSX and ~4.5x for XLSM, ~1.1x faster for XLSB, and ~1.7x faster for XLS — allocating ~87x less for XLSX and XLSM, ~49x less for XLSB and ~15x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~11% faster, while ExcelReader allocates ~46x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 41.625 ms | 0.079 ms | 7.10 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 28.699 ms | 0.022 ms | 23.26 KB |
| Xlsm_ExcelReader_Memory | 41.559 ms | 0.042 ms | 7.10 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 28.687 ms | 0.011 ms | 23.23 KB |
| Xlsb_ExcelReader_Memory | 24.465 ms | 0.008 ms | 8.61 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 17.235 ms | 0.015 ms | 16.93 KB |
| Xls_ExcelReader_Memory | 10.574 ms | 0.026 ms | 11.88 KB |
| Csv_ExcelReader_Memory | 4.762 ms | 0.003 ms | 616 B |

`Csv_ExcelReader_Memory` is both faster (4.762 ms vs. 4.992 ms for `Csv_ExcelReader`) and allocates less (616 B vs. 888 B) than its stream twin.

`Xls_ExcelReader_Memory` is ~2% faster than its stream twin (10.574 ms vs. 10.750 ms for `Xls_ExcelReader`), with near-identical allocation (11.88 KB vs. 12.02 KB). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since mostly caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 33.14 ms, 2.18 MB | 22.97 ms, 2.21 MB | 176.85 ms, 17.41 MB |
| XLSB | 32.56 ms, 2.18 MB | 19.68 ms, 2.20 MB | 59.78 ms, 17.38 MB |

Both formats handle this well: ~5.3x and ~1.8x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 33.06 ms | 2.18 MB |
| Whole read, prefetch | 23.00 ms | 2.21 MB |
| Shared-string table only (open + first row) | 9.37 ms | 747 KB |
| Whole read, stored | 20.69 ms | 2.18 MB |
| Shared-string table only, stored | 3.35 ms | 746 KB |

About 37% of the default read is inflate: without it the whole read drops from 33.06 ms to 20.69 ms, and the shared-string table from 9.37 ms to 3.35 ms. What remains, ~17 ms, is the worksheet scan.

### Inflate

`InflateBenchmark` decompresses one ZIP entry per method with `System.IO.Compression.DeflateStream` and with the library's managed `InflateStream`, which every ZIP-based reader now uses:

| Entry | `DeflateStream` | `InflateStream` | Ratio |
|---|---:|---:|---:|
| XLSX `sheet1.xml` (6.7 MB → 34.4 MB) | 22.232 ms, 280 B | 13.523 ms, 200 B | 0.61 |
| XLSB `sheet1.bin` (3.7 MB → 15.9 MB) | 12.695 ms, 280 B | 9.147 ms, 200 B | 0.72 |
| String-heavy `sharedStrings.xml` | 10.375 ms, 280 B | 5.915 ms, 200 B | 0.57 |

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
| Serial | 171.57 ms, 16.11 KB | 1.0x | 115.39 ms, 13.42 KB | 1.0x |
| Serial, `PrefetchWrite` | 104.92 ms, 18.43 KB | 1.6x | 60.02 ms, 16.25 KB | 1.9x |
| Parallel, 1 batch in flight | 165.22 ms, 48.09 KB | 1.0x | 107.81 ms, 45.48 KB | 1.1x |
| Parallel, 2 | 89.69 ms, 83.40 KB | 1.9x | 57.94 ms, 81.92 KB | 2.0x |
| Parallel, 4 | 46.29 ms, 79.39 KB | 3.7x | 30.62 ms, 77.87 KB | 3.8x |
| Parallel, 8 | 30.56 ms, 74.61 KB | 5.6x | 21.04 ms, 70.78 KB | 5.5x |
| Parallel, 16 | 27.93 ms, 72.66 KB | 6.1x | 18.25 ms, 66.78 KB | 6.3x |

It scales because a write is bound by building rows, not by deflate: `PrefetchWrite` already takes deflate off the calling thread and stops at 1.6x to 1.9x, so compressing chunks in parallel would add nothing on top of it, while formatting the rows in parallel does. Past dop 8 the gain is small, because the machine has 8 physical cores. None of the rows triggers a garbage collection, and every write allocates under 85 KB for 1,000,000 rows.

Each batch is compressed with no reference to the one before it, which costs 0.18% of file size for XLSX (15,481,372 B to 15,508,760 B) and 0.04% for XLSB (16,863,131 B to 16,869,306 B). The process's peak working set rose by 12 MB over the serial write with 4 batches in flight and by 23 MB with 8, about 3 MB per batch, and does not depend on the number of rows. The file sizes and working sets come from a console harness, because BenchmarkDotNet does not report them.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 44.42 ms | 0.04 ms | 16.85 KB |
| XLSM | 44.96 ms | 0.06 ms | 16.85 KB |
| XLSB | 26.61 ms | 0.01 ms | 16.41 KB |
| XLS | 12.31 ms | 0.04 ms | 21.48 KB |
| CSV | 18.63 ms | 0.18 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 8.43 ms | 0.15 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 50.93 ms | 1.98 ms | 1,100.0 | 1,000.0 | 300.0 | 15.25 MB |
| XLSB | 53.12 ms | 0.67 ms | 1,100.0 | 1,000.0 | 300.0 | 15.25 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 888 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 16.41 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.25 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan and is still ~3.5x faster (50.93 ms vs. 176.85 ms). The same allocation holds for XLSB (15.25 MB vs. Sylvan's 17.38 MB, also ~12% less), where matched-work time is ~13% ahead of Sylvan (53.12 ms vs. 59.78 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 8.978 ms | 4.02 MB |
| XLSB | 6.170 ms | 4.02 MB |
| XLS | 6.167 ms | 4.31 MB |
| CSV | 4.688 ms | 4.00 MB |

Relative ordering: CSV at the front, XLS and XLSB close behind it, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 6.170 ms vs. 5.786 ms, XLSX 8.978 ms vs. 8.527 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 15.598 ms | 15.93 KB |
| Native string columns → CSV | 4.963 ms | 416 B |
| Arrow `StringArray` → XLSX | 21.837 ms | 15.98 KB |
| Styled rows → XLSX | 6.318 ms | 16.10 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~16 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 10.98 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 10.32 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 9.67 ms | 4.17 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.9% less than the `class` baseline — and is ~6% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`SheetDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 6.545 ms | 1.00 | 3.67 KB |
| `GetBytes` (text column only) | 6.766 ms | 1.03 | 4.37 KB |
| `GetValue` (every column) | 9.215 ms | 1.41 | 5.01 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 8.954 ms | 1.37 | 1.58 MB |
| `DataTable.Load` | 62.767 ms | 9.59 | 22.84 MB |

The bridge costs ~1.4x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it stays within ~3% of the raw read. `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 9.215 ms to read through `GetValue` alone.

### Arrow conversion

`ArrowConversionBenchmark`: `ToArrowRecordBatch` from [ExcelReader.Arrow](../guide/reading.md#apache-arrow-conversion-opt-in), 100,000 rows into one `RecordBatch`. The typed rows use the generated records (a text name, an int64 id, a timestamp, a double) as CSV and as XLSB. The all-string row is a separate header-less CSV with 8 text columns.

| Scenario | Mean | Allocated |
|---|---:|---:|
| CSV, typed, explicit schema | 11.94 ms | 8.13 MB |
| CSV, typed, inferred schema | 12.51 ms | 14.89 MB |
| XLSB, typed, explicit schema | 16.40 ms | 8.14 MB |
| CSV, 8 text columns, explicit schema | 22.22 ms | 16.27 MB |

On the same file, letting `Excel.InferSchema` guess the schema costs ~5% more time (12.51 ms vs. 11.94 ms) and ~1.8x the allocation of passing an explicit `ExcelColumnSchema[]`. The allocation grows with the sheet, because the batch holds every row at once and there is no streaming variant.

### Encrypted workbooks

`EncryptedWorkbookBenchmark`: agile encryption (AES-256, SHA-512, 100,000 spin iterations — see [Encrypted workbooks](../guide/encryption.md)). The small rows use a 15 KB fixture; the large rows use a generated 300,000-row workbook, encrypted with `Excel.EncryptPackage`.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Small, plain | 0.215 ms | 8.91 KB |
| Small, encrypted, stream | 26.414 ms | 38.13 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 25.961 ms | 40.29 KB |
| Small, encrypted, in memory | 25.546 ms | 53.28 KB |
| Small, encrypted, open only | 25.244 ms | 32.52 KB |
| 300,000 rows, plain | 36.634 ms | 3.67 KB |
| 300,000 rows, encrypted | 63.432 ms | 29.59 KB |
| 300,000 rows, encrypted, open only | 25.605 ms | 28.92 KB |

Nearly all of the cost is a fixed ~25 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 25.24 ms for 15 KB and 25.61 ms for 300,000 rows). After that, decrypting while reading costs ~3% over a plain read: 63.43 ms − 25.61 ms = 37.83 ms, against 36.63 ms unencrypted. On the small fixture, integrity verification and the in-memory path stay within ~3% of the plain encrypted read, inside the spread of the key derivation itself. The fixture is too small to show the extra pass that verification makes over a large stream.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 46.64 ms | 20.91 KB |
| First typed record write | 17.53 ms | 80.81 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 46.64 ms | 20.91 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 42.14 ms | 23.84 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 49.63 ms | 25.97 KB |

`Build` is ~10% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slower than reflection (~6%, with a 0.4 ms standard deviation), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

### Native C ABI reads

`NativeRowReadBenchmark` and `NativeTypedParseBenchmark` call the C ABI entry points from .NET, reading the real-data XLSB workbook (65,535 rows × 14 columns) from memory. This isolates what the ABI itself costs, without any binding's marshalling on top.

| Entry point | Shape handed back | Mean | Managed allocated |
|---|---|---:|---:|
| `xl_rows_next` | one row per call, serialized into a caller buffer | 37.72 ms | 3.71 MB |
| `xl_rows_read_all_blob` | whole sheet in one buffer | 40.31 ms | 21.06 MB |
| `xl_rows_read_all_decoded` | whole sheet as native row structs | 57.21 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 34.30 ms | 3.71 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 35.19 ms | 3.72 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_rows_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Row by row, `xl_rows_next` costs ~1.5x the managed in-memory read of the same file (37.72 ms vs. 24.465 ms for `Xlsb_ExcelReader_Memory` under [In-memory real-data reads](#in-memory-real-data-reads)). Every cell crosses the ABI as text, so each binary XLSB number is formatted to UTF-8 on the way out. `xl_rows_read_all_blob` adds ~7% over it, because the whole sheet is buffered before it is copied out. `xl_rows_read_all_decoded` adds ~52%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` are slightly faster than a row-by-row read (~7–9%) and return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 8.55 ms | 10.06 KB |
| Batches of 10,000 rows | 8.09 ms | 17.00 KB |
| Batches of 1,000 rows | 7.86 ms | 81.34 KB |

Batching costs nothing in time: both batch sizes are slightly faster than one whole-sheet table (~5–8%). A caller can bound its native memory to one batch at a time without paying for it.

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

