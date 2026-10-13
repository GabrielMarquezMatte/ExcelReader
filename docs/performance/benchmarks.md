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
| Cell-by-cell read | 6.847 ms, 3.57 KB | 36.131 ms, 1.89 MB | 106.861 ms, 6.84 MB | - |
| Cell-by-cell read async | 7.391 ms, 3.64 KB | - | - | - |
| Typed row parsing | 10.668 ms, 3.87 MB | 56.150 ms, 10.48 MB | 107.805 ms, 5.70 MB | - |
| Typed row parsing async | 10.833 ms, 3.87 MB | 59.036 ms, 10.48 MB | - | - |
| Typed row parsing, shared strings | 9.216 ms, 2.29 MB | - | - | - |
| Workbook writing | 8.337 ms, 4.02 MB | - | 18.437 ms, 4.02 MB | 15.332 ms, 15.84 MB |
| Workbook writing, shared strings | 8.358 ms, 4.06 MB | - | - | - |

ExcelReader is ~5.3x faster than Sylvan for raw XLSX reads, allocating ~543x less, and ~15.6x faster than OfficeIMO. For typed parsing it is ~5.3x faster than Sylvan (~5.4x async against async). For XLSX writing, ExcelReader is ~1.8x faster than SpreadCheetah while allocating ~3.9x less memory, and ~2.2x faster than OfficeIMO.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~14% faster than the inline-string sheet above (9.216 ms vs. 10.668 ms) and allocates ~41% less (2.29 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | OfficeIMO |
|---|---:|---:|
| Cell-by-cell read | 4.798 ms, 3.82 KB | 11.123 ms, 6.70 MB |
| Cell-by-cell read async | 4.845 ms, 3.89 KB | - |
| Typed row parsing | 7.819 ms, 3.87 MB | - |
| Typed row parsing async | 8.061 ms, 3.87 MB | - |
| Workbook writing | 5.572 ms, 4.02 MB | - |
| Workbook writing, shared strings | 5.595 ms, 4.05 MB | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~1.4x faster than XLSX reads, typed parsing is ~1.4x faster than XLSX parsing, and writing is ~1.5x faster than XLSX writing. Against OfficeIMO's XLSB reader, ExcelReader is ~2.3x faster while allocating ~1,800x less. The XLSB writer is also ~2.8x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 3.916 ms, 2.84 KB | 5.173 ms, 1,717.73 KB | 6.856 ms, 10,499.25 KB |
| Cell-by-cell read async | 3.236 ms, 2.91 KB | - | - |
| Workbook writing | 5.296 ms, 16.03 MB | - | - |

ExcelReader is ~1.3x faster than Sylvan for generated XLS reads while allocating ~605x less memory, and ~1.8x faster than OfficeIMO. The XLS writer is ~1.7x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 4.859 ms vs 8.375 ms in the same run). Most of its 16.31 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.03 MB like the other formats.

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

`SheetParallel.ParseAsync<T>` and `SheetParallel.AggregateAsync` over a `CsvSheet`, on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|
| 1 | 904.0 ms, 671.42 MB | 518.6 ms, 245.76 MB | 784.1 ms, 1.20 MB | 478.1 ms, 1.20 MB |
| 2 | 623.7 ms, 680.13 MB | 382.9 ms, 254.68 MB | 405.4 ms, 1.26 MB | 262.9 ms, 1.28 MB |
| 4 | 343.0 ms, 680.24 MB | 217.2 ms, 254.72 MB | 218.3 ms, 1.37 MB | 137.6 ms, 1.44 MB |
| 8 | 285.2 ms, 680.15 MB | 204.7 ms, 254.63 MB | 158.3 ms, 1.54 MB | 94.6 ms, 1.62 MB |
| 16 | 324.9 ms, 680.32 MB | 189.0 ms, 255.15 MB | 114.4 ms, 1.66 MB | 73.0 ms, 1.69 MB |

These numbers were measured before the rename to `SheetParallel`; the rename did not touch the CSV engine, and a scratch harness comparing the two builds on a 156 MB CSV showed no regression.

The typed path scales ~2.8x on the conversion-heavy corpus (bimodal at dop 16: 86.7 ms standard deviation, median 275.8 ms) and ~2.7x on narrow ints, where it flattens out after dop 4 (217.2 ms vs. 189.0 ms at dop 16). Its ~680 MB is one materialized row object per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~6.9x (conversion-heavy) and ~6.5x (narrow) at dop 16 — because they are not fighting the allocator: at dop 16 the conversion-heavy aggregate is ~2.8x faster than typed while allocating ~410x less (1.66 MB against 680.32 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 784.1 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 490.6 ms, 244.15 MB | 187.6 ms, 255.14 MB | 521.7 ms, 244.15 MB | 429.3 ms, 247.77 MB |
| Conversion-heavy | 626.6 ms, 467.31 MB | 184.9 ms, 474.82 MB | 817.5 ms, 467.31 MB | 605.5 ms, 468.84 MB |

Sequentially, ExcelReader's typed parser is ~1.1x faster than Sep on narrow ints (490.6 ms vs. 521.7 ms), and ~1.3x faster on the conversion-heavy corpus. In parallel ExcelReader is ~2.3x (narrow) and ~3.3x (conversion-heavy) faster than Sep's, while yielding rows in file order.

### Parallel sheet reading

`SheetParallel.AggregateAsync<TAccumulator, TModel>` over the first sheet of the real-data workbook (65,535 rows, 14 columns; XLSX 6.7 MB compressed, 34.4 MB sheet XML; XLSB 3.7 MB, 15.9 MB sheet records), folding every row of a typed model into a per-partition sum (`SheetParallelBenchmark`). One thread inflates the sheet and cuts it into row-aligned chunks; `Dop - 1` workers parse them.

| Dop | XLSX | XLSB |
|---:|---:|---:|
| 1 | 52.60 ms, 8.02 MB | 35.77 ms, 8.02 MB |
| 2 | 36.30 ms, 8.13 MB | 24.48 ms, 8.07 MB |
| 4 | 16.89 ms, 8.11 MB | 13.32 ms, 8.07 MB |
| 8 | 17.06 ms, 8.16 MB | 13.67 ms, 8.09 MB |

At dop 4 XLSX reads ~3.1x faster than sequentially and XLSB ~2.7x. Dop 2 is a single parse worker running alongside the inflating thread. Past dop 4 nothing improves: the one inflating thread is the floor (~14 ms for XLSX, ~10 ms for XLSB, see [Inflate](#inflate), plus the XLSB record walk that finds row boundaries). Allocation barely moves with dop (8.02 to 8.16 MB): it is the typed model's strings.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 41.275 ms, 7.31 KB | 29.616 ms, 23.52 KB | 197.561 ms, 644.23 KB |
| XLSM | 43.649 ms, 7.31 KB | 29.347 ms, 23.48 KB | 194.102 ms, 644.30 KB |
| XLSB | 25.303 ms, 6.87 KB | 17.588 ms, 15.20 KB | 29.771 ms, 338.54 KB |
| XLS | 11.280 ms, 12.02 KB | n/a | 18.379 ms, 185.90 KB |
| CSV | 5.072 ms, 888 B | n/a | 4.585 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~4.8x faster than Sylvan for XLSX and ~4.4x for XLSM, ~1.2x faster for XLSB, and ~1.6x faster for XLS — allocating ~88x less for XLSX and XLSM, ~49x less for XLSB and ~15x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~11% faster, while ExcelReader allocates ~46x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 42.521 ms | 0.905 ms | 7.05 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 29.167 ms | 0.353 ms | 23.03 KB |
| Xlsm_ExcelReader_Memory | 40.925 ms | 0.140 ms | 7.05 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 30.345 ms | 1.268 ms | 23.22 KB |
| Xlsb_ExcelReader_Memory | 27.084 ms | 1.294 ms | 8.55 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 18.140 ms | 0.703 ms | 16.79 KB |
| Xls_ExcelReader_Memory | 10.704 ms | 0.139 ms | 11.88 KB |
| Csv_ExcelReader_Memory | 4.759 ms | 0.022 ms | 616 B |

`Csv_ExcelReader_Memory` is both faster (4.759 ms vs. 5.072 ms for `Csv_ExcelReader`) and allocates less (616 B vs. 888 B) than its stream twin.

`Xls_ExcelReader_Memory` is ~5% faster than its stream twin (10.704 ms vs. 11.280 ms for `Xls_ExcelReader`), with near-identical allocation (11.88 KB vs. 12.02 KB). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since mostly caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 32.84 ms, 2.18 MB | 24.03 ms, 2.21 MB | 186.05 ms, 17.41 MB |
| XLSB | 32.44 ms, 2.18 MB | 21.12 ms, 2.20 MB | 65.22 ms, 17.38 MB |

Both formats handle this well: ~5.7x and ~2.0x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

`XlsxSharedStringHotPathBenchmark` splits the XLSX read into stages on the same workbook. The stored rows re-zip it without compression, so they show what parsing costs once inflate is gone:

| Stage | Mean | Allocated |
|---|---:|---:|
| Whole read | 33.87 ms | 2.18 MB |
| Whole read, prefetch | 23.36 ms | 2.21 MB |
| Shared-string table only (open + first row) | 9.31 ms | 747 KB |
| Whole read, stored | 20.23 ms | 2.18 MB |
| Shared-string table only, stored | 3.42 ms | 746 KB |

About 40% of the default read is inflate: without it the whole read drops from 33.87 ms to 20.23 ms, and the shared-string table from 9.31 ms to 3.42 ms. What remains, ~17 ms, is the worksheet scan.

### Inflate

`InflateBenchmark` decompresses one ZIP entry per method with `System.IO.Compression.DeflateStream` and with the library's managed `InflateStream`, which every ZIP-based reader now uses:

| Entry | `DeflateStream` | `InflateStream` | Ratio |
|---|---:|---:|---:|
| XLSX `sheet1.xml` (6.7 MB → 34.4 MB) | 23.789 ms, 280 B | 13.741 ms, 200 B | 0.58 |
| XLSB `sheet1.bin` (3.7 MB → 15.9 MB) | 13.143 ms, 280 B | 9.734 ms, 200 B | 0.74 |
| String-heavy `sharedStrings.xml` | 10.571 ms, 281 B | 6.106 ms, 200 B | 0.58 |

Without prefetch a read is inflate plus parse, so the faster inflate converts directly; with prefetch the two overlap and the read is bounded by the slower stage, which is why the prefetch columns above moved little.

### Write path: managed deflate and CRC-32

XLSX and XLSB packages are written by the library's own ZIP writer. `CompressionLevel.Fastest`, the default, compresses with a managed fixed-Huffman deflate encoder and checksums with a managed CRC-32 that folds with PCLMULQDQ. Each piece alone, against what `ZipArchive` used before, on the same Ryzen 7 5700X:

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

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 43.03 ms | 0.15 ms | 16.78 KB |
| XLSM | 43.64 ms | 0.28 ms | 16.78 KB |
| XLSB | 28.01 ms | 0.90 ms | 16.34 KB |
| XLS | 12.95 ms | 0.25 ms | 21.48 KB |
| CSV | 19.31 ms | 0.22 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 7.89 ms | 0.03 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 55.51 ms | 0.58 ms | 1,111.1 | 1,000.0 | 333.3 | 15.25 MB |
| XLSB | 55.55 ms | 0.95 ms | 1,111.1 | 1,000.0 | 333.3 | 15.25 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 888 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 16.34 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.25 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan, with 333 Gen2 collections, and is still ~3.4x faster (55.51 ms vs. 186.05 ms). The same allocation holds for XLSB (15.25 MB vs. Sylvan's 17.38 MB, also ~12% less, 333 Gen2 collections), where matched-work time is ~17% ahead of Sylvan (55.55 ms vs. 65.22 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 8.922 ms | 4.02 MB |
| XLSB | 6.113 ms | 4.02 MB |
| XLS | 5.021 ms | 4.03 MB |
| CSV | 4.762 ms | 4.00 MB |

Relative ordering: CSV and XLS close at the front, then XLSB, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 6.113 ms vs. 5.572 ms, XLSX 8.922 ms vs. 8.337 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 14.559 ms | 15.74 KB |
| Native string columns → CSV | 5.098 ms | 416 B |
| Arrow `StringArray` → XLSX | 20.343 ms | 15.79 KB |
| Styled rows → XLSX | 5.411 ms | 15.92 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~16 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 11.22 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 10.11 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 9.65 ms | 4.12 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.9% less than the `class` baseline — and is ~5% faster than the `struct` target, since there's no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### ADO.NET bridge (`IDataReader`)

`DataReaderBenchmark`: the generated 50,000-row XLSX workbook (header plus four columns) read through [`SheetDataReader`](../guide/reading.md#bridge-to-adonet-idatareader), against iterating the same rows directly. The reader is held as `IDataReader`, the way `SqlBulkCopy`, `DataTable.Load` and Dapper hold it.

| Access | Mean | Ratio | Allocated |
|---|---:|---:|---:|
| Raw rows (`row[i].Value`, baseline) | 6.845 ms | 1.00 | 3.62 KB |
| `GetBytes` (text column only) | 7.039 ms | 1.03 | 4.31 KB |
| `GetValue` (every column) | 9.479 ms | 1.39 | 5.01 MB |
| Typed getters (`GetString`/`GetInt32`/`GetDateTime`/`GetDouble`) | 9.094 ms | 1.33 | 1.58 MB |
| `DataTable.Load` | 60.509 ms | 8.85 | 22.84 MB |

The bridge costs ~1.3–1.4x over reading rows directly. `GetValue` boxes every value (the 5 MB), and the typed getters allocate only the text column's `string`s. `GetBytes` copies UTF-8 without making a `string`, so it stays within ~3% of the raw read. `DataTable.Load` is dominated by `DataTable`'s own row storage: the same values cost 9.479 ms to read through `GetValue` alone.

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
| Small, plain | 0.208 ms | 8.84 KB |
| Small, encrypted, stream | 25.880 ms | 38.07 KB |
| Small, encrypted, stream, `VerifyEncryptedIntegrity` | 26.150 ms | 40.16 KB |
| Small, encrypted, in memory | 26.600 ms | 53.24 KB |
| Small, encrypted, open only | 26.378 ms | 32.54 KB |
| 300,000 rows, plain | 38.283 ms | 3.62 KB |
| 300,000 rows, encrypted | 65.485 ms | 29.53 KB |
| 300,000 rows, encrypted, open only | 26.235 ms | 28.92 KB |

Nearly all of the cost is a fixed ~26 ms paid when the workbook is opened: deriving the key through 100,000 SHA-512 iterations, which the format requires so that passwords are slow to brute-force. It does not depend on file size (open-only is 26.38 ms for 15 KB and 26.24 ms for 300,000 rows). After that, decrypting while reading costs ~3% over a plain read: 65.49 ms − 26.24 ms = 39.25 ms, against 38.28 ms unencrypted. On the small fixture, integrity verification and the in-memory path stay within ~3% of the plain encrypted read, inside the spread of the key derivation itself. The fixture is too small to show the extra pass that verification makes over a large stream.

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
| `xl_rows_next` | one row per call, serialized into a caller buffer | 37.93 ms | 3.71 MB |
| `xl_rows_read_all_blob` | whole sheet in one buffer | 42.74 ms | 21.06 MB |
| `xl_rows_read_all_decoded` | whole sheet as native row structs | 56.73 ms | 5.71 MB |
| `xl_parse_typed` | typed native columns | 35.73 ms | 3.71 MB |
| `xl_parse_arrow` | Arrow C Data Interface arrays | 36.93 ms | 3.72 MB |

The allocation column is **managed memory only**. `MemoryDiagnoser` cannot see the native blocks that `xl_rows_read_all_decoded`, `xl_parse_typed` and `xl_parse_arrow` return their data in.

Row by row, `xl_rows_next` costs ~1.4x the managed in-memory read of the same file (37.93 ms vs. 27.084 ms for `Xlsb_ExcelReader_Memory` under [In-memory real-data reads](#in-memory-real-data-reads)). Every cell crosses the ABI as text, so each binary XLSB number is formatted to UTF-8 on the way out. `xl_rows_read_all_blob` adds ~13% over it, because the whole sheet is buffered before it is copied out. `xl_rows_read_all_decoded` adds ~50%, because it allocates a native block per row. `xl_parse_typed` and `xl_parse_arrow` are slightly faster than a row-by-row read (~3–6%) and return columnar native buffers with numbers still binary, so they are the path to use for bulk typed loads.

`ChunkedParseBenchmark` compares `xl_parse_typed` on a whole sheet against draining the same sheet in batches through `xl_typed_reader_open`/`xl_typed_reader_next`. It uses the generated 50,000-row XLSX file, read from disk, with three nullable columns (string, int64, float64):

| Mode | Mean | Managed allocated |
|---|---:|---:|
| Whole sheet (`xl_parse_typed`) | 8.93 ms | 10 KB |
| Batches of 10,000 rows | 8.36 ms | 16.95 KB |
| Batches of 1,000 rows | 8.53 ms | 81.29 KB |

Batching costs nothing in time: both batch sizes are slightly faster than one whole-sheet table (~4–6%). A caller can bound its native memory to one batch at a time without paying for it.

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

