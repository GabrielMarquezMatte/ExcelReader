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
| Cell-by-cell read | 11.478 ms, 11.18 KB | 36.383 ms, 1.89 MB | 106.021 ms, 6.84 MB | - |
| Cell-by-cell read async | 11.210 ms, 13.31 KB | - | - | - |
| Typed row parsing | 22.947 ms, 3.87 MB | 55.957 ms, 10.47 MB | 103.699 ms, 5.69 MB | - |
| Typed row parsing async | 15.523 ms, 3.88 MB | 58.744 ms, 10.48 MB | - | - |
| Typed row parsing, shared strings | 13.048 ms, 2.30 MB | - | - | - |
| Workbook writing | 10.663 ms, 4.02 MB | - | 18.068 ms, 4.02 MB | 15.010 ms, 15.84 MB |
| Workbook writing, shared strings | 10.713 ms, 4.06 MB | - | - | - |

ExcelReader is ~3.2x faster than Sylvan for raw XLSX reads, allocating ~173x less, and ~9.2x faster than OfficeIMO. For typed parsing it is ~3.8x faster than Sylvan on the async path. For XLSX writing, ExcelReader is ~1.4x faster than SpreadCheetah while allocating ~3.9x less memory, and ~1.7x faster than OfficeIMO.

The synchronous typed parse measured 22.947 ms, against 15.523 ms for the async path that runs the same parser. This reproduces in separate processes: the pilot iterations run at ~15–16 ms, then settle at ~22–23 ms after the final tier-1 recompile, so the slowdown is in the code the JIT settles on, not run-to-run noise. Until that is fixed, the async figure shows what the parser itself costs.

OfficeIMO.Excel's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~16% faster than the inline-string sheet's async figure above (13.048 ms vs. 15.523 ms) and allocates ~41% less (2.30 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader | OfficeIMO |
|---|---:|---:|
| Cell-by-cell read | 5.222 ms, 13.28 KB | 10.375 ms, 6.70 MB |
| Cell-by-cell read async | 5.796 ms, 15.84 KB | - |
| Typed row parsing | 8.394 ms, 3.88 MB | - |
| Typed row parsing async | 8.307 ms, 3.88 MB | - |
| Workbook writing | 7.742 ms, 4.02 MB | - |
| Workbook writing, shared strings | 7.223 ms, 4.06 MB | - |

XLSB is the fastest generated Excel format in these results: raw reads are ~2.2x faster than XLSX reads, typed parsing is ~1.9x faster than XLSX parsing (async against async), and writing is ~1.4x faster than XLSX writing. Against OfficeIMO's XLSB reader, ExcelReader is ~2.0x faster while allocating ~517x less. The XLSB writer is also ~1.9x faster than SpreadCheetah's XLSX writer on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan | OfficeIMO |
|---|---:|---:|---:|
| Cell-by-cell read | 3.317 ms, 2.45 KB | 5.181 ms, 1,717.73 KB | 6.803 ms, 10,499.25 KB |
| Cell-by-cell read async | 3.271 ms, 2.52 KB | - | - |
| Workbook writing | 5.272 ms, 16.03 MB | - | - |

ExcelReader is ~1.6x faster than Sylvan for generated XLS reads while allocating ~700x less memory, and ~2.1x faster than OfficeIMO. The XLS writer is ~2.0x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 5.272 ms vs 10.766 ms in the same run). Most of its 16.03 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.03 MB like the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 3.510 ms, 440 B | 7.804 ms, 3.93 KB | 4.567 ms, 37.78 KB |
| Cell-by-cell read async | 2.873 ms, 512 B | - | - |
| Typed row parsing | 4.767 ms, 3.86 MB | 8.256 ms, 3.87 MB | 11.941 ms, 10.95 MB |
| Typed row parsing async | 5.121 ms, 3.86 MB | - | - |
| Row writing | 4.537 ms, 4.00 MB | 6.977 ms, 4.01 MB | 7.002 ms, 4.04 MB |

For raw CSV reads, all three libraries read text through spans. ExcelReader is ~2.2x faster than Sep while allocating ~9x less, and ~1.3x faster than Sylvan.Data.Csv while allocating ~88x less. ExcelReader's CSV read allocations here and in the real-data tables below come from a re-run after the reader's per-enumerator batch arrays moved to `ArrayPool`; allocation does not vary between runs, and the times are from the full run. For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.7x faster than Sep and ~2.5x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`CsvParallel.ParseAsync<T>` and `CsvParallel.AggregateAsync` on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate | Narrow ints, `ref struct` aggregate |
|---:|---:|---:|---:|---:|
| 1 | 914.8 ms, 670.98 MB | 523.2 ms, 245.32 MB | 833.8 ms, 848.53 KB | 490.2 ms, 843.39 KB |
| 2 | 618.3 ms, 683.36 MB | 392.3 ms, 257.93 MB | 453.4 ms, 1.28 MB | 261.9 ms, 1.30 MB |
| 4 | 340.0 ms, 683.48 MB | 211.7 ms, 257.98 MB | 214.8 ms, 1.38 MB | 137.3 ms, 1.46 MB |
| 8 | 272.5 ms, 683.33 MB | 191.6 ms, 257.85 MB | 150.0 ms, 1.57 MB | 95.0 ms, 1.64 MB |
| 16 | 293.9 ms, 683.53 MB | 183.4 ms, 258.36 MB | 110.2 ms, 1.74 MB | 71.8 ms, 1.76 MB |

The typed path scales ~3.4x on the conversion-heavy corpus, stopping at dop 8, and ~2.9x on narrow ints, where dop 16 adds only ~4% over dop 8. The conversion-heavy dop-16 mean carries a 71.1 ms standard deviation (median 257.9 ms), so read it as flat, not slower. Its ~683 MB is one materialized row object per record, and at dop 16 that costs 43,000 Gen0 plus 42,000 Gen1 collections.

The aggregate columns fold each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. They scale better than the typed path — ~7.6x (conversion-heavy) and ~6.8x (narrow) at dop 16 — because they are not fighting the allocator: at dop 16 the conversion-heavy aggregate is ~2.7x faster than typed while allocating ~394x less (1.74 MB against 683.53 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 833.8 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 522.9 ms, 244.15 MB | 187.8 ms, 258.37 MB | 505.1 ms, 244.15 MB | 418.4 ms, 248.03 MB |
| Conversion-heavy | 618.8 ms, 467.31 MB | 307.2 ms, 477.05 MB | 796.6 ms, 467.31 MB | 598.5 ms, 468.85 MB |

Sequentially, Sep is ~4% faster than ExcelReader's typed parser on narrow ints, and ExcelReader is ~1.3x faster than Sep on the conversion-heavy corpus. In parallel ExcelReader is ~2.2x (narrow) and ~1.9x (conversion-heavy) faster than Sep's, while yielding rows in file order. The conversion-heavy parallel row is bimodal in this run (84.8 ms standard deviation, median 344.4 ms), so its ratio is the least stable number in the table.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 64.322 ms, 17.85 KB | 42.850 ms, 37.61 KB | 192.140 ms, 644.23 KB |
| XLSM | 65.006 ms, 17.85 KB | 42.565 ms, 37.63 KB | 197.318 ms, 644.30 KB |
| XLSB | 28.851 ms, 18.80 KB | 15.357 ms, 30.02 KB | 29.582 ms, 338.54 KB |
| XLS | 9.943 ms, 11.30 KB | n/a | 18.635 ms, 185.90 KB |
| CSV | 4.900 ms, 376 B | n/a | 4.477 ms, 40.26 KB |

On this real-data workload, ExcelReader is ~3.0x faster than Sylvan for XLSX and XLSM, essentially tied for XLSB (~1.03x), and ~1.9x faster for XLS — allocating ~36x less for XLSX and XLSM, ~18x less for XLSB and ~16x less for XLS. On CSV, where both sides read text through spans, Sylvan.Data.Csv is ~9% faster, while ExcelReader allocates ~110x less. The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 65.139 ms | 0.756 ms | 7.00 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 43.111 ms | 0.158 ms | 26.76 KB |
| Xlsm_ExcelReader_Memory | 65.624 ms | 0.111 ms | 7.00 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 41.184 ms | 0.083 ms | 26.75 KB |
| Xlsb_ExcelReader_Memory | 28.785 ms | 0.017 ms | 8.69 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 17.019 ms | 0.025 ms | 17.01 KB |
| Xls_ExcelReader_Memory | 9.856 ms | 0.011 ms | 11.30 KB |
| Csv_ExcelReader_Memory | 4.787 ms | 0.004 ms | 312 B |

`Csv_ExcelReader_Memory` is both faster (4.787 ms vs. 4.900 ms for `Csv_ExcelReader`) and allocates less (312 B vs. 376 B) than its stream twin.

`Xls_ExcelReader_Memory` is within ~1% of its stream twin (9.856 ms vs. 9.943 ms for `Xls_ExcelReader`), with identical allocation (11.30 KB both). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; the stream path has since caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 57.34 ms, 2.19 MB | 34.81 ms, 2.21 MB | 178.08 ms, 17.41 MB |
| XLSB | 39.45 ms, 2.19 MB | 25.76 ms, 2.26 MB | 62.27 ms, 17.38 MB |

Both formats handle this well: ~3.1x and ~1.6x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 66.11 ms | 0.22 ms | 27.32 KB |
| XLSM | 65.55 ms | 0.24 ms | 27.32 KB |
| XLSB | 29.30 ms | 0.02 ms | 28.27 KB |
| XLS | 11.73 ms | 0.01 ms | 20.77 KB |
| CSV | 18.52 ms | 0.08 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 13.15 ms | 0.02 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 79.52 ms | 1.45 ms | 1,000.0 | 857.1 | 285.7 | 15.26 MB |
| XLSB | 58.89 ms | 1.08 ms | 1,111.1 | 1,000.0 | 333.3 | 15.26 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 376 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 28.27 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.26 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader allocates ~12% *less* than Sylvan, with 286 Gen2 collections, and is still ~2.2x faster (79.52 ms vs. 178.08 ms). The same allocation holds for XLSB (15.26 MB vs. Sylvan's 17.38 MB, also ~12% less, 333 Gen2 collections), where matched-work time is ~5% ahead of Sylvan (58.89 ms vs. 62.27 ms). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 11.648 ms | 4.02 MB |
| XLSB | 7.710 ms | 4.02 MB |
| XLS | 4.877 ms | 4.03 MB |
| CSV | 4.786 ms | 4.00 MB |

Relative ordering: CSV and XLS close at the front, then XLSB, then XLSX — the record-mapping layer adds little overhead over hand-written cell-by-cell writes (XLSB 7.710 ms vs. 7.742 ms, XLSX 11.648 ms vs. 10.663 ms).

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 20.478 ms | 18.02 KB |
| Native string columns → CSV | 5.088 ms | 416 B |
| Arrow `StringArray` → XLSX | 26.615 ms | 18.06 KB |
| Styled rows → XLSX | 7.577 ms | 18.20 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~18 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 22.95 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 14.91 ms | 1.59 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 13.48 ms | 11.72 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.7% less than the `class` baseline — and is ~10% faster than the `struct` target, since there's no per-row `string` allocation for the text column. (The `class` row is the synchronous typed parse whose tier-1 slowdown is described under [XLSX](#xlsx).) It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 33.70 ms | 28.43 KB |
| First typed record write | 18.38 ms | 84.25 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 33.70 ms | 28.43 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 29.38 ms | 31.37 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 36.44 ms | 33.49 KB |

`Build` is ~13% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slightly slower than reflection (~8%), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

Run the benchmarks locally:

```bash
dotnet run --project tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj --configuration Release -- --filter *
```

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
| vs. [calamine](https://github.com/tafia/calamine) (Rust) | XLSX | 108.5 ms | 279.7 ms | ~2.6x faster |
| vs. calamine (Rust) | XLSB | 66.6 ms | 91.9 ms | ~1.4x faster |
| vs. [DuckDB](https://github.com/duckdb/duckdb) `read_xlsx` (C++) | XLSX | 98.0 ms | 412.3 ms | ~4.2x faster |
| vs. [xlsxio](https://github.com/brechtsanders/xlsxio) (C) | XLSX | 98.0 ms | 497.2 ms | ~5.1x faster |
| vs. [xlnt](https://github.com/tfussell/xlnt) (C++) | XLSX | 98.0 ms | 2,388.5 ms | ~24x faster |

Writing the same 14 columns × 65,535 rows from row-shaped data, the C++ suite measures
`write_sheet` at 100.1–105.3 ms against DuckDB's 824 ms (~8x), libxlsxwriter's 1,177 ms (~11x),
xlsxio's 2,398 ms (~24x) and xlnt's 5,418 ms (~54x). The Rust suite measures 52.4 ms against
rust_xlsxwriter's 322.8 ms (~6.2x) over 7 columns. Caveats for both are in the binding READMEs.

DuckDB, xlsxio and xlnt do not read `.xlsb`, so those comparisons are XLSX-only. calamine is a fast,
well-optimized reader in its own right — the gap there is real but not the order of magnitude seen
against the C/C++ competitors.

