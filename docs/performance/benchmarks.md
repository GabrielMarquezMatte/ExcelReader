# Benchmarks

Throwaway numbers go stale; these are regenerated from the benchmark suite in
`tests/ExcelReader.Benchmarks`. Live results: https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/

## Benchmarks

Benchmarks were run with BenchmarkDotNet v0.15.8 on Windows 10 (22H2), AMD Ryzen 7 5700X, .NET 10.0.11 (SDK 10.0.400). Generated-data benchmarks use 50,000 rows, except the string-heavy reads, which use 65,536. Raw results: [GitHub Pages benchmark dashboard](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/) — `tests/ExcelReader.Benchmarks/BenchmarkDotNet.Artifacts/` is `.gitignore`-excluded and never actually reaches GitHub, so a link to it 404s for every reader.

**Benchmark methodology.** In every table below, the **"Cell-by-cell read"** rows (including the CSV table's, and the per-format rows in "Real data reads" / "String-heavy reads") read ExcelReader's `cell.Value` — a zero-copy `ReadOnlySpan<byte>`, no decode or allocation — against each competitor's own idiomatic read call. For Sylvan, that's the ADO.NET-style `GetString(i)`, which is forced to materialize a UTF-16 `string`; Sylvan's API has no zero-copy accessor, so it cannot avoid that cost. These rows are therefore not matched work: part of the reported gap is "we parse faster" and part is "we skipped an allocation you were never offered a way to skip." Each affected benchmark class also has a `*_Materialized` sibling (calling `cell.GetString()`, the same UTF-16 materialization Sylvan pays) so the matched-work number is measurable — see `tests/ExcelReader.Benchmarks/Shared/BenchmarkAccumulators.cs`. **Those results are published under [Matched-work reads](#matched-work-reads-_materialized), and reading them is the honest way to judge the comparison:** treat the cell-by-cell ratios in the tables below as an upper bound, not a like-for-like number. The **"Typed row parsing"** / **"Typed record writing"** rows are unaffected — both sides already materialize real objects/strings there, so those comparisons are matched work as published.

### XLSX

Compares ExcelReader against established XLSX libraries on the same generated workbook shape.

| Scenario | ExcelReader | Sylvan | SpreadCheetah |
|---|---:|---:|---:|
| Cell-by-cell read | 11.233 ms, 11.13 KB | 37.160 ms, 1.89 MB | - |
| Cell-by-cell read async | 11.879 ms, 13.27 KB | - | - |
| Typed row parsing | 15.741 ms, 3.87 MB | 56.143 ms, 10.47 MB | - |
| Typed row parsing async | 15.352 ms, 3.88 MB | 60.307 ms, 10.48 MB | - |
| Typed row parsing, shared strings | 12.347 ms, 2.30 MB | - | - |
| Workbook writing | 12.261 ms, 4.02 MB | - | 15.548 ms, 15.84 MB |
| Workbook writing, shared strings | 11.890 ms, 4.06 MB | - | - |

ExcelReader is ~3.3x faster than Sylvan for raw XLSX reads, allocating ~174x less, and ~3.6x faster for typed parsing. For XLSX writing, ExcelReader is ~1.3x faster than SpreadCheetah while allocating ~3.9x less memory.

The suite also benchmarks OfficeIMO.Excel (XLSX and XLSB reads in `ReadBenchmark`, plus `ParseBenchmark`, `WriteBenchmark` and `XlsReadBenchmark`). Its numbers join these tables at the next run on the reference machine. OfficeIMO's typed parsing is hand-mapped from its `OpenDataReader`, because `RowsAs<T>` needs an `r` attribute on every row and cell, which the spec makes optional and ExcelReader's writer omits.

Reading a shared-strings XLSX workbook with typed parsing is ~22% faster than the inline-string sheet above (12.347 ms vs. 15.741 ms) and allocates ~41% less (2.30 MB vs. 3.87 MB) — each distinct string decodes once into the shared-string cache instead of once per cell occurrence.

### XLSB (BIFF12)

| Scenario | ExcelReader |
|---|---:|
| Cell-by-cell read | 5.280 ms, 13.23 KB |
| Cell-by-cell read async | 5.809 ms, 15.79 KB |
| Typed row parsing | 7.963 ms, 3.88 MB |
| Typed row parsing async | 8.281 ms, 3.88 MB |
| Workbook writing | 7.515 ms, 4.02 MB |
| Workbook writing, shared strings | 7.148 ms, 4.06 MB |

XLSB is the fastest generated Excel format in these results: raw reads are ~2.1x faster than XLSX reads, typed parsing is ~2.0x faster than XLSX parsing, and writing is ~1.6x faster than XLSX writing. The XLSB writer is also ~2.1x faster than SpreadCheetah on this benchmark while allocating ~75% less memory.

### XLS (BIFF8)

| Scenario | ExcelReader | Sylvan |
|---|---:|---:|
| Cell-by-cell read | 3.328 ms, 3.04 KB | 5.410 ms, 1,717.73 KB |
| Cell-by-cell read async | 3.977 ms, 3.11 KB | - |
| Workbook writing | 5.182 ms, 16.03 MB | - |

ExcelReader is ~1.6x faster than Sylvan for generated XLS reads while allocating ~565x less memory. The XLS writer is ~2.3x faster than the XLSX writer in this benchmark (`XlsWriteBenchmark`, 5.182 ms vs 12.145 ms in the same run). Most of its 16.03 MB is the benchmark's pre-sized 16 MB destination `MemoryStream`; the typed-record table below measures it at 4.03 MB like the other formats.

### CSV

| Scenario | ExcelReader | Sep | Sylvan.Data.Csv |
|---|---:|---:|---:|
| Cell-by-cell read | 3.428 ms, 368 B | 7.697 ms, 3.93 KB | 4.463 ms, 1.61 MB |
| Cell-by-cell read async | 3.657 ms, 440 B | - | - |
| Typed row parsing | 5.860 ms, 3.86 MB | 8.641 ms, 3.87 MB | 12.113 ms, 10.95 MB |
| Typed row parsing async | 5.503 ms, 3.86 MB | - | - |
| Row writing | 4.624 ms, 4.00 MB | 7.075 ms, 4.01 MB | 7.066 ms, 4.04 MB |

For raw CSV reads, ExcelReader is ~2.2x faster than Sep while allocating ~11x less, and ~1.3x faster than Sylvan.Data.Csv while allocating ~4,589x less (1.61 MB vs 368 B). For typed CSV parsing (the more common case — building actual records), ExcelReader is ~1.5x faster than Sep and ~2.1x faster than Sylvan.Data.Csv, with the lowest allocation of the group. For CSV writing, ExcelReader is ~1.5x faster than both Sep and Sylvan.Data.Csv, which are tied; the ~4 MB shown across all three is primarily the benchmark's pre-sized destination `MemoryStream`, not per-row writer state.

### Parallel CSV

`CsvParallel.ParseAsync<T>` and `CsvParallel.AggregateAsync` on generated files (8,000,000 rows of three ints for the narrow corpus, 4,300,000 rows of two strings, a date, two decimals and an int for the conversion-heavy one), by `degreeOfParallelism`. The machine has 8 physical / 16 logical cores.

| Dop | Conversion-heavy, typed | Narrow ints, typed | Conversion-heavy, `ref struct` aggregate |
|---:|---:|---:|---:|
| 1 | 930.0 ms, 670.98 MB | 531.2 ms, 245.32 MB | 834.4 ms, 916.53 KB |
| 2 | 638.8 ms, 679.85 MB | 409.5 ms, 254.38 MB | 442.1 ms, 1.24 MB |
| 4 | 371.2 ms, 679.95 MB | 238.2 ms, 254.44 MB | 228.2 ms, 1.36 MB |
| 8 | 286.0 ms, 679.83 MB | 199.0 ms, 254.34 MB | 162.3 ms, 1.53 MB |
| 16 | 283.9 ms, 680.05 MB | 207.4 ms, 254.83 MB | 120.1 ms, 1.64 MB |

The typed path scales ~3.3x on the conversion-heavy corpus and ~2.7x on narrow ints, both stopping at dop 8. Its ~680 MB is one materialized row object per record, and at dop 16 that costs 42,000 Gen0 plus 41,000 Gen1 collections.

The aggregate column folds each record into a per-partition accumulator through a `ref struct` model, so text columns stay as `ReadOnlySpan<byte>` and no row object is ever materialized. It scales better than the typed path — ~6.9x at dop 16 — because it is not fighting the allocator: at dop 16 it is ~2.4x faster than typed while allocating ~413x less (1.64 MB against 680.05 MB), with **zero** garbage collections at every degree. The remaining allocation is per-partition bookkeeping, not per-row.

Type conversion is the floor under both. A single-threaded pass that allocates nothing still costs 834.4 ms, so parsing bytes into `DateTime`/`decimal`/`int` — not allocation and not I/O — is what the parallelism is actually buying down.

Against [Sep](https://github.com/nietras/Sep)'s own `ParallelEnumerate` (all cores; 8,000,000 narrow rows, 3,000,000 conversion-heavy rows):

| Corpus | ExcelReader sequential | ExcelReader parallel | Sep sequential | Sep parallel |
|---|---:|---:|---:|---:|
| Narrow ints | 563.9 ms, 244.15 MB | 192.2 ms, 254.83 MB | 516.3 ms, 244.15 MB | 430.7 ms, 247.51 MB |
| Conversion-heavy | 657.3 ms, 467.31 MB | 183.6 ms, 474.62 MB | 802.4 ms, 467.31 MB | 624.1 ms, 468.37 MB |

Sequentially, Sep is ~9% faster than ExcelReader's typed parser on narrow ints, and ExcelReader is ~1.2x faster than Sep on the conversion-heavy corpus. In parallel ExcelReader is ~2.2x (narrow) and ~3.4x (conversion-heavy) faster than Sep's, while yielding rows in file order.

### Real data reads

This benchmark reads a real workbook exported in multiple formats.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 65.558 ms, 17.80 KB | 42.198 ms, 37.59 KB | 195.099 ms, 644.23 KB |
| XLSM | 64.966 ms, 17.80 KB | 45.505 ms, 37.58 KB | 203.457 ms, 644.30 KB |
| XLSB | 29.449 ms, 18.76 KB | 16.907 ms, 26.53 KB | 29.931 ms, 338.54 KB |
| XLS | 10.328 ms, 11.89 KB | n/a | 19.200 ms, 185.90 KB |
| CSV | 4.838 ms, 304 B | n/a | 11.043 ms, 35.75 MB |

On this real-data workload, ExcelReader is ~3.0x faster than Sylvan for XLSX, ~3.1x faster for XLSM, essentially tied for XLSB (~1.02x), ~1.9x faster for XLS, and ~2.3x faster for CSV — allocating ~36x less for XLSX, ~36x less for XLSM, ~18x less for XLSB, ~16x less for XLS, and ~123,000x less for CSV (304 B vs 35.75 MB). The prefetch column is the opt-in [`PrefetchDecompression`](../guide/reading.md#prefetch-decompression-xlsxxlsb) option; XLS and CSV are uncompressed, so it does not apply to them.

### In-memory real-data reads

The real-data benchmark also measures the in-memory path for workbook content loaded directly into memory, in the same run as the stream-based rows above (no cross-run comparison needed).

| Method | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Xlsx_ExcelReader_Memory | 65.378 ms | 0.373 ms | 6.95 KB |
| Xlsx_ExcelReader_Memory_Prefetch | 43.067 ms | 0.382 ms | 26.76 KB |
| Xlsm_ExcelReader_Memory | 65.311 ms | 0.369 ms | 6.95 KB |
| Xlsm_ExcelReader_Memory_Prefetch | 43.176 ms | 0.237 ms | 26.75 KB |
| Xlsb_ExcelReader_Memory | 27.747 ms | 0.080 ms | 8.64 KB |
| Xlsb_ExcelReader_Memory_Prefetch | 17.285 ms | 0.064 ms | 16.65 KB |
| Xls_ExcelReader_Memory | 10.069 ms | 0.216 ms | 11.89 KB |
| Csv_ExcelReader_Memory | 4.590 ms | 0.035 ms | 240 B |

`Csv_ExcelReader_Memory` is both faster (4.590 ms vs. 4.838 ms for `Csv_ExcelReader`) and allocates less (240 B vs. 304 B) than its stream twin.

`Xls_ExcelReader_Memory` is now within ~3% of its stream twin (10.069 ms vs. 10.328 ms for `Xls_ExcelReader`), with byte-identical allocation (11.89 KB both). In earlier runs the in-memory path was ~33% faster, thanks to `BiffCursor` caching the enclosing contiguous sector run; on this run the stream path has caught up.

### String-heavy reads

The real-data corpus above is mostly numbers and dates — its shared-string table is only 5 KB across ~910K cells, so it barely exercises shared strings at all. This benchmark uses a generated 65,536-row workbook with 8 text columns and ~190,000 distinct shared strings (a 7.5 MB uncompressed `sharedStrings.xml`), which is closer to a typical business export.

| Format | ExcelReader | ExcelReader, prefetch | Sylvan |
|---|---:|---:|---:|
| XLSX | 56.32 ms, 2.19 MB | 36.96 ms, 2.24 MB | 184.14 ms, 17.41 MB |
| XLSB | 38.97 ms, 2.19 MB | 26.18 ms, 2.27 MB | 66.89 ms, 17.38 MB |

Both formats now handle this well: ~3.3x and ~1.7x faster than Sylvan for XLSX and XLSB respectively, at roughly ~8x less memory in both cases, with no garbage collections in either configuration. XLSB's shared-string path previously materialized its table eagerly (~27 MB here); `ParseSharedStreaming` brought it in line with the XLSX streaming/pooling path, cutting allocation by ~12x on this workload.

### Matched-work reads (`*_Materialized`)

Every read benchmark class has a `*_Materialized` sibling that calls `cell.GetString()` per text cell — the same UTF-16 materialization Sylvan's ADO.NET-style API is forced to pay — while keeping `TryParse`/`TryGetDateTime` for numeric and date cells, exactly mirroring the competitor accumulator. These are the matched-work counterparts to the zero-copy rows above.

These now run in the same suite, on the same Ryzen 7 5700X machine, as the span-based rows above, so the ratios below are directly comparable — no cross-machine caveat needed.

Real-data workbook, per format:

| Format | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| XLSX | 66.18 ms | 0.58 ms | 27.27 KB |
| XLSM | 68.95 ms | 0.40 ms | 27.27 KB |
| XLSB | 30.55 ms | 0.12 ms | 28.23 KB |
| XLS | 12.05 ms | 0.14 ms | 21.36 KB |
| CSV | 19.39 ms | 0.27 ms | 35.71 MB |

Generated 50,000-row XLSX workbook:

| Scenario | Mean | StdDev | Allocated |
|---|---:|---:|---:|
| Cell-by-cell read, materialized | 13.44 ms | 0.08 ms | 1.58 MB |

String-heavy workbook (65,536 rows, ~190,000 distinct shared strings):

| Format | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
|---|---:|---:|---:|---:|---:|---:|
| XLSX | 82.38 ms | 1.66 ms | 1,000.0 | 857.1 | 285.7 | 15.26 MB |
| XLSB | 67.03 ms | 1.22 ms | 1,000.0 | 875.0 | 250.0 | 15.26 MB |

Reading the allocation columns against the tables above gives the honest shape of the tradeoff:

- **CSV real data:** 35.71 MB materialized against 304 B for the span-based read of the same file. Every CSV field is a distinct string, so nothing dedupes — this is where zero-copy reading earns its keep outright.
- **XLSB real data:** 28.23 KB, essentially cheap. That corpus repeats a small set of values, so the shared-string table dedupes and the reader's string cache materializes each distinct value once.
- **String-heavy XLSX:** 15.26 MB materialized, against **Sylvan's 17.41 MB on the same workload** — doing matched work here, ExcelReader now allocates ~12% *less* than Sylvan, with 286 Gen2 collections. The same holds for XLSB (15.26 MB vs. Sylvan's 17.38 MB, also ~12% less, 250 Gen2 collections). This was previously an inversion (ExcelReader allocated ~1.75x *more* than Sylvan here): the per-reader shared-string dedup cache was an unpresized `Dictionary<int,string>`, and at ~190,000 distinct values its resize/rehash churn (several of the largest resizes landing on the LOH) accounted for the entire gap — the strings themselves were never the problem, since both readers retain the same ~190,000 distinct instances. Replacing the dictionary with a `string?[]` indexed by shared-string index (sized exactly from the table's known count, no resizing) cut the allocation in half and cut wall-clock time by 14-16% on this benchmark too, since the churn was costing cycles, not just memory.

The takeaway is not that one column beats the other: it is that ExcelReader's headline read numbers come from a zero-copy path competitors do not expose, and when it does the same work as them, the gap narrows — and here, with the dedup cache fixed, no longer inverts even at high shared-string cardinality.

### Typed record writing

`WriteRecordsAsync` with `ExcelRecordLayout` (the header-plus-one-row-per-object API — see [Write typed records](../guide/writing.md#write-typed-records)) across all four formats, same 50,000-record source:

| Format | Mean | Allocated |
|---|---:|---:|
| XLSX | 14.539 ms | 4.02 MB |
| XLSB | 7.347 ms | 4.02 MB |
| XLS | 5.365 ms | 4.03 MB |
| CSV | 5.022 ms | 4.00 MB |

Relative ordering: CSV and XLS close at the front, then XLSB, then XLSX — the record-mapping layer adds negligible overhead over hand-written cell-by-cell writes.

### Native and Arrow string writes

`WritePathBenchmark`: 100,000 rows of four text columns written from UTF-8 buffers, the way the C ABI's `xl_write_typed` and `ArrowWriteExtensions.WriteRecordBatch` hand them over, plus 100,000 rows that each carry a row style.

| Scenario | Mean | Allocated |
|---|---:|---:|
| Native string columns → XLSX | 21.032 ms | 17.84 KB |
| Native string columns → CSV | 4.957 ms | 384 B |
| Arrow `StringArray` → XLSX | 26.738 ms | 17.88 KB |
| Styled rows → XLSX | 7.948 ms | 18.02 KB |

Text arrives through `IRowWriter.WriteUtf8`, which the XLSX and CSV writers copy through as bytes whenever the text needs no escaping, so none of these allocate per cell: the ~18 KB is per-workbook ZIP and part setup, which CSV does not have.

### Ref struct typed parsing (zero-copy)

A `ref struct` model (see [Parse into a ref struct](../guide/parsing.md#parse-into-a-ref-struct-zero-copy)) extends `ExcelParser.FromAttributes<T>`'s reflection/attribute-driven column mapping to `ref struct` targets, binding a `ReadOnlySpan<byte>` property directly to the cell's raw bytes instead of allocating a `string`. Same generated XLSX workbook, same 50,000 rows, same four columns — only the target type and binding strategy change:

| Target | Mean | Allocated |
|---|---:|---:|
| `class` (`ExcelParser<T>`) | 15.74 ms | 3.87 MB |
| `struct` (`ExcelParser<T>`) | 14.55 ms | 1.58 MB |
| `ref struct` + span binding (`ExcelParser<T>`) | 14.18 ms | 11.63 KB |

Parsing into a `ref struct` with a `ReadOnlySpan<byte>` text column removes essentially all per-row allocation — ~99.7% less than the `class` baseline — and is ~10% faster, since there's no per-row model allocation and no per-row `string` allocation for the text column. It is not AOT/trim-safe (reflection-based, same tradeoff as `ExcelParser.FromAttributes<T>`). It can be consumed with `foreach` or `await foreach` but not through `IEnumerable<T>`/`IAsyncEnumerable<T>`/LINQ — a `ref struct` element can't be boxed through those interfaces.

### Cold start

First use of `ExcelParser.FromAttributes<T>`/`ExcelRecordLayout.FromAttributes<T>` in a process pays a one-time reflection + `Expression.Compile` cost (16 launches, cold JIT, 200 rows):

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse | 34.88 ms | 28.38 KB |
| First typed record write | 19.82 ms | 84.27 KB |

This cost is paid once per type per process and cached thereafter — irrelevant for long-running services, worth knowing for CLI tools or serverless cold starts.

`ExcelParser.Build<T>` has no reflection at all — `configure` only allocates delegates and a `PropertyMap<T>[]` — so it skips this cost. `BuildWithAttributeFallback` still reflects for its attribute-driven half, so it pays close to the same cost as `FromAttributes`:

| Scenario | Mean | Allocated |
|---|---:|---:|
| First typed parse (`ExcelParser.FromAttributes<T>`) | 34.88 ms | 28.38 KB |
| First fluent parse (`ExcelParser.Build<T>`) | 28.73 ms | 31.18 KB |
| First fluent parse (`BuildWithAttributeFallback`) | 37.09 ms | 33.30 KB |

`Build` is ~18% faster than the reflection-based parser here; `BuildWithAttributeFallback` is slightly slower than reflection (~6%), since it runs the same `TypeMapper<T>.GetInfo()` path plus the fluent build on top.

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

