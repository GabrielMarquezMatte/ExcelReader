# Reader performance: where the time goes, and what bounds it

Measured 2026-09-15/16 on a **13th Gen Intel Core i7-1365U** (12 logical / 10 physical cores, P+E),
AVX2, no AVX-512, .NET 10. Corpus: `tests/ExcelReader.Benchmarks/Data/65K_Records_Data.*` — 65,536
rows, 14 columns, numeric and date heavy, with only ~5 KB of shared strings. The string-heavy section
uses `StringHeavyWorkbookGenerator`'s 65,536-row fixture instead.

**This is not the machine the published tables come from.** Those, in `benchmarks.md`, are a Ryzen 7
5700X desktop. A throttling mobile part with P and E cores is exactly the hardware `ARCHITECTURE.md`
describes discarding an earlier parallel-CSV measurement over, so nothing here should be compared
against a published number by absolute milliseconds — only ratios measured within one run are
meaningful, and anything destined for `benchmarks.md` has to be re-run on the documented machine.
The exception is "Second round" below, which was measured on the Ryzen.

The harness was a throwaway console app, not committed. Every table below comes from a single
interleaved round-robin run, where each variant runs once per round so thermal drift hits all of
them equally. **This machine drifts a lot** — the same sequential scan measured between 21 ms and
36 ms across runs — so read the ratios, not the absolute milliseconds. `min` across rounds is used
throughout as the cleanest estimate of true cost.

## The model

With `PrefetchDecompression` off, inflate and parse are serialized, so a read costs `I + P`. With it
on they overlap, so a read costs `max(I, P)` plus handoff overhead. Measuring the inflate on its own
gives `I`, and the two reads then solve for `P`. Predicted and measured agreed to within ~6% on
XLSX, which is what makes the rest of this document trustworthy.

## XLSX — parse-bound, barely

| | |
|---|---|
| I — inflate `xl/worksheets/sheet1.xml`, 6.7 MB → 34.4 MB | 34.4 ms |
| P — parse | 40.6 ms |
| read, prefetched | 44.9 ms |
| floor, `max(I, P)` | 40.6 ms |

Parse decomposed by adding one layer at a time (each without prefetch, so the times stay additive):

| layer | delta | share of P |
|---|---|---|
| L1 scan + emit into accumulator | 34.5 ms | 85.0% |
| L2 + enumerate cells | 0.5 ms | 1.3% |
| L3 + read `Cell.Type` | 0.2 ms | 0.5% |
| L4 + take the zero-copy text span | ~0 ms | ~0% |
| L5 + consumer's `TryParse`/`TryGetDateTime` | 5.8 ms | 14.4% |

Within L1, scanning is ~24.5 ms and emitting into the accumulator ~10.0 ms.

**The whole `Row`/`Cell` ref-struct surface costs ~0.7 ms of 40.6 (1.7%).** Enumeration, construction
and span access are effectively free. There is nothing to win there, and no performance reason to
revisit that design.

## XLSB — inflate-bound, 4% ceiling

| | |
|---|---|
| I — inflate `xl/worksheets/sheet1.bin`, 3.74 MB → 15.9 MB | 18.6–18.9 ms |
| P — parse | 17.0 ms |
| read, prefetched | 19.4–19.6 ms |

`P/I = 0.90`: the parser is already ~10% faster than the decompressor feeding it. An infinitely fast
XLSB parser would move the read from 19.5 ms to 18.7 ms.

**Do not optimize the XLSB parser.** `XlsbReader.Enumerator` is a single forward pass over TLV
records; there is no structural waste left in it, and none of the time is its.

## XLS — no floor, but the format caps the workload

Uncompressed OLE/CFB, so there is no inflate stage at all.

| | |
|---|---|
| read, stream | 11.0 ms |
| read, in-memory | 10.5 ms |
| of which scan + emit | 8.0 ms |
| container overhead (stream vs memory) | 0.4 ms (3–4%) |

BIFF8 is the one format where parallel parsing would convert 1:1 — no inflate wall, and it carries a
native row index (`INDEX` → `DBCELL` → row-block offsets) over seekable uncompressed data, so
partitioning would need no boundary guessing. It is still not worth building: BIFF8 caps at 65,536
rows × 256 columns, this corpus is already at the row cap, and it reads in 11 ms. There is no such
thing as a slow XLS.

## The pattern across the ZIP formats

Read time tracks inflated payload size, not parser quality.

| | inflated payload | read, prefetched |
|---|---|---|
| XLSX | 34.4 MB | ~45 ms |
| XLSB | 15.9 MB | ~19.5 ms |

2.16x the payload, 2.3x the time. XLSB is not faster because its parser is better; it is faster
because there are fewer bytes to inflate. For a caller who controls the producer, writing XLSB is a
larger win than anything available inside the readers.

## String-heavy XLSX — the shared-string prologue dominates

Different corpus, different balance. `sharedStrings.xml` is 2.79 MB → 7.5 MB inflated, 29% of the
payload, and it is loaded to completion before the first row is enumerated.

| stage | without prefetch | with prefetch |
|---|---|---|
| prologue — inflate + parse `sharedStrings.xml` | 34.4 ms | 19.8 ms |
| sheet stage | 49.7 ms | 28.5 ms |
| **total** | **84.1 ms** | **48.4 ms** |

The prologue is 41% of the prefetched read (54% for XLSB on the same corpus). The two stages are
strictly serial today and read two independent ZIP entries — two separate deflate streams, which
*can* inflate concurrently.

Upper bound if they were fully overlapped: `48.4 − min(19.8, 28.5)` = **28.5 ms, 1.69x**. See
"shared-string prefetch" below for why no cheap version reaches it.

Within the stages: the SST stage is inflate-bound (18.6 inflate vs 15.8 parse), so speeding up the
shared-string parser gains nothing. The sheet stage is parse-bound with **8.9 ms above its inflate
floor** — the one place in this document where ordinary parser work still converts 1:1. That 8.9 ms
has not been decomposed.

## CSV — the parser is at the floor; the cost is in the value layer

CSV is the only format with neither an inflate floor nor a format size cap, so every millisecond is
parse and every millisecond saved converts at any file size. That made it the best remaining
candidate. It turned out the parser has nothing left to give.

Generated 50,000-row corpus (2.32 MB), reading three columns:

| stage | |
|---|---|
| framing + emit + span | 2.32 ms |
| raw `IndexOfAny` byte sweep, no framing at all | 2.24 ms |

**1.03x of a bare byte sweep.** `CsvControlScanner` already does the right thing — one `Vector256`
pass producing a mask of delimiter/quote/CR/LF, kept stateful across refills through `Reset`/
`Continue`, so the vector setup amortizes over the whole buffer. (This is the same structural-scan
idea that lost badly in XLSX. It wins here because CSV fields are longer and the scanner is stateful
rather than restarted per short hop.)

**That comparison used a weak baseline.** A sweep that restarts `IndexOfAny` at every hit pays the
vector setup 200,000 times, so 2.24 ms is not a floor — it is a fourth implementation of the same
scan, and a bad one. The layered decomposition below re-measures against a real floor. The
conclusion survives, for a different reason: framing is not 97% of a floor, it is 13% of the read.

So the cost is downstream of framing:

| conversion | per call |
|---|---|
| `TryParse<int>` | 6.2 ns |
| `GetString()` | 17.6 ns |
| `TryGetDateTime` on a text date | see below |

### `TryGetDateTime` did not work on text dates at all

`TryGetDateTime` resolves the cell as an Excel **serial number** — it called `TryGetDouble` first. On
an ISO date string that is a `double.TryParse` that walks 27 characters and fails, so the method
returned `false` after ~69 ns of wasted work. Measured on the corpus: **0 of 50,000 date cells
succeeded**.

There was no other efficient path. `TryParse<T>` is constrained to `IUtf8SpanParsable<T>`, which
`DateTime` does not implement, so the only working option was `GetString()` followed by
`DateTime.Parse` — **163.7 ns per date**, 26x the int parse. A date written by this library's own
`CsvWriter` could not be read back by its own `CsvReader`; a test asserted that limitation rather
than the round-trip its name promised.

`FastDate` (`src/ExcelReader.Core/ValueObjects/FastDate.cs`) now parses ISO-8601 straight from UTF-8:
`yyyy-MM-dd`, optionally `T` or a space, a time, and up to seven fractional digits. A trailing zone
designator is rejected rather than guessed at.

| | before | after |
|---|---|---|
| text date cells resolved | 0 of 50,000 | 50,000 of 50,000 |
| cost per date | 163.7 ns (via `GetString`) | **40.4 ns** |

**4.05x, and the API stopped lying.** Order matters: the text-date attempt runs first, guarded by
`_hasNumber`, because reaching it through `TryGetDouble` costs a full failing double parse on every
date cell — measured at 85.4 ns with the wrong order against 40.4 ns with the right one. The guard
keeps the XLSX serial-date path untouched; its consumer-conversion layer measured 6.3 ms before and
after, inside the run-to-run band.

### Where the CSV read actually goes

A second harness split the read into layers that each add one stage to the same pass, so adjacent
rows subtract. Two groups, interleaved separately because the typed variants allocate a model per
row and their GC cost would otherwise land on their neighbours. Same 50,000-row corpus, same
machine, one run.

Framing prototypes, in memory, no reader machinery:

| layer | | |
|---|---|---|
| L0 `Vector256` masking only, popcount the hits | 0.102 ms | |
| L1 + drain bits, reload the byte to classify it | 0.384 ms | |
| L2 + drain bits, separate masks per class, no reload | 0.389 ms | **0.99x** |
| L3 + emit a 16-byte field descriptor | 0.562 ms | |
| L4 + emit a 32-byte `CellDesc`-shaped descriptor | 0.666 ms | |
| `IndexOfAny` restarted per hit (the old baseline) | 1.748 ms | |

All four field-emitting layers were gated on an identical checksum.

The shipped reader, same corpus, stream source:

| stage | | delta | share |
|---|---|---|---|
| rows, no cell access | 1.476 ms | | 30% |
| + four cell spans | 1.676 ms | +0.200 | 4% |
| + int column | 2.174 ms | +0.498 | 10% |
| + date column | 3.906 ms | +1.732 | 35% |
| + double column | 4.942 ms | +1.036 | 21% |

**The scan is 2% of the read and the date conversion is 35%.** Masking 2.3 MB costs 0.102 ms;
everything above L0 is per-field bookkeeping, not scanning. `Cell` materialization is 1 ns per cell.
Framing at 0.666 ms against the reader's 1.476 ms rows-only figure leaves ~0.8 ms of enumerator
overhead — `MoveNext`, `_acc.Reset`, refill and compaction bookkeeping — which is the largest
undecomposed block left in CSV after the conversions.

**Unexplained:** spans-only is faster from a stream than from an already-resident buffer
(1.676 vs 2.201 ms), but full conversion is faster from the buffer (4.942 vs 4.190 ms). Both
directions reproduced across four runs. No explanation; recorded rather than guessed at.

### The enumerator overhead: register pressure, not memory traffic — 1.16x on wide rows

Wide corpus (50,000 rows × 32 short text fields, 9.8 MB). The reader enumerating rows with no cell
access ran at **2.25x a standalone framing prototype** doing the same scan and writing the same
descriptors: 2.8 ns per field above framing. That overhead is per field, not per row — fitting the
narrow and wide corpora together gives ~5 ns fixed per row and ~2.6 ns per field.

The JIT disassembly of `TryParseSimpleRecord` showed, per field, a load and store of the scanner's
`_mask`, a load and store of `_col`, two loads and a store of the accumulator's `Count`, a load and
store of `_lastCol`, and `_cells` reloaded twice — all loop-carried through memory, where the
prototype keeps them in registers. Three attempts, each measured:

| attempt | result |
|---|---|
| copy the scanner struct into a local for the loop | **slower**, 2.25x → 2.52x |
| accumulator state and column in locals, committed once per record | **no change** |
| drain the vector mask in a call-free leaf function, one call per record | **1.16x** on wide |

The second attempt explains the first two. Its disassembly showed the locals **spilled to the
stack**: the loop still contained calls on cold paths (`GrowCells`, `ThrowColumnLimit`,
`NextScalar`, `SkipByte`), so every variable live across the loop needed a callee-saved register,
and x64 Windows has eight. The store-forwarding chain did not go away; it moved from the heap to
the stack. The prototype has no calls in its loop, so it gets the volatile registers.

`CsvReader.Enumerator.DrainFields` is that loop with no calls in it. It takes vector masks from
`CsvControlScanner.TryTakeMask`, emits delimiter-terminated fields, and closes the record on LF or
a CRLF inside one chunk. It hands back to the existing per-hit path on a quote, a bare CR, a CRLF
split across chunks, a full descriptor array, the column limit, or a buffer tail shorter than a
vector.

A first version drained only delimiters and returned to the per-hit path for every record end. It
won 1.14x on wide rows and **lost ~12% on the narrow benchmark**: a 47-byte record spans one and a
half chunks, so it paid two non-inlined calls per record and still took the slow path for the
newline. Closing the record inside the leaf and loading the next chunk from inside it made the
narrow case neutral.

In-process A/B, separate binaries for the old and new source, alternated three times, min of 12:

| | before | after |
|---|---|---|
| wide, rows only | 5.54–5.90 ms | **4.38–4.70 ms** (~1.28x) |
| wide, 32 cell spans (`ExcelReaderWide` shape) | 7.97–8.47 ms | **6.94–7.48 ms** (~1.16x) |
| narrow, 4 cell spans | 1.50–1.66 ms | 1.53–1.65 ms (neutral) |

BenchmarkDotNet, same machine, one run each: `ExcelReaderWide` 8.939 → 8.128 ms (the new run's
standard deviation was 0.91 ms, so read that as direction only), `ExcelReader` 3.774 → 3.609 ms,
`ExcelReaderAsync` 3.506 → 3.356 ms.

Correctness gate: 3,000 seeded random inputs weighted toward quotes, bare CR, CRLF and rows wider
than the initial descriptor array, each read both from memory and through a stream that returns
1–96 bytes per read so refills land at arbitrary offsets. 5,496,494 rows, byte-identical digests
between the old and new builds. The gate was checked by mutation: shifting the CRLF record end by
one byte changed both the row count and the digest.

### Against a One Billion Row Challenge workload

A scaled-down 1BRC: 10,000,000 rows of `station;temperature`, 413 stations, min/max/sum/count per
station. Every variant shares one open-addressing aggregation table, so the differences are reading
and parsing. All variants produced identical checksums, station counts and row counts. 12 threads on
the i7-1365U (2 performance + 8 efficiency cores).

Re-measured 2026-09-22, after `CsvStructuralScanner` (de49a45) and the `DrainFields` work. Min of
five rounds, three separate processes, all three within a few percent of each other:

| | per row | MB/s | linear extrapolation to 1B |
|---|---|---|---|
| hand-written parser, 1 thread | 43.5 ns | 420 | ~44 s |
| hand-written parser, partitioned by newline | 7.1 ns | 2,600 | ~7 s |
| ExcelReader rows, 1 thread | 60.0 ns | 305 | ~60 s |
| ExcelReader rows, partitioned by the caller | 11.2 ns | 1,630 | ~11 s |
| `Excel.AggregateCsvParallelAsync`, dop 1 | 66.2 ns | 275 | ~66 s |
| `Excel.AggregateCsvParallelAsync`, dop 12 | 11.4 ns | 1,600 | ~11 s |

The corpus is a rebuild, not the original bytes: 183 MB at 19.2 bytes per row against the first
run's 222 MB at 22, because the station names are synthetic and shorter. **What makes the two tables
comparable is the control**: the hand-written parser measured 42.3 ns then and 43.5 ns now. With the
baseline standing still, the reader's movement is the library's, not the corpus's.

| | 2026-09-21 | 2026-09-22 | |
|---|---|---|---|
| ExcelReader rows, 1 thread | 81.3 ns | 60.0 ns | 1.36x |
| ExcelReader rows, partitioned | 18.7 ns | 11.2 ns | 1.67x |
| the library's own parallel path | 241.7 ns (`ParseCsvParallelAsync<T>`) | 11.4 ns (`AggregateCsvParallelAsync`) | — |

The hand-written baseline is deliberately naive (`IndexOf` for the separator, a digit loop, FNV over
the key bytes) and is not a leaderboard entry; the winning entries add SWAR temperature parsing,
branch-free key hashing and memory-mapped input, and run on server hardware. So these rows position
the library against a plain specialized parser on the same machine, not against the leaderboard.

- The general reader costs **1.38x a naive specialized parser**, about 16.5 ns per row of generality
  on this shape (`Row`/`Cell` construction, the 32-byte `CellDesc`, generic `TryParse<double>`,
  double-to-fixed-point conversion). It was ~2x and ~39 ns before the structural scanner. The same
  1.38x holds partitioned, 11.2 ns against 7.1.
- Partitioning converts: 5.4x on 12 mixed threads for raw rows, 5.8x for the aggregation API, 6.1x
  for the hand-written parser.
- **`AggregateCsvParallelAsync` is now at parity with hand-partitioning** — 11.4 ns against 11.2 for
  a caller that splits the buffer itself. The earlier reading, that the library's own parallel path
  ran 13x slower than the caller-partitioned raw path, was measured against
  `ParseCsvParallelAsync<T>`, which allocates a model object and a string per row. The aggregation
  overload does not, and closes the gap. Its cost over the plain single-threaded reader is 1.10x
  (66.2 ns against 60.0), which is the delegate plus partition bookkeeping.
- So the missing public API for partitioned raw rows costs nothing on an aggregation workload.
  It would still matter for a consumer that must see each `Row` itself rather than fold it.

Cost of handing each `Row` to user code instead of an inline loop, same workload, one thread,
interleaved, 5,000,000 rows, identical checksums:

| | per row | |
|---|---|---|
| inline loop | 62.7 ns | 1.00x |
| delegate `(ref TState state, Row row)` | 68.3 ns | 1.09x |
| struct implementing a static-abstract processor interface, constrained generic | 76.8 ns | 1.23x |

That measurement chose the engine behind `Excel.AggregateCsvParallelAsync`: a
`CsvRowAction<TState>(ref TState state, Row row)` delegate, one accumulator per partition, and a
combine folded left to right in file order. It reuses the typed path's speculative partitioning.
The public surface has two shapes over that one engine: an `ICsvAccumulator<TSelf>` type
(`Add(Row)`, `Merge(TSelf)`, parameterless constructor), and a `CsvAggregation<TState>` object
holding `Seed`, `Accumulate` and `Combine`. Both take a `CsvParallelOptions` carrying
`DegreeOfParallelism`, `Reader` and `HeaderRow`.
A partition whose guessed start turns out wrong is read again from its confirmed start with a fresh
accumulator, and the first attempt's accumulator and any exception its callbacks threw are
discarded. That is why the callback may only mutate the accumulator it is handed. Partitions are
sized at a quarter of a worker's share, clamped to 1–64 MB, so accumulator creation and merging stay
per-partition costs, not per-64-KB ones.

Same 1BRC sample, 12 threads, interleaved, identical checksums (the machine was warmer than for the
first table, so compare within this table only):

| | per row |
|---|---|
| caller-partitioned raw rows, inline loop | 25.7 ns, 25.0 ns |
| three loose delegates (the first version of this API) | 31.5 ns, 28.5 ns |
| `ParseCsvParallelAsync<T>` | 209.5 ns |

About 1.15x the hand-partitioned loop — the delegate plus partition bookkeeping — and **7x faster
than the typed parallel path**. Unlike the hand-partitioned loop, it is correct for quoted fields
that span lines.

The two public shapes, measured after the surface was settled, same sample, interleaved, identical
checksums:

| | 12 threads | 1 thread |
|---|---|---|
| caller-partitioned raw rows, inline loop | 24.1 ns, 26.0 ns | — |
| `AggregateCsvParallelAsync(data, CsvAggregation<T>)` | 26.9 ns, 29.0 ns | 84.2 ns, 75.3 ns |
| `AggregateCsvParallelAsync<TAccumulator>(data)` | 34.4 ns, 37.5 ns | 117.6 ns, 95.1 ns |

The accumulator overload costs **~1.3x the aggregation overload** — more than the 1.15–1.23x a
direct interface call measured, because it goes through the same delegate engine and so pays the
delegate call and the interface call on every record. That is the price of keeping one engine; a
second engine specialized on the accumulator type would remove the delegate hop. The one-thread
fallback also runs above the inline loop's ~62 ns, since it enumerates through `MoveNextAsync`.

A prototype (not in the library) put a typed record between `Row` and the accumulator: a
`ref struct Reading : ICsvRecord<Reading>` with `static abstract bool TryParse(Row, out TSelf)`
holding the station as a `ReadOnlySpan<byte>`, and an accumulator whose `Add` takes the record, with
the record parameter declared `allows ref struct`. It compiles — a span taken from the `Row` may flow
into the `out` record — and through the same delegate engine, one thread, interleaved, identical
checksums:

| | per row | |
|---|---|---|
| `CsvAggregation`, `Row` | 68.2 ns, 68.5 ns | 1.00x |
| `ICsvAccumulator<TSelf>`, `Add(Row)` | 99.5 ns, 93.8 ns | 1.37–1.46x |
| record accumulator, `TryParse` into a `ref struct`, `Add(Reading)` | 67.5 ns, 69.6 ns | 0.99–1.02x |

Adding a layer made it faster than the shipped interface. The likely reason — inferred, not
confirmed in the disassembly — is that `TryParse` is a static method on a value type, so the JIT
specializes and inlines the `Row` work into the delegate, and the interface call then carries only a
span and a double instead of the whole `Row`.

The shipped overloads repeat this comparison directly: `CsvModelMap.FromAttributes` and
`CsvModelMap.Generated` bind the header once and call one column parser per column into a
`ref struct` model, the same shape the prototype used. A scaled-down 1BRC sample
(`station;temperature`, 413 stations, one decimal temperature, seeded), 5,000,000 rows, one thread,
interleaved, two runs, identical checksums:

| | per row (two runs) | ratio vs `CsvAggregation` |
|---|---|---|
| `CsvAggregation`, `Row` | 45.4 ns, 44.6 ns | 1.00x |
| record, static `TryParse` | 45.5 ns, 45.2 ns | 1.00–1.01x |
| reflection `CsvModelMap.FromAttributes` | 65.9 ns, 66.4 ns | 1.45–1.49x |
| generated `CsvModelMap.Generated` | 63.3 ns, 62.9 ns | 1.39–1.41x |

Both mapped paths cost more than the record and `CsvAggregation` shapes, and the generated map is
consistently a little cheaper than the reflected one — the same ordering the prototype found. The
per-column cost over the hand-written `TryParse` here is roughly 9–10 ns per column; with two
columns this says nothing yet about wide models, where that cost grows with every mapped column.

The constrained-generic struct, the textbook zero-cost shape, came out slower than the delegate.
The multithreaded version of this comparison was unusable: the same inline partitioned loop
measured 16.9 ns per row when run fourth and 23.8 ns when run seventh, from heat alone.

### The date parser: 1.49x from one vectorized pass over the round-trip form

The date column is 35% of the CSV read, so it was decomposed next. Two candidate cost centres had
never been separated: the digit loops, and building the `DateTime`.

| variant | | |
|---|---|---|
| current `FastDate` | | 1.00x |
| digits only, no `DateTime` built at all | | 1.19–1.23x |
| one leap check + month-table ticks | | 0.83–0.86x |
| branchless civil-days ticks (Hinnant) | | 0.76–0.81x |
| + SWAR 4-digit year on civil-days | | 0.82–0.86x |
| **one vectorized pass over the 27-byte form** | | **1.45–1.49x** |
| + separators validated in the same vectors | | 1.44–1.50x |

**Construction was not the problem.** `DateTime.DaysInMonth` followed by `new DateTime(y, m, d)`
runs two independent leap-year computations, which looked like the obvious waste; removing both
accounts for 19% at most, and every attempt to replace them with hand-written date arithmetic was
**slower**. The BCL's constructor is not worth reimplementing.

**88% of the cost was the digits.** `yyyy-MM-ddTHH:mm:ss.fffffff` is 27 bytes holding 21 digits at
fixed offsets. Two overlapping 16-byte loads cover it (bytes 0..15 and 11..26); one unsigned compare
per load validates every digit position at once against a constant mask; one shuffle per load packs
the digits so they fold in pairs instead of one dependent multiply-add per digit. The seven-digit
fraction loop, which the shape makes free, was the largest single piece.

Anything that is not exactly 27 bytes falls through to the existing parser. Two extensions were
measured and **not** built:

- Validating the separators inside the vectors instead of with five scalar compares: **no
  difference**. Those bytes are in L1 by definition and the branches are perfectly predicted.
- Using the same vectorized date extraction for the 19-byte `yyyy-MM-ddTHH:mm:ss` form: **no
  difference** (7.2 ns either way). Without the fraction there is not enough digit work to pay for
  the load and shuffle.

The masked validation has one trap worth recording. A first version ORed the digit mask and the
separator mask without restricting each to its own positions, which accepts a digit sitting in a
separator slot — `2024503-15T10:20:30.1234567` would have parsed as 2024-01-01. `FastDateTests`
carries one case per separator position for this.

### The double parser: nothing found, and the obvious hypothesis was wrong

| variant | | |
|---|---|---|
| current `FastDouble` | 12.3–15.4 ns | 1.00x |
| split at the dot, two tight digit loops | | 0.81–0.90x |
| BCL `double.TryParse` | | 0.27–0.31x |

`FastDouble`'s loop carries six branches per byte, and the corpus's value column varies in length,
so mispredicted loop exits looked like the cost. Running the same parser over values of uniform
length tests that directly — and it came out **slower**, not faster: 18.7 ns on uniform 8-byte
values against 12.3 ns on the real mixed corpus. Cost tracks digit count at roughly 2.4 ns per
digit, which is the signature of the serial `mantissa * 10 + d` dependency chain, not of
misprediction.

SWAR would break that chain, and this is the one place it would genuinely pay. It needs an 8-byte
load, and the value column averages 5.9 bytes, so the fast path would have to read past the end of
the field. Inside the reader those spans always point into a pooled buffer with slack, but
`FastDouble` would then depend on an invariant nothing enforces. Not built, for 2 ns.

### Parallel CSV converts — 3.46x, and it already ships

`Excel.ParseCsvParallelAsync<T>` partitions by byte range with `CsvBoundaryResolver` confirming row
starts. Typed models, allocating, own interleave:

| | |
|---|---|
| dop 1 (falls through to the sequential path) | 14.202 ms |
| dop 6 | 5.068 ms (2.80x) |
| dop 12 | 4.107 ms (3.46x) |

This is the one place in the library where parallelism converts, and it converts because CSV has no
inflate floor. It is only available on the typed path: `Row` and `Cell` are `ref struct`s and cannot
cross a thread boundary, so raw row enumeration hits the same wall as XLSX.

## Approaches measured and rejected

### SIMD structural index (simdjson stage 1) — 0.86–0.97x, rejected

A prototype replaced the scanner's repeated short `Span.IndexOf` calls with a bitmap of `<`, `>` and
quote positions (one bit per byte, AVX2), added quote-pair attribute extraction, and dropped the
`</row` pre-scan. Both scanners produced identical checksums over 917,504 cells.

It was **slower**, across three independent interleaved runs. Its locate-only core lost 11.07 ms to
12.59 ms even with index construction (2.12 ms) measured separately.

`Span.IndexOf(byte)` is already vectorized, and XLSX hop distances are short (~37 bytes per cell), so
a scan almost always resolves inside the first vector — one load, one compare, one tzcnt. The bitmap
adds 37.5% more memory traffic and a dependent mask load to save a branch that was already cheap.
AVX-512 would halve index construction but not the memory traffic, so it would not flip the verdict.

### Parallel intra-sheet parsing — 4.6x on the scan, 1.29x on the reader, not built

Partitioning an already-inflated sheet by byte range works and is exactly correct: checksums at
degree of parallelism 1/2/4/8/12/16 were byte-identical to the sequential scan. A raw `<` cannot
appear in XML text content, so the first name-boundary-checked `<row` at or after a chunk start is
unambiguously a row start — no confirm-then-reparse cascade like the parallel CSV path needs, and
rows carry no cross-row state.

The scan went from 32.7 ms to 7.1 ms at dop 12 (4.59x). The reader did not move. `P` only barely
exceeds `I`, so parallelising the scan takes `P` to 26.0 ms and the prefetched floor becomes
`max(38.8, 26.0)` = 38.8 ms — 1.29x. Parallelising the *entire* parse gives the identical 1.29x,
because inflate is then all that is left.

It would also only ever serve the typed-parsing layer: `Row`/`Cell` are `ref struct` and cannot cross
a thread boundary.

### Faster inflate — measured separately, no gain

libdeflate was measured against `System.IO.Compression` on this workload before this investigation:
no faster, and slower in places. libdeflate is whole-buffer only — no streaming API — so it forces
the entire sheet to be materialized before parsing starts and destroys the inflate/parse overlap
`PrefetchStream` exists to provide.

**The inflate floor does not move by swapping compression libraries.** Since XLSX and XLSB are both
bounded by it, this is the single most important negative result here.

### Lazy number parsing — a trade, not a win

`EmitScalarValueFast` runs `FastDouble.TryParse` on every number cell at emit time whether or not the
consumer asks for the value. Guarding both call sites (temporarily) to defer it:

| | eager | lazy | |
|---|---|---|---|
| nobody reads a number | 35.3 ms | 29.4 ms | saves 5.8 ms |
| consumer reads every number | 42.4 ms | 44.8 ms | costs 2.4 ms |

Both paths produced a byte-identical accumulator, so the lazy fallback is correct. Parsing at emit
time, while the value bytes are still hot in L1, is cheaper than parsing them later from
`Cell.TryParse`. Deferring skips no work for a consumer that reads numbers, and loses the locality.
Worth having only as an opt-in for consumers that skip numeric columns entirely — never as a default.

### Shared-string prefetch — measured at ~1.1x, not built

The consumer needs the shared-string table at row 1, and in a pull model the consumer's thread is the
thread parsing the sheet, so it blocks there and the sheet parse stops with it.

| version | gain |
|---|---|
| load the SST on a background task at open | ~1% |
| + resolve the shared index lazily at consumption | ~1% |
| overlap the two inflates only | **negative** (49.7 ms vs 48.4) |
| speed up the shared-string parser | 0% (that stage is inflate-bound) |
| incremental SST with a published watermark | **1.11x measured** |
| decouple the sheet parse from the consumer | 1.69x (upper bound, not measured) |

Overlapping the inflates alone loses because the sheet stage is already parse-bound at 28.5 ms
against 20.4 ms of inflate — its inflate is already hidden, so pre-inflating buys nothing and costs
scheduling.

#### The watermark variant, and why it falls short

Shared-string indices are assigned in first-use order, so a sheet written row by row references them
in roughly ascending order. That suggests parsing the SST incrementally on its own thread, publishing
a monotonic "resolved up to index N" watermark, and letting the sheet parse — still on the consumer's
thread — block only when it needs an index the SST has not reached. No producer/consumer split.

The ordering property is real. Measured over `StringHeavy.xlsx` (190,105 unique strings), the highest
index a row needs, by row decile:

| row decile | 0% | 10% | 20% | 30% | 50% | 70% | 90% |
|---|---|---|---|---|---|---|---|
| fraction of the table needed | 20% | 34% | 46% | 56% | 72% | 87% | 100% |

Row 1 alone needs only index 10. A prototype confirmed the consequence: across ~590,000 shared cells,
the sheet stalled exactly **once** — after that the SST parser stayed ahead for the whole file.

But that single stall costs 10.9 ms of the SST's 14.5 ms, hiding only 3.6 ms. The sheet reaches row 1
in microseconds and already demands a fifth of the table. A stall happens whenever
`table fraction / row fraction` exceeds the ratio between the two stages' total costs; in the first
decile that is 3.4 against a ratio of 2.32. The prologue is shortened, not removed.

Serial 46.5 ms → watermark 42.0 ms, **1.11x**. On the numeric corpus (5.9 KB table) it measured
0.98x — threading overhead with nothing to hide.

**In the real reader it would be worse.** The prototype's stage ratio was favourable: its sheet scan
is 33.6 ms against an SST parse of 14.5 ms that excludes inflate, a ratio of 2.32. The reader's real
ratio is 29.3 / 19.8 = 1.48, because its SST stage is inflate-bound at 18.6 ms. A lower ratio means
more stalling, not less. Scaling the measured stall gives ~14.9 ms of the 19.8 ms stage hidden away,
about 4.9 ms off a 48.4 ms read — ~1.1x again, from both directions.

Not worth building: ~1.1x on string-heavy workbooks only, in exchange for a background thread inside
the reader, the `MaxSharedStringBytes` limit being enforced on it, exception propagation across it,
and a public API change. The 1.69x ceiling remains real but sits behind the producer/consumer split,
which is the same constraint that caps parallel parsing — and the watermark, which looked like the
shortcut to it, delivers less than a sixth.

### Split character-class masks in the CSV scanner — 0.99x, rejected

`CsvControlScanner` ORs four `Vector256.Equals` results into one mask, and the record loop then
reloads `buf[stop]` to find out which class the hit belonged to. Keeping the four masks separate and
testing the bit against them removes that second load. It measured 0.389 ms against 0.384 ms — **no
difference**. Three extra `ExtractMostSignificantBits` cost exactly what an L1 hit costs, and the
byte was in L1 by definition, having just been loaded into the vector register.

### A CSV-specific narrow cell descriptor — 0.104 ms, not built

`CellDesc` is 32 bytes because it carries `Number`, `HasNumber` and `SharedIndex` for the binary
formats. CSV never sets any of them. A 16-byte descriptor measured 0.562 ms against 0.666 ms over
200,000 fields — real, and **2% of the read**. Not worth a second descriptor type and a second
`ToCell` path through `Row` and `Cell`.

### Restructuring `FastDouble` around the decimal point — 0.81x, rejected

`FastDouble` runs six branches per byte: dot seen, digit range, leading zero, digit counter,
overflow guard, scale increment. Locating the dot once with `IndexOf` and running two tight digit
loops removes all of that from the per-byte path. It measured **slower**: 17.0 ns against 13.9 ns.
The corpus's value column averages 5.9 bytes, so the `IndexOf` call and the slicing cost more than
the branches they eliminate. `FastDouble` is already 3.4x `double.TryParse` (13.9 ns vs 47.5 ns).

## What landed

Two changes, both verified against the full suite (1535 tests):

- **`ParseRowInWindow` replacing `EnsureRowBuffered` + `ParseRow`.** The old design located `</row`
  first and then walked the same bytes again looking for cells, traversing every row twice. Cells and
  the row end now come off one forward walk; the contiguity guarantee moved to a slow path that
  re-parses after a refill, which is reached only when a row genuinely straddles the buffer window.
  Worth 2.2 ms on the scan in isolation (1.10x).
- **`XlsxXml.ColumnIndex` unrolled to XFD**, the last valid column, with longer refs falling back to
  the original loop. 1.44x in a direct A/B over the 983,040 real `r="..."` values from the corpus
  (4.40 ns → 3.05 ns per ref), agreeing on every one.

Integrated effect: **~1.8 ms, about 5% of L1.** The components measure 3.4 ms in isolation; the gap
is expected, because the `</row` pre-scan was also warming the cache for the walk that followed it.
Microbenchmarks overstate.

A third change, from the CSV investigation: **`FastDate`, a UTF-8 ISO-8601 date parser**, wired into
`Cell.TryGetDateTime` as the first attempt for non-numeric cells. Text dates went from unreadable
(0 of 50,000) to 40.4 ns each, against 163.7 ns for the `GetString()` + `DateTime.Parse` workaround
that was previously the only option. It serves CSV, text dates in XLS, and XLSX `t="d"`.

A fourth: **`FastDate.TryParseRoundTrip`**, a vectorized path for the 27-byte
`yyyy-MM-ddTHH:mm:ss.fffffff` form — the shape this library's own `CsvWriter` emits. 1.45–1.49x over
the scalar parser, gated by a 50,000-value differential and by `FastDateTests` cases covering a
digit in each separator slot. Anything else falls through to the scalar parser unchanged.

### Second round (2026-09-21, Ryzen 7 5700X)

Measured on the published tables' machine, BenchmarkDotNet `--job Medium`, one run before and one
after, each whole benchmark class. Baseline was the previous commit built from a `git archive`.
Full suite at 1759/1759 afterwards.

- **`FastDate` reached the typed path.** `ColumnParserFactory.TryParseDateTimeText` tried
  `Utf8Parser.TryParse(…, 'O')` before the culture parse, and `TryParseDateOnlyText` went straight to
  the culture parse. Both now try `FastDate` first and keep the culture parse as the fallback.
  `CsvParserTests.TextDatesBindExactlyAsTheCultureParserWould` pins nine shapes to the old result.
- **XLSX `t="d"` cells use `FastDate`.** `XlsxReader.Enumerator.TryParseIsoDate` tries it first and
  keeps its char-copy `DateTime.TryParse` for shapes `FastDate` rejects. Years below 100 are rejected
  as before. No benchmark reads `t="d"` cells, so this one is unmeasured; it is the same parser the
  CSV rows below measure.
- **`IRowWriter.WriteUtf8(ReadOnlySpan<byte>)`**, a default interface method that decodes to a
  string. `CsvRowWriter` and `XlsxRowWriter` override it to copy the bytes straight through when they
  are valid UTF-8 and need no escaping, and fall back to `Write(string)` otherwise — always for XLSX
  shared strings. The C ABI and the Arrow writer now call it instead of building a string per cell.
  `WriteUtf8ParityTests` checks 14 values byte-for-byte against `Write(string)` on all three paths.
- **Styled rows format the style id with `Utf8Formatter`** into the row buffer instead of an
  interpolated string per row.

| | before | after | |
|---|---|---|---|
| `CsvParseBenchmark.ExcelParserSync` (typed, 50k rows) | 7.081 ms | 5.979 ms | 1.18x |
| `CsvParallelParseBenchmark.ConversionHeavy`, dop 1 | 1,372.6 ms | 981.9 ms | 1.40x |
| `CsvParallelParseBenchmark.ConversionHeavy`, dop 16 | 331.6 ms | 271.8 ms | 1.22x |
| `CsvParallelParseBenchmark.ConversionHeavyAggregate`, dop 1 | 1,217.1 ms | 845.9 ms | 1.44x |
| `CsvParallelParseBenchmark.ConversionHeavyAggregate`, dop 8 | 240.0 ms | 160.8 ms | 1.49x |
| `WritePathBenchmark.NativeStrings_Csv` | 10.14 ms / 21.36 MB | 4.881 ms / 384 B | 2.08x |
| `WritePathBenchmark.NativeStrings_Xlsx` | 27.16 ms / 21.38 MB | 21.053 ms / 18 KB | 1.29x |
| `WritePathBenchmark.ArrowStrings_Xlsx` | 31.49 ms / 21.38 MB | 26.087 ms / 18 KB | 1.21x |
| `WritePathBenchmark.StyledRows_Xlsx` | 10.07 ms / 7.65 MB | 7.814 ms / 18 KB | 1.29x |

Controls: `NarrowInt`, which has no dates, and Sylvan both stayed inside run-to-run noise (Sylvan
11.97 → 12.78 ms, in the unfavourable direction). Allocation on the typed paths did not change; the
model object per row is what they allocate. The published tables in `benchmarks.md` come from a
third, separate run after these changes.

### Writing doubles (2026-09-21, Ryzen 7 5700X)

`WriteBenchmark` had SpreadCheetah ~8% ahead of the XLSX writer. With compression off the writer
still took 13.6 ms of its 17.5, so the cost was generating XML, not deflate. Per cell, above an
empty row: string ~30 ns, int ~18 ns, double ~92 ns, date ~118 ns. Both of the last two end in
`Utf8Formatter.TryFormat(double)`, the general shortest-round-trip formatter.

`CellFormatter.TryFormatDouble` tries scales 10⁰–10⁴ and takes the first where `round(v·10ᵈ) / 10ᵈ == v`
exactly, for `0 < |v| < 1e9`. That integer, with a decimal point inserted, is the shortest form, so
the output is byte-identical. Anything else, including `-0`, falls through to `Utf8Formatter`. The
XLSX writer (numbers and date serials) and the CSV writer use it.

| per value | `Utf8Formatter` | `TryFormatDouble` | |
|---|---|---|---|
| short decimals (`1.5`, `45293.25`) | 61.6 ns | 15.1 ns | 4.07x |
| integral | 55.2 ns | 12.5 ns | 4.40x |
| serials with a time of day | 99.5 ns | 98.9 ns | 1.01x |

Correctness gate: 6.1M values (short decimals, random doubles across magnitudes, random bit
patterns, edges and their neighbours), zero mismatches. `DoubleFormatTests` keeps 600k of them.

End to end, A/B harness (one executable built against each `ExcelReader.Core.dll`, alternated
three times, min of 25):

| 50k rows | before | after | |
|---|---|---|---|
| XLSX | 17.50 ms | 12.68 ms | 1.38x |
| XLSX, shared strings | 17.08 ms | 12.11 ms | 1.41x |
| CSV | 6.71 ms | 4.23 ms | 1.59x |
| XLSB (control) | 7.65 ms | 7.63 ms | — |
| SpreadCheetah (control) | 15.66 ms | 15.52 ms | — |

BenchmarkDotNet afterwards: XLSX writing 12.261 ms against SpreadCheetah's 15.548 ms, CSV writing
4.624 ms against Sep's 7.075 ms. XLSB writes doubles in binary and is unaffected.

Not extended to serials with a time of day: `days + seconds/86400` has a factor of 3³ in the
denominator, so its shortest form is ~17 digits and no scale reaches it. Only a full shortest-float
algorithm (Ryu-class) would, for ~85 ns per such cell. Not built.

### Write-side decomposition of XLSB and XLS — nothing to take

Same per-column-type method as above, 50k rows:

| | XLSB, `Fastest` | XLSB, no compression | XLS |
|---|---|---|---|
| all four columns | 8.37 ms | 3.98 ms | 5.19 ms |
| empty rows | 1.79 ms | 0.94 ms | 0.53 ms |

XLSB is deflate-bound: more than half its time is compression, and serialization costs 10–17 ns per
cell. `prefetchWrite` already moves that deflate off the caller's thread. XLS costs 5–7 ns per cell.

**RK encoding for non-integral doubles — measured, reverted.** The XLSB writer emits `BrtCellRk` only
for integers. Adding the other two RK forms (truncated IEEE, integer/100), each accepted only when
`Biff12.Rk` decodes it back to the same bits, turned 1.5 and 45293.25 into 4-byte cells: the raw sheet
shrank 8% (3.75 → 3.45 MB) but the compressed one 1.3% (706 → 697 KB), and write time moved ~1%,
inside noise. Deflate was already squeezing those 8-byte doubles. Not worth the diff or the Excel
compatibility check an XLSB writer change needs.

### The native bindings: NativeAOT loses dynamic PGO

The C++ and Rust bindings read `65K_Records_Data.xlsx` in ~110 ms, against ~65 ms for the .NET
reader. One harness ran the same `xl_open_memory` + `xl_parse_typed` exports (14 typed columns) three
ways, plus the equivalent Core-only read:

| | ms |
|---|---|
| Core read, JIT | 70.9 |
| ABI, JIT (`delegate* unmanaged` over the managed assembly) | 84.1 |
| ABI, NativeAOT (P/Invoke into the published library) | 109.3 |
| ABI, NativeAOT, `IlcInstructionSet=x86-64-v3` | 107.3 |
| ABI, NativeAOT, `IlcPgoOptimize` (framework profile) | 109.4 |
| ABI, JIT, `DOTNET_TieredPGO=0` | 106.6 |

The C++ binding's published 110.4 ms matches the AOT row, so its `std::string` conversion is
negligible. The ABI layer costs ~13 ms. The rest, ~25 ms, is **dynamic PGO**: with it off, the JIT
lands exactly on NativeAOT. The instruction set is not the cause.

Where PGO acts, `DOTNET_TieredPGO=0` against the default, JIT:

| | PGO | no PGO |
|---|---|---|
| inflate only | 23.1 ms | 23.2 ms |
| rows only | 61.8 ms | 79.7 ms |
| + 14 cell spans | +1.3 ms | +9.5 ms |

The XLSX scan. `MethodJitInliningFailed` events (runtime keyword `0x1000`, read with an in-process
`EventListener`) showed what PGO inlines into the per-cell path that static heuristics refuse, as
"too many il bytes" or "unprofitable": `ReadCellOpenTagSpan`, `ScanCellAttributes`, `IsXmlSpace` (a
call per attribute byte), `XlsxXml.ColumnIndex`, `IsCellStart`, `FastDouble.TryParse` and
`CellAccumulator.Add`. `[MethodImpl(AggressiveInlining)]` on those seven:

| | before | after |
|---|---|---|
| rows only, JIT without PGO | 79.6 ms | 66.3 ms |
| ABI, JIT without PGO | 107.4 ms | 94.0 ms |
| **ABI, NativeAOT** | **109.9 ms** | **97.2 ms** (1.13x) |
| rows only, JIT with PGO | 60.0–62.2 ms | 60.8–62.8 ms (alternated three times; noise) |

Remaining gap to the PGO'd JIT is ~13 ms. Forcing `Row`'s column lookup inline changed nothing, so
what is left in the cell accessor is block layout and register allocation, which attributes do not
reach. A trace-derived `.mibc` fed to ILC (`MibcFile` items become `--mibc`) is the tool for that; it
needs `dotnet-pgo`, which is not on nuget.org.

Re-running the binding suites against a library published from this tree: the C++ full XLSX read
went from 110.4 to 98.0 ms and Rust's from 123.5 to 108.5 ms, and the C++ `write_sheet` from
136.6 to 100.1–105.3 ms, the last mostly from `TryFormatDouble`.

#### Second pass: XLSB, the ABI layer, and a static profile

Same harness, now over all three formats (14 typed columns, 65K rows), min ms:

| | Core, JIT | ABI, JIT | ABI, NativeAOT |
|---|---|---|---|
| XLSB | 29.2 | 45.3 | 52.2 |
| CSV | 12.2 | 26.9 | 33.7 |

**XLSB record reader.** `Biff12RecordReader.TryReadRecord` runs once per cell and was refused inline
without PGO. Forcing it inline moved the Core XLSB read from 33.3 to 27.1 ms without PGO, faster than
the PGO'd JIT's 28.8. No change with PGO.

**The ABI layer** costs more than the whole CSV read. A CPU-sampled trace (`dotnet-trace`,
`dotnet-sampled-thread-time`) and a counter around `BuildTable` split its ~15 ms on CSV into 0.6 ms
copying the caller's buffer in `xl_open_memory`, 2.5 ms copying the column chunks into the native
table, and ~12 ms of per-cell work. By column type, ABI minus Core: string ~33 ns/cell, date ~24,
double ~24, int ~8. Three changes, worth 2.5–3 ms with PGO and ~1 ms without:

- string cells append their UTF-8 bytes directly instead of `TryFormat` into a scratch buffer and
  then copying;
- date/time/bool columns call the `ColumnParserFactory` readers directly instead of through the
  `ExcelCellReaders` delegate fields;
- the validity bitmap is not touched until a column's first null, which backfills every earlier row;
  a column with no nulls ships no bitmap, so writing a bit per cell was wasted.

Building the table in native memory directly would remove the 2.5 ms copy but needs realloc growth
(which copies too) and a change to what `xl_free_table` frees. Not done.

**Static PGO profile.** `tests/ExcelReader.NativePgoTrainer` runs the typed parse over the three
fixtures; a trace of it becomes `src/ExcelReader.Native/pgo/excelreader.mibc`, which ILC now
consumes on every publish. ABI through NativeAOT, min ms:

| | no profile | full profile | profile without XLSB methods | JIT with PGO |
|---|---|---|---|---|
| XLSX | 95.8 | 85.5 | **84.9** (1.13x) | ~83 |
| XLSB | 42.9 | 44.1 | **38.4** (1.12x) | ~41.4 |
| CSV | 32.1 | 25.8 | **25.6** (1.26x) | ~24.4 |

The profile's XLSB counts made the XLSB parse slower; dropping those methods from it made XLSB faster
than both the unprofiled build and the JIT. The shipped profile excludes them.

**Typed mapping, not changed.** A generated `[ExcelSerializable]` map parses the 50K-row CSV in
5.44 ms against 3.99 ms for a hand-written loop doing the same conversions. The generator's fused
lambdas still invoke `ExcelCellReaders.String`/`DateTimeAuto`, which are delegate fields, so string and
date cells pay two indirect calls. Emitting direct calls through a hidden public helper class
measured ~3%, inside noise. Not worth new public surface.

A prototype of the bigger change, the generator emitting one static whole-row parse over
header-resolved column indices, showed the rest of the gap is not the mapping either. Stepping from the
hand-written loop to the shipped path, min ms of one clean run:

| step | ms | delta |
|---|---|---|
| hand-written loop | 4.05 | |
| + empty-cell checks and a column-index indirection | 4.35 | +0.30 |
| + the whole-row parse as a separate static method | 4.49 | +0.13 |
| + handing each model to a per-row callback | 5.04 | +0.55 |
| + handing it out through an `IEnumerable<T>` iterator | 5.40 | +0.36 |
| shipped generated map | 5.45 | +0.05 |

The per-column delegates and binding loop cost ~0.05 ms over a whole-row parse. The 1.36x is the
API's shape, one model per row delivered to the caller, plus the empty-cell checks any correct mapper
needs. A generator rewrite would recover ~1–5%. Not built.

## Known open items

- The 8.9 ms of sheet parse above the inflate floor in the string-heavy corpus has not been
  decomposed. It is the largest unexamined block.
- `PrefetchStream` costs ~2.7 ms of handoff overhead, which is one full copy of the payload from the
  producer's pooled chunk into the consumer's buffer. Removing it means replacing the `Stream` seam
  with a buffer-exchange protocol — and that seam is where the decompressed-byte limit counters sit,
  so it is a trust boundary, not just a copy.
- **Needs re-measuring.** This reading predates `CsvStructuralScanner` (de49a45, 2026-09-22), which
  moved the enumerator off `CsvControlScanner` and made `Reset` per buffer rather than per record,
  so the write barrier named below is likely already gone. The 1BRC re-run above found 1.36x on
  rows-only enumeration over the same period, so the gap quoted here is stale.
  After `DrainFields`, rows-only enumeration of wide rows still runs ~1.75x the framing prototype.
  What remains is per record rather than per field (the `MoveNext` → `TryParseRecordFromBuffer` →
  `TryParseSimpleRecord` call chain, `BeginRecord`, the write barrier from `CsvControlScanner.Continue`
  storing the buffer reference every record) plus the consumer side: `Row` indexer access costs
  about as much again as enumeration on the wide shape.
- Reading a CSV from an already-resident buffer is slower than from a stream for spans only, and
  faster for full conversion. Reproduced in both directions across four runs, unexplained.

### Resolved: the CSV gap against Sylvan was the benchmark not using this library's date API

`CsvReadBenchmark`'s accumulator parsed the date column with the BCL's
`Utf8Parser.TryParse(..., 'O')` rather than `Cell.TryGetDateTime`. Every competitor in that table
uses its own library's accessor — `GetDateTime`, `GetField<DateTime>` — so ExcelReader was the only
entrant not being measured through its own API. It had no choice: before `FastDate`,
`TryGetDateTime` could not read a text date at all.

Measured at 63.6 ns per date for `Utf8Parser` against 42.0 ns for `TryGetDateTime` on `FastDate`.
Switching the benchmark to the library's own accessor, then tuning the parser, under BenchmarkDotNet:

| `CsvReadBenchmark.ExcelReader` | |
|---|---|
| as published, date via `Utf8Parser.TryParse(…, 'O')` | 4.786 ms |
| date via `Cell.TryGetDateTime` (`FastDate`) | 4.149 ms |
| + `FastDate` tuned (below) | 3.966 ms |
| + vectorized round-trip path in `FastDate` | **3.381 ms** |

Those four figures come from four separate BenchmarkDotNet runs over a session during which the
machine warmed and cooled, so the 4.786 → 3.381 span is not a clean single-run comparison. The
drift-resistant statements are the isolated 1.45–1.49x on a date that is 35% of the read, which
predicts about 1.13x end to end, and the same-run competitor ratio below.

Same run, same corpus, 50,000 rows:

| | mean | ratio | allocated |
|---|---|---|---|
| ExcelReader | 3.381 ms | 1.00 | 368 B |
| ExcelReaderAsync | 3.309 ms | 0.98 | 440 B |
| Sylvan | 4.282 ms | 1.27 | 1,688,737 B |
| Sep | 7.115 ms | 2.10 | 4,024 B |
| CsvHelper | 26.027 ms | 7.70 | 15,073,424 B |

None of this goes to the README until it is re-run on the Ryzen 7 5700X.

**17.1% on the same method.** Sylvan, the control, did not move across the first two runs
(4.460 → 4.489 ms) — a third run reported 6.185 ms with a 1.56 ms standard deviation and is
discarded as an outlier. Against its stable figure ExcelReader now leads by roughly 12%, at 368 B
against 1,688,737 B allocated.

Caveats: six iterations on a throttling mobile CPU give wide error bars (±0.64 ms on ExcelReader), so
the lead over Sylvan sits near the margin. The 17.1% before/after on one method is the sounder
number, being the same code path measured three times. Everything here needs re-running on the
README's machine before publication.

#### Tuning the parser: two of three changes were worth it

Measured over the corpus's 50,000 real date strings, all variants agreeing on every value:

| | per date | |
|---|---|---|
| original (`scale /= 10` per fraction digit) | 31.3 ns | 1.00x |
| fraction scale from a lookup table | 27.2 ns | 1.15x |
| + two-digit fields unrolled | 23.1 ns | 1.35x |
| + SWAR four-digit year | 21.7 ns | 1.44x |

The first two shipped. The seven `scale /= 10` steps were a chain of dependent integer divisions, and
the five two-digit fields did not need a loop at all.

**The SWAR year did not ship.** It buys 1.07x on top — 1.4 ns — in exchange for masked-validation bit
arithmetic, and a first attempt at exactly that code had a bug: `(a | b) <= 9` does not test two
digits, since `2 | 8` is 10. The differential check against the other variants caught it. Six percent
is not worth that surface in a date parser.

### Still open

- A first attempt to attribute this gap in an interleaved in-process harness found no gap at all
  (7.120 ms vs 7.180 ms). That harness was the wrong tool: Sylvan allocates 1.69 MB per operation
  and triggers 265 Gen0 collections per 1000 ops, so its GC cost lands on whichever variant happens
  to be running. Interleaving compares variants with matching allocation profiles; against an
  allocating competitor it needs BenchmarkDotNet's process isolation.
- `FastDate` is now ~23 ns on a full round-trip timestamp. Whether Sylvan's date path is still ahead
  was never measured directly — it was inferred by subtracting totals, which is not evidence.
