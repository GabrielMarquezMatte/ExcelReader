# Reader performance: where the time goes, and what bounds it

Measured 2026-09-15/16 on a **13th Gen Intel Core i7-1365U** (12 logical / 10 physical cores, P+E),
AVX2, no AVX-512, .NET 10. Corpus: `tests/ExcelReader.Benchmarks/Data/65K_Records_Data.*` — 65,536
rows, 14 columns, numeric and date heavy, with only ~5 KB of shared strings. The string-heavy section
uses `StringHeavyWorkbookGenerator`'s 65,536-row fixture instead.

**This is not the machine the README's tables come from.** Those are a Ryzen 7 5700X desktop. A
throttling mobile part with P and E cores is exactly the hardware `ARCHITECTURE.md` describes
discarding an earlier parallel-CSV measurement over, so nothing here should be compared against a
README number by absolute milliseconds — only ratios measured within one run are meaningful, and
anything destined for the README has to be re-run on the documented machine.

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

## Known open items

- The 8.9 ms of sheet parse above the inflate floor in the string-heavy corpus has not been
  decomposed. It is the largest unexamined block.
- `PrefetchStream` costs ~2.7 ms of handoff overhead, which is one full copy of the payload from the
  producer's pooled chunk into the consumer's buffer. Removing it means replacing the `Stream` seam
  with a buffer-exchange protocol — and that seam is where the decompressed-byte limit counters sit,
  so it is a trust boundary, not just a copy.
- `XlsxReader.Enumerator.TryParseIsoDate` still has its own ISO date parse, copying bytes to chars on
  the stack and calling `DateTime.TryParse`. `FastDate` now covers the same shapes and should replace
  it, but that path was not measured, so it was left alone.
- After `DrainFields`, rows-only enumeration of wide rows still runs ~1.75x the framing prototype.
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
