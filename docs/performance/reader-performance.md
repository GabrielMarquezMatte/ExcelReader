# Reader performance: where the time goes, and what bounds it

Measured 2026-09-15/16 on a Ryzen 7 5700X (8C/16T, 12 logical cores visible to the harness), AVX2,
no AVX-512, .NET 10. Corpus: `tests/ExcelReader.Benchmarks/Data/65K_Records_Data.*` — 65,536 rows,
14 columns, numeric and date heavy, with only ~5 KB of shared strings. The string-heavy section uses
`StringHeavyWorkbookGenerator`'s 65,536-row fixture instead.

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

### Shared-string prefetch — no cheap version reaches the 1.69x

The consumer needs the shared-string table at row 1, and in a pull model the consumer's thread is the
thread parsing the sheet, so it blocks there and the sheet parse stops with it.

| version | gain |
|---|---|
| load the SST on a background task at open | ~1% |
| + resolve the shared index lazily at consumption | ~1% |
| overlap the two inflates only | **negative** (49.7 ms vs 48.4) |
| speed up the shared-string parser | 0% (that stage is inflate-bound) |
| decouple the sheet parse from the consumer | 1.69x |

Overlapping the inflates alone loses because the sheet stage is already parse-bound at 28.5 ms
against 20.4 ms of inflate — its inflate is already hidden, so pre-inflating buys nothing and costs
scheduling. The entire prize sits behind a producer/consumer split, which is the same constraint that
caps parallel parsing.

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

## Known open items

- The 8.9 ms of sheet parse above the inflate floor in the string-heavy corpus has not been
  decomposed. It is the largest unexamined block.
- `PrefetchStream` costs ~2.7 ms of handoff overhead, which is one full copy of the payload from the
  producer's pooled chunk into the consumer's buffer. Removing it means replacing the `Stream` seam
  with a buffer-exchange protocol — and that seam is where the decompressed-byte limit counters sit,
  so it is a trust boundary, not just a copy.
- CSV was not measured in this investigation; it already ships a parallel path.
