# docs

Evidence behind decisions in this codebase: benchmark results, memory-footprint measurements,
approaches that were prototyped and rejected, and the reasoning that picked one design over another.

[`STYLEGUIDE.md`](../STYLEGUIDE.md#comments) keeps this content out of the source. A comment asserting
`// 2.4x faster` cannot be checked, cannot be re-run, and is wrong the moment anything around it
changes. A document with the numbers, the machine and the corpus can be reproduced and argued with.

## What belongs here

- A measurement that justified a constant, a buffer size, a cache, or a data structure.
- A design that was tried and abandoned, with the numbers that killed it. **These are the most
  valuable documents here** — they are what stops the same idea being rebuilt next year.
- A performance ceiling and what it is bounded by.

- Usage documentation under [`guide/`](guide/), split out of the root README so the NuGet landing
  page stays short. This is prose for a caller deciding *how* to use the library, not a substitute
  for the `///` comments that document *what* each member does.

## What does not

- Plans, specs and task breakdowns. Those are working artifacts with a short life; they do not get
  committed, and the style guide forbids pointing a comment at one.
- Reference documentation of the public surface. That is `///` XML doc comments, shipped in the
  package — never duplicated in `guide/`, which would then drift.
- Anything the code itself should be saying through naming and flow.

## Writing one

State the machine, the corpus and the tool, so the numbers can be reproduced and disagreed with.
Record what was measured, not only what was concluded — including the runs that came out flat or
negative. Say explicitly what was *not* measured; an untested assumption presented as a result is
the one failure mode that makes a document worse than no document.

Where a finding constrains the code, the code may state the constraint on its own terms (a cap, an
invariant) without restating the evidence for it.

## Index

- [`performance/reader-performance.md`](performance/reader-performance.md) — where read time goes in
  each format, the ceiling on each, and the optimizations measured and rejected.
- [`performance/benchmarks.md`](performance/benchmarks.md) — throughput and allocation against
  Sep, Sylvan, OfficeIMO and SpreadCheetah.
- [`guide/`](guide/) — usage documentation: [reading](guide/reading.md),
  [parsing](guide/parsing.md), [writing](guide/writing.md), [csv](guide/csv.md),
  [encryption](guide/encryption.md).
