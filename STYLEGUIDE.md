# ExcelReader Style Guide

Code style rules. For build/PR/test process, see [CONTRIBUTING.md](CONTRIBUTING.md); for the shape of
the codebase, see [ARCHITECTURE.md](ARCHITECTURE.md).

## Priorities

1. **Correctness** — the code must be right. No shortcuts that sacrifice safety or accuracy.
2. **Performance** — optimize where it matters, but never at the cost of readability.
3. **Readability** — code is read far more than it is written.

---

## Ternaries

Use a ternary only when the entire expression fits on one line and both branches are single values — no function calls, no method chains, no nesting.

```csharp
// OK: two literal values
string label = isEmpty ? "none" : "some";

// OK: one field access each side
int col = rRef.IsEmpty ? _nextCol : XlsxXml.ColumnIndex(rRef);

// NOT OK: nested ternary
string result = a ? b ? "x" : "y" : "z";

// NOT OK: branches contain logic
int value = condition ? ComputeA() : ComputeB();

// Use if instead:
int value;
if (condition)
{
    value = ComputeA();
}
else
{
    value = ComputeB();
}
```

When in doubt, use `if`.

---

## Control Flow — Prefer Early Return

Avoid wrapping the happy path in a condition. Guard and exit early; keep the main logic at the lowest indentation level.

```csharp
// NOT OK: body is nested under the guard
private void AppendDecoded(ReadOnlySpan<byte> src)
{
    if (!src.IsEmpty)
    {
        EnsureValsCapacity(_valLen + src.Length);
        _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
    }
}

// OK: guard exits early
private void AppendDecoded(ReadOnlySpan<byte> src)
{
    if (src.IsEmpty)
    {
        return;
    }
    EnsureValsCapacity(_valLen + src.Length);
    _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
}
```

Always use braces on `if`, even single-line bodies. The linter enforces this.

---

## Scope Nesting — Maximum 3 Levels

Count every block that introduces indentation: `if`, `else`, `for`, `foreach`, `while`, `using`, `try`, `switch`. A function body is level 0. Nesting beyond 3 is a signal to extract a method.

```
level 0: function body
level 1: while (true)
level 2:   if (head.StartsWith("<row"u8))
level 3:     foreach (...)         ← maximum allowed
level 4:       if (...)            ← extract this into a private method
```

When you hit level 4, stop and ask: what is this inner block doing? Name it and move it.

---

## Method Bodies — Always Block Body

Every method, property getter/setter, and constructor must use a block body with braces. Expression bodies (`=>`) are forbidden on methods and properties, no matter how short.

```csharp
// NOT OK
public string SheetName => _sheets[_current].Name;
private static string Foo() => Bar();

// OK
public string SheetName
{
    get
    {
        return _sheets[_current].Name;
    }
}
private static string Foo()
{
    return Bar();
}
```

This rule exists to keep diffs readable and the line count honest. A one-liner expression body hides complexity behind a single line and makes the 70-line rule easy to game.

---

## Function Length — Maximum 70 Lines

Count every non-blank, non-comment line inside the function body (opening and closing brace included). If a function exceeds 70 lines, split it. Nested `ref struct` types with their own methods count each method separately.

70 is the review target. The build enforces a looser backstop through MA0051 (90 lines / 50 statements, see `.editorconfig`), so a method between 71 and 90 lines compiles but still fails review.

Do not game the limit by collapsing logic onto single lines. The limit exists to keep methods focused, not to reward line-golf. A method that packs three operations on one line to stay under 70 is worse than one that splits them across 71 readable lines.

Common splits:

- Parse one thing vs. emit the result of parsing
- Find a boundary vs. process the content within it
- `EnsureCapacity` extracted from the caller

---

## Naming

- Private fields: `_camelCase`
- Private static methods: `PascalCase`
- Local variables: `camelCase`
- Prefer full words over abbreviations (`index` over `idx`, `source` over `src`) unless the abbreviation is idiomatic in the domain (`col`, `buf`, `len`, `pos`)

### Public API naming must be symmetric across formats

The four formats are peers. A factory, option, or capability that exists for one format must use the
same name shape for all of them — `FromXls`/`FromXlsb`/`FromXlsx`/`FromCsv`, not three of those plus a
bare `From`. A caller who learns one format's entry point should be able to guess the others. When you
add a format-specific member, grep the sibling formats and match, or add the missing siblings in the
same change.

---

## Comments

**The default is no comment.** Code is expected to explain itself through naming and execution flow.
Reaching for a comment is the signal to look at the code again first: if a reader cannot follow the
logic, that is a defect in the code, not a gap in its documentation. Rename the variable, extract the
method, flatten the branch. Fix it there, and the comment becomes unnecessary.

A comment is justified only once that has been tried and the code still cannot carry the meaning —
because the reason lives outside the code and no amount of naming can bring it in:

- A constraint imposed by a format spec, a platform, or an external producer.
- An invariant a reader could silently break, where breaking it stays compiling and passes review.

```csharp
// OK: an invariant the type system cannot express
// Decoded output is never longer than its XML source, so src.Length bounds the flat buffer.

// NOT OK: describes what the code does
// Loop through each cell and parse its value.

// NOT OK: the code should have said this itself
// p is the position of the '<' that starts the row's closing tag
int p = ...;          // -> name it rowCloseTagStart and delete the comment
```

Never reference the task, the ticket, or the caller.

### Findings do not belong in code

Measurements, benchmark results, memory-footprint numbers, and the reasoning that picked one
approach over another are **not comments**. They go in [`docs/`](docs/).

This includes the whole class of "why it is this way" content: a throughput figure that justified a
buffer size, an allocation count that motivated pooling, a design that was prototyped and measured
and then rejected. Put it in a document, with the numbers, the machine, and the corpus. Then the
next person can re-run it and disagree with the evidence — which a comment asserting `// 2.4x faster`
never lets them do.

What may stay in the code is the constraint itself, stated without its evidence:

```csharp
// OK: the ceiling, not the benchmark that found it
// ponytail: 32 MB cap; sheets past that fall back to plain allocs.

// NOT OK: a finding, and it rots the moment anything around it changes
// Measured 11.3 ms vs 36.5 ms for Sylvan on the 65K-row corpus, so the span path stays.
```

### A comment must stay true, or it is worse than no comment

Comments carry maintenance cost. Two failure modes are specifically banned:

- **Do not describe a state the code is no longer in.** A header that says "these lookups are
  duplicated across the three readers" is actively misleading once the class exists precisely so they
  are not duplicated. When you refactor, reread the comments on what you touched.
- **Do not cite a file that can disappear.** Never point a comment at a plan, a spec, or a section
  marker ("see step Z4") as the authoritative explanation. Those are working artifacts: they get
  renamed, merged and deleted, and the comment then sends a reader to nothing, which is worse than
  silence. A committed document under [`docs/`](docs/) is different in kind — it is permanent, and
  pointing at one is how a finding stays out of the code. The test is whether the target is expected
  to still exist in a year, not whether a link is involved.

---

## Error Handling

Throw at trust boundaries: constructor arguments, public API parameters, malformed external data. Use `ArgumentOutOfRangeException.ThrowIfNegative` and similar .NET 6+ throw helpers. Do not add defensive checks for conditions that are impossible given the internal invariants.

Match the exception to whose fault it is. A malformed file is `InvalidDataException` (or
`ExcelLimitExceededException` when a configured cap is what rejected it) — never
`ArgumentException`/`ArgumentOutOfRangeException`, which tell the caller they made a mistake when the
input file is what is broken. If hostile input can reach a .NET throw helper, you are missing a
validation step upstream.

---

## Untrusted Input

Everything the parsers read is attacker-controlled: `.xlsx`/`.xlsb`/`.xls`/`.csv` arrive from uploads,
APIs, and mail attachments. These rules are not optional in reader code.

- **Bound every length, offset, count, and size read from the file before it drives an allocation, an
  index, or a loop bound.** A field claiming a multi-GB stream inside a 4 KB file is the canonical
  attack: validate against what the container can actually hold (`source.Length`, the enclosing
  buffer's length, the spec's fixed value) and reject before allocating, not after.

- **When you add a guard for one header field, apply it to every sibling field in that header.** The
  expensive bugs here have all been *inconsistency*, not ignorance — a validated sector count sitting
  three lines above an unvalidated stream size, with a comment on the first one already explaining the
  exact attack. If a field needs the guard, assume its neighbours do too.

- **Use `checked` for arithmetic on values derived from file bytes.** Silent `int` overflow turns a
  bounds check into a wrapped negative that sails past it.

- **Never let a malformed value silently change semantics.** A truncating cast that turns a length into
  a negative sentinel meaning "read everything" is a worse outcome than a thrown exception, because
  nothing reports it.

- **Validate before you allocate, not before you use.** Renting or allocating first and discovering the
  inconsistency during the walk still hands the attacker the allocation.

- New or changed parsing code needs a matching test in `ReaderLimitTests.cs` (forge the malformed
  header) or `FuzzTests.cs`. Assert the exception type and the limit metadata, not merely that
  something threw.

---

## Pooled Buffers

`ArrayPool<T>.Shared` is used throughout the readers. Misuse here is a correctness bug, not a
performance nit.

- Every `Rent` needs a matching `Return` on **every** exit path, including thrown exceptions. A method
  that rents, then throws mid-walk, must `try`/`catch`/`Return`/`rethrow` — a dropped buffer silently
  shrinks the pool for the whole process.
- Never `Return` a buffer that anything still references, and never touch a buffer after returning it.
  Ownership transfers must be explicit in a comment when a rented buffer outlives the renting method.
- Pooled arrays come back oversized and dirty. Bound reads by the logical length you tracked, never by
  `array.Length`.
- Buffers that live beyond one stack frame are returned in `Dispose`; guard against double-return with
  a flag, since `Dispose` may be called twice.

---

## Sync/Async Twins

Several readers keep paired sync and async methods on purpose: an `async` method pays for a state
machine on every row even when the buffer never needs refilling, so the hot path is duplicated rather
than shared.

The cost of that choice is that nothing in the compiler keeps the twins honest. So:

- A behavior fix to one twin **must** be applied to the other in the same change. Read both before you
  edit either.
- Behavior shared by both belongs in a sync helper that neither duplicates — the split should cover
  only the awaiting, not the logic around it.
- Do not add a third variant to work around a bug in one twin.
- Changes to a twinned path need a parity test asserting both produce identical output, including at
  buffer boundaries (a row straddling a refill, a cell forcing buffer growth) — that is exactly where
  the two implementations drift.
- A synchronous method does not take a `CancellationToken`.

---

## Tests and Benchmarks

Harness code is held to the same standard as library code, because a broken harness reports success.

- **Fail loudly, never silently.** A fixture that was never built, a setup that never ran, an input
  that parsed to nothing — assert it. The dangerous case is not the harness that crashes; it is the one
  that measures zero work and publishes a number. Do not rely on the code under test to reject an empty
  input, because some formats accept it.
- **Register new benchmarks with their setup.** When a class targets its `[GlobalSetup]` at specific
  methods, a new `[Benchmark]` absent from every target list runs against an unbuilt fixture.
- **Benchmarks against another library must compare matched work, or say plainly that they do not.**
  If our side reads a zero-copy span while the competitor's API forces a string allocation, that gap is
  partly "we are faster" and partly "we skipped work they cannot skip". Publish a matched-work sibling
  next to it. An unlabelled comparison that flatters us is a defect.
- **State the machine for every published number**, and never derive a ratio across results measured on
  different hardware. Allocation figures are deterministic and comparable across machines; timings are
  not.
- Name tests for what they actually cover. A fixture hand-authored in code is not a "real world" file,
  and calling it one hides the gap it was supposed to close.
- Assert real behavior — exact values, exception types, limit metadata. A test that calls a method and
  asserts no throw documents nothing.

---

## Performance Notes

These rules apply only when a function is on a measured hot path (cell parsing, buffer management):

- Prefer `ReadOnlySpan<byte>` over `string` for text that does not need to be allocated.
- Use `ArrayPool<T>.Shared` for buffers that live beyond one stack frame. Return them in `Dispose`.
- `ref struct` types stay on the stack; use them for row/cell value types.
- Avoid LINQ in hot paths; use `IndexOf` on spans which benefits from SIMD.
- Split cold paths (growth, throw helpers) into `[MethodImpl(MethodImplOptions.NoInlining)]` methods so
  they do not bloat the hot caller's IL and cost it inlining.
- Optimize against a measurement, and record what it showed in [`docs/`](docs/) — never in a comment.
  An optimization with no number behind it is a guess that costs readability. Caches in particular are
  a trade, not a win: document the workload where the cache pays and the one where it does not, with
  the numbers that showed it.

Outside the hot path (workbook loading, shared-strings init), ordinary allocations are fine.
