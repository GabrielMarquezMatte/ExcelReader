# Architecture

A map of how the codebase fits together, not a manual. Start here, then follow the file/type names
into the source — the code comments carry the detailed reasoning.

## The four format families

Each format has its own `Reader` and a writer implementing `IWorkbookWriter<TSheet>`
(`src/ExcelReader.Core/Writer/IWorkbookWriter.cs`):

| Format | Reader | Writer | Sheet/row writer |
|---|---|---|---|
| XLSX | `XlsxReader` | `XlsxWorkbookWriter` | `XlsxSheetWriter`/`XlsxRowWriter` |
| XLSB | `XlsbReader` | `XlsbWorkbookWriter` | `XlsbSheetWriter`/`XlsbRowWriter` |
| XLS  | `XlsReader`  | `XlsWorkbookWriter`  | `XlsSheetWriter`/`XlsRowWriter` |
| CSV  | `CsvReader`  | `CsvWorkbookWriter`  | `CsvSheetWriter`/`CsvRowWriter` |

CSV has one extra layer: `CsvWriter` is the low-level RFC4180 writer (buffered rows straight to the
stream, no sheets/styles/shared-strings machinery); `CsvWorkbookWriter` adapts it to the shared
`IWorkbookWriter<CsvSheetWriter>` contract, exposing exactly one sheet.

On top of all four readers sits the typed-parsing layer (`src/ExcelReader.Core/Parser/`):
`ExcelParser<T>` (reflection/attribute-driven, allocates a model per row) and `RefParser`
(binds a `ref struct` model directly to `Cell.Value` spans — zero allocation for the container and,
for span-typed columns, for the values too). Both consume `Row`/`Cell` from any reader uniformly.

## Shared plumbing

Reader internals that would otherwise be duplicated four times over live in one place:

- **`CellAccumulator`** — pooled per-row cell storage (raw decoded UTF-8 text + a `CellDesc[]`
  describing each cell's column/type/style/offset). Used by every format's enumerator. Also hosts
  the shared BIFF numeric-error-code → text lookup (`#DIV/0!` etc.) for XLS/XLSB.
- **`PooledStreamRowEnumerator`** — abstract base centralizing the pooled-buffer lifecycle
  (`BufferedStreamCursor` + `CellAccumulator`, `Fill`/`FillAsync`/`Ensure` wrappers) that every
  format's enumerator subclasses. `MoveNext`/`MoveNextAsync` stay concrete per format — this only
  removes the buffer-lifecycle boilerplate, not the parsing itself.
- **`BufferedStreamCursor`** — the refill/compact-or-grow cursor behind the XLSX/XLSB/CSV
  forward-only stream enumerators. Has a second constructor for the in-memory-ZIP path that wraps an
  already-fully-decompressed `ReadOnlyMemory<byte>` instead of a `Stream` (`Eof = true` immediately,
  no refills).
- **`WorkbookLookups`** — small lookups that were once duplicated identically across readers: sheet
  name→index, sheet-index bounds checks, date-style flags, shared-string offsets, and (ZIP formats
  only) worksheet entry resolution plus prefetch/limit-counting stream composition. Takes arrays as
  parameters rather than requiring a shared interface, since each format's backing arrays differ in
  shape.
- **`LimitChecks`** — the DoS/resource-limit guards (`ExcelReaderOptions`/`CsvReaderOptions`), including
  the single buffer-growth-cap function (`NextBufferSize`) every pooled buffer in the stack grows
  through, so one limit policy governs all of them consistently.

## Encrypted OOXML workbooks

A password-protected `.xlsx`/`.xlsb` is not a ZIP: it is an OLE/CFB container holding an
`EncryptionInfo` descriptor and an `EncryptedPackage` (the real ZIP, encrypted). `src/ExcelReader.Core/Crypto/`
turns one back into a stream the ordinary readers consume:

- **`CfbContainer`** — the CFB parse, factored out of `XlsCompoundFile` so one parse can yield
  *named* streams. `XlsCompoundFile.OpenWorkbook` is now a thin wrapper over it, which is why the XLS
  path kept its behavior and its fuzz coverage unchanged.
- **`EncryptionDescriptor`** — parses `EncryptionInfo` and dispatches: version 4.4 is agile
  (AES-CBC, 4096-byte segments, XML descriptor); 3.2/4.2 (ECMA-376 standard, AES-ECB/SHA-1, binary
  descriptor) is recognized but rejected with `UnsupportedScheme` rather than implemented, for lack
  of a real fixture to verify a derivation against; anything else is rejected outright. This is the
  one parser in the codebase that uses `XmlReader` rather than hand-rolled scanning; the file header
  explains why.
- **`AgileKeyDerivation`** — password to key, plus the verifier check that distinguishes "wrong
  password" from "corrupt file". (A `StandardKeyDerivation` counterpart is future work, gated on
  real standard-encryption fixtures — see `tests/ExcelReader.Tests/data/encrypted/README.md`.)
- **`DecryptedPackageStream`** — a read-only *seekable* `Stream`. Agile derives a per-segment IV
  from the segment index, so `ZipArchive` finds its central directory without the package ever
  being materialized. One 4 KiB segment is cached, which is enough because ZIP reads are sequential
  within an entry.

`XlsxReader` and `XlsbReader` are untouched by any of this: they receive a stream that happens to
decrypt. Writing encrypted workbooks is not supported — which also means there is no round-trip
check, so the fixtures in `tests/ExcelReader.Tests/data/encrypted/` (paired with plaintext produced
by an independent implementation) are the only correctness oracle for decryption.

## The `excelreader` CLI

`src/ExcelReader.Cli/` is a thin `dotnet tool` shell (`excelreader`) over Core's public API — it
parses no spreadsheet bytes of its own. It splits in two on purpose: `Commands` is a one-line-per-
command adapter whose XML doc comments ConsoleAppFramework's source generator turns into argument
parsing, routing and `--help`, while `CliCommands` holds the bodies as plain functions over explicit
writers. The split keeps the tested surface free of the framework's static output hooks, so the CLI
tests run in parallel like every other test class. ConsoleAppFramework is compile-time only
(`PrivateAssets`), so the published tool's only *runtime* dependency besides ExcelReader.Core is
Spectre.Console, used for `sheets`/`schema`'s tables and `convert`'s progress spinner on a real
terminal.

Rendering follows `Commands`/`CliCommands`'s own split, one level further: `Commands` picks plain vs.
interactive per call (`Console.IsOutputRedirected`/`IsErrorRedirected`), so a script gets the exact
same tab-separated text and stderr line the tool always wrote, unchanged. Two small always-stderr
helpers back that split - `ErrorConsole` (a Spectre `IAnsiConsole` pinned to `Console.Error`, since
Spectre's own default instance targets stdout, which `convert` may be using for the converted bytes
themselves) and `ColorizingErrorWriter` (a `TextWriter` that renders `CliCommands.Execute`'s one-line
failure in red through `ErrorConsole` on a terminal, or passes it through byte-for-byte otherwise).
Both live in `ExcelReader.Cli`, not `CliCommands.cs` - the interactive/plain decision is
`Console`-shaped state, exactly what that file's tests are built to never touch.

## Why readers are split into partial classes

`XlsxReader` and `XlsbReader` are large enough that one file would be unwieldy, so each is split by
concern rather than by size:

- `XlsxReader.cs` / `XlsbReader.cs` — fields, constructors, sheet navigation, dispose.
- `*.Loading.cs` (XLSX only) — one-time workbook-level XML parsing (sheets, shared strings, date1904).
- `*.Memory.cs` — the in-memory path: constructs directly over `ZipMemoryIndex`/`ZipPart` instead of
  a `Stream`/`ZipArchive`, so it never suspends even under `await foreach`.
- `*.Styles.cs` (XLSX only) — builds the cellXfs-index → is-date-style table.
- `*.Enumerator.cs` — the nested `Enumerator`: the actual streaming row/cell parser. By far the
  largest file in each reader.

All partials of one reader share one field set (C# partial classes are one type), so e.g.
`.Loading.cs`'s shared-string parse populates fields the nested `Enumerator` in `.Enumerator.cs`
reads back. `XlsReader` follows a reduced version of the same split (no `.Memory.cs` — the OLE
compound-file container has no in-memory-ZIP equivalent).

## The sync/async twin convention

Hot-path search/refill primitives (e.g. `IndexOf`/`IndexOfAsync`/`IndexOfSlowAsync`,
`EnsureRowBuffered`/`...Async`/`...SlowAsync` in `XlsxReader.Enumerator.cs`) come in three tiers, not
one generic async method:

1. A blocking sync loop for the sync caller.
2. An async method whose common case — the data is already in the buffered window — is a synchronous
   check returning an already-completed `ValueTask`, so no async state machine is allocated on the
   hot path.
3. A separate `...SlowAsync` method holding the actual `await`-in-a-loop, split out so the rare
   awaiting branch doesn't bloat the fast path's IL/JIT inlining.

Once a row is fully buffered, parsing it (`ParseRow`) has no async twin at all — a fully-buffered
span never needs to await, so both `MoveNext` and `MoveNextAsync` call the same synchronous parse.

A parity test suite (`tests/ExcelReader.Tests/SyncAsyncParityTests.cs`) asserts identical cell
snapshots across sync / async-open / `GetAsyncEnumerator` for all four formats, guarding against the
twins drifting apart.

## The `Row`/`Cell` ref-struct lifetime model

`Row` and `Cell` are `public readonly ref struct`s: zero-allocation views aliasing the reader's own
pooled buffers (the accumulator's cell/value arrays, the shared-string buffer, or — for XLSX's bare
`<v>` fast path — the live read buffer directly). Being `ref struct` lets them hold `ReadOnlySpan<T>`
fields, and the compiler physically prevents a caller from doing anything that would outlive them:
they can't be boxed, stored in a field, captured in a closure, or crossed over an `await`/`yield`.

**Validity window:** exactly one enumeration step. `MoveNext`/`MoveNextAsync` resets the accumulator
and may compact or resize the buffer, so a `Row`/`Cell` from step *N* is invalidated the instant step
*N+1* starts.

**Escape hatch:** `Cell.GetString()` materializes the value as a real `string` (deduplicated for
repeated shared-string cells via an internal cache), `TryFormat` copies raw text into a caller-owned
buffer without allocating a `string`, and `TryGetDouble`/`TryParse<T>`/`TryGetDateTime` extract plain
value types — all safe to keep past the row's lifetime since none of them are spans.

## Typed-model instantiation: three tiers, different tradeoffs

`ExcelParser<T>`/`ExcelEnumerable<T>` create one row model per row via `TypeMapInfo<T>.CreateInstance()`
(`Parser/Internal/TypeMapInfo.cs`), and how that instance gets created depends on how `T` was mapped:

- **`[ExcelSerializable]` (source-generated) or a fluent map that called `.Factory(static () => new T())`**
  — the factory is a plain compiled `new T()`, the cheapest possible per-row allocation.
- **A plain struct with no explicit parameterless constructor** — no factory at all; `default(T)` is
  used directly, since it's byte-for-byte what `new T()` would produce.
- **Attribute/reflection-driven mapping of a class `T`** (`TypeMapper<T>.Build`, `Parser/Internal/TypeMapper.cs`)
  — falls back to `Activator.CreateInstance<T>()` per row. This is deliberately *not* an
  `Expression.Compile()`-built factory: that was tried and reverted because the JIT/expression-tree
  compilation cost dominated cold start (first-use latency), which matters more for this reflection
  path's typical caller (one-off scripts, not long-running services) than the steady-state per-row
  cost of `Activator` dispatch.

If per-row allocation throughput matters more than cold start for your workload, prefer
`[ExcelSerializable]` or an explicit `.Factory(...)` over plain attribute/reflection binding on a
class model — it sidesteps `Activator.CreateInstance<T>()` entirely.

## Further reading

- [`SECURITY.md`](SECURITY.md) — supported versions and how to report a vulnerability.
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — build expectations and how to submit a change.
