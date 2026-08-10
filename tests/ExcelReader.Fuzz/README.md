# ExcelReader.Fuzz

Fuzz harnesses for the four reader front ends. Every reader parses untrusted binary input and is
expected to reject malformed data through a small set of documented exception types; anything else
means attacker-controlled bytes reached code that assumed they were well-formed.

## Why this exists

The reader options carry explicit resource limits (`MaxCellBytes`, `MaxSharedStringBytes`, the
buffer-growth cap in `LimitChecks.NextBufferSize`) whose job is to turn a malicious file into an
`ExcelLimitExceededException` instead of an allocation blow-up. Those limits were designed but never
exercised adversarially. This suite is what turns them from *intended* into *demonstrated*.

`FuzzOracle` is the important part: it decides what counts as a bug.

| Accepted | Treated as a defect |
|---|---|
| `InvalidDataException`, `EndOfStreamException` | `IndexOutOfRangeException`, `ArgumentOutOfRangeException` |
| `NotSupportedException` | `NullReferenceException`, `OverflowException` |
| `ExcelLimitExceededException`, `ExcelParseException` | `OutOfMemoryException` — a limit that failed to hold |
| exactly `ArgumentException` / `InvalidOperationException` | anything else, plus hangs (engine timeout) |

`FuzzOracle.SelfCheck()` asserts that polarity in both directions before every smoke run, so a clean
result cannot silently mean "the oracle accepts everything".

## Targets

`xlsx`, `xlsx-memory`, `xlsb`, `xlsb-memory`, `xls`, `csv`, `csv-sniff`.

The `-memory` variants exist because the in-memory ZIP path (`ZipMemoryIndex`) is a different
container parser from the `Stream`/`ZipArchive` one, and `csv-sniff` because dialect detection runs
over untrusted bytes before any reader is constructed.

## Running locally

No native tooling — runs anywhere, and is what gates pull requests:

```bash
dotnet run --project tests/ExcelReader.Fuzz -c Release -- seeds corpus
dotnet run --project tests/ExcelReader.Fuzz -c Release -- check corpus 200 1
#                                                          ^dir  ^mutations ^seed
```

`check` runs every target over the corpus plus deterministic mutations of it (truncation, bit flips,
run overwrites, splices). Failing inputs are written to the temp directory and the seed is printed,
so any failure replays exactly.

Coverage-guided fuzzing (Linux; finds far more, needs sustained runtime):

```bash
dotnet tool install --global SharpFuzz.CommandLine
wget https://raw.githubusercontent.com/Metalnem/libfuzzer-dotnet/master/libfuzzer-dotnet.cc
clang -fsanitize=fuzzer libfuzzer-dotnet.cc -o libfuzzer-dotnet

dotnet publish tests/ExcelReader.Fuzz -c Release -o fuzz-out
chmod +x fuzz-out/ExcelReader.Fuzz

# Seeds FIRST — see the warning below.
dotnet fuzz-out/ExcelReader.Fuzz.dll seeds corpus

sharpfuzz fuzz-out/ExcelReader.Core.dll        # instrument the code under test, not the harness

# target_path is the published apphost, NOT `dotnet`; the target comes from the environment.
FUZZ_TARGET=xlsb ./libfuzzer-dotnet --target_path=fuzz-out/ExcelReader.Fuzz \
  corpus -max_len=131072 -timeout=25 -rss_limit_mb=2048
```

> Two invocation traps, both of which fail in ways that do not name the real cause:
>
> * **`--target_path` must be the apphost, not `dotnet`.** libfuzzer-dotnet hands `--target_args` to
>   the child as a *single* argv entry instead of splitting on spaces, so
>   `--target_path=dotnet --target_args="Foo.dll xlsb"` gives `dotnet` one bogus path. It prints its
>   usage text and exits, which surfaces as `short read: expected 4 bytes, got 0 bytes`.
> * **Select the target with `FUZZ_TARGET`, not `--target_args`.** SharpFuzz's
>   `Fuzzer.LibFuzzer.Run` parses argv itself and treats a lone argument as a single input file to
>   replay, so our own argument gets swallowed as a corpus path.

> **Once `sharpfuzz` has run, that build may only be launched by `libfuzzer-dotnet`.** The
> instrumentation writes coverage into a shared-memory region the engine sets up; running the
> instrumented assembly standalone — `seeds`, `check`, or anything else — dies with
> `AccessViolationException: Attempted to read or write protected memory`. Generate seeds before
> instrumenting, and use a separate ordinary build (`dotnet run --project`) for `check`.

Shrink a bloated corpus with `-merge=1`; minimise a crashing input with `-minimize_crash=1`.

## `corpus/`

Inputs that once crashed a target, kept permanently so those paths stay covered. CI copies them into
the working corpus for both jobs.

| File | Target | Was |
|---|---|---|
| `xls-minifat-overflow.bin` | `xls` | `OverflowException` out of `XlsCompoundFile.ReadIntSectors`. `miniFatSectorCount` (16,777,215 in a 3.4 KB file) was the one header sector count not bounded against the container length, and `ReadIntSectors` multiplies it by `sectorSize` inside a `checked` block. Now rejected as `InvalidDataException`; regression test in `ExcelOpenAndOleErrorTests`. |
| `xlsx-truncated-cellxfs.bin` | `xlsx` | `ArgumentOutOfRangeException` out of `XlsxReader.ParseStyleDateFlags`. A styles part truncated mid-`<cellXfs` open tag left the `'>'` search at -1, which then anchored the search for `</cellXfs>` before being checked. Now returns no date flags; the `IdxOf` helpers also treat a negative anchor as "not found". Regression test in `XlsxReaderTests`. |
| `xlsx-isodate-year-below-100.bin` | `xlsx`, `xlsx-memory` | `OverflowException` ("Not a legal OleAut date") out of `XlsxReader.Enumerator.EmitIsoDate`. A `t="d"` cell holding `0024-02-29T21:00:00.000Z` parsed fine into a `DateTime` but has no OLE automation serial — `ToOADate` only accepts `0100-01-01` and later. Such values are now kept verbatim as text like any other unparseable `t="d"`. Regression test in `NamespacePrefixAndIsoDateTests`. |

## Seeds

Generated by `SeedCorpus`, not committed: a coverage-guided fuzzer needs a structurally valid
starting point per container (random bytes essentially never reach the record parsers of a ZIP- or
OLE-based format), and generating them keeps binaries out of the repo while guaranteeing the seeds
match the current writers. CI additionally seeds from `RealExcel.xlsx`/`RealExcel.xlsb` and
`tests/ExcelReader.Tests/data/` — real producer output is the most valuable seed material there is.

## When a crash is found

1. The engine writes the input to `crash-*` (CI uploads these as artifacts).
2. Reproduce on an **uninstrumented** build — put the file alone in a directory and run
   `dotnet run --project tests/ExcelReader.Fuzz -c Release -- check <that-dir> 0 1`. Reusing
   `fuzz-out/` here would hit the AccessViolationException above instead of the real failure.
3. Add it as a regression test in `ExcelReader.Tests`, then fix.
4. Keep the input in the corpus — it now guards that path.
