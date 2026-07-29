# Contributing

Thanks for considering a contribution. This project takes small, focused pull requests over large
rewrites — see [ARCHITECTURE.md](ARCHITECTURE.md) for the shape of the codebase before diving in, and
[STYLEGUIDE.md](STYLEGUIDE.md) for the code style, which the analyzers only partly enforce.

## Build expectations

- **Warnings are errors.** `Directory.Build.props` sets `TreatWarningsAsErrors`, with a curated
  `AnalysisMode=All` analyzer set (Sonar, Meziantou, Roslynator, AsyncFixer, and more). A PR that
  doesn't build clean locally won't build clean in CI either — run a full build before pushing:

  ```bash
  dotnet build ExcelReader.slnx --configuration Release
  ```

- **Public API changes require a `PublicAPI.Unshipped.txt` entry.** `Microsoft.CodeAnalysis.PublicApiAnalyzers`
  is active (arrives transitively via `Roslyn.Diagnostics.Analyzers`) and fails the build on any
  unrecorded public member. If you add, change, or remove anything public, update **both**
  `src/ExcelReader.Core/PublicAPI/net8.0/PublicAPI.Unshipped.txt` and
  `src/ExcelReader.Core/PublicAPI/net10.0/PublicAPI.Unshipped.txt`. A bot promotes `Unshipped` →
  `Shipped` automatically after each release — don't edit `Shipped.txt` by hand.

- **Tests are required for behavior changes.** Run the suite before opening a PR:

  ```bash
  dotnet test tests/ExcelReader.Tests/ExcelReader.Tests.csproj --configuration Release
  ```

  Untrusted-input paths (the CFB/OLE, BIFF8, BIFF12, and ZIP parsers) get extra scrutiny — new
  parsing code should have a corresponding limit/fuzz-safety test in
  `tests/ExcelReader.Tests/ReaderLimitTests.cs` or `FuzzTests.cs` where relevant. Read
  [STYLEGUIDE.md § Untrusted Input](STYLEGUIDE.md#untrusted-input) before touching a parser: every
  length, offset, and size read from the file must be bounded before it drives an allocation.

## Pull requests

- One focused change per PR — don't batch unrelated fixes into one commit or one PR.
- If a change is user-visible (new API, behavior change, performance claim), mention it in the PR
  description; the README's benchmark tables and changelog are updated separately, not as part of
  every PR.
- CI runs on Linux, Windows, and macOS across .NET 8 and .NET 10 — a change that only builds on one
  OS/TFM combination isn't ready to merge.

## Reporting bugs / requesting features

Use GitHub Issues for bugs and feature requests. For suspected security vulnerabilities, do **not**
open a public issue — see [SECURITY.md](SECURITY.md) for the private reporting channel.

## Code of conduct

This project follows the [Code of Conduct](CODE_OF_CONDUCT.md).
