# C++ and Rust bindings — design spec

Date: 2026-08-18
Status: approved, pending implementation plan

## Goal

Publish two new consumable packages that wrap the existing native C ABI
(`src/ExcelReader.Native/include/excelreader.h`), so C++ and Rust users get
the same NativeAOT shared library the Python package already consumes, without
needing the .NET SDK to build or use it.

Scope for v1, both languages: **open a workbook + `parse_typed` (schema-driven
typed column parse) only.** No write, no Arrow, no row-by-row decode. Matches
the current `excelreader.hpp` and the Python package's own phase-1 scope.

## Repo layout

```
cpp/
  CMakeLists.txt              # INTERFACE target xl::excelreader
  cmake/FetchNativeLib.cmake  # downloads the right native binary for the build platform
  include/                    # symlink -> ../src/ExcelReader.Native/include
  tests/                      # CTest smoke test (successor to teste.cpp)
  README.md

rust/
  excelreader/
    Cargo.toml
    build.rs                  # downloads native binary for the Rust target triple
    src/
      lib.rs                  # extern "C" block mirroring excelreader.h
      error.rs
    tests/
    README.md
```

### `cpp/include` is a symlink

`cpp/include` is a symlink to `../src/ExcelReader.Native/include` (the same
directory that already holds `excelreader.h` and `excelreader.hpp`). One
source of truth — editing the header updates both the native build and the
published C++ package, no manual copy step to forget.

Constraint: git symlinks need `core.symlinks=true` (default on macOS/Linux;
requires Developer Mode or admin on Windows, and `git config core.symlinks
true` + a clone that supports it). Document this in `cpp/README.md`. CI
runners (ubuntu/macos/windows-latest in `cpp.yml`) all support git symlinks
by default; verify during implementation that `actions/checkout@v7` preserves
them on windows-latest specifically (may need `git config --global
core.symlinks true` before checkout, or a `git checkout -- include` re-run
after clone as a fallback if the checkout step doesn't honor it).

## Release pipeline changes (`release.yml`)

Today, `release.yml` builds the native lib per-OS only to package Python
wheels — nothing is ever uploaded as a standalone binary, and no GitHub
Release object is created (the workflow triggers on tag push but never calls
`gh release create`).

New job `publish-native-assets` (matrix: ubuntu-latest, windows-latest,
macos-latest):
1. `dotnet publish` the NativeAOT native library (same publish step
   `build_native.py` already runs).
2. Name the artifact `excelreader-native-<os>-<arch>.<dll|so|dylib>`.
3. Create (or update, if triggered by a re-run) a GitHub Release for the
   pushed tag and upload all three artifacts to it, via
   `softprops/action-gh-release`.

`publish-cpp` and `publish-rust` (new jobs, described below) depend on
`publish-native-assets` so the binaries exist at the release URL before
anything tries to fetch them.

## C++ package

- `cpp/CMakeLists.txt` defines an INTERFACE target `xl::excelreader` (header
  directory = `cpp/include`) plus an IMPORTED SHARED target for the native
  library.
- `cmake/FetchNativeLib.cmake`: at configure time, resolve
  `CMAKE_SYSTEM_NAME`/`CMAKE_SYSTEM_PROCESSOR` to the matching release asset
  name, download it from
  `https://github.com/<org>/<repo>/releases/download/vX.Y.Z/excelreader-native-<os>-<arch>.<ext>`
  into the build tree, and point the imported target at it.
- Consumption: `FetchContent_Declare(excelreader GIT_REPOSITORY ... GIT_TAG
  vX.Y.Z)` then `target_link_libraries(app PRIVATE xl::excelreader)`. No
  vcpkg/Conan submission in v1 — FetchContent/`find_package` after a manual
  `cmake --install` is the only supported path.
- API surface: exactly what `excelreader.hpp` exposes today (`xl::Workbook`,
  `xl::parse_sheet<T>`, `xl::ExcelMapper<T>` specialization pattern). No
  changes to the header's public API — this work only makes it distributable.
- Tests: `cpp/tests/` is the CTest-based successor to the ad hoc
  `teste.cpp`/`teste.c` scratch files at repo root (which get deleted as part
  of this work — they're uncommitted scratch, not part of the package).

## Rust crate

- Single crate `excelreader` (no separate `-sys` crate for v1 — phase-1
  scope is small enough that splitting adds boilerplate without a payoff yet).
- `build.rs`: reads `TARGET` env var, maps it to the matching release asset,
  downloads into `OUT_DIR`, emits `cargo:rustc-link-search` +
  `cargo:rustc-link-lib=dylib=excelreader_native` (or equivalent per-platform
  name). No .NET SDK needed to build the crate.
- `src/lib.rs`: `extern "C"` block mirroring `excelreader.h`
  (`xl_open_file_ex`, `xl_close`, `xl_parse_typed`, `xl_free_table`,
  `xl_last_error_ptr`, status codes) plus a safe wrapper on top:
  - `Workbook::open(path: &str) -> Result<Workbook, Error>`, closing the
    handle on `Drop`.
  - A hand-written `ExcelMapper` trait (mirrors the C++ template
    specialization pattern) — no derive macro in v1, to keep scope matched to
    the C++ side.
  - `parse_sheet::<T>(&workbook) -> Result<TableView<T>, Error>` returning a
    lazy iterator over columnar buffers, freeing the native table on `Drop`.
- Tests: `rust/excelreader/tests/` — integration test against a small fixture
  workbook (reuse an existing test fixture from `tests/`), same shape as the
  Python/C++ smoke tests.

## CI

Two new workflows mirroring `python.yml`'s structure (build native lib
locally for the PR — not via release assets — so CI doesn't depend on a
published release existing):
- `cpp.yml`: build native lib, configure+build `cpp/` with CTest, run tests.
- `rust.yml`: build native lib, point `build.rs` at the locally-built binary
  via an env var override (mirrors Python's `EXCELREADER_NATIVE_LIB`
  escape hatch), `cargo test`.

## Versioning

Both packages version in lockstep with the existing `vX.Y.Z` release tag —
no independent per-language versioning in v1. `cpp/CMakeLists.txt`'s
`project(... VERSION ...)` and `rust/excelreader/Cargo.toml`'s `version` are
both set from the tag the same way `python/pyproject.toml`'s version already
is (`sed` substitution step in the release workflow).

## Out of scope (v1)

- Write support, Arrow C Data Interface, row-by-row (`rows()`/`read_all()`)
  bindings for either language — same phase-1 boundary the Python package
  started with.
- vcpkg/Conan submission for C++, or a `-sys`/safe crate split for Rust.
- Independent versioning per language.
- A derive macro for Rust's `ExcelMapper` trait.
