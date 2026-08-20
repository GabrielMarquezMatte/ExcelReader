# excelreader (C++)

Header-only C++23 wrapper around ExcelReader's native C ABI: opening a workbook (from a path or
memory, with the full open-options surface), sheet navigation, schema inference, and schema-driven
typed table parsing. No writing, no Arrow, no row-by-row decode yet — see the root README's Python
section for what those look like.

## Requirements

- CMake 3.16+, a C++23 compiler (`std::expected`).
- Git symlink support: this package's `include/xl` is a symlink into `src/ExcelReader.Native/include`
  at the repo root. On Windows, either enable Developer Mode (Windows 10 1703+) or clone with
  `git -c core.symlinks=true clone ...` from an elevated shell — otherwise the symlink checks out as
  a plain text file containing the target path, and the build fails with "excelreader.h not found".

## Usage

```cmake
include(FetchContent)
FetchContent_Declare(excelreader
    GIT_REPOSITORY https://github.com/GabrielMarquezMatte/ExcelReader.git
    GIT_TAG v2.1.3
    SOURCE_SUBDIR cpp)
FetchContent_MakeAvailable(excelreader)

target_link_libraries(your_app PRIVATE xl::excelreader)
```

`FetchContent_MakeAvailable` downloads the matching native binary for your platform from that tag's
GitHub Release automatically (see `cmake/FetchNativeLib.cmake`).

## Build notes

Two variables control where `FetchNativeLib.cmake` gets the native binary and which release it
downloads from:

- **`EXCELREADER_NATIVE_LIB`** (environment variable) — path to a locally-built
  `ExcelReader.Native.{dll,so,dylib}` (e.g. from `dotnet publish src/ExcelReader.Native -r win-x64`)
  to use instead of downloading a release asset. Useful for local development and for CI, which
  builds the native lib fresh per PR (see `.github/workflows/cpp.yml`) rather than depending on a
  tag already being released.
- **`EXCELREADER_VERSION`** (CMake cache variable) — the release tag whose assets to download when
  `EXCELREADER_NATIVE_LIB` isn't set. Auto-detected from `git describe --tags --exact-match` when
  left empty; falls back to `v0.0.0` when the checkout isn't exactly on a tag (e.g. a normal
  development branch).

## Example

```cpp
#include <xl/excelreader.hpp>

struct Row { std::string_view Name; double Value; };

template<> struct xl::ExcelMapper<Row> {
    static constexpr auto get_bindings() {
        return std::make_tuple(
            xl::make_field("Name", &Row::Name),
            xl::make_field("Value", &Row::Value));
    }
};

auto workbook = xl::Workbook::open("book.xlsx");
auto table = xl::parse_sheet<Row>(*workbook);
for (const auto& row : *table) { /* ... */ }
```

### Sheets and schema inference

```cpp
for (const auto& name : *workbook->sheet_names()) { /* ... */ }
workbook->move_to_sheet(1);

// Guess a schema from the header row plus a sample of the data, before committing to one.
for (const auto& column : *workbook->infer_schema(1, 100)) {
    // column.name is nullopt when the column must be resolved by column.index instead.
}
```

Every entry point returns `std::expected<T, xl::Error>` — this header throws nothing.

## Bounds and ABI

`TableView::operator[]` is unchecked, like `std::vector`'s. Use `TableView::at(row)`, which returns
`std::optional<T>` and is `nullopt` outside `[0, size())`.

`xl::Workbook::open`/`open_memory` first check the loaded library's `xl_abi_version()` against the
`XL_ABI_VERSION` this header was compiled against, and fail with an explanatory `xl::Error` rather
than reading native memory through a layout that may have changed. `xl::abi_version()` exposes the
loaded revision directly.

## Benchmarks

Google Benchmark suite in `benchmarks/`, opt-in via `-DEXCELREADER_BUILD_BENCHMARKS=ON`. Measured
on Windows 10 (22H2), 16 logical CPUs @ 3.39 GHz, MSVC 19.51 (Release), Google Benchmark v1.9.1,
`--benchmark_repetitions=10` (means shown).

| Benchmark | RealExcel.xlsb (100 rows) | 65K_Records_Data.xlsb (65,535 rows) |
|---|---:|---:|
| `open` | 78.5 µs | 92.5 µs |
| `parse_sheet` (4 or 6 bound columns) | 87.7 µs | 39.9 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 152.6 µs | 1.22 ms |

`open` is nearly flat across a 655x row-count increase (+18%, not +65,435%) — XLSB keeps its
dimensions/index up front, so opening costs header/metadata, not row data. `parse_sheet` scales
linearly with rows × columns, at roughly 608 ns/row here. `infer_schema` scales with its sample
size, not the file's total row count.

A separate opt-in suite (`-DEXCELREADER_BUILD_BENCHMARKS_COMPARE=ON`, gated separately because it
pulls in [xlnt](https://github.com/tfussell/xlnt) and [xlsxio](https://github.com/brechtsanders/xlsxio)
as heavy source builds — xlsxio's own dependencies, expat and minizip, are fetched and compiled
directly against it) compares against both reading `65K_Records_Data.xlsx` in full (all 14
columns, 65,535 rows; neither competitor reads `.xlsb`, so this comparison is xlsx-only). All three
sides decode every cell into an owned value (`std::string` for text, not the zero-copy
`std::string_view` used above) so none gets a zero-copy advantage the others can't take:

| Library | Mean |
|---|---:|
| ExcelReader (`parse_sheet<FullRow>`) | 110.4 ms |
| xlsxio (`xlsxioread_sheet_next_cell_*`) | 509.3 ms |
| xlnt (`worksheet::rows()` + `cell::to_string()`/`value<double>()`) | 2,335.9 ms |

ExcelReader is ~4.6x faster than xlsxio and ~21x faster than xlnt on this workload. xlsxio is a
lean, purpose-built C streaming reader — the same abstraction level as ExcelReader's own native
core — so it was the strongest of the two competitors tested, though still well behind.

Run locally:

```bash
cmake -S cpp -B build -DEXCELREADER_BUILD_BENCHMARKS=ON -DCMAKE_BUILD_TYPE=Release
cmake --build build --config Release --target excelreader_cpp_benchmarks
./build/benchmarks/excelreader_cpp_benchmarks --benchmark_repetitions=10 --benchmark_report_aggregates_only=true
```

Add `-DEXCELREADER_BUILD_BENCHMARKS_COMPARE=ON -DCMAKE_POLICY_VERSION_MINIMUM=3.5` (xlnt's own
`CMakeLists.txt` predates CMake's minimum-version floor) and build/run
`excelreader_cpp_compare_benchmarks` for the xlnt/xlsxio comparison.
