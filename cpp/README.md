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
