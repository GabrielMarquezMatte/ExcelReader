# excelreader (C++)

Header-only C++23 wrapper around ExcelReader's native C ABI (`xl_open_file_ex`, `xl_parse_typed`).
Phase 1: opening a workbook and schema-driven typed table parsing only — no writing, no Arrow, no
row-by-row decode. See the root README's Python section for what those look like; C++ gets the same
scope first.

## Requirements

- CMake 3.16+, a C++23 compiler (`std::expected`).
- Git symlink support: this package's `include/` is a symlink into `src/ExcelReader.Native/include`
  at the repo root. On Windows, either enable Developer Mode (Windows 10 1703+) or clone with
  `git -c core.symlinks=true clone ...` from an elevated shell — otherwise the symlink checks out as
  a plain text file containing the target path, and the build fails with "excelreader.h not found".

## Usage

```cmake
include(FetchContent)
FetchContent_Declare(excelreader
    GIT_REPOSITORY https://github.com/GabrielMarquezMatte/ExcelReader.git
    GIT_TAG v0.1.0
    SOURCE_SUBDIR cpp)
FetchContent_MakeAvailable(excelreader)

target_link_libraries(your_app PRIVATE xl::excelreader)
```

`FetchContent_MakeAvailable` downloads the matching native binary for your platform from that tag's
GitHub Release automatically (see `cmake/FetchNativeLib.cmake`).

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
