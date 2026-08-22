# excelreader (C++)

Header-only C++23 wrapper around ExcelReader's native C ABI: opening a workbook (from a path or
memory, with the full open-options surface), sheet navigation, schema inference, schema-driven
typed table parsing, and schema-driven writing. No Arrow, no row-by-row decode yet — see the root
README's Python section for what those look like.

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

### Arrow export

`<xl/excelreader_arrow.hpp>` is a separate header — including `<xl/excelreader.hpp>` never pulls the
Arrow C Data Interface declarations in. It does not depend on the Apache Arrow C++ library: you get
the raw `ArrowArray`/`ArrowSchema` pair, owned by an RAII `xl::ArrowTable`, to hand to whichever
Arrow implementation you already link.

```cpp
#include <xl/excelreader_arrow.hpp>

auto workbook = xl::Workbook::open("book.xlsx").value();
auto table = xl::parse_arrow<Record>(workbook).value();
// table.array / table.schema are a top-level struct array; both release in ~ArrowTable.
```

## Writing

Two layers, mirroring the two on the reading side.

`xl::write_sheet<T>` uses the same `xl::ExcelMapper<T>` specialization `xl::parse_sheet<T>` reads
with, so a round trip needs one mapping, not two:

```cpp
std::vector<Row> rows = /* ... */;
auto written = xl::write_sheet("out.xlsx", rows);   // format inferred from the extension
if (!written) {
    std::fprintf(stderr, "%s\n", written.error().message.c_str());
}
```

If you already hold columnar buffers, `xl::write_columns` hands them to the ABI with **no copy** —
they are borrowed for the duration of the call and must outlive it:

```cpp
std::vector<int64_t> ids{1, 2, 3};
std::vector<double> values{0.5, 1.5, 2.5};
std::array<xl::ColumnRef, 2> columns{
    xl::i64_column("id", ids),
    xl::f64_column("value", values)};

auto written = xl::write_columns("out.xlsx", XL_FORMAT_XLSX, columns);
```

One constructor per column type: `i64_column`, `f64_column`, `bool_column`, `date_column`,
`time_column`, `timestamp_column`, and `string_column` (which takes an `int32` offsets span of
`rows + 1` entries plus the UTF-8 blob).

A nullable column is a values buffer plus an LSB-first validity bitmap — bit set means the row is
valid — passed as the last argument to any of those constructors. `write_columns` checks the bitmap
is long enough for the row count before calling: the ABI takes it without a length and reads
`(rows + 7) / 8` bytes on trust, so a short one would be a buffer overrun. On the struct side,
declare the field `std::optional<T>` and `write_sheet` builds the bitmap for you.

`xl::WriteOptions` sets the sheet name, the CSV dialect, and the XLS/XLSB and XLSX/XLSB toggles.
`XL_FORMAT_AUTO` is rejected — a file being created has no signature bytes to sniff — so
`xl::format_from_path` returning `XL_FORMAT_AUTO` for an unrecognized extension surfaces as a failed
write rather than a silently chosen format.

`write_sheet` walks the range once and appends each field to its own column buffer, with the
per-field dispatch resolved at compile time. That transpose is the only copy it makes, and it is
what the ABI's columnar shape costs a row-shaped caller; `write_columns` pays nothing.

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

### Writing

`excelreader_cpp_write_benchmarks` (same `-DEXCELREADER_BUILD_BENCHMARKS=ON` flag) measures the two
write layers against each other over 7 columns of the same fixture. Both cases write the same
columns, so the gap between them is only the cost of starting from row-shaped data: `BM_WriteColumns`
is handed buffers that are already columnar, while `BM_WriteSheet` starts from a
`std::vector<Row>` and pays the row-to-column transpose.

| Benchmark | Time | Rows/s |
|---|---:|---:|
| `BM_WriteColumns` (pre-transposed) | 60.2 ms | 1.07 M/s |
| `BM_WriteSheet` (from `std::vector<Row>`) | 67.4 ms | 961 k/s |

The transpose costs ~12% here. It is not free, but it is far from the dominant cost of producing
the file — see the comparison below, where the same two cases over 14 columns land within ~7%.

`excelreader_cpp_write_compare_benchmarks` (under `-DEXCELREADER_BUILD_BENCHMARKS_COMPARE=ON`) puts
that against xlnt, xlsxio and [DuckDB](https://github.com/duckdb/duckdb)'s `excel` extension, all
writing the full 14-column, 65,535-row shape of `65K_Records_Data.xlsx` from the same in-memory
rows.
[libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) — same author as rust_xlsxwriter, and,
like xlsxio, a streaming C writer with no document-model overhead — is measured separately, by
`excelreader_cpp_write_compare_lxw_benchmarks`: xlsxio and libxlsxwriter each vendor their own
incompatible copy of minizip and export the same C symbols (`zipOpen`, `zipOpenNewFileInZip`, ...),
so linking both into one binary let calls cross between the two implementations and corrupted
xlsxio's output — see [Known issue](#known-issue-xlsxio--libxlsxwriter-cannot-share-a-binary) below.
Building both write-compare targets and running them back to back is required to see every
competitor.

| Library | Wall | CPU |
|---|---:|---:|
| ExcelReader (`xl::write_columns`, pre-transposed) | 126.3–128.1 ms | 127.6 ms |
| ExcelReader (`xl::write_sheet<FullRow>`) | 135.7–136.6 ms | 137.5 ms |
| DuckDB (`COPY ... TO ... WITH (FORMAT xlsx)`) | 826.3–863.6 ms | 828.1–843.8 ms |
| libxlsxwriter (`worksheet_write_string`/`_number`) | 1,188.0 ms | 1,187.5 ms |
| xlsxio (`xlsxiowrite_add_cell_*`) | 2,450.6 ms | 1,171.9 ms |
| xlnt (`worksheet::cell().value()` + `save()`) | 5,737.9 ms | 5,703.1 ms |

Same machine as above; Google Benchmark's own iteration counts, no `--benchmark_repetitions` (each
iteration writes a whole 65,535-row file, so the slower cases run once or a handful of times). xlnt
and xlsxio only build into `excelreader_cpp_write_compare_benchmarks`; libxlsxwriter only into
`excelreader_cpp_write_compare_lxw_benchmarks` (see the known issue below) — the ExcelReader and
DuckDB rows appear in both, and the small ranges above are those two independent runs, not repeated
sampling within one run.

`write_sheet` — the matched-work number, since it starts from the same `std::vector<FullRow>` every
competitor is handed — is ~6.1–6.4x faster than DuckDB, ~8.7x faster than libxlsxwriter, ~18.1x
faster than xlsxio, and ~42x faster than xlnt on wall time.

### Known issue: xlsxio + libxlsxwriter cannot share a binary

An earlier version of `excelreader_cpp_write_compare_benchmarks` linked xlnt, xlsxio,
libxlsxwriter and DuckDB into one executable. xlsxio (built against minizip-ng's compat layer,
whose `zipOpenNewFileInZip` takes `uint16_t` extrafield sizes) and libxlsxwriter (which vendors
classic minizip, whose same-named function takes 32-bit `uInt` sizes and starts with `if
(size_extrafield_local > 0xffff) return ZIP_PARAMERROR;` — a check that cannot exist in the
minizip-ng version) both export identical C symbol names from a static library linked into that one
binary. The linker kept exactly one definition of each name, so a call could resolve to the wrong
implementation — a `zipFile` opened by one library's `zipOpen` got handed to the other library's
`zipOpenNewFileInZip`, which read it through an incompatible struct layout. That is what produced
`Error creating file "xl/workbook.xml" inside zip file` on xlsxio's background thread in Release
builds (Debug's different link order happened not to trigger it) — and it was silent otherwise:
`xlsxiowrite_close()` still returned success with the workbook.xml entry missing, so the benchmark
published a timing for a file that was skipping work, not a valid xlsx.

The fix is structural: xlsxio and libxlsxwriter now build into two separate executables
(`excelreader_cpp_write_compare_benchmarks` and `excelreader_cpp_write_compare_lxw_benchmarks`,
both compiled from `benchmark_write_compare.cpp` under an `#ifdef`) that are never linked together.
Every case in both executables also reopens the file it just wrote, outside the timed loop, before
trusting its own timing — a writer that silently drops a required part now fails the benchmark
instead of publishing a number for a broken file.

**libxlsxwriter's number barely moved after the fix** (1,188.0 ms wall vs. 1,226.9 ms pre-fix, ~3%,
inside normal run-to-run noise) and its CPU stayed essentially equal to wall time both before and
after — consistent with the corruption running through xlsxio's calls resolving into
libxlsxwriter's minizip, not the reverse: libxlsxwriter's own output was apparently never affected.
That is a plausible explanation for the asymmetry, not a second confirmed fact — the collision could
in principle run either direction depending on link order, and this project isn't going to re-derive
MSVC's exact symbol-resolution algorithm to be certain which way it went here.

Caveats, none of them optional when quoting these:

- **xlsxio's CPU time moved after the fix**, from 843.8–906.3 ms (pre-fix, some runs missing the
  workbook.xml entry) to a confirmed 1,171.9 ms — about 30% more real work, consistent with the
  entry no longer being silently dropped. Wall time barely moved (2,450.6 ms vs. 2,383.5–2,453.9 ms),
  because xlsxio's wall time is dominated by I/O wait either way: CPU is now ~48% of wall,
  against every other case here being CPU-bound (wall ≈ CPU). Against CPU time the gap from
  `write_sheet` to xlsxio is ~8.5x, not ~18.1x — which number is the honest one depends on what you
  are asking: the wall-time ratio is what a caller waits, the CPU-time ratio is what the library
  costs.
- **ExcelReader does slightly more work here**, not less: it attaches a number format to the two
  date columns so Excel shows a date, while every competitor case writes those as bare serial
  numbers. That difference favours the competitors.
- **xlnt builds a full document model** (styles, formats, formulas) before serializing, which is
  more than this library exposes at all. Its number reflects a different feature set, not only a
  slower path.
- **DuckDB's rows are loaded via its Appender API before the timed region**, so its number measures
  `COPY ... TO ... xlsx` alone — the same treatment `write_columns` gets for its transpose. DuckDB is
  a full analytical query engine doing far more than any Excel-writing library here; this measures
  one narrow slice of it, not "DuckDB" as a whole.

The two ExcelReader cases land within ~6% of each other, which is the interesting internal result:
the row-to-column transpose is nearly free next to the cost of producing the file. Reach for
`write_columns` when your data is already columnar, but `write_sheet` is not the slow path.
