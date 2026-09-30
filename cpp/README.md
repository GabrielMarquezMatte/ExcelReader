# excelreader (C++)

Header-only C++23 wrapper around ExcelReader's native C ABI: opening a workbook (from a path or
memory, with the full open-options surface), sheet navigation, schema inference, schema-driven
typed table parsing, schema-driven writing, row-by-row decoded reads, and Arrow export.

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
    GIT_TAG v4.0.1
    SOURCE_SUBDIR cpp)
FetchContent_MakeAvailable(excelreader)

target_link_libraries(your_app PRIVATE xl::excelreader)
excelreader_copy_native_library(your_app)
```

`FetchContent_MakeAvailable` downloads the matching native binary for your platform from that tag's
GitHub Release automatically (see `cmake/FetchNativeLib.cmake`).

`excelreader_copy_native_library` puts that binary next to your executable once it links. On Windows
it is required: the DLL is loaded by name at startup, and one that sits neither beside the executable
nor on `PATH` makes the process exit with `0xc0000135` (`STATUS_DLL_NOT_FOUND`) before `main` runs.
The call is a no-op elsewhere, where the rpath baked in at link time already resolves the library.

### Install once, `find_package` after

Building `cpp` as its own project installs the headers, the native binary and a config package, so
downstream projects find it without re-downloading anything:

```bash
cmake -S cpp -B build/cpp -DCMAKE_BUILD_TYPE=Release
cmake --install build/cpp --config Release --prefix /your/prefix
```

```cmake
find_package(excelreader 3.0 REQUIRED)
target_link_libraries(your_app PRIVATE xl::excelreader)
```

Point `CMAKE_PREFIX_PATH` at the prefix you installed into. The native binary is resolved once, at
install time — `find_package` never touches the network. The package version comes from
`EXCELREADER_VERSION`, so a checkout that isn't on a tag installs as `0.0.0` and any versioned
`find_package` request against it fails; pass `-DEXCELREADER_VERSION=v3.0.2` to install under a real
version. On Windows the generated import library is installed next to the DLL, and the package ships
`excelreader_copy_native_library` as well, so a `find_package` consumer copies the DLL next to its
executable the same way a `FetchContent` one does.

Pass `-DEXCELREADER_INSTALL=OFF` to skip the install rules. They default off when `cpp` is pulled in
with `add_subdirectory`/`FetchContent`, so a parent project's `install` step never picks them up.

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

## Reading rows one at a time

`Workbook::rows()` returns a cursor over the current sheet. Each row borrows a buffer the cursor
reuses, so iterating allocates nothing per row. A clean end of sheet arrives as an error carrying
`XL_EOF`:

```cpp
auto workbook = xl::Workbook::open("book.xlsx").value();
auto cursor = workbook.rows();
while (auto row = cursor.next_row())
{
    for (auto cell : *row)
    {
        std::print("{}\t", cell.value);
    }
    std::println();
}
```

Each row is invalidated by the next `next_row()`. To hold every row at once, use
`read_all_decoded()`, which decodes the whole remaining sheet in one native call and owns it until
destroyed:

```cpp
auto rows = workbook.read_all_decoded().value();
for (auto row : rows)
{
    std::println("{} cells", row.size());
}
```

`xl::parse_sheet<T>` remains the fastest way to read a sheet whose columns you know.

### Chunked reads

For a sheet too large to hold in memory at once, `xl::typed_reader<T>` is `xl::parse_sheet<T>`
delivered a batch at a time:

```cpp
auto workbook = xl::Workbook::open("book.xlsx").value();
auto reader = xl::typed_reader<Row>(workbook, 1, 10'000).value();
for (auto &batch : reader)
{
    if (!batch) { break; }
    for (const auto &row : *batch) { /* ... */ }
}
```

`typed_reader` returns `std::expected<xl::TypedReader<T>, xl::Error>`; `next()` returns
`std::expected<std::optional<TableView<T>>, Error>`, an empty optional at end of sheet, and
`TypedReader` is also an input range, so the range-`for` above works directly. `batch_size` is rows
per batch — `0` means one unbounded batch, identical to `parse_sheet`; negative is an error. Each
batch is independent and outlives the reader.

A workbook serves one chunked read at a time, of either kind (typed or Arrow — see
[Arrow export](#arrow-export) below), and any other read on it while one is live invalidates that
reader: its next call reports a latched error rather than resuming from the moved cursor. C++ cannot
enforce that at compile time the way the Rust binding does — finish or destroy the reader before
starting another read.

### Encrypted workbooks

`OpenOptions::password(std::string_view)` unlocks a password-protected OOXML workbook
(.xlsx/.xlsb/.xlsm) through `Workbook::open`/`open_memory`:

```cpp
xl::OpenOptions options{};
options.password("hunter2");
auto workbook = xl::Workbook::open("protected.xlsx", XL_FORMAT_AUTO, &options);
if (!workbook) {
    if (workbook.error().code == XL_STATUS_PASSWORD_REQUIRED) { /* no password was supplied */ }
    else if (workbook.error().code == XL_STATUS_PASSWORD_INCORRECT) { /* ask again */ }
    else { /* workbook.error().message - unsupported scheme or a corrupt file; not worth retrying */ }
}
```

Unlike the Rust/Python bindings, this header never throws — `XL_STATUS_PASSWORD_REQUIRED` and
`XL_STATUS_PASSWORD_INCORRECT` (from `excelreader.h`) are reported as ordinary `xl::Error::code`
values on the `std::unexpected` returned from `open`/`open_memory`, distinguishable from any other
failure without parsing `error().message`.

The `password_` string this builds into `xl_open_options::password` must outlive the open call;
keeping it in the `OpenOptions` object (which owns its own copy) is what makes
`options.password("hunter2")` safe to call with a temporary, as above.

Writing an encrypted workbook is a second step: write the plaintext package with `write_columns`/
`write_sheet`, then wrap it with `xl::encrypt_package`, which produces an agile-encrypted (ECMA-376
4.4) CFB container — AES-256-CBC, SHA-512, 100,000 spin iterations, with a `dataIntegrity` HMAC:

```cpp
auto written = xl::write_columns("plain.xlsx", XL_FORMAT_XLSX, columns);
auto encrypted = xl::encrypt_package("plain.xlsx", "secret.xlsx", "hunter2");
if (!encrypted) { /* encrypted.error().message */ }
```

`package_path` is read twice (it is not disposed or removed), so it must already be a finished file.
Encryption parameters are fixed at Excel's own defaults — there are no options — and only XLSX/XLSB
packages can be encrypted, matching what `Workbook::open`/`open_memory` can decrypt.

### Arrow export

`<xl/excelreader_arrow.hpp>` is a separate header — including `<xl/excelreader.hpp>` never pulls the
Arrow C Data Interface declarations in. It does not depend on the Apache Arrow C++ library: you get
the raw `ArrowArray`/`ArrowSchema` pair, owned by an RAII `xl::ArrowTable`, to hand to whichever
Arrow implementation you already link.

```cpp
#include <xl/excelreader_arrow.hpp>

auto workbook = xl::Workbook::open("book.xlsx");
auto table = xl::parse_arrow<Row>(*workbook);
// table->array / table->schema are a top-level struct array; both release in ~ArrowTable.
```

Batched: `xl::arrow_stream<T>` is `xl::parse_arrow<T>` delivered a batch at a time. `schema()` and
`next()` return RAII guards (`ArrowSchemaGuard`, `ArrowArrayGuard`) that release exactly once, and
each batch outlives the stream that produced it:

```cpp
auto stream = xl::arrow_stream<Row>(*workbook, 1, 10'000);
while (true)
{
    auto batch = stream->next();
    if (!batch || !batch->has_value()) { break; }
    // (*batch)->array is this one batch
}
```

Same `header_row`/`batch_size` meaning, and the same one-chunked-read-at-a-time rule, as
`xl::typed_reader` above.

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
| `open` (`open_memory`) | 20.0 µs | 1.24 ms |
| `parse_sheet` (4 or 6 bound columns) | 50.3 µs | 27.3 ms |
| `typed_reader`, batches of 1,000 / 10,000 rows | - | 26.1 ms / 26.5 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 73.7 µs | 605.7 µs |

`open` here is `open_memory`, which copies the caller's buffer before returning (the caller may free
it afterwards), so it scales with file size: 1.24 ms is mostly copying the 3.7 MB fixture. Opening by
path reads only the package directory and stays nearly flat — see the Rust suite below.
`parse_sheet` scales linearly with rows × columns, at roughly 420 ns/row here. Draining the same
sheet through the batched typed reader costs the same as one whole-sheet `parse_sheet` while holding
only one batch in memory at a time. `infer_schema` scales with its sample size, not the file's total
row count.

A separate opt-in suite (`-DEXCELREADER_BUILD_BENCHMARKS_COMPARE=ON`, gated separately because it
pulls in [xlnt](https://github.com/tfussell/xlnt) and [xlsxio](https://github.com/brechtsanders/xlsxio)
as heavy source builds — xlsxio's own dependencies, expat and minizip, are fetched and compiled
directly against it) compares against both reading `65K_Records_Data.xlsx` in full (all 14
columns, 65,535 rows; neither competitor reads `.xlsb`, so this comparison is xlsx-only). All three
sides decode every cell into an owned value (`std::string` for text, not the zero-copy
`std::string_view` used above) so none gets a zero-copy advantage the others can't take:

| Library | Mean |
|---|---:|
| ExcelReader (`parse_sheet<FullRow>`) | 73.7 ms |
| DuckDB (`read_xlsx`, summing every column in SQL) | 433.6 ms |
| xlsxio (`xlsxioread_sheet_next_cell_*`) | 505.0 ms |
| xlnt (`worksheet::rows()` + `cell::to_string()`/`value<double>()`) | 2,494.2 ms |

ExcelReader is ~5.9x faster than DuckDB, ~6.9x faster than xlsxio and ~34x faster than xlnt on this
workload. xlsxio is a
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

### CSV against zsv

`excelreader_cpp_csv_compare_benchmarks` (`-DEXCELREADER_BUILD_BENCHMARKS_ZSV=ON`) reads
`65K_Records_Data.csv` (8.2 MB, 65,535 rows × 14 columns, CRLF, UTF-8 BOM) against
[zsv](https://github.com/liquidaty/zsv) v1.4.3, a SIMD CSV parser in C, with both of its engines:
the default (`compat`) and the branchless SIMD one (`fast`, `scan_engine = 3`). zsv uses GCC vector
extensions MSVC cannot compile, so this target needs GCC or Clang, and it is its own executable
because the DuckDB comparison above only links under MSVC.

Two workloads:

- **Typed:** `parse_sheet<FullRow>` against zsv plus the same conversions done by hand (`std::string`
  for text, ISO dates, `std::from_chars` for the integers and doubles), one owned `FullRow` per row.
  All three produce the same checksum.
- **Cells:** every cell's bytes counted. ExcelReader goes through `RowCursor` (`xl_next_row_view`), zsv
  through its row handler. zsv's checksum is 3 bytes higher because it keeps the BOM in the first
  header cell.

Measured on the same machine as above, GCC 16.2.0 (MSYS2 UCRT64, `-O3 -mavx2` for zsv), Release,
`--benchmark_repetitions=10` (means shown):

| Workload | ExcelReader | zsv `compat` | zsv `fast` |
|---|---:|---:|---:|
| Typed, 14 columns into `FullRow` | 30.4 ms | 25.7 ms | 22.7 ms |
| Cells, byte count | 11.9 ms | 5.9 ms | 3.5 ms |

**zsv is faster on both.** Typed, it is ~1.2x (`compat`) and ~1.3x (`fast`) faster than ExcelReader.
Counting cells, it is ~2.0x and ~3.4x faster.

The cells gap is the C ABI, not the CSV parser. `xl_next_row_view` copies every cell's value into a
native row the handle owns, one call per row, and that alone costs ~10.6 ms here (calling it without
reading any cell, measured separately). The managed reader reads the same file in ~5 ms under
BenchmarkDotNet (`RealDataReadBenchmark.Csv_ExcelReader`), but that is a different harness and
runtime, so it is context rather than a like-for-like number.

The ExcelReader rows were 34.2 ms (typed) and 14.3 ms (cells) before the NativeAOT library's static
PGO profile was regenerated (a namespace refactor had left it matching none of the hot methods) and
its x64 build was pinned to `x86-64-v3`. The cells row was 30.5 ms before that, when `xl::RowView`
still decoded each cell through `std::optional` helpers. The rows are from separate 10-repetition
runs (typed CVs 2–5%, cells CVs ≤1.3%). A first run had CVs up to 20% and was discarded.

Run locally, on Windows from an MSYS2 UCRT64 shell with `gcc` and `ninja`:

```bash
cmake -S cpp -B build-zsv -G Ninja -DCMAKE_BUILD_TYPE=Release -DCMAKE_POLICY_DEFAULT_CMP0168=NEW \
  -DEXCELREADER_BUILD_BENCHMARKS=ON -DEXCELREADER_BUILD_BENCHMARKS_ZSV=ON
cmake --build build-zsv --target excelreader_cpp_csv_compare_benchmarks
./build-zsv/benchmarks/excelreader_cpp_csv_compare_benchmarks --benchmark_repetitions=10 --benchmark_report_aggregates_only=true
```

The MinGW Makefiles generator failed here inside FetchContent's sub-build. Configuring with Ninja and
`CMAKE_POLICY_DEFAULT_CMP0168=NEW` worked. Under MinGW the executable links its runtime statically,
so it does not pick up another toolchain's `libstdc++-6.dll` from `PATH`.

### Writing

`excelreader_cpp_write_benchmarks` (same `-DEXCELREADER_BUILD_BENCHMARKS=ON` flag) measures the two
write layers against each other over 7 columns of the same fixture. Both cases write the same
columns, so the gap between them is only the cost of starting from row-shaped data: `BM_WriteColumns`
is handed buffers that are already columnar, while `BM_WriteSheet` starts from a
`std::vector<Row>` and pays the row-to-column transpose.

| Benchmark | Time | Rows/s |
|---|---:|---:|
| `BM_WriteColumns` (pre-transposed) | 44.2 ms | 1.48 M/s |
| `BM_WriteSheet` (from `std::vector<Row>`) | 50.9 ms | 1.29 M/s |

The transpose costs ~15% here. It is not free, but it is not the dominant cost of producing the
file — see the comparison below, where the same two cases over 14 columns land within ~15–17%.

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
| ExcelReader (`xl::write_columns`, pre-transposed) | 85.1–85.9 ms | 85.1–87.1 ms |
| ExcelReader (`xl::write_sheet<FullRow>`) | 97.8–100.1 ms | 96.4–100.4 ms |
| DuckDB (`COPY ... TO ... WITH (FORMAT xlsx)`) | 833.3–869.5 ms | 828.1–875.0 ms |
| libxlsxwriter (`worksheet_write_string`/`_number`) | 1,183.8 ms | 1,187.5 ms |
| xlsxio (`xlsxiowrite_add_cell_*`) | 2,453.8 ms | 843.8 ms |
| xlnt (`worksheet::cell().value()` + `save()`) | 5,640.7 ms | 5,640.6 ms |

Same machine as above; Google Benchmark's own iteration counts, no `--benchmark_repetitions` (each
iteration writes a whole 65,535-row file, so the slower cases run once or a handful of times). xlnt
and xlsxio only build into `excelreader_cpp_write_compare_benchmarks`; libxlsxwriter only into
`excelreader_cpp_write_compare_lxw_benchmarks` (see the known issue below) — the ExcelReader and
DuckDB rows appear in both, and the small ranges above are those two independent runs, not repeated
sampling within one run.

`write_sheet` — the matched-work number, since it starts from the same `std::vector<FullRow>` every
competitor is handed — is ~8.5–8.7x faster than DuckDB, ~12.1x faster than libxlsxwriter, ~25x
faster than xlsxio, and ~56x faster than xlnt on wall time.

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

- **xlsxio's CPU time is not stable across runs.** It read 843.8–906.3 ms before the fix (some
  runs missing the workbook.xml entry), 1,171.9 ms and 687.5 ms in the two runs after it,
  1,015.6 ms in the next and 843.8 ms in this one, while its wall time stayed at ~2.4 s throughout. Its wall time is dominated by I/O wait —
  CPU is well under half of wall, against every other case here being CPU-bound (wall ≈ CPU) — so
  quote the wall-time ratio, which is what a caller waits. Against this run's CPU time the gap from
  `write_sheet` to xlsxio is ~8.4x, not ~25x.
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

The two ExcelReader cases land within ~12–22% of each other. The transpose's share grew as producing
the file got cheaper, but it is still a minority of the cost. Reach for
`write_columns` when your data is already columnar, but `write_sheet` is not the slow path.
