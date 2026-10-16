![ExcelReader logo](https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png)

# excelreader (C++)

Header-only C++23 wrapper around ExcelReader's native C ABI: opening a workbook (from a path,
memory, a random-access `xl::Source` or an `xl::InputStream`, with the full open-options surface), sheet navigation, schema inference, schema-driven
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
    GIT_TAG v7.0.0
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
find_package(excelreader 7.0 REQUIRED)
target_link_libraries(your_app PRIVATE xl::excelreader)
```

Point `CMAKE_PREFIX_PATH` at the prefix you installed into. The native binary is resolved once, at
install time — `find_package` never touches the network. The package version comes from
`EXCELREADER_VERSION`, so a checkout that isn't on a tag installs as `0.0.0` and any versioned
`find_package` request against it fails; pass `-DEXCELREADER_VERSION=v7.0.0` to install under a real
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
auto table = xl::parse<Row>(workbook->sheet(0));
for (const auto& row : *table) { /* ... */ }
```

### Sheets and schema inference

```cpp
for (xl::Sheet sheet : *workbook->sheets()) { std::println("{}", sheet.name().value()); }

// Guess a schema from the header row plus a sample of the data, before committing to one.
for (const auto& column : *workbook->sheet(1).infer_schema(1, 100)) {
    // column.name is nullopt when the column must be resolved by column.index instead.
}
```

Every entry point returns `std::expected<T, xl::Error>` — this header throws nothing.

A workbook hands out sheets. An `xl::Sheet` is the workbook handle and an index: it owns nothing,
it is cheap to copy, and every read on it opens its own cursor. `workbook->sheet(i)` makes no native
call; an index outside the workbook is reported by the first read on the returned sheet.

```cpp
auto workbook = xl::Workbook::open("report.xlsx");
if (!workbook) { /* workbook.error().message */ }

auto sheets = workbook->sheets().value();
for (xl::Sheet sheet : sheets)
{
    std::println("{} {}", sheet.index(), sheet.name().value());
}

if (auto totals = workbook->sheet_by_name("Totals").value())
{
    auto cursor = totals->rows().value();
    while (auto row = cursor.next_row())
    {
        // ...
    }
}
```

One open workbook can be read from several threads, which loads the shared-string table once:

```cpp
std::vector<std::thread> threads;
for (xl::Sheet sheet : sheets)
{
    threads.emplace_back([sheet] { auto table = xl::parse<Row>(sheet); /* ... */ });
}
for (std::thread &thread : threads)
{
    thread.join();
}
```

A cursor, a typed reader and an Arrow stream are each used by one thread at a time, and keep
working after the workbook is destroyed. A `Sheet` does not: take what you need from it first.

## Reading rows one at a time

`Sheet::rows()` returns a cursor over that sheet. Each row borrows a buffer the cursor
reuses, so iterating allocates nothing per row. A clean end of sheet arrives as an error carrying
`XL_EOF`:

```cpp
auto workbook = xl::Workbook::open("book.xlsx").value();
auto cursor = workbook.sheet(0).rows().value();
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
`Sheet::read_all()`, which decodes the whole sheet in one native call and owns it until
destroyed:

```cpp
auto rows = workbook.sheet(0).read_all().value();
for (auto row : rows)
{
    std::println("{} cells", row.size());
}
```

`xl::parse<T>` remains the fastest way to read a sheet whose columns you know.

### Chunked reads

For a sheet too large to hold in memory at once, `xl::parse_batches<T>` is `xl::parse<T>`
delivered a batch at a time:

```cpp
auto workbook = xl::Workbook::open("book.xlsx").value();
auto reader = xl::parse_batches<Row>(workbook.sheet(0), {.batch_size = 10'000}).value();
for (auto &batch : reader)
{
    if (!batch) { break; }
    for (const auto &row : *batch) { /* ... */ }
}
```

`parse_batches` returns `std::expected<xl::TypedReader<T>, xl::Error>`; `next()` returns
`std::expected<std::optional<TableView<T>>, Error>`, an empty optional at end of sheet, and
`TypedReader` is also an input range, so the range-`for` above works directly. `batch_size` is rows
per batch and must be positive (`0` or negative is `XL_INVALID_ARGUMENT`; read a whole sheet with
`parse`). Each batch is independent and outlives the reader.

Several typed readers and Arrow streams may be open on one workbook at once; each read is
independent.

### Encrypted workbooks

`OpenOptions::password` unlocks a password-protected OOXML workbook (.xlsx/.xlsb/.xlsm) through
any `Workbook::open` overload:

```cpp
auto workbook = xl::Workbook::open("protected.xlsx", {.password = "hunter2"});
if (!workbook) {
    if (workbook.error().code == XL_STATUS_PASSWORD_REQUIRED) { /* no password was supplied */ }
    else if (workbook.error().code == XL_STATUS_PASSWORD_INCORRECT) { /* ask again */ }
    else { /* workbook.error().message - unsupported scheme or a corrupt file; not worth retrying */ }
}
```

Unlike the Rust/Python bindings, this header never throws — `XL_STATUS_PASSWORD_REQUIRED` and
`XL_STATUS_PASSWORD_INCORRECT` (from `excelreader.h`) are reported as ordinary `xl::Error::code`
values on the `std::unexpected` returned from `open`, distinguishable from any other failure without
parsing `error().message`.

Writing an encrypted workbook is a second step: write the plaintext package with `write_columns`/
`write_sheet`, then wrap it with `xl::encrypt_package`, which produces an agile-encrypted (ECMA-376
4.4) CFB container — AES-256-CBC, SHA-512, 100,000 spin iterations, with a `dataIntegrity` HMAC:

```cpp
auto written = xl::write_columns("plain.xlsx", columns);
auto encrypted = xl::encrypt_package("plain.xlsx", "secret.xlsx", "hunter2");
if (!encrypted) { /* encrypted.error().message */ }
```

`package_path` is read twice (it is not disposed or removed), so it must already be a finished file.
Encryption parameters are fixed at Excel's own defaults — there are no options — and only XLSX/XLSB
packages can be encrypted, matching what `Workbook::open` can decrypt.

### Reading from your own bytes

Derive from `xl::Source` to read from anything you can read at an offset; `read_at` is called from
several threads, in cached 4 MiB blocks. Derive from `xl::InputStream` for a stream that cannot
seek: a CSV is read as it arrives, an XLSX, XLSB or XLS is read whole first.

```cpp
class StdinStream final : public xl::InputStream
{
public:
    std::expected<std::size_t, std::string> read(std::span<std::byte> buffer) override
    {
        std::cin.read(reinterpret_cast<char *>(buffer.data()), static_cast<std::streamsize>(buffer.size()));
        if (std::cin.bad())
        {
            return std::unexpected(std::string("reading stdin failed"));
        }
        return static_cast<std::size_t>(std::cin.gcount());
    }
};

auto workbook = xl::Workbook::open(std::make_unique<StdinStream>(), {.format = XL_FORMAT_CSV});
```

The workbook owns the object you pass and deletes it once nothing needs it. A returned error
string or a thrown exception becomes the `xl::Error` of the call that needed the bytes.

### Arrow export

`<xl/excelreader_arrow.hpp>` is a separate header — including `<xl/excelreader.hpp>` never pulls the
Arrow C Data Interface declarations in. It does not depend on the Apache Arrow C++ library: you get
the raw `ArrowArray`/`ArrowSchema` pair, owned by an RAII `xl::ArrowTable`, to hand to whichever
Arrow implementation you already link.

```cpp
#include <xl/excelreader_arrow.hpp>

auto workbook = xl::Workbook::open("book.xlsx");
auto table = xl::to_arrow<Row>(workbook->sheet(0));
// table->array / table->schema are a top-level struct array; both release in ~ArrowTable.
```

Batched: `xl::to_arrow_stream<T>` is `xl::to_arrow<T>` delivered a batch at a time. `next()` returns
an `xl::ArrowTable` holding the batch and its own copy of the schema (`schema()` returns one whose
`array` is empty); each releases exactly once and outlives the stream that produced it:

```cpp
auto stream = xl::to_arrow_stream<Row>(workbook->sheet(0), {.batch_size = 10'000});
while (true)
{
    auto batch = stream->next();
    if (!batch || !batch->has_value()) { break; }
    // (*batch)->array is this one batch
}
```

Same `header_row`/`batch_size` meaning as `xl::parse_batches` above.

## Writing

Two layers, mirroring the two on the reading side.

`xl::write_sheet<T>` uses the same `xl::ExcelMapper<T>` specialization `xl::parse<T>` reads
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

auto written = xl::write_columns("out.xlsx", columns);
```

Both functions also write into memory: pass a `std::vector<std::uint8_t>&` instead of a path and it
is replaced with the file's bytes. There is no extension to infer the format from, so name it:

```cpp
std::vector<std::uint8_t> bytes;
auto written = xl::write_sheet(bytes, rows, {.format = XL_FORMAT_XLSX});
```

One constructor per column type: `i64_column`, `f64_column`, `bool_column`, `date_column`,
`time_column`, `timestamp_column`, and `string_column` (which takes an `int32` offsets span of
`rows + 1` entries plus the UTF-8 blob).

A nullable column is a values buffer plus an LSB-first validity bitmap — bit set means the row is
valid — passed as the last argument to any of those constructors. `write_columns` checks the bitmap
is long enough for the row count before calling: the ABI takes it without a length and reads
`(rows + 7) / 8` bytes on trust, so a short one would be a buffer overrun. On the struct side,
declare the field `std::optional<T>` and `write_sheet` builds the bitmap for you.

`xl::WriteOptions` sets the format, the sheet name, the CSV dialect, and the XLS/XLSB and XLSX/XLSB
toggles. Its `format` defaults to `XL_FORMAT_AUTO`, which takes the format from the path's
extension (.xlsx, .xlsb, .xls or .csv, in any case); an unrecognized extension, or AUTO with a vector
destination, fails the write rather than silently choosing a format.

For XLSX and XLSB, `degree_of_parallelism` formats and compresses batches of `batch_rows` rows on
several threads: `0` uses every processor, `1` (the default) writes serially, and `batch_rows = 0`
keeps the library default of 8,192. The columns must not change while the write runs. With
`use_shared_strings` on, two parallel runs give the same content in files that differ byte for byte.
XLS, CSV and `xl::WriterHandle` ignore both fields.

```cpp
xl::write_columns("big.xlsx", columns, {.degree_of_parallelism = 0});
```

`write_sheet` walks the range once and appends each field to its own column buffer, with the
per-field dispatch resolved at compile time. That transpose is the only copy it makes, and it is
what the ABI's columnar shape costs a row-shaped caller; `write_columns` pays nothing.

### Writing cell by cell

`xl::WriterHandle` writes a workbook as calls arrive, sheet by sheet, without holding it in memory
first. Call `finish()` at the end: destruction finishes too, but cannot report a failure. Opened over
a `std::vector<std::uint8_t>&`, the handle fills the vector when it finishes.

```cpp
auto writer = xl::WriterHandle::open("out.xlsx").value();
writer.start_sheet("Totals");
writer.start_row();
writer.write(std::string_view("widgets"));
writer.write(int64_t{7});
writer.write(std::optional<double>{}); // a blank cell
writer.end_row();
writer.end_sheet();
if (auto done = writer.finish(); !done) { /* done.error().message */ }
```

Every call returns `std::expected<void, xl::Error>`; one out of order fails instead of corrupting the
file.

## Aggregating a CSV in parallel

`xl::aggregate_csv<Acc>` folds a CSV file or buffer across worker threads. Each partition starts from
a copy of the seed, `accumulate` sees each record, and `combine` merges partitions:

```cpp
struct RowCount
{
    int64_t rows = 0;
    int32_t accumulate(xl::RowView) noexcept { ++rows; return XL_OK; }
    int32_t combine(RowCount &next) noexcept { rows += next.rows; return XL_OK; }
};

auto counted = xl::aggregate_csv("big.csv", RowCount{}, {.csv_delimiter = ';'}, /*header_row*/ 1);
```

The dialect (`csv_delimiter`, `csv_quote`, `csv_detect_bom`) and `max_cell_bytes` come from
`xl::OpenOptions`; the header row and the degree of parallelism (0 = every processor) follow it.

## Bounds and ABI

`TableView::operator[]` is unchecked, like `std::vector`'s. Use `TableView::at(row)`, which returns
`std::optional<T>` and is `nullopt` outside `[0, size())`.

Every entry point that starts native work (`xl::Workbook::open`, `write_columns`, `write_sheet`,
`WriterHandle::open`, `encrypt_package` and `aggregate_csv`) first checks the loaded library's
`xl_abi_version()` against the `XL_ABI_VERSION` this header was compiled against, and fails with an
explanatory `xl::Error` rather than reading native memory through a layout that may have changed.
`xl::abi_version()` exposes the loaded revision directly.

## Benchmarks

Google Benchmark suite in `benchmarks/`, opt-in via `-DEXCELREADER_BUILD_BENCHMARKS=ON`. Measured
on Windows 10 (22H2), 16 logical CPUs @ 3.39 GHz, MSVC 19.51 (Release), Google Benchmark v1.9.1,
`--benchmark_repetitions=10` (means shown).

| Benchmark | RealExcel.xlsb (100 rows) | 65K_Records_Data.xlsb (65,535 rows) |
|---|---:|---:|
| `open` (from memory) | 23.4 µs | 1.17 ms |
| `parse` (4 or 6 bound columns) | 44.8 µs | 23.3 ms |
| `parse_batches`, batches of 1,000 / 10,000 rows | - | 22.4 ms / 22.9 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 68.8 µs | 572.7 µs |

`open` here is the byte-span overload, which copies the caller's buffer before returning (the caller may free
it afterwards), so it scales with file size: 1.17 ms is mostly copying the 3.7 MB fixture. Opening by
path reads only the package directory and stays nearly flat — see the Rust suite below.
`parse` scales linearly with rows × columns, at roughly 360 ns/row here. Draining the same
sheet through the batched typed reader costs the same as one whole-sheet `parse` while holding
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
| ExcelReader (`parse<FullRow>`) | 61.3 ms |
| DuckDB (`read_xlsx`, summing every column in SQL) | 413.8 ms |
| xlsxio (`xlsxioread_sheet_next_cell_*`) | 497.4 ms |
| xlnt (`worksheet::rows()` + `cell::to_string()`/`value<double>()`) | 2,335.0 ms |

ExcelReader is ~6.7x faster than DuckDB, ~8.1x faster than xlsxio and ~38x faster than xlnt on this
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

- **Typed:** `parse<FullRow>` against zsv plus the same conversions done by hand (`std::string`
  for text, ISO dates, `std::from_chars` for the integers and doubles), one owned `FullRow` per row.
  All three produce the same checksum.
- **Cells:** every cell's bytes counted. ExcelReader goes through `RowCursor` (`xl_rows_next_view`), zsv
  through its row handler. zsv's checksum is 3 bytes higher because it keeps the BOM in the first
  header cell.

Measured on the same machine as above, GCC 16.2.0 (MSYS2 UCRT64, `-O3 -mavx2` for zsv), Release,
`--benchmark_repetitions=10` (means shown):

| Workload | ExcelReader | zsv `compat` | zsv `fast` |
|---|---:|---:|---:|
| Typed, 14 columns into `FullRow` | 27.4 ms | 22.6 ms | 20.0 ms |
| Cells, byte count | 11.6 ms | 5.8 ms | 3.5 ms |

**zsv is faster on both.** Typed, it is ~1.2x (`compat`) and ~1.4x (`fast`) faster than ExcelReader.
Counting cells, it is ~2.0x and ~3.3x faster.

The cells gap is the C ABI, not the CSV parser. `xl_rows_next_view` copies every cell's value into a
native row the cursor owns, one call per row, and that alone costs ~10.6 ms here (calling it without
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
| `BM_WriteColumns` (pre-transposed) | 33.2 ms | 1.98 M/s |
| `BM_WriteSheet` (from `std::vector<Row>`) | 40.3 ms | 1.63 M/s |

The transpose costs ~21% here. It is not free, but it is not the dominant cost of producing the
file — see the comparison below, where the same two cases over 14 columns land within ~14–16%.

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
| ExcelReader (`xl::write_columns`, pre-transposed) | 64.2–66.9 ms | 63.9–68.2 ms |
| ExcelReader (`xl::write_sheet<FullRow>`) | 74.8–76.0 ms | 74.7–76.4 ms |
| DuckDB (`COPY ... TO ... WITH (FORMAT xlsx)`) | 831.4–831.6 ms | 828.1 ms |
| libxlsxwriter (`worksheet_write_string`/`_number`) | 1,206.7 ms | 1,187.5 ms |
| xlsxio (`xlsxiowrite_add_cell_*`) | 2,405.1 ms | 1,125.0 ms |
| xlnt (`worksheet::cell().value()` + `save()`) | 5,525.3 ms | 5,515.6 ms |

Same machine as above; Google Benchmark's own iteration counts, no `--benchmark_repetitions` (each
iteration writes a whole 65,535-row file, so the slower cases run once or a handful of times). xlnt
and xlsxio only build into `excelreader_cpp_write_compare_benchmarks`; libxlsxwriter only into
`excelreader_cpp_write_compare_lxw_benchmarks` (see the known issue below) — the ExcelReader and
DuckDB rows appear in both, and the small ranges above are those two independent runs, not repeated
sampling within one run.

`write_sheet` — the matched-work number, since it starts from the same `std::vector<FullRow>` every
competitor is handed — is ~11x faster than DuckDB, ~16x faster than libxlsxwriter, ~32x
faster than xlsxio, and ~74x faster than xlnt on wall time.

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
  1,015.6 ms and 843.8 ms in the next two, 953.1 ms in the one after and 1,125.0 ms in the latest, while its wall time stayed at ~2.4 s throughout. Its wall time is dominated by I/O wait —
  CPU is well under half of wall, against every other case here being CPU-bound (wall ≈ CPU) — so
  quote the wall-time ratio, which is what a caller waits. Against this run's CPU time the gap from
  `write_sheet` to xlsxio is ~15x, not ~32x.
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

The two ExcelReader cases land within ~14–16% of each other. The transpose's share grew as producing
the file got cheaper, but it is still a minority of the cost. Reach for
`write_columns` when your data is already columnar, but `write_sheet` is not the slow path.
