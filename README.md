# ExcelReader

[![CI](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/ci.yml)
[![CodeQL](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml/badge.svg?branch=master)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/codeql.yml)
[![Release](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml/badge.svg)](https://github.com/GabrielMarquezMatte/ExcelReader/actions/workflows/release.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![Downloads](https://img.shields.io/nuget/dt/ExcelReader.NET.svg)](https://www.nuget.org/packages/ExcelReader.NET)
[![License](https://img.shields.io/github/license/GabrielMarquezMatte/ExcelReader.svg)](LICENSE)
[![Benchmarks](https://img.shields.io/badge/benchmarks-GitHub%20Pages-informational)](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/)

High-performance Excel reading and writing for .NET 10. Reads `.xlsx`, `.xlsb`, `.xls`, and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.

ExcelReader is built for streaming spreadsheet workloads where low allocations matter. It reads worksheet rows as lightweight `ref struct` values, resolves shared strings, recognizes date styles, and handles sparse cells. Workbooks open from a path, a `Stream`, or a caller-owned `ReadOnlyMemory<byte>` — the in-memory path decompresses parts directly, without a `ZipArchive` or an intermediate stream.

## Install

```bash
dotnet add package ExcelReader.NET
```

```csharp
using ExcelReader.Core.Reader;

using IExcelRowReader reader = Excel.Open("book.xlsx");
foreach (Row row in reader)
{
    foreach (RowCell cell in row.Cells)
    {
        Console.Write(cell.Value.GetString());
        Console.Write('\t');
    }
    Console.WriteLine();
}
```

## Guide

- [Reading rows](docs/guide/reading.md) — opening a workbook, enumerating rows, auto-detecting the format, async reads, prefetch decompression, and the ADO.NET `IDataReader` bridge.
- [Typed parsing](docs/guide/parsing.md) — binding rows to your own types with attributes, the compile-time generator (Native AOT / trimming), the fluent API, required columns, custom converters, and the zero-copy `ref struct` path.
- [Writing workbooks](docs/guide/writing.md) — XLSX, XLSB (BIFF12) and XLS (BIFF8) writers, cell styles, typed records, and prefetch compression.
- [CSV](docs/guide/csv.md) — reading, dialect sniffing, parallel parsing, and writing delimited text.
- [Encrypted workbooks](docs/guide/encryption.md) — opening password-protected packages and encrypting written ones.
- [Benchmarks](docs/performance/benchmarks.md) — throughput and allocation against CsvHelper, Sep, Sylvan, MiniExcel and SpreadCheetah. [Live results](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/).

## Command line

```bash
dotnet tool install --global ExcelReader.NET.Cli
```

```bash
excelreader sheets  book.xlsb                        # 0<TAB>Sheet1
excelreader schema  book.xlsb --sample-size 500      # 0<TAB>Id<TAB>Int64
excelreader convert book.xlsb --output data.csv      # convert to another format, by extension
excelreader convert book.xlsb --output data.xlsx     # .xlsx, .xlsb, .xls and .csv all work
excelreader convert book.xlsb --format xlsx | head -c 4   # to stdout, --format picks it instead
```

Flags: `--sheet|-s <name|index>`, `--header-row N` (0 = no header), `--sample-size N`,
`--output|-o <file>`, `--format|-f <xlsx|xlsb|xls|csv>` (defaults to `--output`'s extension, or csv
for stdout), `--delimiter|-d <char>` (csv only). Run `excelreader <command> --help` for the full list.

`sheets` and `schema` write plain tab-separated text. `convert` reports progress on stderr while it
runs, unless stderr is redirected. Exit codes are `0` ok and `1` failure; results go to stdout and
errors to stderr, so `convert` is safe to pipe.

## Notes

- Reads `.xlsx`, `.xlsb` (BIFF12), `.xls` (BIFF8), and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.
- Reads one sheet at a time (XLSX/XLSB/XLS); use `MoveToSheet(index)` or `TryMoveToSheet(name)` to switch sheets. CSV has no sheets.
- Missing cells in sparse rows are exposed as empty cells.
- String conversion allocates only when you call `GetString()`.
- The XLSX scanner accepts the SpreadsheetML shapes commonly emitted by non-Excel producers, including single-quoted attributes, comments in `sheetData`, and CDATA text runs.
- Readers bound untrusted input by default: 512 MB total decompressed ZIP data, 32 MB per cell/row value buffer, and 128 MB for shared strings. Pass `ExcelReaderOptions` to the `Excel.From*`/`Excel.Open*` factories to tune these limits; set a limit to `0` to opt out and restore unlimited behavior for that limit. `CsvReader` has its own `CsvReaderOptions.MaxCellBytes` (default 32 MB) for the same purpose.
- The XLSX writer emits a compact workbook with strings, numbers, booleans, dates, and blank cells; shared strings are opt-in.
- The XLSB writer emits BIFF12 workbook parts inside the standard XLSB ZIP package; shared strings are opt-in.
- The XLS writer buffers records in memory and assembles the OLE container at `EndAsync`; choose it when write throughput matters more than peak allocation.
- `Excel.EncryptPackage`/`EncryptPackageAsync` wrap a written XLSX/XLSB package in agile ECMA-376 encryption; see [Encrypted workbooks](docs/guide/encryption.md).

## Build

```bash
dotnet restore ExcelReader.slnx
dotnet build ExcelReader.slnx --configuration Release
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj --configuration Release
```

## Other languages

ExcelReader ships a NativeAOT shared library with a C ABI, so non-.NET languages can read and write
XLSX, XLSB, XLS and CSV without a .NET runtime installed.

- C ABI header: [`src/ExcelReader.Native/include/excelreader.h`](src/ExcelReader.Native/include/excelreader.h)
- Python package: [`python/`](python/README.md)
- C++ package: header-only CMake wrapper, `xl::Workbook`/`xl::parse_sheet`/`xl::write_sheet` over the same ABI — see [`cpp/README.md`](cpp/README.md).
- Rust crate: safe `Workbook`/`parse_sheet`/`write_sheet` bindings, downloadable via `cargo add excelreader` — see [`rust/excelreader/README.md`](rust/excelreader/README.md).

```python
from excelreader import open_workbook

with open_workbook("book.xlsx") as workbook:
    for row in workbook.rows():
        print([cell.value for cell in row])
```

Writing goes through one export, `xl_write_typed`: a whole sheet in a single call, from columnar
buffers the ABI borrows rather than copies. All three bindings expose it — Python as
`write_workbook`/`write_arrow`/`write_pandas`/`write_polars`, C++ as `xl::write_columns` and
`xl::write_sheet<T>`, Rust as `writer::write_columns` and `writer::write_sheet`. In C++ and Rust the
same struct mapping drives both directions, so reading a sheet and writing it back needs one
mapping, not two:

```rust
use excelreader::writer::write_sheet;
use excelreader::XL_FORMAT_XLSX;

write_sheet("out.xlsx", XL_FORMAT_XLSX, &rows, None)?;
```

```cpp
auto written = xl::write_sheet("out.xlsx", rows);   // format inferred from the extension
```

Row-by-row decoded reads are available from all three bindings — Python as `Workbook.rows()`, C++
as `xl::Workbook::rows()`, Rust as `Workbook::rows()`. Python additionally exposes
`read_all_columnar()` over `xl_read_all_blob`, which the other two do not wrap.

The Arrow export is available from Python
(`to_arrow`/`to_record_batch`), C++ (`xl::parse_arrow<T>`, in the separate `<xl/excelreader_arrow.hpp>`
header — no Apache Arrow C++ dependency, you get the raw C Data Interface pair), and Rust
(`excelreader::arrow::parse_arrow`, behind the `arrow` cargo feature, returning an `arrow::array::RecordBatch`).

All three bindings can also read a sheet a batch at a time instead of the whole sheet in one call, so
peak memory is one batch rather than one sheet — Python as `iter_parse_typed`/
`to_record_batch_reader` (plus `iter_pandas`/`iter_polars`), C++ as `xl::typed_reader<T>`/
`xl::arrow_stream<T>`, Rust as `Workbook::typed_chunks`/`arrow::parse_arrow_stream`. The .NET reader
above already streams row-by-row by construction, so there is nothing to add here for it; see each
binding's README for the chunked-reading semantics (one chunked read per workbook, `batch_size`
meaning, and what invalidates a live one).

Encrypting a written package goes through one export too, `xl_encrypt_package` — wrap a finished
plaintext XLSX/XLSB package in agile ECMA-376 encryption, given a password. All three bindings
expose it: Python as `encrypt_package`, C++ as `xl::encrypt_package`, Rust as
`writer::encrypt_package`; see each binding's README for the encrypted-workbooks section.

## Contributing

See [ARCHITECTURE.md](ARCHITECTURE.md) for a map of the codebase, [STYLEGUIDE.md](STYLEGUIDE.md) for
the code style, and [CONTRIBUTING.md](CONTRIBUTING.md) for build expectations and how to submit a
change. Security issues should go through the private channel in [SECURITY.md](SECURITY.md), not a
public issue.

## License

ExcelReader is licensed under the MIT License. See [LICENSE](LICENSE).
