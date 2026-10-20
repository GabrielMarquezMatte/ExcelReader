![ExcelReader logo](https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png)

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

It moves data, not presentation. Reading gives each cell's value and type, with a formula cell read as
its last calculated result, plus sheet names, visibility and Excel tables. Formula text, merged cells,
comments, hyperlinks, charts, images and conditional formatting are not read. Writing covers values,
column widths and cell styles limited to bold, italic and a number format. To edit an existing workbook
or produce formatted reports, use a library built for document manipulation.

## Install

```bash
dotnet add package ExcelReader.NET
```

```csharp
using ExcelReader.Core.Reader;

using IExcelWorkbook workbook = Excel.Open("book.xlsx");
foreach (Row row in workbook.FirstSheet)
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
- [Writing workbooks](docs/guide/writing.md) — XLSX, XLSB (BIFF12) and XLS (BIFF8) writers, cell styles, typed records, prefetch compression, and parallel record writes.
- [CSV](docs/guide/csv.md) — reading, dialect sniffing, parallel parsing, and writing delimited text.
- [Encrypted workbooks](docs/guide/encryption.md) — opening password-protected packages and encrypting written ones.
- [Reading untrusted files](docs/guide/untrusted-input.md) — the resource limits that bound what one hostile file can cost, and how to reject one cleanly.
- [Migrating to v7](docs/guide/migrating-to-v7.md) — porting `CsvParallel` code to `SheetParallel`, which now reads XLSX and XLSB on several threads too.
- [Migrating to v6](docs/guide/migrating-to-v6.md) — porting v5 reader code to workbooks and sheets, and reading sheets in parallel.
- [Benchmarks](docs/performance/benchmarks.md) — throughput and allocation against Sep, Sylvan, ExcelDataReader, OfficeIMO and SpreadCheetah. [Live results](https://gabrielmarquezmatte.github.io/ExcelReader/dev/bench/).

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
for stdout), `--delimiter|-d <char>` (csv output), `--input-delimiter <char>` (csv input). Run
`excelreader <command> --help` for the full list.

`--sheet` matches a sheet name first and falls back to a 0-based index, so a sheet literally named
`2` is reachable by name. A `.csv` input has its dialect sniffed unless `--input-delimiter` names one,
so a `;`-separated export reads as columns rather than as one field; `--delimiter` stays on the output
side, so `convert data.csv -o out.csv -d ';'` still means "read as sniffed, write semicolons".
`sheets` and `schema` write plain tab-separated text. `convert` reports progress on stderr while it
runs, unless stderr is redirected. Exit codes are `0` ok and `1` failure; results go to stdout and
errors to stderr, so `convert` is safe to pipe.

## Notes

- Reads `.xlsx`, `.xlsb` (BIFF12), `.xls` (BIFF8), and `.csv`; writes `.xlsx`, `.xlsb`, `.xls`, and `.csv`.
- `Excel.Open`/`OpenAsync` auto-detect XLSX/XLSB/XLS from the file signature. CSV has none, so pass `ExcelFileFormat.Csv` to open delimited text through the same entry point; `ExcelReaderOptions.Csv` carries the dialect, and `CsvReaderOptions.SniffDialect` infers it from the source.
- A workbook hands out its sheets: `FirstSheet`, `Sheets[index]` or `SheetAt(index)`, and `TryGetSheet(name)`. CSV is a workbook with one unnamed sheet. Coming from v5? See [Migrating to v6](docs/guide/migrating-to-v6.md).
- Sheets of one workbook can be read in parallel, one enumerator per thread. One sheet of any format (XLSX, XLSB or CSV) can also be read on several threads with `SheetParallel`.
- `sheet.Visibility` reports whether a sheet is hidden (`Visible`, `Hidden`, `VeryHidden`). Hidden sheets still enumerate their rows — filter on the value if you want them skipped. The writers take the same value via `AddSheet(name, visibility)`, and reject a workbook whose sheets would all be hidden.
- Missing cells in sparse rows are exposed as empty cells.
- String conversion allocates only when you call `GetString()`.
- The XLSX scanner accepts the SpreadsheetML shapes commonly emitted by non-Excel producers, including single-quoted attributes, comments in `sheetData`, and CDATA text runs.
- Readers bound untrusted input by default: 512 MB total decompressed ZIP data, 32 MB per cell/row value buffer, and 128 MB for shared strings. Pass `ExcelReaderOptions` to the `Excel.From*`/`Excel.Open*` factories to tune these limits; set a limit to `0` to opt out and restore unlimited behavior for that limit. `CsvReader` has its own `CsvReaderOptions.MaxCellBytes` (default 32 MB) for the same purpose.
- Each ZIP entry of an XLSX or XLSB package is checked against its CRC-32 as it is read, so a damaged file throws `InvalidDataException` instead of returning wrong cells. `ExcelReaderOptions.VerifyZipChecksums = false` turns the check off.
- The XLSX writer emits a compact workbook with strings, numbers, booleans, dates, and blank cells; shared strings are opt-in.
- The XLSB writer emits BIFF12 workbook parts inside the standard XLSB ZIP package; shared strings are opt-in.
- `WriteRecordsParallelAsync` and `WriteTableParallelAsync` format and compress batches of records on several cores and write them in order, for XLSX and XLSB. On a Ryzen 7 5700X a 1,000,000-row sheet writes in 46 ms as XLSX with 4 batches in flight and 31 ms with 8, against 172 ms serially; XLSB takes 31 ms and 21 ms against 115 ms. Memory stays at about 3 MB per batch in flight. See [Write records in parallel](docs/guide/writing.md#write-records-in-parallel-xlsxxlsb) and the [benchmarks](docs/performance/benchmarks.md#parallel-record-writes).
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
- Python package: `pip install excelreader-native` — see [`python/README.md`](python/README.md).
- C++ package: header-only CMake wrapper, `xl::Workbook`/`xl::parse_sheet`/`xl::write_sheet` over the same ABI — see [`cpp/README.md`](cpp/README.md).
- Rust crate: safe `Workbook`/`parse_sheet`/`write_sheet` bindings, downloadable via `cargo add excelreader` — see [`rust/excelreader/README.md`](rust/excelreader/README.md).

```python
from excelreader import open_workbook

with open_workbook("book.xlsx") as workbook:
    for row in workbook.sheets[0].rows():
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

Row-by-row decoded reads are available from all three bindings — Python as `Sheet.rows()`, C++
as `xl::Sheet::rows()`, Rust as `Sheet::rows()`. Python additionally exposes
`read_all_columnar()` over `xl_rows_read_all_blob`, which the other two do not wrap.

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
