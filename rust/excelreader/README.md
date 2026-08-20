# excelreader (Rust)

Read Excel/CSV workbooks via ExcelReader's native ABI: opening a workbook (from a path or memory,
with the full open-options surface), sheet navigation, schema inference, and schema-driven typed
parse. No writing, no Arrow, no row-by-row decode yet - see the root README's Python section for
what those look like.

## Usage

```toml
[dependencies]
excelreader = "2.1"
```

```rust
use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Name")]
    name: String,
    #[excel(name = "Count")]
    count: u32,
}

let mut workbook = Workbook::open("book.xlsx")?;
let table = parse_sheet::<Row>(&mut workbook, 1)?;
for row in table.iter() { /* ... */ }
```

`parse_sheet` takes `&mut Workbook` because it consumes the workbook's shared row cursor.

### Field types

| Field type | Column |
|---|---|
| `String` | `XL_T_STRING` |
| `i8`..`i64`, `isize`, `u8`..`u64`, `usize` | `XL_T_I64` |
| `f32`, `f64` | `XL_T_F64` |
| `bool` | `XL_T_BOOL` |
| `Date` / `chrono::NaiveDate` | `XL_T_DATE` |
| `Time` / `chrono::NaiveTime` | `XL_T_TIME` |
| `Timestamp` / `chrono::NaiveDateTime` | `XL_T_TIMESTAMP` |

`Option<T>` of any of the above makes the field `None` when the cell is null; a non-`Option` field is
left at its `Default` instead.

Narrower integers convert through `TryFrom` and panic on a value that does not fit, rather than
wrapping silently. `Date`/`Time`/`Timestamp` are dependency-free newtypes over the exact wire
representation; enable the `chrono` feature to use `chrono`'s calendar types directly:

```toml
excelreader = { version = "2.1", features = ["chrono"] }
```

### Sheets and schema inference

```rust
let mut workbook = Workbook::open("book.xlsx")?;
for (index, name) in workbook.sheet_names()?.iter().enumerate() {
    println!("{index}: {name}");
}
workbook.move_to_sheet(1)?;

// Guess a schema from the header row plus a sample of the data, before committing to one.
for column in workbook.infer_schema(1, 100)? {
    println!("{:?} -> type {}", column.name, column.column_type);
}
```

`Workbook::open_with` takes an explicit format and `OpenOptions`; `Workbook::open_memory` reads from
a byte slice. Note that format sniffing does not detect CSV - pass `XL_FORMAT_CSV` explicitly.

## Bounds and panics

`TableView::get` returns `Option<T>` and is `None` outside `0..len()`. The `column_*` accessors used
by generated bindings panic on an out-of-range row, a column type that does not match the binding, or
a string the native library returned as non-UTF-8 - each is a contract violation rather than
recoverable input.

## Build notes

`build.rs` downloads the native `ExcelReader.Native` binary matching your target from the crate's
matching GitHub Release. Set `EXCELREADER_NATIVE_LIB_DIR` to a directory containing a locally-built
copy instead (e.g. from `dotnet publish ../../src/ExcelReader.Native -r win-x64`) to skip that
download - useful when building from source before a release exists yet.

Every constructor first checks the loaded library's `xl_abi_version()` against the `XL_ABI_VERSION`
this crate was compiled against, and fails with a explanatory error rather than reading native memory
through a layout that may have changed.

## Benchmarks

Criterion suite in `benches/`. Measured on Windows 10 (22H2), 16 logical CPUs @ 3.39 GHz,
rustc 1.97.1 (Release), Criterion 0.5, 100 samples per benchmark (medians shown).

`benches/parse_bench.rs` - `open`/`parse_sheet`/`infer_schema`, same methodology as the C++ suite:

| Benchmark | RealExcel.xlsb (100 rows) | 65K_Records_Data.xlsb (65,535 rows) |
|---|---:|---:|
| `open` | 79.4 µs | 92.8 µs |
| `parse_sheet` (2 or 6 bound columns) | 79.0 µs | 39.3 ms |
| `infer_schema` (sample 100 / 1,000 rows) | 147.9 µs | 1.19 ms |

`open` stays nearly flat across the 655x row-count jump (+17%) - XLSB's header carries its
dimensions/index, so opening costs metadata, not row data. `parse_sheet` scales linearly with
rows × columns; `infer_schema` scales with its sample size, not the file's total row count.

`benches/compare_bench.rs` compares against [calamine](https://github.com/tafia/calamine) reading
`65K_Records_Data.{xlsx,xlsb}` in full (all 14 columns, 65,535 rows). Both sides decode every cell
into an owned value (`String` for text) and fold it into one accumulator, so neither side gets a
zero-copy advantage the other can't take:

| Format | ExcelReader (`parse_sheet::<FullRow>`) | calamine (`worksheet_range` + `Data` match) |
|---|---:|---:|
| XLSX | 123.5 ms | 273.4 ms |
| XLSB | 70.6 ms | 84.9 ms |

ExcelReader is ~2.2x faster than calamine for XLSX and ~1.2x faster for XLSB on this workload -
calamine is a fast, well-optimized reader in its own right, so the gap is real but not the order
of magnitude seen against slower libraries.

Run locally:

```bash
cd rust
EXCELREADER_NATIVE_LIB_DIR=/path/to/native/lib/dir cargo bench -p excelreader
```

`EXCELREADER_NATIVE_LIB_DIR` should point at a directory containing a locally-built
`ExcelReader.Native.{dll,so,dylib}` - see [Build notes](#build-notes) above. Pass
`--bench parse_bench` or `--bench compare_bench` to run one suite only.
