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
