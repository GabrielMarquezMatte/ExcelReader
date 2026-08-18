# excelreader (Rust)

Read Excel/CSV workbooks via ExcelReader's native ABI. Phase 1: opening a workbook and
schema-driven typed parse only - no writing, no Arrow, no row-by-row decode.

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
    count: i64,
}

let workbook = Workbook::open("book.xlsx")?;
let table = parse_sheet::<Row>(&workbook, 1)?;
for row in table.iter() { /* ... */ }
```

## Build notes

`build.rs` downloads the native `ExcelReader.Native` binary matching your target from the crate's
matching GitHub Release. Set `EXCELREADER_NATIVE_LIB_DIR` to a directory containing a locally-built
copy instead (e.g. from `dotnet publish ../../src/ExcelReader.Native -r win-x64`) to skip that
download - useful when building from source before a release exists yet.
