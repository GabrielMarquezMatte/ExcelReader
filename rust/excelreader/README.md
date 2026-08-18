# excelreader (Rust)

Read Excel/CSV workbooks via ExcelReader's native ABI. Phase 1: opening a workbook and
schema-driven typed parse only - no writing, no Arrow, no row-by-row decode.

## Usage

```toml
[dependencies]
excelreader = "2.1"
```

```rust
use excelreader::workbook::{column_str, column_i64, parse_sheet, ColumnBinding, ExcelMapper, Workbook};
use excelreader::{XL_T_STRING, XL_T_I64};

#[derive(Default)]
struct Row { name: String, count: i64 }

impl ExcelMapper for Row {
    fn bindings() -> Vec<ColumnBinding<Self>> {
        vec![
            ColumnBinding { name: "Name", xl_type: XL_T_STRING, assign: |r, c, i| r.name = column_str(c, i).to_string() },
            ColumnBinding { name: "Count", xl_type: XL_T_I64, assign: |r, c, i| r.count = column_i64(c, i) },
        ]
    }
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
