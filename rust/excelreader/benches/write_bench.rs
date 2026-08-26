//! Write benchmarks over tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsx (65,535 data
//! rows, 14 columns), the same fixture the C++, Python and .NET suites use.
//!
//! WORK IS NOT MATCHED across all three groups, and the difference is structural rather than an
//! oversight:
//!
//!   * `columns` hands the ABI buffers that are already columnar. Nothing is transposed, nothing
//!     is copied. This is the ceiling, and no cell-at-a-time API can be compared to it fairly.
//!   * `sheet` starts from a Vec<Row> and pays the row-to-column transpose. This is the
//!     matched-work sibling: it does the same job rust_xlsxwriter does, from the same starting
//!     shape.
//!   * `rust_xlsxwriter` writes cell by cell through an API that also owns styling and formula
//!     support this library does not expose.
//!
//! Read `sheet` against `rust_xlsxwriter`. Read `columns` only against `sheet`, as the cost of
//! having row-shaped data in the first place.
//!
//! Machine: state the CPU, OS and toolchain version alongside any number published from this file.

use criterion::{black_box, criterion_group, criterion_main, Criterion};
use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};
use excelreader::writer::{write_columns, write_sheet, Column, OwnedColumn};
use excelreader::{Date, XL_FORMAT_XLSX};
use std::path::{Path, PathBuf};

const FIXTURE: &str = "../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsx";

#[derive(Default, Clone, ExcelMapper)]
struct Row {
    #[excel(name = "Region")]
    region: String,
    #[excel(name = "Country")]
    country: String,
    #[excel(name = "Item Type")]
    item_type: String,
    #[excel(name = "Order Date")]
    order_date: Date,
    #[excel(name = "Order ID")]
    order_id: i64,
    #[excel(name = "Units Sold")]
    units_sold: i64,
    #[excel(name = "Total Revenue")]
    total_revenue: f64,
}

/// Reads the fixture once into row structs. Panics rather than silently benchmarking nothing when
/// the fixture is missing or empty - a suite that measures an empty input is worse than no suite.
fn load_rows() -> Vec<Row> {
    let path = Path::new(FIXTURE);
    assert!(
        path.exists(),
        "missing benchmark fixture {FIXTURE} - run from rust/excelreader, and check the file is \
         present in the repo"
    );
    let mut workbook = Workbook::open(FIXTURE).expect("the fixture must open");
    let table = parse_sheet::<Row>(&mut workbook, 1).expect("the fixture must parse");
    let rows: Vec<Row> = table.iter().collect();
    assert!(!rows.is_empty(), "the fixture parsed to zero rows");
    rows
}

fn output_path(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!("excelreader-bench-{}-{name}", std::process::id()));
    path
}

fn benchmark_write(criterion: &mut Criterion) {
    let rows = load_rows();
    // Transposed once, outside the measured region: the `columns` group exists to measure the
    // write, not the transpose the `sheet` group already covers.
    let owned: Vec<OwnedColumn> =
        <Row as excelreader::writer::ExcelWriter>::to_columns(&rows).expect("transpose must succeed");
    let borrowed: Vec<Column<'_>> = owned.iter().map(OwnedColumn::as_column).collect();

    let mut group = criterion.benchmark_group("write_xlsx_65k");
    group.sample_size(20);

    group.bench_function("columns", |b| {
        let path = output_path("columns.xlsx");
        let target = path.to_str().expect("temp path must be UTF-8");
        b.iter(|| {
            write_columns(target, XL_FORMAT_XLSX, black_box(&borrowed), None)
                .expect("write_columns must succeed");
        });
        std::fs::remove_file(&path).ok();
    });

    group.bench_function("sheet", |b| {
        let path = output_path("sheet.xlsx");
        let target = path.to_str().expect("temp path must be UTF-8");
        b.iter(|| {
            write_sheet(target, XL_FORMAT_XLSX, black_box(&rows), None)
                .expect("write_sheet must succeed");
        });
        std::fs::remove_file(&path).ok();
    });

    group.bench_function("rust_xlsxwriter", |b| {
        let path = output_path("xlsxwriter.xlsx");
        b.iter(|| {
            let mut workbook = rust_xlsxwriter::Workbook::new();
            let sheet = workbook.add_worksheet();
            for (index, row) in black_box(&rows).iter().enumerate() {
                let r = (index + 1) as u32;
                sheet.write_string(r, 0, &row.region).unwrap();
                sheet.write_string(r, 1, &row.country).unwrap();
                sheet.write_string(r, 2, &row.item_type).unwrap();
                sheet.write_number(r, 3, row.order_date.days_since_epoch as f64).unwrap();
                sheet.write_number(r, 4, row.order_id as f64).unwrap();
                sheet.write_number(r, 5, row.units_sold as f64).unwrap();
                sheet.write_number(r, 6, row.total_revenue).unwrap();
            }
            workbook.save(&path).unwrap();
        });
        std::fs::remove_file(&path).ok();
    });

    group.finish();
}

criterion_group!(benches, benchmark_write);
criterion_main!(benches);