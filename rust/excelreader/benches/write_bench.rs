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

use criterion::{criterion_group, criterion_main, Criterion};
use excelreader::ReadOptions;
use excelreader::{
    write_columns, write_sheet, Column, ColumnData, Date, ExcelMapper, Workbook, WriteOptions,
};
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
    let workbook = Workbook::open(FIXTURE).expect("the fixture must open");
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<Row>(&ReadOptions::new())
        .expect("the fixture must parse");
    let rows: Vec<Row> = table.iter().collect();
    assert!(!rows.is_empty(), "the fixture parsed to zero rows");
    rows
}

fn output_path(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!("excelreader-bench-{}-{name}", std::process::id()));
    path
}

/// One string column's offsets and concatenated bytes.
fn strings<'a>(values: impl Iterator<Item = &'a str>) -> (Vec<i32>, Vec<u8>) {
    let mut offsets = vec![0];
    let mut data = Vec::new();
    for value in values {
        data.extend_from_slice(value.as_bytes());
        offsets.push(data.len() as i32);
    }
    (offsets, data)
}

fn str_column<'a>(name: &'a str, (offsets, data): &'a (Vec<i32>, Vec<u8>)) -> Column<'a> {
    Column::new(Some(name), ColumnData::Str { offsets, data })
}

fn benchmark_write(criterion: &mut Criterion) {
    let rows = load_rows();
    let region = strings(rows.iter().map(|r| r.region.as_str()));
    let country = strings(rows.iter().map(|r| r.country.as_str()));
    let item_type = strings(rows.iter().map(|r| r.item_type.as_str()));
    let order_date: Vec<i32> = rows.iter().map(|r| r.order_date.days_since_epoch).collect();
    let order_id: Vec<i64> = rows.iter().map(|r| r.order_id).collect();
    let units_sold: Vec<i64> = rows.iter().map(|r| r.units_sold).collect();
    let total_revenue: Vec<f64> = rows.iter().map(|r| r.total_revenue).collect();
    let column = |name, data| Column::new(Some(name), data);
    let borrowed = [
        str_column("Region", &region),
        str_column("Country", &country),
        str_column("Item Type", &item_type),
        column("Order Date", ColumnData::Date(&order_date)),
        column("Order ID", ColumnData::I64(&order_id)),
        column("Units Sold", ColumnData::I64(&units_sold)),
        column("Total Revenue", ColumnData::F64(&total_revenue)),
    ];
    let options = WriteOptions::new();

    let mut group = criterion.benchmark_group("write_xlsx_65k");
    group.sample_size(20);

    group.bench_function("columns", |b| {
        let path = output_path("columns.xlsx");
        b.iter(|| {
            write_columns(&path, std::hint::black_box(&borrowed), &options)
                .expect("write_columns must succeed");
        });
        std::fs::remove_file(&path).ok();
    });

    group.bench_function("sheet", |b| {
        let path = output_path("sheet.xlsx");
        b.iter(|| {
            write_sheet(&path, std::hint::black_box(&rows), &options)
                .expect("write_sheet must succeed");
        });
        std::fs::remove_file(&path).ok();
    });

    group.bench_function("rust_xlsxwriter", |b| {
        let path = output_path("xlsxwriter.xlsx");
        b.iter(|| {
            let mut workbook = rust_xlsxwriter::Workbook::new();
            let sheet = workbook.add_worksheet();
            for (index, row) in std::hint::black_box(&rows).iter().enumerate() {
                let r = (index + 1) as u32;
                sheet.write_string(r, 0, &row.region).unwrap();
                sheet.write_string(r, 1, &row.country).unwrap();
                sheet.write_string(r, 2, &row.item_type).unwrap();
                sheet
                    .write_number(r, 3, row.order_date.days_since_epoch as f64)
                    .unwrap();
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
