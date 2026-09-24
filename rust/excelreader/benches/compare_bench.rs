
use calamine::{open_workbook_auto, Data, Reader};
use criterion::{criterion_group, criterion_main, Criterion};
use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};
use excelreader::Date;

#[derive(Default, ExcelMapper)]
struct FullRow {
    #[excel(name = "Region")]
    region: String,
    #[excel(name = "Country")]
    country: String,
    #[excel(name = "Item Type")]
    item_type: String,
    #[excel(name = "Sales Channel")]
    sales_channel: String,
    #[excel(name = "Order Priority")]
    order_priority: String,
    #[excel(name = "Order Date")]
    order_date: Date,
    #[excel(name = "Order ID")]
    order_id: i64,
    #[excel(name = "Ship Date")]
    ship_date: Date,
    #[excel(name = "Units Sold")]
    units_sold: i64,
    #[excel(name = "Unit Price")]
    unit_price: f64,
    #[excel(name = "Unit Cost")]
    unit_cost: f64,
    #[excel(name = "Total Revenue")]
    total_revenue: f64,
    #[excel(name = "Total Cost")]
    total_cost: f64,
    #[excel(name = "Total Profit")]
    total_profit: f64,
}

fn accumulate_full_row(row: &FullRow) -> i64 {
    row.region.len() as i64
        + row.country.len() as i64
        + row.item_type.len() as i64
        + row.sales_channel.len() as i64
        + row.order_priority.len() as i64
        + row.order_date.days_since_epoch as i64
        + row.order_id
        + row.ship_date.days_since_epoch as i64
        + row.units_sold
        + row.unit_price as i64
        + row.unit_cost as i64
        + row.total_revenue as i64
        + row.total_cost as i64
        + row.total_profit as i64
}

fn xlsx_path() -> String {
    concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsx"
    )
    .to_string()
}

fn xlsb_path() -> String {
    concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb"
    )
    .to_string()
}

fn accumulate_data_cell(acc: &mut i64, cell: &Data) {
    match cell {
        Data::String(s) => *acc += s.len() as i64,
        Data::Int(i) => *acc += *i,
        Data::Float(f) => *acc += *f as i64,
        Data::DateTime(dt) => *acc += dt.as_f64() as i64,
        _ => {}
    }
}

fn bench_excelreader(c: &mut Criterion, group: &str, path: String, format: i32) {
    c.bench_function(group, |b| {
        b.iter_batched(
            || Workbook::open_with(&path, format, None).expect("open must succeed"),
            |mut workbook| {
                let table =
                    parse_sheet::<FullRow>(&mut workbook, 1).expect("parse_sheet must succeed");
                let mut acc = 0i64;
                for row in table.iter() {
                    acc += accumulate_full_row(&row);
                }
                std::hint::black_box(acc);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

fn bench_excelreader_parse_only(c: &mut Criterion, group: &str, path: String, format: i32) {
    c.bench_function(group, |b| {
        b.iter_batched(
            || Workbook::open_with(&path, format, None).expect("open must succeed"),
            |mut workbook| {
                let table =
                    parse_sheet::<FullRow>(&mut workbook, 1).expect("parse_sheet must succeed");
                std::hint::black_box(table.len());
                std::hint::black_box(table);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

fn bench_calamine(c: &mut Criterion, group: &str, path: String) {
    c.bench_function(group, |b| {
        b.iter(|| {
            let mut workbook = open_workbook_auto(&path).expect("open must succeed");
            let range = workbook
                .worksheet_range_at(0)
                .expect("sheet 0 must exist")
                .expect("worksheet_range must succeed");
            let mut acc = 0i64;
            for row in range.rows() {
                for cell in row {
                    accumulate_data_cell(&mut acc, cell);
                }
            }
            std::hint::black_box(acc);
        });
    });
}

fn bench_excelreader_xlsx(c: &mut Criterion) {
    bench_excelreader(c, "excelreader_xlsx_full", xlsx_path(), excelreader::XL_FORMAT_XLSX);
}

fn bench_calamine_xlsx(c: &mut Criterion) {
    bench_calamine(c, "calamine_xlsx_full", xlsx_path());
}

fn bench_excelreader_xlsb(c: &mut Criterion) {
    bench_excelreader(c, "excelreader_xlsb_full", xlsb_path(), excelreader::XL_FORMAT_XLSB);
}

fn bench_calamine_xlsb(c: &mut Criterion) {
    bench_calamine(c, "calamine_xlsb_full", xlsb_path());
}

fn bench_excelreader_xlsx_parse_only(c: &mut Criterion) {
    bench_excelreader_parse_only(
        c,
        "excelreader_xlsx_parse_only",
        xlsx_path(),
        excelreader::XL_FORMAT_XLSX,
    );
}

fn bench_excelreader_xlsb_parse_only(c: &mut Criterion) {
    bench_excelreader_parse_only(
        c,
        "excelreader_xlsb_parse_only",
        xlsb_path(),
        excelreader::XL_FORMAT_XLSB,
    );
}

criterion_group!(
    benches,
    bench_excelreader_xlsx,
    bench_calamine_xlsx,
    bench_excelreader_xlsb,
    bench_calamine_xlsb,
    bench_excelreader_xlsx_parse_only,
    bench_excelreader_xlsb_parse_only
);
criterion_main!(benches);
