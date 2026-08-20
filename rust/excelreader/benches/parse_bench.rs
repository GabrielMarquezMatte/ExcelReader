use criterion::{black_box, criterion_group, criterion_main, Criterion};
use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna3")]
    coluna3: i64,
}

// Matches tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb's header: Region, Country, Item
// Type, Sales Channel, Order Priority, Order Date, Order ID, Ship Date, Units Sold, Unit Price,
// Unit Cost, Total Revenue, Total Cost, Total Profit. Only a subset is bound - enough to exercise
// string/date/int/float decoding, not every column.
#[derive(Default, ExcelMapper)]
struct LargeRow {
    #[excel(name = "Region")]
    region: String,
    #[excel(name = "Country")]
    country: String,
    #[excel(name = "Order Date")]
    order_date: excelreader::Date,
    #[excel(name = "Order ID")]
    order_id: i64,
    #[excel(name = "Units Sold")]
    units_sold: i64,
    #[excel(name = "Total Profit")]
    total_profit: f64,
}

fn fixture_path() -> String {
    concat!(env!("CARGO_MANIFEST_DIR"), "/../../RealExcel.xlsb").to_string()
}

// Isolates per-row cost from the fixed open overhead the 100-row RealExcel.xlsb benchmarks below
// can't separate out.
fn large_fixture_path() -> String {
    concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb"
    )
    .to_string()
}

// Isolates Workbook::open's cost: file read, container parse, header decode - no row typing.
fn bench_open(c: &mut Criterion) {
    let path = fixture_path();
    c.bench_function("open", |b| {
        b.iter(|| {
            let workbook = Workbook::open(black_box(&path)).expect("open must succeed");
            black_box(workbook);
        });
    });
}

// Isolates parse_sheet's cost on top of an already-open workbook: schema resolution and the
// columnar typed decode over all 100 rows.
fn bench_parse_sheet(c: &mut Criterion) {
    let path = fixture_path();
    c.bench_function("parse_sheet", |b| {
        b.iter_batched(
            || Workbook::open(&path).expect("open must succeed"),
            |mut workbook| {
                let table =
                    parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
                black_box(table);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

// Isolates Workbook::infer_schema's cost: sampling the header + up to 100 rows to guess a schema,
// independent of any typed parse.
fn bench_infer_schema(c: &mut Criterion) {
    let path = fixture_path();
    c.bench_function("infer_schema", |b| {
        b.iter_batched(
            || Workbook::open(&path).expect("open must succeed"),
            |workbook| {
                let schema = workbook
                    .infer_schema(1, 100)
                    .expect("infer_schema must succeed");
                black_box(schema);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

fn bench_open_large(c: &mut Criterion) {
    let path = large_fixture_path();
    c.bench_function("open_large", |b| {
        b.iter(|| {
            let workbook = Workbook::open(black_box(&path)).expect("open must succeed");
            black_box(workbook);
        });
    });
}

fn bench_parse_sheet_large(c: &mut Criterion) {
    let path = large_fixture_path();
    c.bench_function("parse_sheet_large", |b| {
        b.iter_batched(
            || Workbook::open(&path).expect("open must succeed"),
            |mut workbook| {
                let table = parse_sheet::<LargeRow>(&mut workbook, 1)
                    .expect("parse_sheet must succeed");
                black_box(table);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

fn bench_infer_schema_large(c: &mut Criterion) {
    let path = large_fixture_path();
    c.bench_function("infer_schema_large", |b| {
        b.iter_batched(
            || Workbook::open(&path).expect("open must succeed"),
            |workbook| {
                let schema = workbook
                    .infer_schema(1, 1000)
                    .expect("infer_schema must succeed");
                black_box(schema);
            },
            criterion::BatchSize::SmallInput,
        );
    });
}

criterion_group!(
    benches,
    bench_open,
    bench_parse_sheet,
    bench_infer_schema,
    bench_open_large,
    bench_parse_sheet_large,
    bench_infer_schema_large
);
criterion_main!(benches);
