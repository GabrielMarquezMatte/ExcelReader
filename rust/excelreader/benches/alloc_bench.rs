use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicU64, Ordering};

use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};

static BYTES_ALLOCATED: AtomicU64 = AtomicU64::new(0);

struct CountingAllocator;

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        BYTES_ALLOCATED.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        // Freeing doesn't need instrumentation to answer "bytes allocated" - same scope choice as
        // the C++ side (total_allocated_bytes there, not net/peak usage).
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        if new_size > layout.size() {
            BYTES_ALLOCATED.fetch_add((new_size - layout.size()) as u64, Ordering::Relaxed);
        }
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: CountingAllocator = CountingAllocator;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Coluna1")]
    coluna1: String,
    #[excel(name = "Coluna3")]
    coluna3: i64,
}

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

fn large_fixture_path() -> String {
    concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb"
    )
    .to_string()
}

/// Runs `routine` `iterations` times, measuring ONLY `routine` itself — matching
/// criterion::Bencher::iter_batched's own contract (setup excluded, see parse_bench.rs) even though
/// this harness is otherwise independent of criterion. Reports min/median/max bytes allocated
/// across the run: median is what to compare between runs, min/max show how flat (or not) the
/// distribution actually is, standing in for the variance report criterion normally provides.
fn measure<S, R, O>(name: &str, iterations: usize, mut setup: S, mut routine: R)
where
    S: FnMut() -> O,
    R: FnMut(O),
{
    let mut samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let input = setup();
        let before = BYTES_ALLOCATED.load(Ordering::Relaxed);
        routine(input);
        let after = BYTES_ALLOCATED.load(Ordering::Relaxed);
        samples.push(after.saturating_sub(before));
    }
    samples.sort_unstable();
    let min = samples[0];
    let max = samples[samples.len() - 1];
    let median = samples[samples.len() / 2];
    println!("{name:<28} min={min:>10} bytes  median={median:>10} bytes  max={max:>10} bytes  (n={iterations})");
}

fn main() {
    let small_path = fixture_path();
    let large_path = large_fixture_path();
    const ITERATIONS: usize = 200;
    const ITERATIONS_LARGE: usize = 30; // the large fixture's own per-call cost dwarfs the loop overhead already at this count

    println!("Bytes allocated per call, Rust binding (System allocator, this process only — excludes the NativeAOT side of the FFI boundary):\n");

    measure(
        "open",
        ITERATIONS,
        || (),
        |()| {
            let workbook = Workbook::open(&small_path).expect("open must succeed");
            std::hint::black_box(workbook);
        },
    );

    measure(
        "parse_sheet",
        ITERATIONS,
        || Workbook::open(&small_path).expect("open must succeed"),
        |mut workbook| {
            let table = parse_sheet::<Row>(&mut workbook, 1).expect("parse_sheet must succeed");
            std::hint::black_box(table);
        },
    );

    measure(
        "infer_schema",
        ITERATIONS,
        || Workbook::open(&small_path).expect("open must succeed"),
        |workbook| {
            let schema = workbook.infer_schema(1, 100).expect("infer_schema must succeed");
            std::hint::black_box(schema);
        },
    );

    measure(
        "open_large",
        ITERATIONS_LARGE,
        || (),
        |()| {
            let workbook = Workbook::open(&large_path).expect("open must succeed");
            std::hint::black_box(workbook);
        },
    );

    measure(
        "parse_sheet_large",
        ITERATIONS_LARGE,
        || Workbook::open(&large_path).expect("open must succeed"),
        |mut workbook| {
            let table = parse_sheet::<LargeRow>(&mut workbook, 1).expect("parse_sheet must succeed");
            std::hint::black_box(table);
        },
    );

    measure(
        "infer_schema_large",
        ITERATIONS_LARGE,
        || Workbook::open(&large_path).expect("open must succeed"),
        |workbook| {
            let schema = workbook.infer_schema(1, 1000).expect("infer_schema must succeed");
            std::hint::black_box(schema);
        },
    );
}
