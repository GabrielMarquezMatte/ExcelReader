// Write benchmarks over tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb (65,535 data rows),
// the same fixture the Rust, Python and .NET suites use.
//
// Both cases write the SAME seven columns, so the only difference between them is where the data
// starts out - which is the one thing this file is measuring:
//
//   * write_columns is handed buffers that are already columnar - nothing is transposed and
//     nothing is copied. It is the ceiling.
//   * write_sheet starts from a std::vector<Row> and pays the row-to-column transpose. It is what
//     a row-shaped caller actually experiences, and the number to compare against any cell-at-a-
//     time writer.
//
// The gap between them is therefore the cost of holding row-shaped data, not a difference in how
// much gets written. For the comparison against other libraries, see benchmark_write_compare.cpp.
//
// State the CPU, OS and compiler version alongside any number published from this file.

#include <xl/excelreader.hpp>

#include <benchmark/benchmark.h>

#include <cstdio>
#include <filesystem>
#include <string>
#include <vector>

struct Row
{
    std::string region;
    std::string country;
    std::string item_type;
    std::chrono::sys_days order_date;
    int64_t order_id;
    int64_t units_sold;
    double total_revenue;
};

template <>
struct xl::ExcelMapper<Row>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Region", &Row::region),
            xl::make_field("Country", &Row::country),
            xl::make_field("Item Type", &Row::item_type),
            xl::make_field("Order Date", &Row::order_date),
            xl::make_field("Order ID", &Row::order_id),
            xl::make_field("Units Sold", &Row::units_sold),
            xl::make_field("Total Revenue", &Row::total_revenue));
    }
};

// Reads the fixture once into row structs. Aborts rather than silently benchmarking an empty
// input - a suite that measures nothing is worse than no suite.
static const std::vector<Row> &fixture_rows()
{
    static const std::vector<Row> rows = []
    {
        auto workbook = xl::Workbook::open(EXCELREADER_LARGE_FIXTURE_PATH);
        if (!workbook.has_value())
        {
            std::fprintf(stderr, "missing or unreadable fixture %s\n", EXCELREADER_LARGE_FIXTURE_PATH);
            std::abort();
        }
        auto table = xl::parse_sheet<Row>(*workbook);
        if (!table.has_value() || table->size() == 0)
        {
            std::fprintf(stderr, "fixture %s parsed to zero rows\n", EXCELREADER_LARGE_FIXTURE_PATH);
            std::abort();
        }
        return table->to_vector();
    }();
    return rows;
}

static std::filesystem::path bench_path(std::string_view name)
{
    return std::filesystem::temp_directory_path() /
           std::filesystem::path(std::string("excelreader-bench-") + std::string(name));
}

namespace
{
    // The offsets/blob pair an XL_T_STRING column needs. xl::write_sheet builds one of these
    // internally; the columnar benchmark below builds its own so both cases start from the same
    // shape.
    struct StringBuffer
    {
        std::vector<int32_t> offsets{0};
        std::vector<uint8_t> data{};

        void reserve(size_t count)
        {
            offsets.reserve(count + 1);
        }

        void push(std::string_view value)
        {
            const uint8_t *bytes = reinterpret_cast<const uint8_t *>(value.data());
            data.insert(data.end(), bytes, bytes + value.size());
            offsets.push_back(static_cast<int32_t>(data.size()));
        }
    };
}

static void BM_WriteSheet(benchmark::State &state)
{
    const std::vector<Row> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("sheet.xlsx");
    for (auto _ : state)
    {
        auto result = xl::write_sheet(path.string(), XL_FORMAT_XLSX, rows);
        benchmark::DoNotOptimize(result);
        if (!result.has_value())
        {
            state.SkipWithError("write_sheet failed");
            break;
        }
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_WriteSheet);

// Columnar buffers, transposed from fixture_rows() once (function-local static, same caching
// rationale as fixture_rows() itself). BM_WriteColumns previously rebuilt these on every call
// outside its `for (auto _ : state)` loop - harmless for timing (PauseTiming/ResumeTiming aren't
// even needed there), but Google Benchmark's dedicated memory-measurement run
// (BenchmarkRunner::RunMemoryManager) still executes that untimed setup code, so every one of its
// (up to 16) calls re-transposed all 65,535 rows, and the *total* over that run - not amortized per
// iteration - dominated BM_WriteColumns' allocated-bytes figure with a cost the benchmark isn't
// meant to isolate (that's BM_WriteSheet's job). Caching means only the first call pays it.
//
// ALL SEVEN columns, the same set BM_WriteSheet writes. An earlier version of this benchmark wrote
// only four, which made it look ~2x faster when a third of that gap was simply three fewer columns
// of work.
struct WriteColumnsFixture
{
    StringBuffer region;
    StringBuffer country;
    StringBuffer item_type;
    std::vector<int32_t> order_dates;
    std::vector<int64_t> order_ids;
    std::vector<int64_t> units;
    std::vector<double> revenue;
};

static const WriteColumnsFixture &write_columns_fixture()
{
    static const WriteColumnsFixture fixture = []
    {
        const std::vector<Row> &rows = fixture_rows();
        WriteColumnsFixture built;
        built.region.reserve(rows.size());
        built.country.reserve(rows.size());
        built.item_type.reserve(rows.size());
        built.order_dates.reserve(rows.size());
        built.order_ids.reserve(rows.size());
        built.units.reserve(rows.size());
        built.revenue.reserve(rows.size());
        for (const Row &row : rows)
        {
            built.region.push(row.region);
            built.country.push(row.country);
            built.item_type.push(row.item_type);
            built.order_dates.push_back(static_cast<int32_t>(row.order_date.time_since_epoch().count()));
            built.order_ids.push_back(row.order_id);
            built.units.push_back(row.units_sold);
            built.revenue.push_back(row.total_revenue);
        }
        return built;
    }();
    return fixture;
}

static void BM_WriteColumns(benchmark::State &state)
{
    const std::vector<Row> &rows = fixture_rows();
    const WriteColumnsFixture &columns_fixture = write_columns_fixture();

    const std::array<xl::ColumnRef, 7> columns{
        xl::string_column("Region", columns_fixture.region.offsets, columns_fixture.region.data),
        xl::string_column("Country", columns_fixture.country.offsets, columns_fixture.country.data),
        xl::string_column("Item Type", columns_fixture.item_type.offsets, columns_fixture.item_type.data),
        xl::date_column("Order Date", columns_fixture.order_dates),
        xl::i64_column("Order ID", columns_fixture.order_ids),
        xl::i64_column("Units Sold", columns_fixture.units),
        xl::f64_column("Total Revenue", columns_fixture.revenue)};

    const std::filesystem::path path = bench_path("columns.xlsx");
    for (auto _ : state)
    {
        auto result = xl::write_columns(path.string(), XL_FORMAT_XLSX, columns);
        benchmark::DoNotOptimize(result);
        if (!result.has_value())
        {
            state.SkipWithError("write_columns failed");
            break;
        }
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_WriteColumns);