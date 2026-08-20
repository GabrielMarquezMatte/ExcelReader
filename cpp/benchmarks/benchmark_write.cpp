// Write benchmarks over tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb (65,535 data rows),
// the same fixture the Rust, Python and .NET suites use.
//
// WORK IS NOT MATCHED between the two cases, by construction:
//
//   * write_columns is handed buffers that are already columnar - nothing is transposed and
//     nothing is copied. It is the ceiling.
//   * write_sheet starts from a std::vector<Row> and pays the row-to-column transpose. It is what
//     a row-shaped caller actually experiences, and the number to compare against any cell-at-a-
//     time writer.
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

static void BM_WriteColumns(benchmark::State &state)
{
    const std::vector<Row> &rows = fixture_rows();

    // Transposed once, outside the measured region: this case exists to measure the write, not the
    // transpose BM_WriteSheet already covers.
    std::vector<int32_t> region_offsets{0};
    std::vector<uint8_t> region_data;
    std::vector<int64_t> order_ids;
    std::vector<int64_t> units;
    std::vector<double> revenue;
    order_ids.reserve(rows.size());
    units.reserve(rows.size());
    revenue.reserve(rows.size());
    region_offsets.reserve(rows.size() + 1);
    for (const Row &row : rows)
    {
        const uint8_t *bytes = reinterpret_cast<const uint8_t *>(row.region.data());
        region_data.insert(region_data.end(), bytes, bytes + row.region.size());
        region_offsets.push_back(static_cast<int32_t>(region_data.size()));
        order_ids.push_back(row.order_id);
        units.push_back(row.units_sold);
        revenue.push_back(row.total_revenue);
    }

    const std::array<xl::ColumnRef, 4> columns{
        xl::string_column("Region", region_offsets, region_data),
        xl::i64_column("Order ID", order_ids),
        xl::i64_column("Units Sold", units),
        xl::f64_column("Total Revenue", revenue)};

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