#include <xl/excelreader.hpp>

#include <benchmark/benchmark.h>

#include <chrono>
#include <cstdint>
#include <string_view>
#include <fstream>

namespace
{
    struct Row
    {
        std::string_view Coluna1;
        std::chrono::year_month_day Coluna2;
        int64_t Coluna3;
        double Coluna16;
    };
}

template <>
struct xl::ExcelMapper<Row>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Coluna1", &Row::Coluna1),
            xl::make_field("Coluna2", &Row::Coluna2),
            xl::make_field("Coluna3", &Row::Coluna3),
            xl::make_field("Coluna16", &Row::Coluna16));
    }
};

namespace
{
    // Matches tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb's header: Region, Country,
    // Item Type, Sales Channel, Order Priority, Order Date, Order ID, Ship Date, Units Sold, Unit
    // Price, Unit Cost, Total Revenue, Total Cost, Total Profit. Only a subset is bound here - enough
    // to exercise string/date/int/float decoding, not every column.
    struct LargeRow
    {
        std::string_view Region;
        std::string_view Country;
        std::chrono::year_month_day OrderDate;
        int64_t OrderId;
        int64_t UnitsSold;
        double TotalProfit;
    };
}

template <>
struct xl::ExcelMapper<LargeRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Region", &LargeRow::Region),
            xl::make_field("Country", &LargeRow::Country),
            xl::make_field("Order Date", &LargeRow::OrderDate),
            xl::make_field("Order ID", &LargeRow::OrderId),
            xl::make_field("Units Sold", &LargeRow::UnitsSold),
            xl::make_field("Total Profit", &LargeRow::TotalProfit));
    }
};

std::expected<std::vector<std::uint8_t>, std::string> read_file_to_buffer(std::string_view path)
{
    std::ifstream file(path.data(), std::ios::binary);
    if (!file)
    {
        return std::unexpected("Failed to open fixture file for reading");
    }
    std::vector<std::uint8_t> buffer;
    buffer.assign(std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>());
    return buffer;
}

// Cached once per process (function-local static, thread-safe init since C++11) rather than
// re-read on every call to a BM_* function below. Google Benchmark invokes the whole benchmark
// function an EXTRA, separate time to measure memory (BenchmarkRunner::RunMemoryManager), and that
// call still executes everything outside the `for (auto _ : state)` loop too - so a fresh
// read_file_to_buffer() there was inflating every *_Large benchmark's allocated-bytes total by the
// fixture's own size (~15 MB), drowning out what the operation under test actually allocates per
// call. Caching means only the first BM_* invocation across the whole binary pays that cost; every
// later one - including the dedicated memory-measurement run - sees zero new allocation for it.
const std::expected<std::vector<std::uint8_t>, std::string> &small_fixture_buffer()
{
    static const std::expected<std::vector<std::uint8_t>, std::string> buffer =
        read_file_to_buffer(EXCELREADER_FIXTURE_PATH);
    return buffer;
}

const std::expected<std::vector<std::uint8_t>, std::string> &large_fixture_buffer()
{
    static const std::expected<std::vector<std::uint8_t>, std::string> buffer =
        read_file_to_buffer(EXCELREADER_LARGE_FIXTURE_PATH);
    return buffer;
}

// Isolates xl::Workbook::open's cost: file read, container parse, header decode - no row typing.
static void BM_Open(benchmark::State &state)
{
    const auto &buffer_result = small_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        benchmark::DoNotOptimize(workbook);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_Open);

// Isolates xl::parse_sheet<Row>'s cost on top of an already-open workbook: schema resolution and
// the columnar typed decode over all 100 rows.
static void BM_ParseSheet(benchmark::State &state)
{
    const auto &buffer_result = small_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        state.PauseTiming();
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
        state.ResumeTiming();

        auto table = xl::parse_sheet<Row>(*workbook);
        benchmark::DoNotOptimize(table);
        if (!table.has_value())
        {
            state.SkipWithError(table.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_ParseSheet);

// Isolates xl::Workbook::infer_schema's cost: sampling the header + up to 100 rows to guess a
// schema, independent of any typed parse.
static void BM_InferSchema(benchmark::State &state)
{
    const auto &buffer_result = small_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        state.PauseTiming();
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
        state.ResumeTiming();

        auto schema = workbook->infer_schema(1, 100);
        benchmark::DoNotOptimize(schema);
        if (!schema.has_value())
        {
            state.SkipWithError(schema.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_InferSchema);

// --- Large fixture (65,535 data rows) - isolates per-row cost from the fixed open overhead the
// 100-row RealExcel.xlsb benchmarks above can't separate out.

static void BM_Open_Large(benchmark::State &state)
{
    const auto &buffer_result = large_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        benchmark::DoNotOptimize(workbook);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_Open_Large);

static void BM_ParseSheet_Large(benchmark::State &state)
{
    const auto &buffer_result = large_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        state.PauseTiming();
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
        state.ResumeTiming();

        auto table = xl::parse_sheet<LargeRow>(*workbook);
        benchmark::DoNotOptimize(table);
        if (!table.has_value())
        {
            state.SkipWithError(table.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_ParseSheet_Large);

// What this DOES measure: wall-clock time and C++-side allocation (operator new/delete only) over
// the full read. What this does NOT measure: the peak-memory ceiling, because the allocation harness
// sees only C++-side allocations and reports a cumulative total, not bytes-live-at-once. The columnar
// buffers (the bulk of what batching bounds) are allocated inside NativeAOT and invisible here. The
// ceiling itself is proven instead by TypedParseSessionTests.NextBatch_Should_Bound_The_Largest_Live_Table_To_Roughly_One_Batch
// in tests/ExcelReader.Tests/TypedParseSessionTests.cs, which sums live NativeTable bytes directly.
// Read that test for the ceiling; read this benchmark for the wall-clock cost and C++ allocation trade-off.
static void BM_TypedReaderBatched_Large(benchmark::State &state)
{
    const auto &buffer_result = large_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        state.PauseTiming();
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
        state.ResumeTiming();

        auto reader = xl::typed_reader<LargeRow>(*workbook, 1, state.range(0));
        if (!reader.has_value())
        {
            state.SkipWithError(reader.error().message.c_str());
            return;
        }
        int64_t rows = 0;
        for (auto &batch : *reader)
        {
            if (!batch.has_value())
            {
                state.SkipWithError(batch.error().message.c_str());
                return;
            }
            rows += batch->size();
        }
        benchmark::DoNotOptimize(rows);
    }
}
BENCHMARK(BM_TypedReaderBatched_Large)->Arg(1000)->Arg(10000)->Arg(0);

static void BM_InferSchema_Large(benchmark::State &state)
{
    const auto &buffer_result = large_fixture_buffer();
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error().c_str());
        return;
    }
    for (auto _ : state)
    {
        state.PauseTiming();
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSB);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message.c_str());
            return;
        }
        state.ResumeTiming();

        auto schema = workbook->infer_schema(1, 1000);
        benchmark::DoNotOptimize(schema);
        if (!schema.has_value())
        {
            state.SkipWithError(schema.error().message.c_str());
            return;
        }
    }
}
BENCHMARK(BM_InferSchema_Large);
