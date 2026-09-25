#include <xl/excelreader.hpp>

#include <benchmark/benchmark.h>
// zsv's headers are plain C99: no extern "C" guard, and `restrict`, which GCC and Clang spell
// `__restrict` in C++.
#define restrict __restrict
extern "C"
{
#include <zsv.h>
}
#undef restrict

#include <charconv>
#include <chrono>
#include <cstdint>
#include <fstream>
#include <string>
#include <string_view>
#include <vector>

namespace
{
    struct FullRow
    {
        std::string Region;
        std::string Country;
        std::string ItemType;
        std::string SalesChannel;
        std::string OrderPriority;
        std::chrono::year_month_day OrderDate;
        int64_t OrderId;
        std::chrono::year_month_day ShipDate;
        int64_t UnitsSold;
        double UnitPrice;
        double UnitCost;
        double TotalRevenue;
        double TotalCost;
        double TotalProfit;
    };
}

template <>
struct xl::ExcelMapper<FullRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Region", &FullRow::Region),
            xl::make_field("Country", &FullRow::Country),
            xl::make_field("Item Type", &FullRow::ItemType),
            xl::make_field("Sales Channel", &FullRow::SalesChannel),
            xl::make_field("Order Priority", &FullRow::OrderPriority),
            xl::make_field("Order Date", &FullRow::OrderDate),
            xl::make_field("Order ID", &FullRow::OrderId),
            xl::make_field("Ship Date", &FullRow::ShipDate),
            xl::make_field("Units Sold", &FullRow::UnitsSold),
            xl::make_field("Unit Price", &FullRow::UnitPrice),
            xl::make_field("Unit Cost", &FullRow::UnitCost),
            xl::make_field("Total Revenue", &FullRow::TotalRevenue),
            xl::make_field("Total Cost", &FullRow::TotalCost),
            xl::make_field("Total Profit", &FullRow::TotalProfit));
    }
};

namespace
{
    constexpr unsigned char kZsvEngineCompat = 0;
    constexpr unsigned char kZsvEngineFast = 3;
    constexpr size_t kColumns = 14;

    std::vector<std::uint8_t> read_fixture()
    {
        std::ifstream file(EXCELREADER_CSV_FIXTURE_PATH, std::ios::binary);
        return {std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>()};
    }

    int64_t days_since_epoch(std::chrono::year_month_day ymd)
    {
        return std::chrono::sys_days{ymd}.time_since_epoch().count();
    }

    int64_t accumulate_full_row(const FullRow &row)
    {
        return static_cast<int64_t>(row.Region.size())
            + static_cast<int64_t>(row.Country.size())
            + static_cast<int64_t>(row.ItemType.size())
            + static_cast<int64_t>(row.SalesChannel.size())
            + static_cast<int64_t>(row.OrderPriority.size())
            + days_since_epoch(row.OrderDate)
            + row.OrderId
            + days_since_epoch(row.ShipDate)
            + row.UnitsSold
            + static_cast<int64_t>(row.UnitPrice)
            + static_cast<int64_t>(row.UnitCost)
            + static_cast<int64_t>(row.TotalRevenue)
            + static_cast<int64_t>(row.TotalCost)
            + static_cast<int64_t>(row.TotalProfit);
    }

    template <typename T>
    bool parse_number(std::string_view text, T &out)
    {
        const auto [end, error] = std::from_chars(text.data(), text.data() + text.size(), out);
        return error == std::errc{} && end == text.data() + text.size();
    }

    bool parse_date(std::string_view text, std::chrono::year_month_day &out)
    {
        int year = 0;
        unsigned month = 0;
        unsigned day = 0;
        if (text.size() != 10 || text[4] != '-' || text[7] != '-'
            || !parse_number(text.substr(0, 4), year)
            || !parse_number(text.substr(5, 2), month)
            || !parse_number(text.substr(8, 2), day))
        {
            return false;
        }
        out = std::chrono::year_month_day{std::chrono::year{year}, std::chrono::month{month}, std::chrono::day{day}};
        return out.ok();
    }

    struct ZsvRun
    {
        zsv_parser parser{};
        bool header_seen{};
        bool failed{};
        std::vector<FullRow> rows;
        int64_t cell_bytes{};
    };

    std::string_view cell_text(zsv_parser parser, size_t index)
    {
        const zsv_cell cell = zsv_get_cell(parser, index);
        return {reinterpret_cast<const char *>(cell.str), cell.len};
    }

    // The same conversions parse_sheet<FullRow> makes: text into owned std::string, ISO dates, integers
    // and doubles, one FullRow per data row, header skipped.
    void zsv_typed_row(void *context)
    {
        auto *run = static_cast<ZsvRun *>(context);
        if (!run->header_seen)
        {
            run->header_seen = true;
            return;
        }
        const size_t count = zsv_cell_count(run->parser);
        // zsv reports the empty line after the file's final CRLF as a one-cell row; ExcelReader skips
        // blank lines.
        if (count == 1 && zsv_get_cell(run->parser, 0).len == 0)
        {
            return;
        }
        if (count != kColumns)
        {
            run->failed = true;
            return;
        }
        FullRow &row = run->rows.emplace_back();
        row.Region = cell_text(run->parser, 0);
        row.Country = cell_text(run->parser, 1);
        row.ItemType = cell_text(run->parser, 2);
        row.SalesChannel = cell_text(run->parser, 3);
        row.OrderPriority = cell_text(run->parser, 4);
        const bool ok = parse_date(cell_text(run->parser, 5), row.OrderDate)
            && parse_number(cell_text(run->parser, 6), row.OrderId)
            && parse_date(cell_text(run->parser, 7), row.ShipDate)
            && parse_number(cell_text(run->parser, 8), row.UnitsSold)
            && parse_number(cell_text(run->parser, 9), row.UnitPrice)
            && parse_number(cell_text(run->parser, 10), row.UnitCost)
            && parse_number(cell_text(run->parser, 11), row.TotalRevenue)
            && parse_number(cell_text(run->parser, 12), row.TotalCost)
            && parse_number(cell_text(run->parser, 13), row.TotalProfit);
        run->failed |= !ok;
    }

    void zsv_cells_row(void *context)
    {
        auto *run = static_cast<ZsvRun *>(context);
        const size_t count = zsv_cell_count(run->parser);
        for (size_t i = 0; i < count; ++i)
        {
            run->cell_bytes += static_cast<int64_t>(zsv_get_cell(run->parser, i).len);
        }
    }

    bool run_zsv(const std::span<const std::uint8_t> buffer, unsigned char engine, void (*row_handler)(void *), ZsvRun &run)
    {
        zsv_opts opts{};
        opts.row_handler = row_handler;
        opts.ctx = &run;
        opts.scan_engine = engine;
        run.parser = zsv_new(&opts);
        if (run.parser == nullptr)
        {
            return false;
        }
        const bool parsed = zsv_parse_bytes(run.parser, buffer.data(), buffer.size()) == zsv_status_ok
            && zsv_finish(run.parser) == zsv_status_ok;
        zsv_delete(run.parser);
        return parsed && !run.failed;
    }
}

static void BM_ExcelReader_Csv_Full(benchmark::State &state)
{
    const std::vector<std::uint8_t> buffer = read_fixture();
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer, XL_FORMAT_CSV);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message);
            return;
        }
        auto table = xl::parse_sheet<FullRow>(*workbook);
        if (!table.has_value())
        {
            state.SkipWithError(table.error().message);
            return;
        }
        int64_t acc = 0;
        for (const FullRow &row : *table)
        {
            acc += accumulate_full_row(row);
        }
        benchmark::DoNotOptimize(acc);
        state.counters["checksum"] = static_cast<double>(acc);
    }
}
BENCHMARK(BM_ExcelReader_Csv_Full);

static void BM_Zsv_Csv_Full(benchmark::State &state, unsigned char engine)
{
    const std::vector<std::uint8_t> buffer = read_fixture();
    for (auto _ : state)
    {
        ZsvRun run;
        if (!run_zsv(buffer, engine, zsv_typed_row, run))
        {
            state.SkipWithError("zsv parse failed");
            return;
        }
        int64_t acc = 0;
        for (const FullRow &row : run.rows)
        {
            acc += accumulate_full_row(row);
        }
        benchmark::DoNotOptimize(acc);
        state.counters["checksum"] = static_cast<double>(acc);
    }
}
BENCHMARK_CAPTURE(BM_Zsv_Csv_Full, compat, kZsvEngineCompat);
BENCHMARK_CAPTURE(BM_Zsv_Csv_Full, fast, kZsvEngineFast);

static void BM_ExcelReader_Csv_Cells(benchmark::State &state)
{
    const std::vector<std::uint8_t> buffer = read_fixture();
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer, XL_FORMAT_CSV);
        if (!workbook.has_value())
        {
            state.SkipWithError(workbook.error().message);
            return;
        }
        xl::RowCursor cursor = workbook->rows();
        int64_t acc = 0;
        while (true)
        {
            auto row = cursor.next_row();
            if (!row.has_value())
            {
                if (row.error().code != XL_EOF)
                {
                    state.SkipWithError(row.error().message);
                    return;
                }
                break;
            }
            for (const xl::CellView cell : *row)
            {
                acc += static_cast<int64_t>(cell.value.size());
            }
        }
        benchmark::DoNotOptimize(acc);
        state.counters["checksum"] = static_cast<double>(acc);
    }
}
BENCHMARK(BM_ExcelReader_Csv_Cells);

static void BM_Zsv_Csv_Cells(benchmark::State &state, unsigned char engine)
{
    const std::vector<std::uint8_t> buffer = read_fixture();
    for (auto _ : state)
    {
        ZsvRun run;
        if (!run_zsv(buffer, engine, zsv_cells_row, run))
        {
            state.SkipWithError("zsv parse failed");
            return;
        }
        benchmark::DoNotOptimize(run.cell_bytes);
        state.counters["checksum"] = static_cast<double>(run.cell_bytes);
    }
}
BENCHMARK_CAPTURE(BM_Zsv_Csv_Cells, compat, kZsvEngineCompat);
BENCHMARK_CAPTURE(BM_Zsv_Csv_Cells, fast, kZsvEngineFast);
