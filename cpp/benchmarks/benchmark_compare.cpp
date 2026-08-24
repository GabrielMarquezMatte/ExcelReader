// Compares ExcelReader against xlnt (https://github.com/tfussell/xlnt), xlsxio
// (https://github.com/brechtsanders/xlsxio) and DuckDB's (https://github.com/duckdb/duckdb)
// "excel" extension reading the full row shape of
// tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsx: all 14 columns, 65,535 data rows. None
// of the three competitors reads .xlsb, so this fixture is xlsx-only, unlike the RealExcel.xlsb
// fixture the other benchmarks use.
//
// All four sides decode every cell into an owned value and fold it into one accumulator - the
// ExcelReader side uses std::string bindings rather than the zero-copy std::string_view used
// elsewhere in this suite, so no side gets an allocation-free advantage the others can't take.
// DuckDB's case expresses the same accumulator as a single SQL aggregate query rather than a C++
// loop, which is the idiomatic way to make a SQL engine touch every cell, not a concession to it.
// Same methodology as BenchmarkAccumulators.cs (the .NET benchmark suite's ExcelReader-vs-Sylvan
// comparison).

#include <xl/excelreader.hpp>

#include <benchmark/benchmark.h>
#include <duckdb.hpp>
#include <xlnt/xlnt.hpp>
#include <xlsxio_read.h>

#include <chrono>
#include <cstdint>
#include <cstring>
#include <string>
#include <fstream>

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

std::expected<std::vector<std::uint8_t>, std::string> read_file_to_buffer(std::string_view path)
{
    std::ifstream file(path.data(), std::ios::binary);
    if (!file)
    {
        return std::unexpected("Failed to open fixture file for reading");
    }
    std::vector<std::uint8_t> buffer((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    return buffer;
}

namespace
{
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
}

static void BM_ExcelReader_Xlsx_Full(benchmark::State &state)
{
    auto buffer_result = read_file_to_buffer(EXCELREADER_XLSX_FIXTURE_PATH);
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error());
        return;
    }
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSX);
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
    }
}
BENCHMARK(BM_ExcelReader_Xlsx_Full);

// Same native call as BM_ExcelReader_Xlsx_Full (xl::parse_sheet is exactly one xl_parse_typed FFI
// call - it returns a lazy TableView, no per-row work happens until iteration), but stops before
// the for-loop. The gap against BM_ExcelReader_Xlsx_Full's number is exactly what materializing one
// FullRow (with an owned std::string per text column) per row costs on the C++ side - the same
// question rust/excelreader/benches/compare_bench.rs's *_parse_only benchmarks answer for Rust.
static void BM_ExcelReader_Xlsx_ParseOnly(benchmark::State &state)
{
    auto buffer_result = read_file_to_buffer(EXCELREADER_XLSX_FIXTURE_PATH);
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error());
        return;
    }
    for (auto _ : state)
    {
        auto workbook = xl::Workbook::open_memory(buffer_result.value(), XL_FORMAT_XLSX);
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
        benchmark::DoNotOptimize(table);
    }
}
BENCHMARK(BM_ExcelReader_Xlsx_ParseOnly);

static void BM_Xlnt_Xlsx_Full(benchmark::State &state)
{
    auto buffer_result = read_file_to_buffer(EXCELREADER_XLSX_FIXTURE_PATH);
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error());
        return;
    }
    for (auto _ : state)
    {
        xlnt::workbook wb;
        wb.load(buffer_result.value());
        xlnt::worksheet ws = wb.active_sheet();

        int64_t acc = 0;
        bool first_row = true;
        for (const auto &row : ws.rows(false))
        {
            if (first_row)
            {
                // Row 1 is the header - ExcelReader's parse_sheet consumes it as schema, not data.
                first_row = false;
                continue;
            }
            for (const auto &cell : row)
            {
                if (!cell.has_value())
                {
                    continue;
                }
                switch (cell.data_type())
                {
                case xlnt::cell_type::shared_string:
                case xlnt::cell_type::inline_string:
                case xlnt::cell_type::formula_string:
                    acc += static_cast<int64_t>(cell.to_string().size());
                    break;
                case xlnt::cell_type::number:
                    acc += static_cast<int64_t>(cell.value<double>());
                    break;
                default:
                    break;
                }
            }
        }
        benchmark::DoNotOptimize(acc);
    }
}
BENCHMARK(BM_Xlnt_Xlsx_Full);

namespace
{
    // Column order matches the fixture's header exactly - xlsxio has no named-column lookup, only
    // positional streaming, so each column's typed accessor is called in that fixed order.
    //
    // xlsxioread_sheet_next_cell() only marks a row as finished internally when it returns NULL -
    // the call *after* the last real cell, not the last real cell itself. Stopping right after the
    // 14th (known) column without making that trailing call leaves the row-end bookkeeping stale,
    // and the next xlsxioread_sheet_next_row() desyncs every column by one from there on. Draining
    // to NULL at the end of every row (like the header skip above) avoids that.
    int64_t accumulate_xlsxio_row(xlsxioreadersheet sheet)
    {
        int64_t acc = 0;
        char *text = nullptr;
        int64_t ivalue = 0;
        double fvalue = 0.0;
        time_t tvalue = 0;

        // Region, Country, Item Type, Sales Channel, Order Priority
        for (int i = 0; i < 5; ++i)
        {
            if (xlsxioread_sheet_next_cell_string(sheet, &text) && text)
            {
                acc += static_cast<int64_t>(std::strlen(text));
            }
            if (text)
            {
                xlsxioread_free(text);
                text = nullptr;
            }
        }
        // Order Date
        if (xlsxioread_sheet_next_cell_datetime(sheet, &tvalue))
        {
            acc += static_cast<int64_t>(tvalue);
        }
        // Order ID
        if (xlsxioread_sheet_next_cell_int(sheet, &ivalue))
        {
            acc += ivalue;
        }
        // Ship Date
        if (xlsxioread_sheet_next_cell_datetime(sheet, &tvalue))
        {
            acc += static_cast<int64_t>(tvalue);
        }
        // Units Sold
        if (xlsxioread_sheet_next_cell_int(sheet, &ivalue))
        {
            acc += ivalue;
        }
        // Unit Price, Unit Cost, Total Revenue, Total Cost, Total Profit
        for (int i = 0; i < 5; ++i)
        {
            if (xlsxioread_sheet_next_cell_float(sheet, &fvalue))
            {
                acc += static_cast<int64_t>(fvalue);
            }
        }
        for (XLSXIOCHAR *extra; (extra = xlsxioread_sheet_next_cell(sheet)) != nullptr;)
        {
            xlsxioread_free(extra);
        }
        return acc;
    }
}

static void BM_Xlsxio_Xlsx_Full(benchmark::State &state)
{
    auto buffer_result = read_file_to_buffer(EXCELREADER_XLSX_FIXTURE_PATH);
    if (!buffer_result.has_value())
    {
        state.SkipWithError(buffer_result.error());
        return;
    }
    auto &buffer = buffer_result.value();
    for (auto _ : state)
    {
        xlsxioreader handle = xlsxioread_open_memory(buffer.data(), buffer.size(), 0);
        if (!handle)
        {
            state.SkipWithError("xlsxioread_open failed");
            return;
        }
        xlsxioreadersheet sheet = xlsxioread_sheet_open(handle, nullptr, XLSXIOREAD_SKIP_NONE);
        if (!sheet)
        {
            xlsxioread_close(handle);
            state.SkipWithError("xlsxioread_sheet_open failed");
            return;
        }

        // Row 1 is the header - ExcelReader's parse_sheet consumes it as schema, not data. xlsxio's
        // row cursor only advances correctly once every cell of the current row has been read (its
        // SAX resume loop tracks column position), so the header's cells must be drained here, not
        // just skipped - advancing past a partially-read row desyncs every subsequent typed read.
        xlsxioread_sheet_next_row(sheet);
        for (XLSXIOCHAR *header_cell; (header_cell = xlsxioread_sheet_next_cell(sheet)) != nullptr;)
        {
            xlsxioread_free(header_cell);
        }

        int64_t acc = 0;
        while (xlsxioread_sheet_next_row(sheet))
        {
            acc += accumulate_xlsxio_row(sheet);
        }
        benchmark::DoNotOptimize(acc);

        xlsxioread_sheet_close(sheet);
        xlsxioread_close(handle);
    }
}
BENCHMARK(BM_Xlsxio_Xlsx_Full);

// DuckDB (https://github.com/duckdb/duckdb) reads via its "excel" extension's read_xlsx() table
// function, invoked here as a single aggregate query rather than pulled apart row by row in C++:
// DuckDB is a SQL engine, and an aggregate over every column is the idiomatic way to make it
// decode every cell, not an artificial concession to it. The expression matches
// accumulate_full_row() exactly - same columns, same "text length, numeric value" split, same
// epoch-days conversion for the two date columns - so the three sides remain the same "touch every
// cell into one accumulator" methodology, just expressed once in SQL instead of once per row.
static void BM_DuckDB_Xlsx_Full(benchmark::State &state)
{
    duckdb::DuckDB db(nullptr);
    duckdb::Connection con(db);

    // INSTALL pulls the extension from DuckDB's extension repository on first use and caches it
    // locally afterward; done once per process, outside the timed region.
    auto setup = con.Query("INSTALL excel; LOAD excel;");
    if (setup->HasError())
    {
        state.SkipWithError(setup->GetError());
        return;
    }

    const std::string query =
        "SELECT sum(length(\"Region\")) + sum(length(\"Country\")) + sum(length(\"Item Type\")) + "
        "sum(length(\"Sales Channel\")) + sum(length(\"Order Priority\")) + "
        "sum(CAST(\"Order Date\" - DATE '1970-01-01' AS BIGINT)) + sum(\"Order ID\") + "
        "sum(CAST(\"Ship Date\" - DATE '1970-01-01' AS BIGINT)) + sum(\"Units Sold\") + "
        "sum(CAST(\"Unit Price\" AS BIGINT)) + sum(CAST(\"Unit Cost\" AS BIGINT)) + "
        "sum(CAST(\"Total Revenue\" AS BIGINT)) + sum(CAST(\"Total Cost\" AS BIGINT)) + "
        "sum(CAST(\"Total Profit\" AS BIGINT)) AS acc FROM read_xlsx('" +
        std::string(EXCELREADER_XLSX_FIXTURE_PATH) + "', header = true)";

    for (auto _ : state)
    {
        auto result = con.Query(query);
        if (result->HasError())
        {
            state.SkipWithError(result->GetError());
            return;
        }
        int64_t acc = result->GetValue<int64_t>(0, 0);
        benchmark::DoNotOptimize(acc);
    }
}
BENCHMARK(BM_DuckDB_Xlsx_Full);
