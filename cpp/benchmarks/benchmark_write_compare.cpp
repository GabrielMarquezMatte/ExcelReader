
#include <xl/excelreader.hpp>

#include <benchmark/benchmark.h>
#include <duckdb.hpp>

#ifdef EXCELREADER_BENCH_XLSXIO
#include <xlnt/xlnt.hpp>
#include <xlsxio_write.h>
#endif
#ifdef EXCELREADER_BENCH_LIBXLSXWRITER
#include <xlsxwriter.h>
#endif

#include <array>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <filesystem>
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
        std::chrono::sys_days OrderDate;
        int64_t OrderId;
        std::chrono::sys_days ShipDate;
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
    constexpr std::array<const char *, 14> kHeaders{
        "Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date",
        "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue",
        "Total Cost", "Total Profit"};

    const std::vector<FullRow> &fixture_rows()
    {
        static const std::vector<FullRow> rows = []
        {
            auto workbook = xl::Workbook::open(EXCELREADER_XLSX_FIXTURE_PATH, XL_FORMAT_XLSX);
            if (!workbook.has_value())
            {
                std::fprintf(stderr, "missing or unreadable fixture %s\n", EXCELREADER_XLSX_FIXTURE_PATH);
                std::abort();
            }
            auto table = xl::parse_sheet<FullRow>(*workbook);
            if (!table.has_value() || table->size() == 0)
            {
                std::fprintf(stderr, "fixture %s parsed to zero rows\n", EXCELREADER_XLSX_FIXTURE_PATH);
                std::abort();
            }
            return table->to_vector();
        }();
        return rows;
    }

    std::filesystem::path bench_path(std::string_view name)
    {
        const auto ticks = std::chrono::steady_clock::now().time_since_epoch().count();
        return std::filesystem::temp_directory_path() /
               std::filesystem::path("excelreader-write-compare-" + std::to_string(ticks) + "-" +
                                      std::string(name));
    }

    int32_t days(std::chrono::sys_days value)
    {
        return static_cast<int32_t>(value.time_since_epoch().count());
    }

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

static void BM_ExcelReader_WriteSheet(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("sheet.xlsx");
    for (auto _ : state)
    {
        auto result = xl::write_sheet(path.string(), XL_FORMAT_XLSX, rows);
        if (!result.has_value())
        {
            state.SkipWithError(result.error().message.c_str());
            return;
        }
        benchmark::DoNotOptimize(result);
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_ExcelReader_WriteSheet);

static void BM_ExcelReader_WriteColumns(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();

    StringBuffer region;
    StringBuffer country;
    StringBuffer item_type;
    StringBuffer sales_channel;
    StringBuffer order_priority;
    std::vector<int32_t> order_date;
    std::vector<int64_t> order_id;
    std::vector<int32_t> ship_date;
    std::vector<int64_t> units_sold;
    std::vector<double> unit_price;
    std::vector<double> unit_cost;
    std::vector<double> total_revenue;
    std::vector<double> total_cost;
    std::vector<double> total_profit;

    const size_t count = rows.size();
    for (StringBuffer *buffer : {&region, &country, &item_type, &sales_channel, &order_priority})
    {
        buffer->reserve(count);
    }
    order_date.reserve(count);
    order_id.reserve(count);
    ship_date.reserve(count);
    units_sold.reserve(count);
    unit_price.reserve(count);
    unit_cost.reserve(count);
    total_revenue.reserve(count);
    total_cost.reserve(count);
    total_profit.reserve(count);

    for (const FullRow &row : rows)
    {
        region.push(row.Region);
        country.push(row.Country);
        item_type.push(row.ItemType);
        sales_channel.push(row.SalesChannel);
        order_priority.push(row.OrderPriority);
        order_date.push_back(days(row.OrderDate));
        order_id.push_back(row.OrderId);
        ship_date.push_back(days(row.ShipDate));
        units_sold.push_back(row.UnitsSold);
        unit_price.push_back(row.UnitPrice);
        unit_cost.push_back(row.UnitCost);
        total_revenue.push_back(row.TotalRevenue);
        total_cost.push_back(row.TotalCost);
        total_profit.push_back(row.TotalProfit);
    }

    const std::array<xl::ColumnRef, 14> columns{
        xl::string_column(kHeaders[0], region.offsets, region.data),
        xl::string_column(kHeaders[1], country.offsets, country.data),
        xl::string_column(kHeaders[2], item_type.offsets, item_type.data),
        xl::string_column(kHeaders[3], sales_channel.offsets, sales_channel.data),
        xl::string_column(kHeaders[4], order_priority.offsets, order_priority.data),
        xl::date_column(kHeaders[5], order_date),
        xl::i64_column(kHeaders[6], order_id),
        xl::date_column(kHeaders[7], ship_date),
        xl::i64_column(kHeaders[8], units_sold),
        xl::f64_column(kHeaders[9], unit_price),
        xl::f64_column(kHeaders[10], unit_cost),
        xl::f64_column(kHeaders[11], total_revenue),
        xl::f64_column(kHeaders[12], total_cost),
        xl::f64_column(kHeaders[13], total_profit)};

    const std::filesystem::path path = bench_path("columns.xlsx");
    for (auto _ : state)
    {
        auto result = xl::write_columns(path.string(), XL_FORMAT_XLSX, columns);
        if (!result.has_value())
        {
            state.SkipWithError(result.error().message.c_str());
            return;
        }
        benchmark::DoNotOptimize(result);
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_ExcelReader_WriteColumns);

#ifdef EXCELREADER_BENCH_XLSXIO
static void BM_Xlnt_Write(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("xlnt.xlsx");
    for (auto _ : state)
    {
        xlnt::workbook workbook;
        xlnt::worksheet sheet = workbook.active_sheet();

        for (uint32_t column = 0; column < kHeaders.size(); ++column)
        {
            sheet.cell(column + 1, 1).value(kHeaders[column]);
        }

        uint32_t row_index = 2;
        for (const FullRow &row : rows)
        {
            sheet.cell(1, row_index).value(row.Region);
            sheet.cell(2, row_index).value(row.Country);
            sheet.cell(3, row_index).value(row.ItemType);
            sheet.cell(4, row_index).value(row.SalesChannel);
            sheet.cell(5, row_index).value(row.OrderPriority);
            sheet.cell(6, row_index).value(static_cast<double>(days(row.OrderDate)));
            sheet.cell(7, row_index).value(static_cast<double>(row.OrderId));
            sheet.cell(8, row_index).value(static_cast<double>(days(row.ShipDate)));
            sheet.cell(9, row_index).value(static_cast<double>(row.UnitsSold));
            sheet.cell(10, row_index).value(row.UnitPrice);
            sheet.cell(11, row_index).value(row.UnitCost);
            sheet.cell(12, row_index).value(row.TotalRevenue);
            sheet.cell(13, row_index).value(row.TotalCost);
            sheet.cell(14, row_index).value(row.TotalProfit);
            ++row_index;
        }

        workbook.save(path.string());
        benchmark::DoNotOptimize(sheet);
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_Xlnt_Write);

static void BM_Xlsxio_Write(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("xlsxio.xlsx");
    const std::string target = path.string();
    for (auto _ : state)
    {
        xlsxiowriter handle = xlsxiowrite_open(target.c_str(), "Sheet1");
        if (!handle)
        {
            state.SkipWithError("xlsxiowrite_open failed");
            return;
        }
        xlsxiowrite_set_detection_rows(handle, 0);

        for (const char *header : kHeaders)
        {
            xlsxiowrite_add_column(handle, header, 0);
        }

        for (const FullRow &row : rows)
        {
            xlsxiowrite_add_cell_string(handle, row.Region.c_str());
            xlsxiowrite_add_cell_string(handle, row.Country.c_str());
            xlsxiowrite_add_cell_string(handle, row.ItemType.c_str());
            xlsxiowrite_add_cell_string(handle, row.SalesChannel.c_str());
            xlsxiowrite_add_cell_string(handle, row.OrderPriority.c_str());
            xlsxiowrite_add_cell_int(handle, days(row.OrderDate));
            xlsxiowrite_add_cell_int(handle, row.OrderId);
            xlsxiowrite_add_cell_int(handle, days(row.ShipDate));
            xlsxiowrite_add_cell_int(handle, row.UnitsSold);
            xlsxiowrite_add_cell_float(handle, row.UnitPrice);
            xlsxiowrite_add_cell_float(handle, row.UnitCost);
            xlsxiowrite_add_cell_float(handle, row.TotalRevenue);
            xlsxiowrite_add_cell_float(handle, row.TotalCost);
            xlsxiowrite_add_cell_float(handle, row.TotalProfit);
            xlsxiowrite_next_row(handle);
        }

        if (xlsxiowrite_close(handle) != 0)
        {
            state.SkipWithError("xlsxiowrite_close failed");
            return;
        }
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_Xlsxio_Write);
#endif // EXCELREADER_BENCH_XLSXIO

#ifdef EXCELREADER_BENCH_LIBXLSXWRITER
static void BM_Libxlsxwriter_Write(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("libxlsxwriter.xlsx");
    const std::string target = path.string();
    for (auto _ : state)
    {
        lxw_workbook *workbook = workbook_new(target.c_str());
        if (!workbook)
        {
            state.SkipWithError("workbook_new failed");
            return;
        }
        lxw_worksheet *sheet = workbook_add_worksheet(workbook, nullptr);

        for (lxw_col_t column = 0; column < static_cast<lxw_col_t>(kHeaders.size()); ++column)
        {
            worksheet_write_string(sheet, 0, column, kHeaders[column], nullptr);
        }

        lxw_row_t row_index = 1;
        for (const FullRow &row : rows)
        {
            worksheet_write_string(sheet, row_index, 0, row.Region.c_str(), nullptr);
            worksheet_write_string(sheet, row_index, 1, row.Country.c_str(), nullptr);
            worksheet_write_string(sheet, row_index, 2, row.ItemType.c_str(), nullptr);
            worksheet_write_string(sheet, row_index, 3, row.SalesChannel.c_str(), nullptr);
            worksheet_write_string(sheet, row_index, 4, row.OrderPriority.c_str(), nullptr);
            worksheet_write_number(sheet, row_index, 5, static_cast<double>(days(row.OrderDate)), nullptr);
            worksheet_write_number(sheet, row_index, 6, static_cast<double>(row.OrderId), nullptr);
            worksheet_write_number(sheet, row_index, 7, static_cast<double>(days(row.ShipDate)), nullptr);
            worksheet_write_number(sheet, row_index, 8, static_cast<double>(row.UnitsSold), nullptr);
            worksheet_write_number(sheet, row_index, 9, row.UnitPrice, nullptr);
            worksheet_write_number(sheet, row_index, 10, row.UnitCost, nullptr);
            worksheet_write_number(sheet, row_index, 11, row.TotalRevenue, nullptr);
            worksheet_write_number(sheet, row_index, 12, row.TotalCost, nullptr);
            worksheet_write_number(sheet, row_index, 13, row.TotalProfit, nullptr);
            ++row_index;
        }

        if (workbook_close(workbook) != LXW_NO_ERROR)
        {
            state.SkipWithError("workbook_close failed");
            return;
        }
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_Libxlsxwriter_Write);
#endif // EXCELREADER_BENCH_LIBXLSXWRITER

static void BM_DuckDB_Write(benchmark::State &state)
{
    const std::vector<FullRow> &rows = fixture_rows();
    const std::filesystem::path path = bench_path("duckdb.xlsx");
    const std::string target = path.string();

    duckdb::DuckDB db(nullptr);
    duckdb::Connection con(db);

    auto setup = con.Query("INSTALL excel; LOAD excel;");
    if (setup->HasError())
    {
        state.SkipWithError(setup->GetError().c_str());
        return;
    }

    auto create = con.Query(
        "CREATE TABLE fixture ("
        "\"Region\" VARCHAR, \"Country\" VARCHAR, \"Item Type\" VARCHAR, "
        "\"Sales Channel\" VARCHAR, \"Order Priority\" VARCHAR, \"Order Date\" DATE, "
        "\"Order ID\" BIGINT, \"Ship Date\" DATE, \"Units Sold\" BIGINT, "
        "\"Unit Price\" DOUBLE, \"Unit Cost\" DOUBLE, \"Total Revenue\" DOUBLE, "
        "\"Total Cost\" DOUBLE, \"Total Profit\" DOUBLE)");
    if (create->HasError())
    {
        state.SkipWithError(create->GetError().c_str());
        return;
    }

    {
        duckdb::Appender appender(con, "fixture");
        for (const FullRow &row : rows)
        {
            appender.AppendRow(
                row.Region.c_str(), row.Country.c_str(), row.ItemType.c_str(),
                row.SalesChannel.c_str(), row.OrderPriority.c_str(),
                duckdb::date_t(days(row.OrderDate)), row.OrderId, duckdb::date_t(days(row.ShipDate)),
                row.UnitsSold, row.UnitPrice, row.UnitCost, row.TotalRevenue, row.TotalCost,
                row.TotalProfit);
        }
        appender.Close();
    }

    const std::string copy_query =
        "COPY fixture TO '" + target + "' WITH (FORMAT xlsx, HEADER true)";
    for (auto _ : state)
    {
        auto result = con.Query(copy_query);
        if (result->HasError())
        {
            state.SkipWithError(result->GetError().c_str());
            return;
        }
        benchmark::DoNotOptimize(result);
    }
    state.SetItemsProcessed(static_cast<int64_t>(state.iterations()) * static_cast<int64_t>(rows.size()));
    std::filesystem::remove(path);
}
BENCHMARK(BM_DuckDB_Write);
