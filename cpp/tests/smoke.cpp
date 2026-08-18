#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string_view>

struct Row {
    std::string_view Coluna1;
    std::chrono::year_month_day Coluna2;
    int64_t Coluna3;
    double Coluna16;
};

template<> struct xl::ExcelMapper<Row> {
    static constexpr auto get_bindings() {
        return std::make_tuple(
            xl::make_field("Coluna1", &Row::Coluna1),
            xl::make_field("Coluna2", &Row::Coluna2),
            xl::make_field("Coluna3", &Row::Coluna3),
            xl::make_field("Coluna16", &Row::Coluna16)
        );
    }
};

#define CHECK(cond, msg) \
    do { if (!(cond)) { std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); return 1; } } while (0)

int main() {
    xl::OpenOptions options { .prefetch_decompression = 1};
    auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
    CHECK(workbook.has_value(), "xl::Workbook::open must succeed on the RealExcel.xlsb fixture");

    auto table = xl::parse_sheet<Row>(*workbook);
    CHECK(table.has_value(), "xl::parse_sheet<Row> must succeed");
    CHECK(table->size() == 100, "RealExcel.xlsb has 100 data rows");

    auto it = table->begin();
    Row first = *it;
    CHECK(first.Coluna1 == "Valor1", "first row's Coluna1 must be Valor1");
    CHECK(first.Coluna2 == std::chrono::year{2026}/std::chrono::month{1}/std::chrono::day{1}, "first row's Coluna2 must be 2026-01-01");
    CHECK(first.Coluna3 == 1, "first row's Coluna3 must be 1");
    CHECK(first.Coluna16 == 0.1, "first row's Coluna16 must be 0.1");

    std::printf("OK: C++ smoke test passed\n");
    return 0;
}
