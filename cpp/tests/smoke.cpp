#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string_view>

struct Row {
    std::string_view Coluna1;
    int64_t Coluna3;
};

template<> struct xl::ExcelMapper<Row> {
    static constexpr auto get_bindings() {
        return std::make_tuple(
            xl::make_field("Coluna1", &Row::Coluna1),
            xl::make_field("Coluna3", &Row::Coluna3));
    }
};

#define CHECK(cond, msg) \
    do { if (!(cond)) { std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); return 1; } } while (0)

int main() {
    auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
    CHECK(workbook.has_value(), "xl::Workbook::open must succeed on the RealExcel.xlsb fixture");

    auto table = xl::parse_sheet<Row>(*workbook);
    CHECK(table.has_value(), "xl::parse_sheet<Row> must succeed");
    CHECK(table->size() == 100, "RealExcel.xlsb has 100 data rows");

    auto it = table->begin();
    Row first = *it;
    CHECK(first.Coluna1 == "Valor1", "first row's Coluna1 must be Valor1");
    CHECK(first.Coluna3 == 1, "first row's Coluna3 must be 1");

    std::printf("OK: C++ smoke test passed\n");
    return 0;
}
