#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string_view>

struct Row
{
    std::string_view Coluna1;
    std::chrono::year_month_day Coluna2;
    int64_t Coluna3;
    double Coluna16;
};

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

struct AliasRow
{
    std::string_view Coluna1;
};

template <>
struct xl::ExcelMapper<AliasRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field({"ThisColumnDoesNotExist", "Coluna1"}, &AliasRow::Coluna1));
    }
};

#define CHECK(cond, msg)                                                         \
    do                                                                           \
    {                                                                            \
        if (!(cond))                                                             \
        {                                                                        \
            std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); \
            return 1;                                                            \
        }                                                                        \
    } while (0)

static int test_parse(xl::Workbook &workbook)
{
    auto table = xl::parse_sheet<Row>(workbook);
    CHECK(table.has_value(), "xl::parse_sheet<Row> must succeed");
    CHECK(table->size() == 100, "RealExcel.xlsb has 100 data rows");

    auto it = table->begin();
    Row first = *it;
    CHECK(first.Coluna1 == "Valor1", "first row's Coluna1 must be Valor1");
    CHECK(first.Coluna2 == std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1}, "first row's Coluna2 must be 2026-01-01");
    CHECK(first.Coluna3 == 1, "first row's Coluna3 must be 1");
    CHECK(first.Coluna16 == 0.1, "first row's Coluna16 must be 0.1");

    // at() is the bounds-checked counterpart to operator[]; before it existed, every out-of-range
    // row read past the columnar buffers and returned garbage.
    CHECK(table->at(0).has_value(), "at(0) must be in bounds");
    CHECK(table->at(table->size() - 1).has_value(), "at(size() - 1) must be in bounds");
    CHECK(!table->at(table->size()).has_value(), "at(size()) must be out of bounds");
    CHECK(!table->at(-1).has_value(), "at(-1) must be out of bounds");
    CHECK(!table->at(INT64_MAX).has_value(), "at(INT64_MAX) must be out of bounds");
    CHECK(table->at(0)->Coluna1 == "Valor1", "at(0) must decode the same row as operator[]");
    return 0;
}

static int test_sheets(xl::Workbook &workbook)
{
    auto count = workbook.sheet_count();
    CHECK(count.has_value(), "sheet_count must succeed");
    CHECK(*count >= 1, "the fixture has at least one sheet");

    auto names = workbook.sheet_names();
    CHECK(names.has_value(), "sheet_names must succeed");
    CHECK(names->size() == static_cast<size_t>(*count), "sheet_names must return one name per sheet");

    auto current = workbook.sheet_name();
    CHECK(current.has_value(), "sheet_name must succeed");
    CHECK(*current == names->front(), "the first sheet is selected before any move_to_sheet");

    CHECK(workbook.move_to_sheet(0).has_value(), "move_to_sheet(0) must succeed");
    CHECK(!workbook.move_to_sheet(*count).has_value(), "an index past the last sheet must fail");

    // Reading it is enough - which system the fixture uses is not this test's business.
    CHECK(workbook.is_date1904().has_value(), "is_date1904 must succeed");
    return 0;
}

static int test_infer_schema(const xl::Workbook &workbook)
{
    auto schema = workbook.infer_schema(1, 100);
    CHECK(schema.has_value(), "infer_schema must succeed");
    CHECK(!schema->empty(), "the fixture has columns to infer");
    CHECK(schema->front().name.has_value(), "the first column must carry a header name");
    CHECK(*schema->front().name == "Coluna1", "the first column must be Coluna1");
    CHECK(schema->front().type == XL_T_STRING, "Coluna1 must be inferred as a string column");
    return 0;
}

static int test_parse_with_alias(xl::Workbook &workbook)
{
    auto table = xl::parse_sheet<AliasRow>(workbook);
    CHECK(table.has_value(), "xl::parse_sheet<AliasRow> must succeed by resolving the second candidate name");
    CHECK(table->size() == 100, "RealExcel.xlsb has 100 data rows");
    AliasRow first = *table->begin();
    return 0;
}

int main()
{
    CHECK(xl::abi_version() == XL_ABI_VERSION, "the linked native library must speak this header's ABI revision");

    xl::OpenOptions options{.prefetch_decompression = 1};
    auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
    CHECK(workbook.has_value(), "xl::Workbook::open must succeed on the RealExcel.xlsb fixture");

    auto missing = xl::Workbook::open("does-not-exist.xlsx");
    CHECK(!missing.has_value(), "opening a missing file must fail");
    CHECK(!missing.error().message.empty(), "a failure must carry the native error detail");

    if (int failed = test_sheets(*workbook))
    {
        return failed;
    }
    if (int failed = test_infer_schema(*workbook))
    {
        return failed;
    }
    // Last: parse_sheet consumes the workbook's shared row cursor.
    if (int failed = test_parse(*workbook))
    {
        return failed;
    }

    std::printf("OK: C++ smoke test passed\n");
    return 0;
}
