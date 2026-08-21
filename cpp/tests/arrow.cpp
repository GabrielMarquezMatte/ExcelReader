#include <xl/excelreader_arrow.hpp>

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <string_view>

struct Record
{
    std::string_view Coluna1;
    int64_t Coluna3;
};

template <>
struct xl::ExcelMapper<Record>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Coluna1", &Record::Coluna1),
            xl::make_field("Coluna3", &Record::Coluna3));
    }
};

#define CHECK(cond, msg)                                                         \
    do                                                                           \
    {                                                                            \
        if (!(cond))                                                            \
        {                                                                       \
            std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); \
            return 1;                                                           \
        }                                                                       \
    } while (0)

int main()
{
    xl::OpenOptions options{.prefetch_decompression = 1};
    auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
    CHECK(workbook.has_value(), "xl::Workbook::open must succeed on the RealExcel.xlsb fixture");

    auto table = xl::parse_arrow<Record>(*workbook);
    CHECK(table.has_value(), "xl::parse_arrow<Record> must succeed");

    // The export hands back ONE top-level struct array whose children are the columns.
    CHECK(std::strcmp(table->schema.format, "+s") == 0, "top level must be a struct array");
    CHECK(table->schema.n_children == 2, "must have two child columns");
    CHECK(table->array.n_children == 2, "must have two child arrays");
    CHECK(std::strcmp(table->schema.children[0]->name, "Coluna1") == 0, "first column must be named Coluna1");
    CHECK(std::strcmp(table->schema.children[0]->format, "u") == 0, "Coluna1 must be utf8");
    CHECK(std::strcmp(table->schema.children[1]->format, "l") == 0, "Coluna3 must be int64");
    CHECK(table->array.length == 100, "RealExcel.xlsb has 100 data rows");

    // Destructor must release both; running under a leak checker in CI is what proves it, but a
    // move-then-destroy here at least exercises the moved-from path being inert.
    {
        xl::ArrowTable moved = std::move(*table);
        CHECK(moved.array.release != nullptr, "moved-to table must still own a release callback");
        CHECK(table->array.release == nullptr, "moved-from table must be released/inert");
    }

    std::printf("OK: C++ arrow test passed\n");
    return 0;
}
