#include <xl/excelreader.hpp>

#include <cstdio>

int main()
{
    auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSX, nullptr);
    if (!workbook.has_value())
    {
        std::fprintf(stderr, "FAIL: the installed package could not open %s\n", EXCELREADER_FIXTURE_PATH);
        return 1;
    }

    auto cursor = workbook->rows();
    int rows = 0;
    while (cursor.next_row().has_value())
    {
        ++rows;
    }

    if (rows == 0)
    {
        std::fprintf(stderr, "FAIL: the installed package read no rows\n");
        return 1;
    }

    std::printf("read %d rows through the installed package\n", rows);
    return 0;
}
