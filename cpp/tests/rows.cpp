#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <string>
#include <vector>

namespace
{
    int failures = 0;

    void check(bool condition, const char *what)
    {
        if (!condition)
        {
            std::fprintf(stderr, "FAILED: %s\n", what);
            ++failures;
        }
    }

    std::string fixture()
    {
        return EXCELREADER_XLSX_FIXTURE_PATH;
    }

    std::vector<std::vector<std::string>> read_with_cursor(const std::string &path)
    {
        auto workbook = xl::Workbook::open(path);
        check(workbook.has_value(), "open the fixture");
        if (!workbook.has_value())
        {
            return {};
        }

        std::vector<std::vector<std::string>> all;
        auto cursor = workbook->rows();
        while (true)
        {
            auto row = cursor.next_row();
            if (!row.has_value())
            {
                check(row.error().code == XL_EOF, "the cursor stops on XL_EOF, not a real error");
                break;
            }
            std::vector<std::string> cells;
            for (auto cell : *row)
            {
                cells.emplace_back(cell.value);
            }
            all.push_back(std::move(cells));
        }
        return all;
    }
}

int main()
{
    const auto rows = read_with_cursor(fixture());
    check(rows.size() > 1, "the fixture has a header plus data rows");

    bool found_coluna1 = false;
    for (const auto &cell : rows.front())
    {
        if (cell == "Coluna1")
        {
            found_coluna1 = true;
        }
    }
    check(found_coluna1, "the header row contains Coluna1");

    // A cell far larger than the cursor's initial buffer forces the XL_BUFFER_TOO_SMALL retry.
    {
        const auto dir = std::filesystem::temp_directory_path() / "excelreader-cpp-rows-test";
        std::filesystem::create_directories(dir);
        const auto path = dir / "wide.csv";
        {
            std::ofstream out(path, std::ios::binary);
            out << "a\n" << std::string(200000, 'x') << "\n";
        }

        auto workbook = xl::Workbook::open(path.string(), XL_FORMAT_CSV);
        check(workbook.has_value(), "open the oversized CSV");
        if (workbook.has_value())
        {
            auto cursor = workbook->rows();
            auto header = cursor.next_row();
            check(header.has_value() && (*header)[0].value == "a", "header survives");

            auto big = cursor.next_row();
            check(big.has_value(), "the oversized row is returned, not lost");
            if (big.has_value())
            {
                check((*big)[0].value.size() == 200000, "the oversized row is complete");
                check((*big)[0].type == xl::CellType::String, "the oversized cell is a string");
            }
        }
        std::error_code ignored;
        std::filesystem::remove(path, ignored);
    }

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return EXIT_FAILURE;
    }
    std::puts("rows: OK");
    return EXIT_SUCCESS;
}
