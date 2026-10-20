#include <xl/excelreader.hpp>

#include <charconv>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <stdexcept>
#include <tuple>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <string>
#include <utility>
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
        auto cursor = workbook->sheet(0).rows();
        check(cursor.has_value(), "open a cursor");
        if (!cursor.has_value())
        {
            return {};
        }
        while (true)
        {
            auto row = cursor->next_row();
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

    using Cell = std::tuple<int32_t, xl::CellType, std::optional<double>, std::string>;

    std::vector<Cell> collect(xl::RowView row)
    {
        std::vector<Cell> cells;
        for (auto cell : row)
        {
            cells.emplace_back(cell.column, cell.type, cell.as_number(), cell.text());
        }
        return cells;
    }

    std::string shortest(double value)
    {
        char buffer[64];
        const auto result = std::to_chars(buffer, buffer + sizeof(buffer), value);
        return std::string(buffer, result.ptr);
    }

    void for_each_row_matches_the_cursor(const std::string &path)
    {
        auto workbook = xl::Workbook::open(path);
        check(workbook.has_value(), "open the fixture for for_each_row");
        if (!workbook.has_value())
        {
            return;
        }
        const xl::Sheet sheet = workbook->sheet(0);

        std::vector<std::vector<Cell>> pushed;
        auto walked = sheet.for_each_row([&](xl::RowView row)
                                         {
                                             pushed.push_back(collect(row));
                                             return XL_OK; });
        check(walked.has_value(), "for_each_row succeeds");

        std::vector<std::vector<Cell>> pulled;
        auto cursor = sheet.rows();
        check(cursor.has_value(), "open a cursor to compare against");
        if (!cursor.has_value())
        {
            return;
        }
        for (auto row = cursor->next_row(); row.has_value(); row = cursor->next_row())
        {
            pulled.push_back(collect(*row));
        }

        check(pushed.size() == pulled.size(), "for_each_row sees as many rows as the cursor");
        for (size_t i = 0; i < pushed.size() && i < pulled.size(); ++i)
        {
            check(pushed[i].size() == pulled[i].size(), "same cell count per row");
            for (size_t j = 0; j < pushed[i].size() && j < pulled[i].size(); ++j)
            {
                const auto &[column, type, number, text] = pushed[i][j];
                const auto &[cursor_column, cursor_type, cursor_number, cursor_text] = pulled[i][j];
                check(column == cursor_column && type == cursor_type, "same column and type");
                if (!text.empty() && !cursor_text.empty())
                {
                    check(text == cursor_text, "same text when both carry text");
                }
                else
                {
                    check(number == cursor_number, "same number otherwise");
                }
            }
        }
    }

    void number_only_cell_formats_its_text()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the xlsb fixture");
        if (!workbook.has_value())
        {
            return;
        }
        int seen = 0;
        auto walked = workbook->sheet(0).for_each_row([&](xl::RowView row)
                                                      {
                                                          for (auto cell : row)
                                                          {
                                                              if (cell.value.empty() && cell.number.has_value())
                                                              {
                                                                  check(cell.text() == shortest(*cell.number), "a number-only cell formats its number");
                                                                  ++seen;
                                                              }
                                                          }
                                                          return XL_OK; });
        check(walked.has_value(), "for_each_row over the xlsb fixture succeeds");
        check(seen > 0, "the xlsb fixture has a cell that carries only a number");
    }

    void positive_code_stops_the_walk()
    {
        auto workbook = xl::Workbook::open(fixture());
        if (!workbook.has_value())
        {
            check(false, "open the fixture");
            return;
        }
        int calls = 0;
        auto walked = workbook->sheet(0).for_each_row([&](xl::RowView)
                                                      { return ++calls == 3 ? 7 : XL_OK; });
        check(!walked.has_value() && walked.error().code == 7, "a positive code comes back as the error code");
        check(calls == 3, "a positive code stops the walk");
    }

    void non_positive_code_is_an_error()
    {
        auto workbook = xl::Workbook::open(fixture());
        if (!workbook.has_value())
        {
            check(false, "open the fixture");
            return;
        }
        auto walked = workbook->sheet(0).for_each_row([](xl::RowView)
                                                      { return -1; });
        check(!walked.has_value(), "a code <= 0 is still an error");
    }

    void nested_call_inside_the_callback_works()
    {
        auto workbook = xl::Workbook::open(fixture());
        if (!workbook.has_value())
        {
            check(false, "open the fixture");
            return;
        }
        const xl::Sheet sheet = workbook->sheet(0);
        const auto name = sheet.name();
        check(name.has_value(), "read the sheet name");
        auto walked = sheet.for_each_row([&](xl::RowView)
                                         {
                                             const auto nested = sheet.name();
                                             check(nested.has_value() && name.has_value() && *nested == *name, "a nested call returns the right name");
                                             return XL_OK; });
        check(walked.has_value(), "for_each_row with a nested call succeeds");
    }

    void exception_is_reported_and_the_workbook_stays_usable()
    {
        auto workbook = xl::Workbook::open(fixture());
        if (!workbook.has_value())
        {
            check(false, "open the fixture");
            return;
        }
        const xl::Sheet sheet = workbook->sheet(0);
        auto thrown = sheet.for_each_row([](xl::RowView) -> int32_t
                                         { throw std::runtime_error("boom"); });
        check(!thrown.has_value(), "an exception in the callback is reported as an error");
        auto again = sheet.for_each_row([](xl::RowView)
                                        { return XL_OK; });
        check(again.has_value(), "the workbook stays usable after an exception");
    }

    void empty_row_is_empty()
    {
        std::vector<uint8_t> bytes;
        {
            xl::WriteOptions options{};
            options.format = XL_FORMAT_XLSX;
            auto writer = xl::WriterHandle::open(bytes, options);
            check(writer.has_value(), "open a writer");
            if (!writer.has_value())
            {
                return;
            }
            check(writer->start_sheet("S1").has_value(), "start sheet");
            check(writer->start_row().has_value() && writer->end_row().has_value(), "write an empty row");
            check(writer->start_row().has_value(), "start a row");
            check(writer->write(std::string_view("x")).has_value(), "write a cell");
            check(writer->end_row().has_value() && writer->end_sheet().has_value(), "end the sheet");
            check(writer->finish().has_value(), "finish");
        }
        auto workbook = xl::Workbook::open(std::span<const uint8_t>(bytes), xl::OpenOptions{.format = XL_FORMAT_XLSX});
        check(workbook.has_value(), "open the written workbook");
        if (!workbook.has_value())
        {
            return;
        }
        std::vector<size_t> sizes;
        auto walked = workbook->sheet(0).for_each_row([&](xl::RowView row)
                                                      {
                                                          size_t iterated = 0;
                                                          for (auto cell : row)
                                                          {
                                                              (void)cell;
                                                              ++iterated;
                                                          }
                                                          check(iterated == row.size(), "iterating a row visits size() cells");
                                                          sizes.push_back(row.size());
                                                          return XL_OK; });
        check(walked.has_value(), "for_each_row over the written workbook succeeds");
        check(!sizes.empty() && sizes.front() == 0, "the first row seen is empty");
    }
}

int main()
{
    for_each_row_matches_the_cursor(fixture());
    for_each_row_matches_the_cursor(EXCELREADER_FIXTURE_PATH);
    number_only_cell_formats_its_text();
    positive_code_stops_the_walk();
    non_positive_code_is_an_error();
    nested_call_inside_the_callback_works();
    exception_is_reported_and_the_workbook_stays_usable();
    empty_row_is_empty();

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

    {
        const auto dir = std::filesystem::temp_directory_path() / "excelreader-cpp-rows-test";
        std::filesystem::create_directories(dir);
        const auto path = dir / "wide.csv";
        {
            std::ofstream out(path, std::ios::binary);
            out << "a\n" << std::string(200000, 'x') << "\n";
        }

        auto workbook = xl::Workbook::open(path.string(), {.format = XL_FORMAT_CSV});
        check(workbook.has_value(), "open the oversized CSV");
        if (workbook.has_value())
        {
            auto cursor = workbook->sheet(0).rows();
            check(cursor.has_value(), "open a cursor on the oversized CSV");
            if (!cursor.has_value())
            {
                return 1;
            }
            auto header = cursor->next_row();
            check(header.has_value() && (*header)[0].value == "a", "header survives");

            auto big = cursor->next_row();
            check(big.has_value(), "the oversized row is returned, not lost");
            if (big.has_value())
            {
                check((*big)[0].value.size() == 200000, "the oversized row is complete");
                check((*big)[0].type == xl::CellType::String, "the oversized cell is a string");
            }
        }

        auto limited = xl::Workbook::open(path.string(), {.format = XL_FORMAT_CSV, .max_cell_bytes = 1000});
        check(limited.has_value(), "open the oversized CSV under a cell limit");
        if (limited.has_value())
        {
            auto all = limited->sheet(0).read_all();
            check(!all.has_value(), "max_cell_bytes bounds a CSV cell too");
        }
        std::error_code ignored;
        std::filesystem::remove(path, ignored);
    }

    {
        auto workbook = xl::Workbook::open(fixture());
        check(workbook.has_value(), "open for read_all");
        if (workbook.has_value())
        {
            auto decoded = workbook->sheet(0).read_all();
            check(decoded.has_value(), "read_all succeeds");
            if (decoded.has_value())
            {
                check(decoded->size() == rows.size(), "same row count as the cursor");
                for (size_t i = 0; i < decoded->size() && i < rows.size(); ++i)
                {
                    const auto row = (*decoded)[i];
                    check(row.size() == rows[i].size(), "same cell count");
                    size_t j = 0;
                    for (auto cell : row)
                    {
                        if (j < rows[i].size())
                        {
                            check(std::string(cell.value) == rows[i][j], "same cell value");
                        }
                        ++j;
                    }
                }
            }
        }
    }

    {
        auto workbook = xl::Workbook::open(fixture());
        if (workbook.has_value())
        {
            auto cursor = workbook->sheet(0).rows();
            check(cursor.has_value(), "open a cursor to drain");
            while (cursor.has_value() && cursor->next_row().has_value())
            {
            }
            auto decoded = workbook->sheet(0).read_all();
            check(decoded.has_value(), "a read after another cursor drained is not an error");
            check(decoded.has_value() && decoded->size() == rows.size(), "and it still reads every row, from the start");
        }
    }

    {
        auto workbook = xl::Workbook::open(fixture());
        if (workbook.has_value())
        {
            auto decoded = workbook->sheet(0).read_all();
            check(decoded.has_value(), "open a set to move");
            if (decoded.has_value())
            {
                xl::DecodedRows moved = std::move(*decoded);
                check(moved.size() > 0, "the moved-to object owns the rows");
            }
        }
    }

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return EXIT_FAILURE;
    }
    std::puts("rows: OK");
    return EXIT_SUCCESS;
}
