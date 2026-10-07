#include <xl/excelreader.hpp>

#include <atomic>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

namespace
{
    std::atomic<int> failures = 0;

    void check(bool condition, const char *what)
    {
        if (!condition)
        {
            std::fprintf(stderr, "FAILED: %s\n", what);
            ++failures;
        }
    }

    std::vector<uint8_t> build(int sheet_count, int row_count)
    {
        auto writer = xl::WriterHandle::open_memory(XL_FORMAT_XLSX);
        check(writer.has_value(), "open the in-memory writer");
        if (!writer.has_value())
        {
            return {};
        }
        for (int s = 0; s < sheet_count; ++s)
        {
            const std::string name = "sheet" + std::to_string(s);
            check(writer->start_sheet(name).has_value(), "start a sheet");
            for (int r = 0; r < row_count; ++r)
            {
                const std::string value = "s" + std::to_string(s) + "-r" + std::to_string(r);
                check(writer->start_row().has_value(), "start a row");
                check(writer->write(value).has_value(), "write a cell");
                check(writer->end_row().has_value(), "end a row");
            }
            check(writer->end_sheet().has_value(), "end a sheet");
        }
        auto bytes = writer->bytes();
        check(bytes.has_value(), "read the written bytes");
        return bytes.has_value() ? std::move(*bytes) : std::vector<uint8_t>{};
    }

    std::vector<std::string> drain(xl::RowCursor &cursor)
    {
        std::vector<std::string> values;
        while (true)
        {
            auto row = cursor.next_row();
            if (!row.has_value())
            {
                check(row.error().code == XL_EOF, "the cursor stops on XL_EOF, not a real error");
                return values;
            }
            values.emplace_back((*row)[0].value);
        }
    }

    std::vector<std::string> read_sheet(xl::Sheet sheet)
    {
        auto cursor = sheet.rows();
        check(cursor.has_value(), "open a cursor");
        return cursor.has_value() ? drain(*cursor) : std::vector<std::string>{};
    }
}

int main()
{
    const std::vector<uint8_t> two = build(2, 3);

    {
        auto workbook = xl::Workbook::open_memory(two);
        check(workbook.has_value(), "open the two-sheet workbook");
        if (workbook.has_value())
        {
            const auto second = read_sheet(workbook->sheet(1));
            check(second.size() == 3 && second.front() == "s1-r0", "sheet 1 reads its own rows");
            const auto first = read_sheet(workbook->sheet(0));
            check(first.size() == 3 && first.front() == "s0-r0", "sheet 0 reads its own rows");

            auto sheets = workbook->sheets();
            check(sheets.has_value() && sheets->size() == 2, "sheets() lists both sheets");
            if (sheets.has_value() && sheets->size() == 2)
            {
                check((*sheets)[1].index() == 1, "a sheet knows its index");
                auto name = (*sheets)[1].name();
                check(name.has_value() && *name == "sheet1", "a sheet knows its name");
                auto visibility = (*sheets)[0].visibility();
                check(visibility.has_value() && *visibility == xl::SheetVisibility::Visible, "a written sheet is visible");
            }

            auto exact = workbook->sheet_by_name("sheet1");
            check(exact.has_value() && exact->has_value() && (*exact)->index() == 1, "sheet_by_name finds an exact name");
            auto other_case = workbook->sheet_by_name("SHEET1");
            check(other_case.has_value() && other_case->has_value() && (*other_case)->index() == 1, "sheet_by_name ignores case");
            auto missing = workbook->sheet_by_name("nope");
            check(missing.has_value() && !missing->has_value(), "a missing name is an empty optional, not an error");

            auto past_end = workbook->sheet(2).rows();
            check(!past_end.has_value() && past_end.error().code == XL_ERROR, "a sheet index past the end fails at the first read");
            auto negative = workbook->sheet(-1).rows();
            check(!negative.has_value() && negative.error().code == XL_INVALID_ARGUMENT, "a negative sheet index is an invalid argument");

            auto decoded = workbook->sheet(1).read_all_decoded();
            check(decoded.has_value() && decoded->size() == 3, "read_all_decoded reads the whole named sheet");
            auto partial = workbook->sheet(1).rows();
            check(partial.has_value() && partial->next_row().has_value(), "open a cursor and read one row");
            if (partial.has_value())
            {
                auto rest = partial->read_all_decoded();
                check(rest.has_value() && rest->size() == 2, "the cursor's read_all_decoded returns the rows it has left");
            }
            auto schema = workbook->sheet(1).infer_schema(0, 10);
            check(schema.has_value() && schema->size() == 1, "infer_schema samples the named sheet");
        }
    }

    {
        auto workbook = xl::Workbook::open_memory(two);
        check(workbook.has_value(), "open for two cursors");
        if (workbook.has_value())
        {
            auto one = workbook->sheet(1).rows();
            auto another = workbook->sheet(1).rows();
            check(one.has_value() && another.has_value(), "two cursors open on one sheet");
            if (one.has_value() && another.has_value())
            {
                check(one->next_row().has_value(), "the first cursor advances");
                check(another->next_row().has_value(), "the second cursor advances on its own");
                check(drain(*one).size() == 2, "the first cursor has two rows left");
                check(drain(*another).size() == 2, "the second cursor has two rows left");
            }
        }
    }

    {
        auto workbook = xl::Workbook::open_memory(two);
        check(workbook.has_value(), "open for the moved cursor");
        if (workbook.has_value())
        {
            auto cursor = workbook->sheet(0).rows();
            check(cursor.has_value(), "open a cursor to move");
            if (cursor.has_value())
            {
                xl::RowCursor moved = std::move(*cursor);
                check(cursor->handle() == nullptr, "a moved-from cursor holds nothing");
                check(drain(moved).size() == 3, "the moved-to cursor reads every row");
            }
            check(read_sheet(workbook->sheet(0)).size() == 3, "the workbook still reads after the cursors are gone");
        }
    }

    {
        auto workbook = xl::Workbook::open_memory(two);
        check(workbook.has_value(), "open for the move-assigned cursor");
        if (workbook.has_value())
        {
            auto a = workbook->sheet(0).rows();
            auto b = workbook->sheet(1).rows();
            check(a.has_value() && b.has_value(), "open two cursors to move-assign");
            if (a.has_value() && b.has_value())
            {
                check(a->next_row().has_value(), "the overwritten cursor advances");
                *a = std::move(*b);
                check(b->handle() == nullptr, "a move-assigned-from cursor holds nothing");
                const auto values = drain(*a);
                check(values.size() == 3 && values.front() == "s1-r0", "the move-assigned cursor reads the other sheet from the top");
            }
        }
    }

    {
        const std::vector<uint8_t> empty_book = build(1, 0);
        auto workbook = xl::Workbook::open_memory(empty_book);
        check(workbook.has_value(), "open the empty workbook");
        if (workbook.has_value())
        {
            auto decoded = workbook->sheet(0).read_all_decoded();
            check(decoded.has_value() && decoded->empty(), "an empty sheet decodes to no rows");
        }
    }

    {
        std::optional<xl::RowCursor> survivor;
        std::optional<xl::Sheet> stale;
        {
            auto workbook = xl::Workbook::open_memory(two);
            check(workbook.has_value(), "open for the outliving cursor");
            if (workbook.has_value())
            {
                xl::Workbook relocated = std::move(*workbook);
                xl::Sheet sheet = relocated.sheet(1);
                check(read_sheet(sheet).size() == 3, "a sheet taken after the workbook moved still reads");
                auto cursor = sheet.rows();
                check(cursor.has_value(), "open a cursor before the workbook closes");
                if (cursor.has_value())
                {
                    survivor.emplace(std::move(*cursor));
                }
                stale = sheet;
            }
        }
        check(survivor.has_value() && drain(*survivor).size() == 3, "a cursor reads every row after its workbook closed");
        if (stale.has_value())
        {
            auto after_close = stale->rows();
            check(!after_close.has_value() && after_close.error().code == XL_INVALID_HANDLE, "a sheet of a closed workbook reports an invalid handle");
        }
    }

    {
        constexpr int sheet_count = 6;
        constexpr int row_count = 500;
        const std::vector<uint8_t> many = build(sheet_count, row_count);
        auto workbook = xl::Workbook::open_memory(many);
        check(workbook.has_value(), "open the many-sheet workbook");
        if (workbook.has_value())
        {
            std::vector<std::vector<std::string>> sequential;
            for (int s = 0; s < sheet_count; ++s)
            {
                sequential.push_back(read_sheet(workbook->sheet(s)));
            }
            check(sequential[3].size() == static_cast<size_t>(row_count) && sequential[3].front() == "s3-r0", "the sequential read sees sheet 3");

            std::vector<std::vector<std::string>> parallel(static_cast<size_t>(sheet_count) * 2);
            {
                std::vector<std::thread> threads;
                for (size_t i = 0; i < parallel.size(); ++i)
                {
                    threads.emplace_back([&parallel, &workbook, i]
                                         { parallel[i] = read_sheet(workbook->sheet(static_cast<int32_t>(i % sheet_count))); });
                }
                for (std::thread &thread : threads)
                {
                    thread.join();
                }
            }
            bool same = true;
            for (size_t i = 0; i < parallel.size(); ++i)
            {
                same = same && parallel[i] == sequential[i % sheet_count];
            }
            check(same, "every sheet read on its own thread matches the sequential read");
        }
    }

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures.load());
        return 1;
    }
    std::puts("OK: sheets");
    return 0;
}
