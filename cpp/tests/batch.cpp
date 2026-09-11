#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

struct Row
{
    std::string_view Coluna1;
    int64_t Coluna3;
};

template <>
struct xl::ExcelMapper<Row>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("Coluna1", &Row::Coluna1),
            xl::make_field("Coluna3", &Row::Coluna3));
    }
};

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

    // Row::Coluna1 is a string_view into the batch's own buffers, so every comparison copies it
    // out - the batch it points into is freed before the next one is compared.
    using OwnedRow = std::pair<std::string, int64_t>;

    std::vector<OwnedRow> whole_sheet()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the whole-sheet read");
        if (!workbook.has_value())
        {
            return {};
        }
        auto table = xl::parse_sheet<Row>(*workbook);
        check(table.has_value(), "parse_sheet succeeds");
        if (!table.has_value())
        {
            return {};
        }

        std::vector<OwnedRow> rows;
        for (Row row : *table)
        {
            rows.emplace_back(std::string(row.Coluna1), row.Coluna3);
        }
        return rows;
    }

    // The load-bearing property: chunked output equals whole-sheet output, at every batch size.
    // The sizes that are not multiples of 8 are where a validity-bitmap boundary bug surfaces.
    void test_batches_equal_whole_sheet(const std::vector<OwnedRow> &expected)
    {
        const int64_t rows = static_cast<int64_t>(expected.size());
        for (int64_t batch_size : {int64_t{1}, int64_t{7}, int64_t{8}, int64_t{9}, int64_t{1000},
                                   int64_t{0}, rows + 10})
        {
            auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
            check(workbook.has_value(), "open the fixture for a batched read");
            if (!workbook.has_value())
            {
                return;
            }
            auto reader = xl::typed_reader<Row>(*workbook, 1, batch_size);
            check(reader.has_value(), "typed_reader opens");
            if (!reader.has_value())
            {
                return;
            }

            std::vector<OwnedRow> actual;
            int64_t batches = 0;
            while (true)
            {
                auto batch = reader->next();
                check(batch.has_value(), "a batch reads without error");
                if (!batch.has_value() || !batch->has_value())
                {
                    break;
                }
                ++batches;
                if (batch_size > 0)
                {
                    check((*batch)->size() <= batch_size, "no batch exceeds the requested size");
                }
                for (Row row : **batch)
                {
                    actual.emplace_back(std::string(row.Coluna1), row.Coluna3);
                }
            }

            check(actual == expected, "batched rows match the whole-sheet read");

            // Pins the batching itself: an implementation that ignored batch_size and returned one
            // big batch would satisfy every assertion above.
            const int64_t want = batch_size == 0 ? 1 : (rows + batch_size - 1) / batch_size;
            check(batches == want, "batch count matches ceil(rows / batch_size)");
        }
    }

    void test_range_for_matches_next(const std::vector<OwnedRow> &expected)
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the range-for read");
        if (!workbook.has_value())
        {
            return;
        }
        auto reader = xl::typed_reader<Row>(*workbook, 1, 8);
        check(reader.has_value(), "typed_reader opens for the range-for read");
        if (!reader.has_value())
        {
            return;
        }

        int64_t seen = 0;
        for (auto &batch : *reader)
        {
            check(batch.has_value(), "range-for yields a batch without error");
            if (!batch.has_value())
            {
                break;
            }
            seen += batch->size();
        }
        check(seen == static_cast<int64_t>(expected.size()), "range-for sees every row");
    }

    void test_second_reader_is_rejected()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the second-reader test");
        if (!workbook.has_value())
        {
            return;
        }
        auto first = xl::typed_reader<Row>(*workbook, 1, 8);
        check(first.has_value(), "the first reader opens");
        auto second = xl::typed_reader<Row>(*workbook, 1, 8);
        check(!second.has_value(), "a second reader on the same workbook is rejected");
        if (!second.has_value())
        {
            check(second.error().code == XL_ERROR, "the rejection is XL_ERROR");
            check(!second.error().message.empty(), "the rejection carries a message");
        }
    }

    void test_negative_batch_size_is_rejected()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the negative-batch test");
        if (!workbook.has_value())
        {
            return;
        }
        auto reader = xl::typed_reader<Row>(*workbook, 1, -1);
        check(!reader.has_value(), "a negative batch size is rejected");
        if (!reader.has_value())
        {
            check(reader.error().code == XL_INVALID_ARGUMENT, "the rejection is XL_INVALID_ARGUMENT");
        }
    }

    // ABI rule: any other read on the workbook invalidates a live reader, and the failure latches
    // rather than silently resuming from the moved cursor.
    void test_foreign_read_latches_the_error()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the invalidation test");
        if (!workbook.has_value())
        {
            return;
        }
        auto reader = xl::typed_reader<Row>(*workbook, 1, 4);
        check(reader.has_value(), "the reader opens before the foreign read");
        if (!reader.has_value())
        {
            return;
        }
        auto first = reader->next();
        check(first.has_value() && first->has_value(), "the first batch reads before the foreign read");

        auto stolen = xl::parse_sheet<Row>(*workbook);
        check(stolen.has_value(), "the foreign whole-sheet read succeeds");

        auto after = reader->next();
        check(!after.has_value(), "the invalidated reader fails");
        auto again = reader->next();
        check(!again.has_value(), "the failure latches on every later call");
        if (!after.has_value() && !again.has_value())
        {
            check(after.error().message == again.error().message, "the latched message is stable");
        }
    }
}

int main()
{
    std::vector<OwnedRow> expected = whole_sheet();
    check(expected.size() == 100, "RealExcel.xlsb has 100 data rows");
    test_batches_equal_whole_sheet(expected);
    test_range_for_matches_next(expected);
    test_second_reader_is_rejected();
    test_negative_batch_size_is_rejected();
    test_foreign_read_latches_the_error();

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return 1;
    }
    std::printf("all batch checks passed\n");
    return 0;
}
