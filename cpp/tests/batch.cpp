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

    using OwnedRow = std::pair<std::string, int64_t>;

    std::vector<OwnedRow> whole_sheet()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the whole-sheet read");
        if (!workbook.has_value())
        {
            return {};
        }
        auto table = xl::parse<Row>(workbook->sheet(0));
        check(table.has_value(), "parse succeeds");
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

    void test_batches_equal_whole_sheet(const std::vector<OwnedRow> &expected)
    {
        const int64_t rows = static_cast<int64_t>(expected.size());
        for (int64_t batch_size : {int64_t{1}, int64_t{7}, int64_t{8}, int64_t{9}, int64_t{1000}, rows + 10})
        {
            auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
            check(workbook.has_value(), "open the fixture for a batched read");
            if (!workbook.has_value())
            {
                return;
            }
            auto reader = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = batch_size});
            check(reader.has_value(), "parse_batches opens");
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
                check((*batch)->size() <= batch_size, "no batch exceeds the requested size");
                for (Row row : **batch)
                {
                    actual.emplace_back(std::string(row.Coluna1), row.Coluna3);
                }
            }

            check(actual == expected, "batched rows match the whole-sheet read");

            const int64_t want = (rows + batch_size - 1) / batch_size;
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
        auto reader = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = 8});
        check(reader.has_value(), "parse_batches opens for the range-for read");
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

    void test_second_reader_is_allowed()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the second-reader test");
        if (!workbook.has_value())
        {
            return;
        }
        auto first = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = 8});
        check(first.has_value(), "the first reader opens");
        auto second = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = 8});
        check(second.has_value(), "a second reader on the same workbook opens");
    }

    void test_non_positive_batch_size_is_rejected()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the non-positive-batch test");
        if (!workbook.has_value())
        {
            return;
        }
        for (int64_t batch_size : {int64_t{0}, int64_t{-1}})
        {
            auto reader = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = batch_size});
            check(!reader.has_value(), "a batch size below 1 is rejected");
            if (!reader.has_value())
            {
                check(reader.error().code == XL_INVALID_ARGUMENT, "the rejection is XL_INVALID_ARGUMENT");
            }
        }
    }

    void test_foreign_read_leaves_the_reader_alone()
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the foreign-read test");
        if (!workbook.has_value())
        {
            return;
        }
        auto reader = xl::parse_batches<Row>(workbook->sheet(0), {.batch_size = 4});
        check(reader.has_value(), "the reader opens before the foreign read");
        if (!reader.has_value())
        {
            return;
        }
        auto first = reader->next();
        check(first.has_value() && first->has_value(), "the first batch reads before the foreign read");

        auto stolen = xl::parse<Row>(workbook->sheet(0));
        check(stolen.has_value(), "the foreign whole-sheet read succeeds");

        auto after = reader->next();
        check(after.has_value() && after->has_value(), "the reader keeps reading after a whole-sheet read");
    }

    const xl::ColumnRef *find_column(const xl::Table &table, std::string_view name)
    {
        for (const xl::ColumnRef &column : table.columns())
        {
            if (column.name == name)
            {
                return &column;
            }
        }
        return nullptr;
    }

    void test_runtime_schema_matches_the_typed_parse(const std::vector<OwnedRow> &expected)
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the runtime-schema parse");
        if (!workbook.has_value())
        {
            return;
        }
        auto schema = workbook->sheet(0).infer_schema();
        check(schema.has_value() && !schema->empty(), "infer_schema finds columns");
        if (!schema.has_value())
        {
            return;
        }

        auto table = xl::parse(workbook->sheet(0), *schema);
        check(table.has_value(), "parse over an inferred schema succeeds");
        if (!table.has_value())
        {
            return;
        }
        check(table->size() == static_cast<int64_t>(expected.size()), "the runtime-schema parse reads every row");
        check(table->columns().size() == schema->size(), "one column per inferred column");

        const xl::ColumnRef *coluna1 = find_column(*table, "Coluna1");
        const xl::ColumnRef *coluna3 = find_column(*table, "Coluna3");
        check(coluna1 != nullptr && coluna1->type == XL_T_STRING, "Coluna1 is a string column");
        check(coluna3 != nullptr && coluna3->type == XL_T_I64, "Coluna3 is an i64 column");
        if (coluna1 == nullptr || coluna3 == nullptr || coluna1->type != XL_T_STRING || coluna3->type != XL_T_I64)
        {
            return;
        }
        const auto *offsets = static_cast<const int32_t *>(coluna1->values);
        const auto *text = reinterpret_cast<const char *>(coluna1->data);
        const auto *numbers = static_cast<const int64_t *>(coluna3->values);
        bool same = true;
        for (size_t row = 0; row < expected.size(); ++row)
        {
            const std::string_view cell(text + offsets[row], static_cast<size_t>(offsets[row + 1] - offsets[row]));
            same = same && cell == expected[row].first && numbers[row] == expected[row].second;
        }
        check(same, "the runtime-schema parse matches the typed parse cell for cell");

        std::vector<uint8_t> bytes;
        check(xl::write_columns(bytes, table->columns(), {.format = XL_FORMAT_XLSX}).has_value(),
              "a runtime-schema table writes back as it was read");
        auto reopened = xl::Workbook::open(bytes);
        check(reopened.has_value(), "the written workbook opens");
        if (!reopened.has_value())
        {
            return;
        }
        auto again = xl::parse(reopened->sheet(0), *schema);
        check(again.has_value() && again->size() == table->size(), "the written workbook parses with the same schema");

        auto by_index = workbook->sheet(0).infer_schema(0);
        check(by_index.has_value() && !(*by_index)[0].name.has_value(), "a schema without a header has no names");
        if (by_index.has_value())
        {
            auto unnamed = xl::parse(workbook->sheet(0), *by_index, {.header_row = 0});
            check(unnamed.has_value() && unnamed->size() == static_cast<int64_t>(expected.size()) + 1,
                  "unnamed columns resolve by index, and with no header the header row is data");
        }

        auto empty = xl::parse(workbook->sheet(0), std::span<const xl::InferredColumn>{});
        check(!empty.has_value() && empty.error().code == XL_INVALID_ARGUMENT, "an empty schema is rejected");

        xl::Table moved = std::move(*table);
        check(find_column(moved, "Coluna3") != nullptr, "a moved table keeps its column names");
    }

    void test_runtime_schema_batches_equal_whole_sheet(const std::vector<OwnedRow> &expected)
    {
        auto workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH);
        check(workbook.has_value(), "open the fixture for the runtime-schema batched read");
        if (!workbook.has_value())
        {
            return;
        }
        auto schema = workbook->sheet(0).infer_schema();
        check(schema.has_value(), "infer_schema succeeds for the batched read");
        if (!schema.has_value())
        {
            return;
        }

        auto reader = xl::parse_batches(workbook->sheet(0), *schema, {.batch_size = 8});
        check(reader.has_value(), "parse_batches over an inferred schema opens");
        if (!reader.has_value())
        {
            return;
        }
        schema->clear();

        std::vector<int64_t> coluna3;
        int batches = 0;
        for (auto &batch : *reader)
        {
            check(batch.has_value(), "a runtime-schema batch reads without error");
            if (!batch.has_value())
            {
                return;
            }
            ++batches;
            check(batch->size() <= 8, "no batch exceeds the requested batch size");
            const xl::ColumnRef *column = find_column(*batch, "Coluna3");
            check(column != nullptr && column->type == XL_T_I64, "every batch keeps the schema's names and types");
            if (column == nullptr || column->type != XL_T_I64)
            {
                return;
            }
            const auto *values = static_cast<const int64_t *>(column->values);
            coluna3.insert(coluna3.end(), values, values + column->length);
        }
        check(batches == 13, "100 rows at batch size 8 is 13 batches");
        bool same = coluna3.size() == expected.size();
        for (size_t row = 0; same && row < expected.size(); ++row)
        {
            same = coluna3[row] == expected[row].second;
        }
        check(same, "the batches concatenate to the whole-sheet read");

        auto rejected = xl::parse_batches(workbook->sheet(0), *schema, {.batch_size = 0});
        check(!rejected.has_value() && rejected.error().code == XL_INVALID_ARGUMENT, "a batch size below 1 is rejected");
    }
}

int main()
{
    std::vector<OwnedRow> expected = whole_sheet();
    check(expected.size() == 100, "RealExcel.xlsb has 100 data rows");
    test_batches_equal_whole_sheet(expected);
    test_range_for_matches_next(expected);
    test_second_reader_is_allowed();
    test_non_positive_batch_size_is_rejected();
    test_foreign_read_leaves_the_reader_alone();
    test_runtime_schema_matches_the_typed_parse(expected);
    test_runtime_schema_batches_equal_whole_sheet(expected);

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return 1;
    }
    std::printf("all batch checks passed\n");
    return 0;
}
