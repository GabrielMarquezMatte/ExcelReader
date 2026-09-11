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

    // An out-of-range header_row must fail cleanly, leaving no half-built ArrowTable behind - this
    // matters here more than on the happy path because ~ArrowTable calls through the release
    // function pointers it holds, so a half-initialized table on the failure path would mean the
    // destructor walks into garbage.
    auto failed = xl::parse_arrow<Record>(*workbook, 1'000'000);
    CHECK(!failed.has_value(), "xl::parse_arrow<Record> must fail for an out-of-range header_row");

    // --- Chunked: xl::arrow_stream -------------------------------------------------------------
    {
        auto stream_workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(stream_workbook.has_value(), "opening a workbook for the Arrow stream must succeed");

        auto stream = xl::arrow_stream<Record>(*stream_workbook, 1, 8);
        CHECK(stream.has_value(), "xl::arrow_stream<Record> must succeed");

        auto schema = stream->schema();
        CHECK(schema.has_value(), "the stream must report a schema");
        CHECK(schema->schema.release != nullptr, "the schema must be a live, owned struct");
        CHECK(std::strcmp(schema->schema.format, "+s") == 0, "the stream schema must be a struct array");
        CHECK(schema->schema.n_children == 2, "the stream schema must have two child columns");

        int64_t rows = 0;
        int64_t batches = 0;
        while (true)
        {
            auto batch = stream->next();
            CHECK(batch.has_value(), "a stream batch must read without error");
            if (!batch->has_value())
            {
                break;
            }
            ++batches;
            CHECK((*batch)->array.length <= 8, "no batch may exceed the requested batch size");
            rows += (*batch)->array.length;
        }
        CHECK(rows == 100, "the stream must deliver all 100 data rows");
        CHECK(batches == 13, "100 rows at batch size 8 is 13 batches");

        // One chunked read per workbook: a second stream must be rejected rather than quietly
        // sharing the row cursor.
        auto second = xl::arrow_stream<Record>(*stream_workbook, 1, 8);
        CHECK(!second.has_value(), "a second stream on one workbook must be rejected");
    }

    // A batch outlives the stream that produced it, and releasing it twice is a no-op - the
    // ownership rule a hand-written Arrow consumer has to get right.
    {
        auto abandoned_workbook = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(abandoned_workbook.has_value(), "opening a workbook for the abandonment test must succeed");

        xl::ArrowArrayGuard kept;
        {
            auto stream = xl::arrow_stream<Record>(*abandoned_workbook, 1, 4);
            CHECK(stream.has_value(), "the abandoned stream must open");
            auto batch = stream->next();
            CHECK(batch.has_value() && batch->has_value(), "one batch must read before abandonment");
            kept = std::move(**batch);
        }
        CHECK(kept.array.length == 4, "a batch must outlive the stream that produced it");
        kept.release();
        kept.release();
        CHECK(kept.array.release == nullptr, "releasing must null the release callback");
    }

    std::printf("OK: C++ arrow test passed\n");
    return 0;
}
