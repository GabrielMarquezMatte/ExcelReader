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

    // --- batch_size edge cases and content parity with xl::parse_arrow -------------------------
    {
        auto wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(wb.has_value(), "opening a workbook for the batch_size=0 stream must succeed");
        auto stream = xl::arrow_stream<Record>(*wb, 1, 0);
        CHECK(stream.has_value(), "xl::arrow_stream<Record> with batch_size=0 must succeed");
        auto batch = stream->next();
        CHECK(batch.has_value() && batch->has_value(), "batch_size=0 must yield one batch");
        CHECK((*batch)->array.length == 100, "the one batch must hold all 100 rows");
        auto end = stream->next();
        CHECK(end.has_value() && !end->has_value(), "the stream ends after the single batch");
    }

    {
        auto wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(wb.has_value(), "opening a workbook for the negative-batch-size stream must succeed");
        auto stream = xl::arrow_stream<Record>(*wb, 1, -1);
        CHECK(!stream.has_value(), "a negative batch size must be rejected");
        if (!stream.has_value())
        {
            CHECK(stream.error().code == XL_INVALID_ARGUMENT, "the rejection is XL_INVALID_ARGUMENT");
        }
    }

    {
        auto whole_wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(whole_wb.has_value(), "opening a workbook for the whole-sheet comparison must succeed");
        auto whole = xl::parse_arrow<Record>(*whole_wb);
        CHECK(whole.has_value(), "the whole-sheet parse_arrow must succeed for comparison");

        auto stream_wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(stream_wb.has_value(), "opening a workbook for the streamed comparison must succeed");
        auto stream = xl::arrow_stream<Record>(*stream_wb, 1, 4);
        CHECK(stream.has_value(), "xl::arrow_stream<Record> at batch_size=4 must succeed");

        auto first = stream->next();
        CHECK(first.has_value() && first->has_value(), "the first streamed batch must read");
        CHECK((*first)->array.n_children == whole->array.n_children,
              "batch child-array count must match the whole-sheet read");
        CHECK((*first)->array.length == 4, "batch size 4's first batch holds 4 rows");

        const auto *whole_coluna3 = static_cast<const int64_t *>(whole->array.children[1]->buffers[1]);
        const auto *batch_coluna3 = static_cast<const int64_t *>((*first)->array.children[1]->buffers[1]);
        CHECK(batch_coluna3[0] == whole_coluna3[0], "first row's Coluna3 must match the whole-sheet read");

        const auto *whole_offsets = static_cast<const int32_t *>(whole->array.children[0]->buffers[1]);
        const auto *whole_data = static_cast<const char *>(whole->array.children[0]->buffers[2]);
        const auto *batch_offsets = static_cast<const int32_t *>((*first)->array.children[0]->buffers[1]);
        const auto *batch_data = static_cast<const char *>((*first)->array.children[0]->buffers[2]);
        std::string_view whole_first(whole_data + whole_offsets[0], whole_offsets[1] - whole_offsets[0]);
        std::string_view batch_first(batch_data + batch_offsets[0], batch_offsets[1] - batch_offsets[0]);
        CHECK(batch_first == whole_first, "first row's Coluna1 must match the whole-sheet read");

        int64_t rows = (*first)->array.length;
        while (true)
        {
            auto next_batch = stream->next();
            CHECK(next_batch.has_value(), "a stream batch must read without error");
            if (!next_batch->has_value())
            {
                break;
            }
            rows += (*next_batch)->array.length;
        }
        CHECK(rows == whole->array.length, "streamed total row count must match the whole-sheet read");
    }

    // --- foreign read invalidates a live stream, and the failure latches ------------------------
    {
        auto wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(wb.has_value(), "opening a workbook for the foreign-read test must succeed");
        auto stream = xl::arrow_stream<Record>(*wb, 1, 4);
        CHECK(stream.has_value(), "the stream opens before the foreign read");
        auto first = stream->next();
        CHECK(first.has_value() && first->has_value(), "the first batch reads before the foreign read");

        auto stolen = xl::parse_sheet<Record>(*wb);
        CHECK(stolen.has_value(), "the foreign whole-sheet read succeeds");

        auto after = stream->next();
        CHECK(!after.has_value(), "the invalidated stream fails");
        auto again = stream->next();
        CHECK(!again.has_value(), "the failure latches on every later call");
        if (!after.has_value() && !again.has_value())
        {
            CHECK(after.error().message == again.error().message, "the latched message is stable");
        }
    }

    // --- one chunked read per workbook, across kinds --------------------------------------------
    {
        auto wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(wb.has_value(), "opening a workbook for the reader-then-stream test must succeed");
        auto reader = xl::typed_reader<Record>(*wb, 1, 8);
        CHECK(reader.has_value(), "the typed reader opens first");
        auto stream = xl::arrow_stream<Record>(*wb, 1, 8);
        CHECK(!stream.has_value(), "an Arrow stream must be rejected while a typed reader is live");
    }
    {
        auto wb = xl::Workbook::open(EXCELREADER_FIXTURE_PATH, XL_FORMAT_XLSB, &options);
        CHECK(wb.has_value(), "opening a workbook for the stream-then-reader test must succeed");
        auto stream = xl::arrow_stream<Record>(*wb, 1, 8);
        CHECK(stream.has_value(), "the Arrow stream opens first");
        auto reader = xl::typed_reader<Record>(*wb, 1, 8);
        CHECK(!reader.has_value(), "a typed reader must be rejected while an Arrow stream is live");
    }

    std::printf("OK: C++ arrow test passed\n");
    return 0;
}
