#include <xl/excelreader.hpp>

#include <algorithm>
#include <atomic>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <iterator>
#include <memory>
#include <stdexcept>
#include <string>
#include <thread>
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

    std::vector<std::byte> read_file(const char *path)
    {
        std::ifstream in(path, std::ios::binary);
        std::vector<char> chars((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        std::vector<std::byte> bytes(chars.size());
        std::memcpy(bytes.data(), chars.data(), chars.size());
        return bytes;
    }

    struct Counters
    {
        std::atomic<int> reads{0};
        std::atomic<int> destroyed{0};
    };

    class MemorySource final : public xl::Source
    {
    public:
        MemorySource(std::vector<std::byte> bytes, Counters &counters) : bytes_(std::move(bytes)), counters_(counters) {}
        ~MemorySource() override { counters_.destroyed.fetch_add(1); }

        std::uint64_t size() const override { return bytes_.size(); }

        std::expected<std::size_t, std::string> read_at(std::uint64_t offset, std::span<std::byte> buffer) const override
        {
            counters_.reads.fetch_add(1);
            if (offset >= bytes_.size())
            {
                return std::size_t{0};
            }
            const std::size_t count = std::min<std::size_t>(buffer.size(), bytes_.size() - offset);
            std::memcpy(buffer.data(), bytes_.data() + offset, count);
            return count;
        }

    private:
        std::vector<std::byte> bytes_;
        Counters &counters_;
    };

    class FailingSource final : public xl::Source
    {
    public:
        explicit FailingSource(Counters &counters) : counters_(counters) {}
        ~FailingSource() override { counters_.destroyed.fetch_add(1); }

        std::uint64_t size() const override { return 1024; }
        std::expected<std::size_t, std::string> read_at(std::uint64_t, std::span<std::byte>) const override
        {
            return std::unexpected(std::string("disk on fire"));
        }

    private:
        Counters &counters_;
    };

    class ThrowingSource final : public xl::Source
    {
    public:
        explicit ThrowingSource(Counters &counters) : counters_(counters) {}
        ~ThrowingSource() override { counters_.destroyed.fetch_add(1); }

        std::uint64_t size() const override { return 1024; }
        std::expected<std::size_t, std::string> read_at(std::uint64_t, std::span<std::byte>) const override
        {
            throw std::runtime_error("boom");
        }

    private:
        Counters &counters_;
    };

    class ErrorStream final : public xl::InputStream
    {
    public:
        std::expected<std::size_t, std::string> read(std::span<std::byte>) override
        {
            return std::unexpected(std::string("pipe broke"));
        }
    };

    class ChunkStream final : public xl::InputStream
    {
    public:
        explicit ChunkStream(std::vector<std::byte> bytes) : bytes_(std::move(bytes)) {}

        std::expected<std::size_t, std::string> read(std::span<std::byte> buffer) override
        {
            const std::size_t count = std::min<std::size_t>({buffer.size(), std::size_t{7}, bytes_.size() - position_});
            std::memcpy(buffer.data(), bytes_.data() + position_, count);
            position_ += count;
            return count;
        }

    private:
        std::vector<std::byte> bytes_;
        std::size_t position_ = 0;
    };

    std::vector<std::byte> build_sheets(int sheet_count, int row_count)
    {
        auto writer = xl::WriterHandle::open_memory(XL_FORMAT_XLSX);
        std::vector<std::byte> bytes;
        if (!writer.has_value())
        {
            return bytes;
        }
        for (int s = 0; s < sheet_count; ++s)
        {
            (void)writer->start_sheet("sheet" + std::to_string(s));
            for (int r = 0; r < row_count; ++r)
            {
                (void)writer->start_row();
                (void)writer->write("s" + std::to_string(s) + "-r" + std::to_string(r));
                (void)writer->end_row();
            }
            (void)writer->end_sheet();
        }
        auto written = writer->bytes();
        if (written.has_value())
        {
            bytes.resize(written->size());
            std::memcpy(bytes.data(), written->data(), written->size());
        }
        return bytes;
    }

    std::vector<std::string> first_column(const xl::Workbook &workbook, int32_t sheet)
    {
        std::vector<std::string> values;
        auto cursor = workbook.sheet(sheet).rows();
        while (cursor.has_value())
        {
            auto row = cursor->next_row();
            if (!row.has_value())
            {
                break;
            }
            values.emplace_back((*row)[0].value);
        }
        return values;
    }

    std::size_t row_count(const xl::Workbook &workbook)
    {
        auto decoded = workbook.sheet(0).read_all_decoded();
        return decoded.has_value() ? decoded->size() : 0;
    }
}

int main()
{
    const std::vector<std::byte> bytes = read_file(EXCELREADER_XLSX_FIXTURE_PATH);
    auto reference = xl::Workbook::open(EXCELREADER_XLSX_FIXTURE_PATH);
    check(reference.has_value(), "open the fixture by path");
    const std::size_t expected_rows = reference.has_value() ? row_count(*reference) : 0;
    check(expected_rows > 1, "the fixture has rows");

    Counters counters;
    {
        auto workbook = xl::Workbook::open_source(std::make_unique<MemorySource>(bytes, counters));
        check(workbook.has_value(), "open_source opens the fixture");
        check(workbook.has_value() && row_count(*workbook) == expected_rows, "open_source reads every row");
    }
    check(counters.destroyed.load() == 1, "the source is destroyed exactly once");

    const int blocks = static_cast<int>((bytes.size() + (4u << 20) - 1) / (4u << 20));
    check(counters.reads.load() <= blocks, "the default cache fetches each block once");

    Counters uncached;
    {
        xl::OpenOptions options;
        options.source_block_size = -1;
        auto workbook = xl::Workbook::open_source(std::make_unique<MemorySource>(bytes, uncached), XL_FORMAT_XLSX, &options);
        check(workbook.has_value() && row_count(*workbook) == expected_rows, "open_source without the cache reads every row");
    }
    check(uncached.reads.load() > blocks, "without the cache every read reaches read_at");

    {
        Counters failing;
        auto workbook = xl::Workbook::open_source(std::make_unique<FailingSource>(failing), XL_FORMAT_XLSX);
        check(!workbook.has_value() && workbook.error().message.find("disk on fire") != std::string::npos,
              "the source's error message reaches the caller");
        check(failing.destroyed.load() == 1, "a source whose read failed is destroyed exactly once");
    }
    {
        Counters throwing;
        auto workbook = xl::Workbook::open_source(std::make_unique<ThrowingSource>(throwing), XL_FORMAT_XLSX);
        check(!workbook.has_value() && workbook.error().message.find("boom") != std::string::npos,
              "an exception from the source becomes an error");
        check(throwing.destroyed.load() == 1, "a source that threw is destroyed exactly once");
    }
    {
        auto workbook = xl::Workbook::open_stream(std::make_unique<ErrorStream>(), XL_FORMAT_XLSX);
        check(!workbook.has_value() && workbook.error().message.find("pipe broke") != std::string::npos,
              "the stream's error message reaches the caller");
    }
    {
        constexpr int sheet_count = 4;
        const std::vector<std::byte> many = build_sheets(sheet_count, 300);
        check(!many.empty(), "build the multi-sheet workbook");
        Counters parallel_counters;
        auto reference_many = xl::Workbook::open_memory(std::span<const uint8_t>(reinterpret_cast<const uint8_t *>(many.data()), many.size()));
        auto workbook = xl::Workbook::open_source(std::make_unique<MemorySource>(many, parallel_counters), XL_FORMAT_XLSX);
        check(reference_many.has_value() && workbook.has_value(), "open the multi-sheet workbook both ways");
        if (reference_many.has_value() && workbook.has_value())
        {
            std::vector<std::vector<std::string>> parallel(sheet_count * 2);
            {
                std::vector<std::thread> threads;
                for (std::size_t i = 0; i < parallel.size(); ++i)
                {
                    threads.emplace_back([&parallel, &workbook, i]
                                         { parallel[i] = first_column(*workbook, static_cast<int32_t>(i % sheet_count)); });
                }
                for (std::thread &thread : threads)
                {
                    thread.join();
                }
            }
            bool same = true;
            for (std::size_t i = 0; i < parallel.size(); ++i)
            {
                const auto expected = first_column(*reference_many, static_cast<int32_t>(i % sheet_count));
                same = same && expected.size() == 300 && parallel[i] == expected;
            }
            check(same, "every sheet of a source read on its own thread matches the in-memory read");
        }
    }
    {
        const std::string csv = "name,qty\nwidget,7\n";
        std::vector<std::byte> csv_bytes(csv.size());
        std::memcpy(csv_bytes.data(), csv.data(), csv.size());
        auto workbook = xl::Workbook::open_stream(std::make_unique<ChunkStream>(csv_bytes), XL_FORMAT_CSV);
        check(workbook.has_value() && row_count(*workbook) == 2, "open_stream reads a CSV");
    }
    {
        auto workbook = xl::Workbook::open_stream(std::make_unique<ChunkStream>(bytes));
        check(workbook.has_value() && row_count(*workbook) == expected_rows, "open_stream reads an XLSX");
    }

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return EXIT_FAILURE;
    }
    std::puts("sources: OK");
    return EXIT_SUCCESS;
}
