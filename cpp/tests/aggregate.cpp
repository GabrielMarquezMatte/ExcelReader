#include <xl/excelreader.hpp>

#include <atomic>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <limits>
#include <memory>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>

#define CHECK(cond, msg)                                                         \
    do                                                                           \
    {                                                                            \
        if (!(cond))                                                             \
        {                                                                        \
            std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); \
            return 1;                                                            \
        }                                                                        \
    } while (0)

namespace
{
    constexpr std::string_view kCsv = "a,b\n1,2\n3,4\n5,6\n";

    std::atomic<int> g_constructed{0};
    std::atomic<int> g_destructed{0};

    struct RowCounter
    {
        int64_t rows = 0;
        int64_t combines = 0;

        RowCounter() { g_constructed.fetch_add(1, std::memory_order_relaxed); }
        RowCounter(const RowCounter &other) : rows(other.rows), combines(other.combines)
        {
            g_constructed.fetch_add(1, std::memory_order_relaxed);
        }
        RowCounter(RowCounter &&other) noexcept : rows(other.rows), combines(other.combines)
        {
            g_constructed.fetch_add(1, std::memory_order_relaxed);
        }
        ~RowCounter() { g_destructed.fetch_add(1, std::memory_order_relaxed); }

        int32_t accumulate(xl::RowView row) noexcept
        {
            if (!row.empty())
            {
                ++rows;
            }
            return XL_OK;
        }

        int32_t combine(RowCounter &next) noexcept
        {
            rows += next.rows;
            ++combines;
            return XL_OK;
        }
    };

    static_assert(xl::CsvAccumulator<RowCounter>, "RowCounter must satisfy xl::CsvAccumulator");

    struct MoveOnly
    {
        std::unique_ptr<int> state;
        int32_t accumulate(xl::RowView) noexcept { return XL_OK; }
        int32_t combine(MoveOnly &) noexcept { return XL_OK; }
    };

    template <typename Acc>
    concept Aggregatable = requires(std::span<const uint8_t> data) { xl::aggregate_csv<Acc>(data); };

    static_assert(Aggregatable<RowCounter>, "a copyable accumulator must be accepted");
    static_assert(!Aggregatable<MoveOnly>,
                  "every partition copies the seed, so a move-only accumulator must fail the constraint cleanly");

    constexpr int64_t kLargeRowCount = 800'000;

    void write_repeated_rows(const std::filesystem::path &path, int64_t count)
    {
        std::string content;
        content.reserve(static_cast<size_t>(count) * 4);
        for (int64_t i = 0; i < count; ++i)
        {
            content += "1,2\n";
        }
        std::ofstream out(path, std::ios::binary);
        out.write(content.data(), static_cast<std::streamsize>(content.size()));
    }
}

static int test_aggregate_csv_memory()
{
    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv(data, RowCounter{}, {}, 1, 1);
    CHECK(result.has_value(), "aggregate_csv must succeed on a well-formed in-memory CSV");
    CHECK(result->rows == 3, "the fixture has 3 data rows after the header");
    return 0;
}

static int test_aggregate_csv_file()
{
    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate.csv";
    {
        std::ofstream out(path, std::ios::binary);
        out.write(kCsv.data(), static_cast<std::streamsize>(kCsv.size()));
    }

    auto result = xl::aggregate_csv(path.string(), RowCounter{}, {}, 1, 1);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "aggregate_csv must succeed on a well-formed CSV file");
    CHECK(result->rows == 3, "the fixture has 3 data rows after the header");
    return 0;
}

static int test_aggregate_default_options()
{
    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv<RowCounter>(data);
    CHECK(result.has_value(), "aggregate_csv must succeed with default options");
    CHECK(result->rows == 4, "with no header configured, all 4 rows are data rows");
    return 0;
}

static int test_aggregate_csv_file_partitioned()
{
    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-large.csv";
    write_repeated_rows(path, kLargeRowCount);

    auto result = xl::aggregate_csv(path.string(), RowCounter{}, {}, 0, 8);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "aggregate_csv must succeed on a large multi-partition CSV file");
    CHECK(result->rows == kLargeRowCount,
          "every one of the generated rows must be counted exactly once");
    CHECK(result->combines >= 1,
          "a source this large at degree_of_parallelism=8 must be split into more than one "
          "partition, so combine must run at least once");
    return 0;
}

static int test_aggregate_quoted_records_straddling_chunks()
{
    std::string content;
    for (int record = 0; record < 8; ++record)
    {
        content += "1,\"";
        for (int filler = 0; filler < 40000; ++filler)
        {
            content += "BOOM,notint,x\n";
        }
        content += "\"\n";
    }

    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-quoted.csv";
    {
        std::ofstream out(path, std::ios::binary);
        out.write(content.data(), static_cast<std::streamsize>(content.size()));
    }

    auto result = xl::aggregate_csv(path.string(), RowCounter{}, {}, 0, 8);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "a file of quoted multi-line records must aggregate");
    CHECK(result->rows == 8, "each quoted record counts exactly once, however the chunks fell");
    CHECK(result->combines >= 1,
          "this fixture must really be split, or the quoted-straddle case is never exercised");
    return 0;
}

static int test_aggregate_reports_a_throwing_accumulate()
{
    struct Throws
    {
        int32_t accumulate(xl::RowView) { throw std::runtime_error("boom"); }
        int32_t combine(Throws &) noexcept { return XL_OK; }
    };

    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv<Throws>(data);
    CHECK(!result.has_value(), "a throwing accumulate must produce an error, not a value");
    CHECK(std::string_view(result.error().message).find("threw") != std::string_view::npos,
          "the error must say a callback threw");
    return 0;
}

static int test_aggregate_with_a_seed()
{
    struct Summer
    {
        int64_t total = 0;
        int64_t step = 0;

        int32_t accumulate(xl::RowView row) noexcept
        {
            if (!row.empty())
            {
                total += step;
            }
            return XL_OK;
        }

        int32_t combine(Summer &next) noexcept
        {
            total += next.total;
            return XL_OK;
        }
    };

    const int64_t step = 10;
    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv(data, Summer{0, step}, {}, 1);
    CHECK(result.has_value(), "aggregate_csv with a seed must succeed");
    CHECK(result->total == 30, "3 data rows at step 10 must total 30");

    auto no_options = xl::aggregate_csv(data, Summer{0, step});
    CHECK(no_options.has_value(), "a seed with default options must succeed");
    CHECK(no_options->total == 40, "with no header configured, all 4 rows are data rows");

    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-seed.csv";
    {
        std::ofstream out(path, std::ios::binary);
        out.write(kCsv.data(), static_cast<std::streamsize>(kCsv.size()));
    }
    auto from_file = xl::aggregate_csv(path.string(), Summer{0, step});
    std::filesystem::remove(path);
    CHECK(from_file.has_value(), "aggregate_csv of a file with a seed must succeed");
    CHECK(from_file->total == 40, "with no header configured, all 4 rows are data rows");
    return 0;
}

static int test_aggregate_reports_a_callback_status_of_int32_max()
{
    struct ReturnsIntMax
    {
        int32_t accumulate(xl::RowView) noexcept { return (std::numeric_limits<int32_t>::max)(); }
        int32_t combine(ReturnsIntMax &) noexcept { return XL_OK; }
    };

    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv<ReturnsIntMax>(data, {}, {}, 0, 1);
    CHECK(!result.has_value(), "a nonzero callback status must abort the run");
    CHECK(result.error().code == (std::numeric_limits<int32_t>::max)(),
          "the caller's own status must come back verbatim");

    const std::string_view message(result.error().message);
    CHECK(message.find("threw") == std::string_view::npos,
          "INT32_MAX from a callback that did not throw must NOT be reported as a thrown exception");
    CHECK(message.find("2147483647") != std::string_view::npos,
          "the error must name the status the callback returned");
    return 0;
}

static int test_aggregate_destroys_every_state()
{
    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-balance.csv";
    write_repeated_rows(path, kLargeRowCount);

    g_constructed.store(0);
    g_destructed.store(0);
    {
        auto result = xl::aggregate_csv(path.string(), RowCounter{}, {}, 0, 8);
        CHECK(result.has_value(), "the run must succeed");
        CHECK(result->combines >= 1,
              "the fixture must really be split, or free_state never runs and the balance is "
              "vacuous");
    }
    std::filesystem::remove(path);

    CHECK(g_constructed.load() > 0, "the run must have constructed at least one accumulator");
    CHECK(g_constructed.load() == g_destructed.load(),
          "every constructed accumulator, including the one returned, must be destroyed");
    return 0;
}

static int test_aggregate_takes_the_dialect_from_open_options()
{
    struct CellCounter
    {
        int64_t cells = 0;

        int32_t accumulate(xl::RowView row) noexcept
        {
            cells += static_cast<int64_t>(row.size());
            return XL_OK;
        }

        int32_t combine(CellCounter &next) noexcept
        {
            cells += next.cells;
            return XL_OK;
        }
    };

    constexpr std::string_view semicolons = "a;b\n1;2\n3;4\n";
    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(semicolons.data()), semicolons.size());
    auto split = xl::aggregate_csv(data, CellCounter{}, {.csv_delimiter = ';'}, 1, 1);
    CHECK(split.has_value() && split->cells == 4, "csv_delimiter must split each data row into two cells");

    auto unsplit = xl::aggregate_csv(data, CellCounter{}, {}, 1, 1);
    CHECK(unsplit.has_value() && unsplit->cells == 2, "the default delimiter leaves each row one cell");
    return 0;
}

int main()
{
    if (int failed = test_aggregate_csv_memory())
    {
        return failed;
    }
    if (int failed = test_aggregate_csv_file())
    {
        return failed;
    }
    if (int failed = test_aggregate_default_options())
    {
        return failed;
    }
    if (int failed = test_aggregate_csv_file_partitioned())
    {
        return failed;
    }
    if (int failed = test_aggregate_quoted_records_straddling_chunks())
    {
        return failed;
    }
    if (int failed = test_aggregate_reports_a_throwing_accumulate())
    {
        return failed;
    }
    if (int failed = test_aggregate_reports_a_callback_status_of_int32_max())
    {
        return failed;
    }
    if (int failed = test_aggregate_with_a_seed())
    {
        return failed;
    }
    if (int failed = test_aggregate_takes_the_dialect_from_open_options())
    {
        return failed;
    }
    if (int failed = test_aggregate_destroys_every_state())
    {
        return failed;
    }

    std::printf("OK: C++ aggregate test passed\n");
    return 0;
}
