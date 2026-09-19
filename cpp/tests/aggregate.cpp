#include <xl/excelreader.hpp>

#include <atomic>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <limits>
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

    template <typename Seed>
    concept SeedAccepted = requires(std::span<const uint8_t> data, Seed &seed) {
        xl::aggregate_csv_memory<RowCounter>(data, seed);
    };

    static_assert(SeedAccepted<decltype([n = 0] { return RowCounter{}; })>,
                  "a const-callable seed must be accepted");
    static_assert(!SeedAccepted<decltype([n = 0]() mutable { return RowCounter{}; })>,
                  "a mutable seed must fail the constraint cleanly, not hard-error inside a shim");

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
    xl::CsvParallelOptions options{};
    options.header_row = 1;
    options.degree_of_parallelism = 1;

    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv_memory<RowCounter>(data, &options);
    CHECK(result.has_value(), "aggregate_csv_memory must succeed on a well-formed in-memory CSV");
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

    xl::CsvParallelOptions options{};
    options.header_row = 1;
    options.degree_of_parallelism = 1;

    auto result = xl::aggregate_csv_file<RowCounter>(path.string(), &options);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "aggregate_csv_file must succeed on a well-formed CSV file");
    CHECK(result->rows == 3, "the fixture has 3 data rows after the header");
    return 0;
}

static int test_aggregate_default_options()
{
    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv_memory<RowCounter>(data);
    CHECK(result.has_value(), "aggregate_csv_memory must succeed with default (null) options");
    CHECK(result->rows == 4, "with no header configured, all 4 rows are data rows");
    return 0;
}

static int test_aggregate_csv_file_partitioned()
{
    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-large.csv";
    write_repeated_rows(path, kLargeRowCount);

    xl::CsvParallelOptions options{};
    options.degree_of_parallelism = 8;

    auto result = xl::aggregate_csv_file<RowCounter>(path.string(), &options);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "aggregate_csv_file must succeed on a large multi-partition CSV file");
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

    xl::CsvParallelOptions options{};
    options.degree_of_parallelism = 8;
    auto result = xl::aggregate_csv_file<RowCounter>(path.string(), &options);
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
    auto result = xl::aggregate_csv_memory<Throws>(data);
    CHECK(!result.has_value(), "a throwing accumulate must produce an error, not a value");
    CHECK(std::string_view(result.error().message).find("threw") != std::string_view::npos,
          "the error must say a callback threw");
    return 0;
}

static int test_aggregate_with_a_seed_callable()
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
    xl::CsvParallelOptions options{};
    options.header_row = 1;
    auto result = xl::aggregate_csv_memory<Summer>(data, [step] { return Summer{0, step}; }, &options);
    CHECK(result.has_value(), "aggregate_csv_memory with a seed callable must succeed");
    CHECK(result->total == 30, "3 data rows at step 10 must total 30");

    auto no_options = xl::aggregate_csv_memory<Summer>(data, [step] { return Summer{0, step}; });
    CHECK(no_options.has_value(), "a seed callable with default (null) options must succeed");
    CHECK(no_options->total == 40, "with no header configured, all 4 rows are data rows");

    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-seed.csv";
    {
        std::ofstream out(path, std::ios::binary);
        out.write(kCsv.data(), static_cast<std::streamsize>(kCsv.size()));
    }
    auto from_file = xl::aggregate_csv_file<Summer>(path.string(), [step] { return Summer{0, step}; });
    std::filesystem::remove(path);
    CHECK(from_file.has_value(), "aggregate_csv_file with a seed callable must succeed");
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

    xl::CsvParallelOptions options{};
    options.degree_of_parallelism = 1;

    std::span<const uint8_t> data(reinterpret_cast<const uint8_t *>(kCsv.data()), kCsv.size());
    auto result = xl::aggregate_csv_memory<ReturnsIntMax>(data, &options);
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
        xl::CsvParallelOptions options{};
        options.degree_of_parallelism = 8;

        auto result = xl::aggregate_csv_file<RowCounter>(path.string(), &options);
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
    if (int failed = test_aggregate_with_a_seed_callable())
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
