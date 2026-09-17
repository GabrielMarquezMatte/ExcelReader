#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
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

    struct RowCounter
    {
        int64_t rows = 0;
        int64_t combines = 0;

        int32_t accumulate(const xl_row *) noexcept
        {
            ++rows;
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
    constexpr int64_t kRowCount = 800'000; 
    std::string content;
    content.reserve(static_cast<size_t>(kRowCount) * 4);
    for (int64_t i = 0; i < kRowCount; ++i)
    {
        content += "1,2\n";
    }

    const std::filesystem::path path =
        std::filesystem::temp_directory_path() / "excelreader-cpp-aggregate-large.csv";
    {
        std::ofstream out(path, std::ios::binary);
        out.write(content.data(), static_cast<std::streamsize>(content.size()));
    }

    xl::CsvParallelOptions options{};
    options.degree_of_parallelism = 8;

    auto result = xl::aggregate_csv_file<RowCounter>(path.string(), &options);
    std::filesystem::remove(path);

    CHECK(result.has_value(), "aggregate_csv_file must succeed on a large multi-partition CSV file");
    CHECK(result->rows == kRowCount, "every one of the generated rows must be counted exactly once");
    CHECK(result->combines >= 1,
          "a source this large at degree_of_parallelism=8 must be split into more than one "
          "partition, so combine must run at least once");
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

    std::printf("OK: C++ aggregate test passed\n");
    return 0;
}
