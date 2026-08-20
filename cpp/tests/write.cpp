#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <string_view>
#include <array>
#include <filesystem>
#include <string>
#include <vector>
#include <chrono>

#define CHECK(cond, msg)                                                         \
    do                                                                           \
    {                                                                            \
        if (!(cond))                                                             \
        {                                                                        \
            std::fprintf(stderr, "FAIL: %s (%s:%d)\n", msg, __FILE__, __LINE__); \
            return 1;                                                            \
        }                                                                        \
    } while (0)

struct WrittenRow
{
    std::string_view texto;
    int64_t inteiro;
    double numero;
    std::chrono::year_month_day data;
    std::chrono::microseconds hora;
    std::chrono::system_clock::time_point instante;
};

template <>
struct xl::ExcelMapper<WrittenRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("texto", &WrittenRow::texto),
            xl::make_field("inteiro", &WrittenRow::inteiro),
            xl::make_field("numero", &WrittenRow::numero),
            xl::make_field("data", &WrittenRow::data),
            xl::make_field("hora", &WrittenRow::hora),
            xl::make_field("instante", &WrittenRow::instante));
    }
};

// A path under the system temp directory. Deleted by each test that creates it.
static std::filesystem::path temp_path(std::string_view name)
{
    return std::filesystem::temp_directory_path() /
           std::filesystem::path(std::string("excelreader-cpp-") + std::string(name));
}

static int test_write_columns_round_trip()
{
    const std::vector<int32_t> offsets{0, 3, 6};
    const std::vector<uint8_t> blob{'u', 'm', 'a', 'd', 'o', 'i'};
    const std::vector<int64_t> inteiros{1, 2};
    const std::vector<double> numeros{0.5, 1.5};
    const std::vector<int32_t> datas{20454, 20455}; // 2026-01-01, 2026-01-02
    const std::vector<int64_t> horas{3600000000, 7200000000};
    const std::vector<int64_t> instantes{1767225600000000, 1767312000000000};

    const std::array<xl::ColumnRef, 6> columns{
        xl::string_column("texto", offsets, blob),
        xl::i64_column("inteiro", inteiros),
        xl::f64_column("numero", numeros),
        xl::date_column("data", datas),
        xl::time_column("hora", horas),
        xl::timestamp_column("instante", instantes)};

    const std::filesystem::path path = temp_path("columns.xlsx");
    auto written = xl::write_columns(path.string(), XL_FORMAT_XLSX, columns);
    CHECK(written.has_value(), "write_columns must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse_sheet<WrittenRow>(*workbook);
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 2, "two rows were written");
    
        WrittenRow first = *table->begin();
        CHECK(first.texto == "uma", "row 0's string must round-trip");
        CHECK(first.inteiro == 1, "row 0's int64 must round-trip");
        CHECK(first.numero == 0.5, "row 0's double must round-trip");
        CHECK(first.data == std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1},
              "row 0's date must round-trip");
        CHECK(first.hora == std::chrono::microseconds{3600000000}, "row 0's time must round-trip");
        auto second_row = table->at(1);
        CHECK(second_row.has_value(), "row 1 must be in bounds");
        CHECK(second_row->texto == "doi", "row 1's string must round-trip");
        CHECK(second_row->inteiro == 2, "row 1's int64 must round-trip");
    }
    std::filesystem::remove(path);
    return 0;
}

static int test_write_columns_rejects_bad_input()
{
    const std::vector<int64_t> two{1, 2};
    const std::vector<int64_t> three{1, 2, 3};
    const std::vector<uint8_t> empty_bitmap{};
    const std::filesystem::path path = temp_path("rejected.xlsx");

    std::array<xl::ColumnRef, 2> mismatched{xl::i64_column("a", two), xl::i64_column("b", three)};
    CHECK(!xl::write_columns(path.string(), XL_FORMAT_XLSX, mismatched).has_value(),
          "columns of different lengths must be rejected");

    std::array<xl::ColumnRef, 2> partial_header{xl::i64_column("a", two), xl::i64_column("", two)};
    CHECK(!xl::write_columns(path.string(), XL_FORMAT_XLSX, partial_header).has_value(),
          "a partial header row must be rejected");

    // Two rows need one byte of bitmap. Hand it a non-null pointer with zero length: the ABI takes
    // the bitmap without a length, so this is exactly the overrun the wrapper exists to refuse.
    std::array<xl::ColumnRef, 1> short_bitmap{xl::i64_column("a", two)};
    short_bitmap[0].validity = reinterpret_cast<const uint8_t *>(two.data());
    short_bitmap[0].validity_len = 0;
    CHECK(!xl::write_columns(path.string(), XL_FORMAT_XLSX, short_bitmap).has_value(),
          "a validity bitmap shorter than the row count must be rejected");

    std::array<xl::ColumnRef, 1> fine{xl::i64_column("a", two)};
    CHECK(!xl::write_columns(path.string(), XL_FORMAT_AUTO, fine).has_value(),
          "XL_FORMAT_AUTO must be rejected: a new file has no signature bytes to sniff");

    std::span<const xl::ColumnRef> none{};
    CHECK(!xl::write_columns(path.string(), XL_FORMAT_XLSX, none).has_value(),
          "an empty column set must be rejected");

    std::filesystem::remove(path);
    return 0;
}

static int test_write_options()
{
    xl::WriteOptions defaults{};
    xl_write_options raw = defaults.to_c();
    CHECK(raw.struct_size == static_cast<int32_t>(sizeof(xl_write_options)), "to_c must fill struct_size");
    CHECK(raw.sheet_name == nullptr, "an empty sheet_name must lower to NULL, meaning Sheet1");
    CHECK(raw.sheet_name_len == 0, "an empty sheet_name must lower to length 0");
    CHECK(raw.csv_delimiter == 0, "an unset csv_delimiter must lower to 0");
    CHECK(raw.date1904 == XL_OPT_DEFAULT, "an unset tri-state must lower to XL_OPT_DEFAULT");

    xl::WriteOptions configured{};
    configured.sheet_name = "Dados";
    configured.csv_delimiter = ';';
    configured.use_shared_strings = XL_OPT_TRUE;
    xl_write_options set = configured.to_c();
    CHECK(set.sheet_name_len == 5, "sheet_name_len must be the byte length");
    CHECK(set.sheet_name != nullptr, "a non-empty sheet_name must lower to a pointer");
    CHECK(set.csv_delimiter == ';', "csv_delimiter must pass through unchanged");
    CHECK(set.use_shared_strings == XL_OPT_TRUE, "use_shared_strings must pass through unchanged");
    return 0;
}

static int test_format_from_path()
{
    CHECK(xl::format_from_path("out.xlsx") == XL_FORMAT_XLSX, ".xlsx must resolve to XL_FORMAT_XLSX");
    CHECK(xl::format_from_path("out.XLSB") == XL_FORMAT_XLSB, "the extension match must be case-insensitive");
    CHECK(xl::format_from_path("out.xls") == XL_FORMAT_XLS, ".xls must resolve to XL_FORMAT_XLS");
    CHECK(xl::format_from_path("out.csv") == XL_FORMAT_CSV, ".csv must resolve to XL_FORMAT_CSV");
    CHECK(xl::format_from_path("out.txt") == XL_FORMAT_AUTO, "an unknown extension must resolve to AUTO");
    CHECK(xl::format_from_path("out") == XL_FORMAT_AUTO, "no extension at all must resolve to AUTO");
    // A dot in a directory name is not an extension.
    CHECK(xl::format_from_path("v1.2/report") == XL_FORMAT_AUTO, "a dot before the last separator is not an extension");
    return 0;
}

int main()
{
    if (test_write_options() != 0)
    {
        return 1;
    }
    if (test_format_from_path() != 0)
    {
        return 1;
    }
    if (test_write_columns_round_trip() != 0)
    {
        return 1;
    }
    if (test_write_columns_rejects_bad_input() != 0)
    {
        return 1;
    }
    std::printf("OK\n");
    return 0;
}