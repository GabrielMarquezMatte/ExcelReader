#include <xl/excelreader.hpp>

#include <cstdint>
#include <cstdio>
#include <string_view>
#include <array>
#include <filesystem>
#include <string>
#include <vector>
#include <chrono>
#include <utility>

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

struct NarrowRow
{
    uint8_t small;
};

template <>
struct xl::ExcelMapper<NarrowRow>
{
    static constexpr auto get_bindings() { return std::make_tuple(xl::make_field("small", &NarrowRow::small)); }
};

struct WideRow
{
    uint64_t big;
};

template <>
struct xl::ExcelMapper<WideRow>
{
    static constexpr auto get_bindings() { return std::make_tuple(xl::make_field("big", &WideRow::big)); }
};

struct NamedRow
{
    std::string name = "fallback";
    std::string_view view = "fallback";
};

template <>
struct xl::ExcelMapper<NamedRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(xl::make_field("name", &NamedRow::name), xl::make_field("view", &NamedRow::view));
    }
};

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
    const std::vector<int32_t> datas{20454, 20455}; 
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
    auto written = xl::write_columns(path.string(), columns);
    CHECK(written.has_value(), "write_columns must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse<WrittenRow>(workbook->sheet(0));
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

static int test_encrypt_package_round_trip()
{
    const std::vector<int32_t> offsets{0, 3, 6};
    const std::vector<uint8_t> blob{'u', 'm', 'a', 'd', 'o', 'i'};
    const std::vector<int64_t> inteiros{1, 2};
    const std::vector<double> numeros{0.5, 1.5};
    const std::vector<int32_t> datas{20454, 20455};
    const std::vector<int64_t> horas{3600000000, 7200000000};
    const std::vector<int64_t> instantes{1767225600000000, 1767312000000000};

    const std::array<xl::ColumnRef, 6> columns{
        xl::string_column("texto", offsets, blob),
        xl::i64_column("inteiro", inteiros),
        xl::f64_column("numero", numeros),
        xl::date_column("data", datas),
        xl::time_column("hora", horas),
        xl::timestamp_column("instante", instantes)};

    const std::filesystem::path plain_path = temp_path("encrypt-plain.xlsx");
    const std::filesystem::path encrypted_path = temp_path("encrypt-cipher.xlsx");
    auto written = xl::write_columns(plain_path.string(), columns);
    CHECK(written.has_value(), "write_columns must succeed");

    auto encrypted = xl::encrypt_package(plain_path.string(), encrypted_path.string(), "hunter2");
    CHECK(encrypted.has_value(), "encrypt_package must succeed");

    {
        auto workbook = xl::Workbook::open(encrypted_path.string(), {.password = "hunter2"});
        CHECK(workbook.has_value(), "the encrypted file must open with the right password");
        auto table = xl::parse<WrittenRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the decrypted file must parse back");
        CHECK(table->size() == 2, "two rows were written");
        WrittenRow first = *table->begin();
        CHECK(first.texto == "uma", "the round-tripped row must match what was written");
        CHECK(first.inteiro == 1, "row 0's int64 must round-trip");
    }

    auto wrong_password = xl::Workbook::open(encrypted_path.string());
    CHECK(!wrong_password.has_value(), "opening without a password must fail");
    CHECK(wrong_password.error().code == XL_STATUS_PASSWORD_REQUIRED,
         "opening without a password must report XL_STATUS_PASSWORD_REQUIRED");

    std::filesystem::remove(plain_path);
    std::filesystem::remove(encrypted_path);
    return 0;
}

static int test_encrypt_package_rejects_an_empty_password()
{
    const std::filesystem::path plain_path = temp_path("encrypt-empty-pw.xlsx");
    const std::array<xl::ColumnRef, 1> columns{xl::i64_column("inteiro", std::vector<int64_t>{1})};
    auto written = xl::write_columns(plain_path.string(), columns);
    CHECK(written.has_value(), "write_columns must succeed");

    auto encrypted = xl::encrypt_package(plain_path.string(), temp_path("encrypt-empty-pw-out.xlsx").string(), "");
    CHECK(!encrypted.has_value(), "encrypt_package must reject an empty password");

    std::filesystem::remove(plain_path);
    return 0;
}

static int test_write_columns_to_memory_round_trip()
{
    const std::vector<int32_t> offsets{0, 3, 6};
    const std::vector<uint8_t> blob{'u', 'm', 'a', 'd', 'o', 'i'};
    const std::vector<int64_t> inteiros{1, 2};
    const std::vector<double> numeros{0.5, 1.5};
    const std::vector<int32_t> datas{20454, 20455}; 
    const std::vector<int64_t> horas{3600000000, 7200000000};
    const std::vector<int64_t> instantes{1767225600000000, 1767312000000000};

    const std::array<xl::ColumnRef, 6> columns{
        xl::string_column("texto", offsets, blob),
        xl::i64_column("inteiro", inteiros),
        xl::f64_column("numero", numeros),
        xl::date_column("data", datas),
        xl::time_column("hora", horas),
        xl::timestamp_column("instante", instantes)};

    std::vector<uint8_t> bytes;
    CHECK(xl::write_columns(bytes, columns, {.format = XL_FORMAT_XLSX}).has_value(),
          "write_columns into a vector must succeed");
    CHECK(!bytes.empty(), "write_columns must fill the vector");

    auto workbook = xl::Workbook::open(bytes, {.format = XL_FORMAT_XLSX});
    CHECK(workbook.has_value(), "the returned bytes must open");
    auto table = xl::parse<WrittenRow>(workbook->sheet(0));
    CHECK(table.has_value(), "the returned bytes must parse back");
    CHECK(table->size() == 2, "two rows were written");

    WrittenRow first = *table->begin();
    CHECK(first.texto == "uma", "row 0's string must round-trip");
    CHECK(first.inteiro == 1, "row 0's int64 must round-trip");
    return 0;
}

static int test_write_columns_rejects_bad_input()
{
    const std::vector<int64_t> two{1, 2};
    const std::vector<int64_t> three{1, 2, 3};
    const std::vector<uint8_t> empty_bitmap{};
    const std::filesystem::path path = temp_path("rejected.xlsx");

    std::array<xl::ColumnRef, 2> mismatched{xl::i64_column("a", two), xl::i64_column("b", three)};
    CHECK(!xl::write_columns(path.string(), mismatched).has_value(),
          "columns of different lengths must be rejected");

    std::array<xl::ColumnRef, 2> partial_header{xl::i64_column("a", two), xl::i64_column("", two)};
    CHECK(!xl::write_columns(path.string(), partial_header).has_value(),
          "a partial header row must be rejected");

    std::array<xl::ColumnRef, 1> short_bitmap{xl::i64_column("a", two)};
    short_bitmap[0].validity = reinterpret_cast<const uint8_t *>(two.data());
    short_bitmap[0].validity_len = 0;
    CHECK(!xl::write_columns(path.string(), short_bitmap).has_value(),
          "a validity bitmap shorter than the row count must be rejected");

    std::array<xl::ColumnRef, 1> fine{xl::i64_column("a", two)};
    CHECK(!xl::write_columns(temp_path("rejected.txt").string(), fine).has_value(),
          "XL_FORMAT_AUTO with an unknown extension must be rejected");
    std::vector<uint8_t> untouched{42};
    CHECK(!xl::write_columns(untouched, fine).has_value(),
          "XL_FORMAT_AUTO must be rejected for memory: there is no extension to infer it from");
    CHECK(untouched.size() == 1 && untouched[0] == 42, "a failed write leaves the vector untouched");

    std::span<const xl::ColumnRef> none{};
    CHECK(!xl::write_columns(path.string(), none).has_value(),
          "an empty column set must be rejected");

    std::filesystem::remove(path);
    return 0;
}

static int test_write_columns_in_parallel()
{
    CHECK(xl::WriteOptions{}.degree_of_parallelism == 1, "the wrapper must default to a serial write");

    const std::vector<int32_t> offsets{0, 1, 2, 3, 4, 5, 6, 7};
    const std::vector<uint8_t> blob{'a', 'b', 'c', 'd', 'e', 'f', 'g'};
    const std::vector<int64_t> values{0, 10, 20, 30, 40, 50, 60};
    const std::array<xl::ColumnRef, 2> columns{xl::string_column("texto", offsets, blob), xl::i64_column("inteiro", values)};

    for (const int32_t format : {XL_FORMAT_XLSX, XL_FORMAT_XLSB})
    {
        std::vector<uint8_t> bytes;
        CHECK(xl::write_columns(bytes, columns, {.format = format, .degree_of_parallelism = 4, .batch_rows = 2}).has_value(),
              "a parallel write must succeed");

        auto workbook = xl::Workbook::open(bytes, {.format = format});
        CHECK(workbook.has_value(), "the parallel output must open");
        auto cursor = workbook->sheet(0).rows();
        CHECK(cursor.has_value(), "the parallel output must have a row cursor");
        int64_t index = 0;
        for (auto row = cursor->next_row(); row.has_value(); row = cursor->next_row(), ++index)
        {
            std::vector<std::string> cells;
            for (auto cell : *row)
            {
                cells.emplace_back(cell.value);
            }
            CHECK(cells.size() == 2, "every row must have both cells");
            if (index > 0)
            {
                CHECK(cells[0] == std::string(1, static_cast<char>('a' + index - 1)), "the string column must round-trip in order");
                CHECK(cells[1] == std::to_string((index - 1) * 10), "the int64 column must round-trip in order");
            }
        }
        CHECK(index == 8, "the header and seven data rows must be written");
    }
    return 0;
}

static int test_write_options()
{
    const std::vector<int64_t> values{1, 2};
    const std::array<xl::ColumnRef, 2> columns{xl::i64_column("a", values), xl::i64_column("b", values)};

    std::vector<uint8_t> bytes;
    CHECK(xl::write_columns(bytes, columns, {.format = XL_FORMAT_CSV, .csv_delimiter = ';'}).has_value(),
          "writing a CSV into a vector must succeed");
    const std::string_view text(reinterpret_cast<const char *>(bytes.data()), bytes.size());
    CHECK(text.find("a;b") != std::string_view::npos, "WriteOptions::csv_delimiter must reach the output");
    return 0;
}

static int test_format_from_extension()
{
    const std::vector<int64_t> values{1, 2};
    const std::array<xl::ColumnRef, 1> columns{xl::i64_column("a", values)};
    const std::array<std::pair<std::string_view, int32_t>, 4> cases{{{"ext.xlsx", XL_FORMAT_XLSX},
                                                                     {"ext.XLSB", XL_FORMAT_XLSB},
                                                                     {"ext.xls", XL_FORMAT_XLS},
                                                                     {"ext.csv", XL_FORMAT_CSV}}};
    for (const auto &[name, format] : cases)
    {
        const std::filesystem::path path = temp_path(name);
        CHECK(xl::write_columns(path.string(), columns).has_value(), "a known extension must pick the format");
        {
            auto workbook = xl::Workbook::open(path.string(), {.format = format});
            CHECK(workbook.has_value(), "the file must open as the format its extension names");
            auto rows = workbook->sheet(0).read_all();
            CHECK(rows.has_value() && rows->size() == 3, "the header and both rows must read back");
        }
        std::filesystem::remove(path);
    }
    CHECK(!xl::write_columns(temp_path("no-extension").string(), columns).has_value(),
          "a path without an extension must be rejected");
    CHECK(!xl::write_columns("ab", columns).has_value(), "a name shorter than any extension must be rejected");
    return 0;
}

struct FlagRow
{
    bool ativo;
};

template <>
struct xl::ExcelMapper<FlagRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(xl::make_field("ativo", &FlagRow::ativo));
    }
};

static int test_bool_round_trip()
{
    const std::vector<uint8_t> flags{1, 0, 1};
    const std::array<xl::ColumnRef, 1> columns{xl::bool_column("ativo", flags)};

    const std::filesystem::path path = temp_path("bools.xlsx");
    CHECK(xl::write_columns(path.string(), columns).has_value(),
          "writing a bool column must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse<FlagRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 3, "three rows were written");

        CHECK(table->at(0)->ativo, "row 0 was written true");
        CHECK(!table->at(1)->ativo, "row 1 was written false");
        CHECK(table->at(2)->ativo, "row 2 was written true");
    }
    std::filesystem::remove(path);
    return 0;
}

struct NullableRow
{
    std::optional<int64_t> quantidade;
};

template <>
struct xl::ExcelMapper<NullableRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(xl::make_field("quantidade", &NullableRow::quantidade));
    }
};

static int test_optional_round_trip()
{
    const std::vector<int64_t> valores{10, 0, 30};
    const std::vector<uint8_t> validity{0b00000101};
    const std::array<xl::ColumnRef, 1> columns{xl::i64_column("quantidade", valores, validity)};

    const std::filesystem::path path = temp_path("nullable.xlsx");
    CHECK(xl::write_columns(path.string(), columns).has_value(),
          "writing a nullable column must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse<NullableRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 3, "three rows were written");

        CHECK(table->at(0)->quantidade == 10, "row 0 must round-trip its value");
        CHECK(!table->at(1)->quantidade.has_value(), "row 1 was written null and must come back empty");
        CHECK(table->at(2)->quantidade == 30, "row 2 must round-trip its value");
    }
    std::filesystem::remove(path);
    return 0;
}

struct FullRow
{
    std::string texto;
    int64_t inteiro;
    double numero;
    bool ativo;
    std::chrono::year_month_day data;
    std::chrono::microseconds hora;
    std::chrono::system_clock::time_point instante;
    std::optional<int64_t> opcional;
};

template <>
struct xl::ExcelMapper<FullRow>
{
    static constexpr auto get_bindings()
    {
        return std::make_tuple(
            xl::make_field("texto", &FullRow::texto),
            xl::make_field("inteiro", &FullRow::inteiro),
            xl::make_field("numero", &FullRow::numero),
            xl::make_field("ativo", &FullRow::ativo),
            xl::make_field("data", &FullRow::data),
            xl::make_field("hora", &FullRow::hora),
            xl::make_field("instante", &FullRow::instante),
            xl::make_field("opcional", &FullRow::opcional));
    }
};

static int test_write_sheet_round_trip()
{
    const std::vector<FullRow> rows{
        FullRow{"uma", 1, 0.5, true,
                std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1},
                std::chrono::microseconds{3600000000},
                std::chrono::system_clock::time_point{std::chrono::microseconds{1767225600000000}},
                std::optional<int64_t>{7}},
        FullRow{"duas", 2, 1.5, false,
                std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{2},
                std::chrono::microseconds{7200000000},
                std::chrono::system_clock::time_point{std::chrono::microseconds{1767312000000000}},
                std::nullopt}};

    const std::filesystem::path path = temp_path("sheet.xlsx");
    CHECK(xl::write_sheet(path.string(), rows).has_value(), "write_sheet must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse<FullRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 2, "two rows were written");

        auto first = table->at(0);
        CHECK(first.has_value(), "row 0 must be in bounds");
        CHECK(first->texto == "uma", "row 0's string must round-trip");
        CHECK(first->inteiro == 1, "row 0's int64 must round-trip");
        CHECK(first->numero == 0.5, "row 0's double must round-trip");
        CHECK(first->ativo, "row 0's bool must round-trip");
        CHECK(first->data == std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1},
              "row 0's date must round-trip");
        CHECK(first->hora == std::chrono::microseconds{3600000000}, "row 0's time must round-trip");
        CHECK(first->opcional == 7, "row 0's optional must round-trip its value");

        auto second = table->at(1);
        CHECK(second.has_value(), "row 1 must be in bounds");
        CHECK(second->texto == "duas", "row 1's string must round-trip");
        CHECK(!second->ativo, "row 1's bool must round-trip as false");
        CHECK(!second->opcional.has_value(), "row 1's nullopt must come back empty");
    }
    std::filesystem::remove(path);
    return 0;
}

static int test_write_sheet_to_memory_round_trip()
{
    const std::vector<FullRow> rows{
        FullRow{"uma", 1, 0.5, true,
                std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1},
                std::chrono::microseconds{3600000000},
                std::chrono::system_clock::time_point{std::chrono::microseconds{1767225600000000}},
                std::optional<int64_t>{7}},
        FullRow{"duas", 2, 1.5, false,
                std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{2},
                std::chrono::microseconds{7200000000},
                std::chrono::system_clock::time_point{std::chrono::microseconds{1767312000000000}},
                std::nullopt}};

    std::vector<uint8_t> bytes;
    CHECK(xl::write_sheet(bytes, rows, {.format = XL_FORMAT_XLSX}).has_value(),
          "write_sheet into a vector must succeed");

    auto workbook = xl::Workbook::open(bytes);
    CHECK(workbook.has_value(), "the returned bytes must open");
    auto table = xl::parse<FullRow>(workbook->sheet(0));
    CHECK(table.has_value(), "the returned bytes must parse back");
    CHECK(table->size() == 2, "two rows were written");

    auto first = table->at(0);
    CHECK(first.has_value(), "row 0 must be in bounds");
    CHECK(first->texto == "uma", "row 0's string must round-trip");
    CHECK(first->opcional == 7, "row 0's optional must round-trip its value");
    auto second = table->at(1);
    CHECK(second.has_value(), "row 1 must be in bounds");
    CHECK(!second->opcional.has_value(), "row 1's nullopt must come back empty");
    return 0;
}

static int test_write_sheet_options_and_csv()
{
    const std::vector<FullRow> rows{};

    const std::filesystem::path path = temp_path("named.xlsx");
    CHECK(xl::write_sheet(path.string(), rows, {.format = XL_FORMAT_XLSX, .sheet_name = "Dados"}).has_value(),
          "writing an empty sheet with a custom name must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto name = workbook->sheet(0).name();
        CHECK(name.has_value(), "the sheet name must be readable");
        CHECK(*name == "Dados", "WriteOptions::sheet_name must reach the file");
    }
    std::filesystem::remove(path);

    const std::filesystem::path csv = temp_path("sheet.csv");
    CHECK(xl::write_sheet(csv.string(), rows).has_value(), "a .csv path must infer XL_FORMAT_CSV");
    std::filesystem::remove(csv);
    return 0;
}

static int test_writer_handle()
{
    static const auto write_str = [](xl_writer_handle *handle, std::string_view value)
    {
        return xl_write_string(handle, reinterpret_cast<const uint8_t *>(value.data()),
                               static_cast<int32_t>(value.size()));
    };

    const std::filesystem::path path = temp_path("writer_handle.xlsx");
    xl_writer_handle *handle = nullptr;
    const std::string c_path = path.string();
    int status = xl_open_write_handle(reinterpret_cast<const uint8_t *>(c_path.data()),
                                      static_cast<int32_t>(c_path.size()), XL_FORMAT_XLSX, nullptr, &handle);
    CHECK(status == XL_OK, "xl_open_write_handle must succeed");
    CHECK(handle != nullptr, "the returned handle must be non-null");

    status = xl_start_sheet(handle, reinterpret_cast<const uint8_t *>("Planilha1"), 9);
    CHECK(status == XL_OK, "xl_start_sheet must succeed");

    status = xl_start_row(handle);
    CHECK(status == XL_OK, "xl_start_row must succeed for the header row");
    for (std::string_view header : {"texto", "inteiro", "numero", "data", "hora", "instante"})
    {
        CHECK(write_str(handle, header) == XL_OK, "writing a header cell must succeed");
    }
    status = xl_end_row(handle);
    CHECK(status == XL_OK, "xl_end_row must succeed for the header row");

    status = xl_start_row(handle);
    CHECK(status == XL_OK, "xl_start_row must succeed for the data row");
    CHECK(write_str(handle, "uma") == XL_OK, "xl_write_string must succeed");
    CHECK(xl_write_int64(handle, 1) == XL_OK, "xl_write_int64 must succeed");
    CHECK(xl_write_float64(handle, 0.5) == XL_OK, "xl_write_float64 must succeed");
    CHECK(xl_write_date(handle, 20454) == XL_OK, "xl_write_date must succeed"); 
    CHECK(xl_write_time(handle, 3600000000) == XL_OK, "xl_write_time must succeed");
    CHECK(xl_write_timestamp(handle, 1767225600000000) == XL_OK, "xl_write_timestamp must succeed");
    status = xl_end_row(handle);
    CHECK(status == XL_OK, "xl_end_row must succeed for the data row");

    status = xl_end_sheet(handle);
    CHECK(status == XL_OK, "xl_end_sheet must succeed");
    status = xl_close_write_handle(handle);
    CHECK(status == XL_OK, "xl_close_write_handle must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the file xl_close_write_handle produced must open");
        auto name = workbook->sheet(0).name();
        CHECK(name.has_value() && *name == "Planilha1", "xl_start_sheet's name must reach the file");
        auto table = xl::parse<WrittenRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 1, "exactly one data row was written");
        WrittenRow first = *table->begin();
        CHECK(first.texto == "uma", "the streamed string cell must round-trip");
        CHECK(first.inteiro == 1, "the streamed int64 cell must round-trip");
        CHECK(first.numero == 0.5, "the streamed float64 cell must round-trip");
        CHECK(first.data == std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1},
              "the streamed date cell must round-trip");
        CHECK(first.hora == std::chrono::microseconds{3600000000}, "the streamed time cell must round-trip");
    }
    std::filesystem::remove(path);
    return 0;
}

static int test_writer_handle_rejects_bad_handle()
{
    xl_writer_handle *bogus = reinterpret_cast<xl_writer_handle *>(static_cast<std::uintptr_t>(0x1));
    CHECK(xl_start_sheet(bogus, reinterpret_cast<const uint8_t *>("x"), 1) == XL_INVALID_HANDLE,
          "xl_start_sheet on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_start_row(bogus) == XL_INVALID_HANDLE, "xl_start_row on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_write_int64(bogus, 1) == XL_INVALID_HANDLE, "xl_write_int64 on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_end_row(bogus) == XL_INVALID_HANDLE, "xl_end_row on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_end_sheet(bogus) == XL_INVALID_HANDLE, "xl_end_sheet on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_close_write_handle(bogus) == XL_INVALID_HANDLE, "xl_close_write_handle on a bad handle must return XL_INVALID_HANDLE");
    CHECK(xl_close_write_handle(nullptr) == XL_INVALID_HANDLE, "xl_close_write_handle on a null handle must return XL_INVALID_HANDLE");
    return 0;
}

static int test_writer_handle_rejects_out_of_order_calls()
{
    const std::filesystem::path path = temp_path("writer_handle_order.xlsx");
    const std::string c_path = path.string();
    xl_writer_handle *handle = nullptr;
    int status = xl_open_write_handle(reinterpret_cast<const uint8_t *>(c_path.data()),
                                      static_cast<int32_t>(c_path.size()), XL_FORMAT_XLSX, nullptr, &handle);
    CHECK(status == XL_OK, "xl_open_write_handle must succeed");

    CHECK(xl_start_row(handle) == XL_ERROR, "xl_start_row before xl_start_sheet must fail");
    CHECK(xl_write_int64(handle, 1) == XL_ERROR, "a cell write before xl_start_row must fail");
    CHECK(xl_end_row(handle) == XL_ERROR, "xl_end_row without an open row must fail");
    CHECK(xl_end_sheet(handle) == XL_ERROR, "xl_end_sheet without an open sheet must fail");

    status = xl_start_sheet(handle, reinterpret_cast<const uint8_t *>("S"), 1);
    CHECK(status == XL_OK, "xl_start_sheet must still succeed after the earlier rejected calls");
    status = xl_close_write_handle(handle);
    CHECK(status == XL_OK, "xl_close_write_handle must still succeed after the earlier rejected calls");

    std::filesystem::remove(path);
    return 0;
}

static int test_writer_handle_class_round_trip()
{
    const std::filesystem::path path = temp_path("writer_handle_class.xlsx");
    {
        auto handle = xl::WriterHandle::open(path.string());
        CHECK(handle.has_value(), "WriterHandle::open must succeed");

        CHECK(handle->start_sheet("Planilha1").has_value(), "start_sheet must succeed");

        CHECK(handle->start_row().has_value(), "start_row must succeed for the header row");
        for (std::string_view header : {"texto", "inteiro", "numero", "ativo", "data", "hora", "instante", "opcional"})
        {
            CHECK(handle->write(header).has_value(), "writing a header cell must succeed");
        }
        CHECK(handle->end_row().has_value(), "end_row must succeed for the header row");

        CHECK(handle->start_row().has_value(), "start_row must succeed for the data row");
        CHECK(handle->write(std::string_view("uma")).has_value(), "write(string_view) must succeed");
        CHECK(handle->write(int64_t{1}).has_value(), "write(int64_t) must succeed");
        CHECK(handle->write(0.5).has_value(), "write(double) must succeed");
        CHECK(handle->write(true).has_value(), "write(bool) must succeed");
        CHECK(handle->write(std::chrono::year{2026} / std::chrono::month{1} / std::chrono::day{1}).has_value(),
              "write(year_month_day) must succeed");
        CHECK(handle->write(std::chrono::microseconds{3600000000}).has_value(), "write(microseconds) must succeed");
        CHECK(handle->write(std::chrono::system_clock::time_point{std::chrono::microseconds{1767225600000000}})
                  .has_value(),
              "write(time_point) must succeed");
        CHECK(handle->write(std::optional<int64_t>{7}).has_value(), "write(optional<int64_t>{value}) must succeed");
        CHECK(handle->end_row().has_value(), "end_row must succeed for the data row");

        CHECK(handle->start_row().has_value(), "start_row must succeed for the null-cell row");
        CHECK(handle->write(std::optional<std::string_view>{}).has_value(),
              "write(optional<string_view>{}) must succeed");
        CHECK(handle->write(std::optional<int64_t>{}).has_value(), "write(optional<int64_t>{}) must succeed");
        CHECK(handle->write_null(XL_T_F64).has_value(), "write_null(XL_T_F64) must succeed");
        for (int i = 0; i < 5; ++i)
        {
            CHECK(handle->write_null(XL_T_STRING).has_value(), "padding the null-cell row out must succeed");
        }
        CHECK(handle->end_row().has_value(), "end_row must succeed for the null-cell row");

        CHECK(handle->end_sheet().has_value(), "end_sheet must succeed");
        CHECK(handle->finish().has_value(), "finish must succeed");
        CHECK(!handle->start_row().has_value(), "a finished handle accepts no more writes");
    }
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the file WriterHandle produced must open");
        auto table = xl::parse<FullRow>(workbook->sheet(0));
        CHECK(table.has_value(), "the written file must parse back");
        CHECK(table->size() == 2, "two data rows were written");

        auto first = table->at(0);
        CHECK(first.has_value(), "row 0 must be in bounds");
        CHECK(first->texto == "uma", "the streamed string cell must round-trip");
        CHECK(first->inteiro == 1, "the streamed int64 cell must round-trip");
        CHECK(first->numero == 0.5, "the streamed double cell must round-trip");
        CHECK(first->ativo, "the streamed bool cell must round-trip");
        CHECK(first->opcional == 7, "the streamed optional cell must round-trip its value");

        auto second = table->at(1);
        CHECK(second.has_value(), "row 1 must be in bounds");
        CHECK(second->texto.empty(), "write(optional<string_view>{}) must have written a blank cell");
        CHECK(!second->opcional.has_value(), "write(optional<int64_t>{}) must have written a blank cell");
    }
    std::filesystem::remove(path);
    return 0;
}

static int test_writer_handle_class_to_memory_round_trip()
{
    std::vector<uint8_t> bytes;
    auto handle = xl::WriterHandle::open(bytes, {.format = XL_FORMAT_XLSX});
    CHECK(handle.has_value(), "WriterHandle::open over a vector must succeed");

    CHECK(handle->start_sheet("Dados").has_value(), "start_sheet must succeed");
    CHECK(handle->start_row().has_value(), "start_row must succeed");
    CHECK(handle->write(std::string_view("uma")).has_value(), "write(string_view) must succeed");
    CHECK(handle->write(int64_t{3}).has_value(), "write(int64_t) must succeed");
    CHECK(handle->end_row().has_value(), "end_row must succeed");
    CHECK(handle->end_sheet().has_value(), "end_sheet must succeed");
    CHECK(bytes.empty(), "the vector is filled only when the handle finishes");

    CHECK(handle->finish().has_value(), "finish must succeed");
    CHECK(!bytes.empty(), "finish must fill the vector");
    CHECK(handle->finish().has_value(), "a second finish does nothing");

    auto workbook = xl::Workbook::open(bytes, {.format = XL_FORMAT_XLSX});
    CHECK(workbook.has_value(), "the bytes finish wrote must open");
    auto sheet_name = workbook->sheet(0).name();
    CHECK(sheet_name.has_value() && *sheet_name == "Dados", "start_sheet's name must reach the bytes");
    return 0;
}

static std::vector<std::string> read_cells(const xl::Workbook &workbook, int32_t sheet)
{
    std::vector<std::string> rows;
    auto cursor = workbook.sheet(sheet).rows();
    if (!cursor.has_value())
    {
        return rows;
    }
    for (auto row = cursor->next_row(); row.has_value(); row = cursor->next_row())
    {
        std::string line;
        for (auto cell : *row)
        {
            line += (line.empty() ? "" : "|") + std::string(cell.value);
        }
        rows.push_back(std::move(line));
    }
    return rows;
}

static int test_writer_handle_writes_whole_sheets_beside_streamed_ones()
{
    const std::vector<NamedRow> sales{{"widget", "w"}, {"gadget", "g"}};
    const std::vector<int32_t> offsets{0, 1, 2, 3, 4, 5};
    const std::vector<uint8_t> blob{'a', 'b', 'c', 'd', 'e'};
    const std::vector<int64_t> values{10, 20, 30, 40, 50};
    const std::array<xl::ColumnRef, 2> stock{xl::string_column("name", offsets, blob), xl::i64_column("qty", values)};

    for (const int32_t format : {XL_FORMAT_XLSX, XL_FORMAT_XLSB})
    {
        for (const int32_t degree : {1, 4})
        {
            std::vector<uint8_t> bytes;
            auto handle = xl::WriterHandle::open(bytes, {.format = format, .degree_of_parallelism = degree, .batch_rows = 2});
            CHECK(handle.has_value(), "WriterHandle::open over a vector must succeed");

            CHECK(handle->write_sheet("Vendas", sales).has_value(), "write_sheet must succeed");
            CHECK(handle->start_sheet("Resumo").has_value(), "start_sheet must succeed");
            CHECK(handle->start_row().has_value(), "start_row must succeed");
            CHECK(handle->write(std::string_view("total")).has_value(), "write must succeed");
            CHECK(handle->end_row().has_value(), "end_row must succeed");
            CHECK(handle->end_sheet().has_value(), "end_sheet must succeed");
            CHECK(handle->write_columns("Estoque", stock).has_value(), "write_columns must succeed");
            CHECK(handle->finish().has_value(), "finish must succeed");

            auto workbook = xl::Workbook::open(bytes, {.format = format});
            CHECK(workbook.has_value(), "the bytes finish wrote must open");
            CHECK(workbook->sheet(0).name() == "Vendas", "write_sheet's name must reach the bytes");
            CHECK(workbook->sheet(1).name() == "Resumo", "start_sheet's name must reach the bytes");
            CHECK(workbook->sheet(2).name() == "Estoque", "write_columns's name must reach the bytes");
            CHECK(read_cells(*workbook, 0) == (std::vector<std::string>{"name|view", "widget|w", "gadget|g"}),
                  "write_sheet must write a header and every row");
            CHECK(read_cells(*workbook, 1) == (std::vector<std::string>{"total"}), "the streamed sheet must sit between them");
            CHECK(read_cells(*workbook, 2) == (std::vector<std::string>{"name|qty", "a|10", "b|20", "c|30", "d|40", "e|50"}),
                  "write_columns must write a header and every row in order");
        }
    }
    return 0;
}

static int test_writer_handle_write_columns_fails_while_a_sheet_is_open()
{
    const std::vector<int64_t> values{1};
    const std::array<xl::ColumnRef, 1> columns{xl::i64_column("qty", values)};
    std::vector<uint8_t> bytes;
    auto handle = xl::WriterHandle::open(bytes, {.format = XL_FORMAT_XLSX});
    CHECK(handle.has_value(), "WriterHandle::open over a vector must succeed");

    CHECK(handle->start_sheet("Aberta").has_value(), "start_sheet must succeed");
    CHECK(!handle->write_columns("Tabela", columns).has_value(), "write_columns must fail while a sheet is open");
    CHECK(handle->end_sheet().has_value(), "end_sheet must succeed");
    CHECK(handle->write_columns("Tabela", columns).has_value(), "write_columns must succeed once the sheet is ended");
    CHECK(handle->finish().has_value(), "finish must succeed");

    auto workbook = xl::Workbook::open(bytes, {.format = XL_FORMAT_XLSX});
    CHECK(workbook.has_value() && workbook->sheet(1).name() == "Tabela", "the table must be the second sheet");
    return 0;
}

static int test_writer_handle_writes_excel_tables_the_workbook_reads_back()
{
    const std::vector<NamedRow> sales{{"widget", "w"}, {"gadget", "g"}};
    const std::vector<int32_t> offsets{0, 1, 2, 3};
    const std::vector<uint8_t> blob{'a', 'b', 'c'};
    const std::vector<int64_t> values{10, 20, 30};
    const std::array<xl::ColumnRef, 2> stock{xl::string_column("name", offsets, blob), xl::i64_column("qty", values)};

    for (const int32_t format : {XL_FORMAT_XLSX, XL_FORMAT_XLSB})
    {
        std::vector<uint8_t> bytes;
        auto handle = xl::WriterHandle::open(bytes, {.format = format, .degree_of_parallelism = 4, .batch_rows = 1});
        CHECK(handle.has_value(), "WriterHandle::open over a vector must succeed");
        CHECK(handle->start_sheet("Notas").has_value(), "start_sheet must succeed");
        CHECK(handle->end_sheet().has_value(), "end_sheet must succeed");
        CHECK(handle->write_sheet("Vendas", sales, xl::TableOptions{.name = "Sales"}).has_value(),
              "write_sheet as a table must succeed");
        CHECK(handle->write_columns("Estoque", stock, xl::TableOptions{.name = "Stock", .style = std::nullopt, .row_stripes = false})
                  .has_value(),
              "write_columns as a table must succeed");
        CHECK(handle->finish().has_value(), "finish must succeed");

        auto workbook = xl::Workbook::open(bytes, {.format = format});
        CHECK(workbook.has_value(), "the bytes finish wrote must open");
        auto tables = workbook->tables();
        CHECK(tables.has_value() && tables->size() == 2, "both tables must be listed");
        CHECK((*tables)[0].name == "Sales" && (*tables)[1].name == "Stock", "tables keep their names and order");
        const xl::TableInfo &first = (*tables)[0];
        CHECK(first.index == 0 && first.sheet == 1 && first.first_row == 0 && first.last_row == 2 && first.first_column == 0 &&
                  first.last_column == 1 && first.header_row_count == 1 && first.totals_row_count == 0,
              "the table's range must cover its header and both rows");

        auto found = workbook->table_by_name("sales");
        CHECK(found.has_value() && found->has_value(), "table_by_name must ignore case");
        CHECK((*found)->name() == "Vendas", "a table's sheet reports the sheet holding it");
        CHECK(read_cells(*workbook, (*found)->index()) == (std::vector<std::string>{"name|view", "widget|w", "gadget|g"}),
              "the table must read its header and rows");
        CHECK(read_cells(*workbook, workbook->table(1).index()) == (std::vector<std::string>{"name|qty", "a|10", "b|20", "c|30"}),
              "table() must read the second table");

        auto missing = workbook->table_by_name("missing");
        CHECK(missing.has_value() && !missing->has_value(), "an unknown table name finds nothing");
        CHECK(!workbook->table(2).rows().has_value(), "a table index past the end fails at the first read");
    }

    std::vector<uint8_t> xls;
    auto handle = xl::WriterHandle::open(xls, {.format = XL_FORMAT_XLS});
    CHECK(handle.has_value(), "WriterHandle::open over a vector must succeed");
    CHECK(!handle->write_sheet("Vendas", sales, xl::TableOptions{.name = "Sales"}).has_value(), "XLS has no tables");
    CHECK(handle->write_sheet("Vendas", sales).has_value(), "a plain sheet must still be written");
    return 0;
}

static int test_writer_handle_class_rejects_auto_for_memory()
{
    std::vector<uint8_t> bytes;
    auto handle = xl::WriterHandle::open(bytes);
    CHECK(!handle.has_value(), "WriterHandle::open over a vector must reject XL_FORMAT_AUTO");
    return 0;
}

static std::expected<xl::TableView<NarrowRow>, xl::Error> parse_narrow(const std::vector<int64_t> &values,
                                                                       std::vector<uint8_t> &bytes)
{
    std::array<xl::ColumnRef, 1> columns{xl::i64_column("small", values)};
    if (!xl::write_columns(bytes, columns, {.format = XL_FORMAT_XLSX}).has_value())
    {
        return std::unexpected(xl::Error{XL_ERROR, "write_columns failed"});
    }
    auto workbook = xl::Workbook::open(std::span<const uint8_t>(bytes));
    if (!workbook.has_value())
    {
        return std::unexpected(workbook.error());
    }
    auto sheets = workbook->sheets();
    if (!sheets.has_value())
    {
        return std::unexpected(sheets.error());
    }
    return xl::parse<NarrowRow>(sheets->front());
}

static int test_narrow_field_rejects_a_value_it_cannot_hold()
{
    std::vector<uint8_t> bytes;
    auto too_big = parse_narrow({1, 300}, bytes);
    CHECK(!too_big.has_value(), "300 must not be read into a uint8_t");
    CHECK(too_big.error().code == XL_INVALID_ARGUMENT, "an out-of-range value is an invalid argument");
    CHECK(too_big.error().message.find("small") != std::string::npos, "the error must name the column");

    auto negative = parse_narrow({-1}, bytes);
    CHECK(!negative.has_value(), "-1 must not be read into a uint8_t");

    auto limits = parse_narrow({0, 255}, bytes);
    CHECK(limits.has_value(), "values at the uint8_t limits must parse");
    CHECK((*limits)[0].small == 0 && (*limits)[1].small == 255, "values at the limits must round-trip");

    (void)parse_narrow({300}, bytes);
    auto workbook = xl::Workbook::open(std::span<const uint8_t>(bytes));
    CHECK(workbook.has_value(), "the written bytes must open");
    auto batches = xl::parse_batches<NarrowRow>(workbook->sheets()->front());
    CHECK(batches.has_value(), "parse_batches must open");
    auto batch = batches->next();
    CHECK(!batch.has_value(), "a batch holding 300 must fail for a uint8_t field");
    return 0;
}

static int test_unsigned_values_past_int64_are_refused_on_write()
{
    std::vector<uint8_t> bytes;
    const std::vector<WideRow> rows{{static_cast<uint64_t>(INT64_MAX) + 1}};
    auto written = xl::write_sheet(bytes, rows, {.format = XL_FORMAT_XLSX});
    CHECK(!written.has_value(), "a uint64_t above INT64_MAX must not be written as a negative number");

    const std::vector<WideRow> limit{{static_cast<uint64_t>(INT64_MAX)}};
    CHECK(xl::write_sheet(bytes, limit, {.format = XL_FORMAT_XLSX}).has_value(), "INT64_MAX must be written");
    auto workbook = xl::Workbook::open(std::span<const uint8_t>(bytes));
    CHECK(workbook.has_value(), "the written bytes must open");
    auto table = xl::parse<WideRow>(workbook->sheets()->front());
    CHECK(table.has_value() && (*table)[0].big == static_cast<uint64_t>(INT64_MAX), "INT64_MAX must round-trip");

    std::vector<uint8_t> streamed;
    auto handle = xl::WriterHandle::open(streamed, {.format = XL_FORMAT_XLSX});
    CHECK(handle.has_value(), "the writer handle must open");
    CHECK(handle->start_sheet("S").has_value() && handle->start_row().has_value(), "the sheet and row must start");
    CHECK(!handle->write(UINT64_MAX).has_value(), "WriterHandle::write must refuse a uint64_t above INT64_MAX");
    return 0;
}

static int test_a_valid_empty_string_replaces_the_default()
{
    const int32_t offsets[] = {0, 0};
    xl_column columns[] = {
        xl_column{XL_T_STRING, 1, offsets, nullptr, nullptr, 0},
        xl_column{XL_T_STRING, 1, offsets, nullptr, nullptr, 0},
    };
    const xl_table table{2, 1, columns};
    const NamedRow row = xl::detail::row_at<NamedRow>(table, 0);
    CHECK(row.name.empty(), "an empty std::string must overwrite the default");
    CHECK(row.view.empty(), "an empty std::string_view must overwrite the default");
    return 0;
}

int main()
{
    if (test_writer_handle_writes_whole_sheets_beside_streamed_ones() != 0)
    {
        return 1;
    }
    if (test_writer_handle_write_columns_fails_while_a_sheet_is_open() != 0)
    {
        return 1;
    }
    if (test_writer_handle_writes_excel_tables_the_workbook_reads_back() != 0)
    {
        return 1;
    }
    if (test_narrow_field_rejects_a_value_it_cannot_hold() != 0)
    {
        return 1;
    }
    if (test_unsigned_values_past_int64_are_refused_on_write() != 0)
    {
        return 1;
    }
    if (test_a_valid_empty_string_replaces_the_default() != 0)
    {
        return 1;
    }
    if (test_write_options() != 0)
    {
        return 1;
    }
    if (test_write_columns_in_parallel() != 0)
    {
        return 1;
    }
    if (test_format_from_extension() != 0)
    {
        return 1;
    }
    if (test_write_columns_round_trip() != 0)
    {
        return 1;
    }
    if (test_encrypt_package_round_trip() != 0)
    {
        return 1;
    }
    if (test_encrypt_package_rejects_an_empty_password() != 0)
    {
        return 1;
    }
    if (test_write_columns_to_memory_round_trip() != 0)
    {
        return 1;
    }
    if (test_write_columns_rejects_bad_input() != 0)
    {
        return 1;
    }
    if (test_bool_round_trip() != 0)
    {
        return 1;
    }
    if (test_optional_round_trip() != 0)
    {
        return 1;
    }
    if (test_write_sheet_round_trip() != 0)
    {
        return 1;
    }
    if (test_write_sheet_to_memory_round_trip() != 0)
    {
        return 1;
    }
    if (test_write_sheet_options_and_csv() != 0)
    {
        return 1;
    }
    if (test_writer_handle() != 0)
    {
        return 1;
    }
    if (test_writer_handle_rejects_bad_handle() != 0)
    {
        return 1;
    }
    if (test_writer_handle_rejects_out_of_order_calls() != 0)
    {
        return 1;
    }
    if (test_writer_handle_class_round_trip() != 0)
    {
        return 1;
    }
    if (test_writer_handle_class_to_memory_round_trip() != 0)
    {
        return 1;
    }
    if (test_writer_handle_class_rejects_auto_for_memory() != 0)
    {
        return 1;
    }
    std::printf("OK\n");
    return 0;
}