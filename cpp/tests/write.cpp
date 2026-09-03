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
    auto written = xl::write_columns(plain_path.string(), XL_FORMAT_XLSX, columns);
    CHECK(written.has_value(), "write_columns must succeed");

    auto encrypted = xl::encrypt_package(plain_path.string(), encrypted_path.string(), "hunter2");
    CHECK(encrypted.has_value(), "encrypt_package must succeed");

    {
        // Scoped so workbook/table release their handle on encrypted_path before it's removed
        // below - same reason test_writer_handle_class_round_trip scopes its WriterHandle.
        xl::OpenOptions options;
        options.password("hunter2");
        auto workbook = xl::Workbook::open(encrypted_path.string(), XL_FORMAT_AUTO, &options);
        CHECK(workbook.has_value(), "the encrypted file must open with the right password");
        auto table = xl::parse_sheet<WrittenRow>(*workbook);
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
    auto written = xl::write_columns(plain_path.string(), XL_FORMAT_XLSX, columns);
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

    auto bytes = xl::write_columns_to_memory(XL_FORMAT_XLSX, columns);
    CHECK(bytes.has_value(), "write_columns_to_memory must succeed");
    CHECK(!bytes->empty(), "write_columns_to_memory must return non-empty bytes");

    auto workbook = xl::Workbook::open_memory(*bytes, XL_FORMAT_XLSX);
    CHECK(workbook.has_value(), "the returned bytes must open");
    auto table = xl::parse_sheet<WrittenRow>(*workbook);
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
    // One byte per row, 0 or 1 - XL_T_BOOL's wire layout, not a bit-packed bitmap.
    const std::vector<uint8_t> flags{1, 0, 1};
    const std::array<xl::ColumnRef, 1> columns{xl::bool_column("ativo", flags)};

    const std::filesystem::path path = temp_path("bools.xlsx");
    CHECK(xl::write_columns(path.string(), XL_FORMAT_XLSX, columns).has_value(),
          "writing a bool column must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse_sheet<FlagRow>(*workbook);
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
    // LSB-first: bit 0 and bit 2 set, bit 1 clear - row 1 is null.
    const std::vector<uint8_t> validity{0b00000101};
    const std::array<xl::ColumnRef, 1> columns{xl::i64_column("quantidade", valores, validity)};

    const std::filesystem::path path = temp_path("nullable.xlsx");
    CHECK(xl::write_columns(path.string(), XL_FORMAT_XLSX, columns).has_value(),
          "writing a nullable column must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse_sheet<NullableRow>(*workbook);
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
    // The format is inferred from the .xlsx extension by this overload.
    CHECK(xl::write_sheet(path.string(), rows).has_value(), "write_sheet must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto table = xl::parse_sheet<FullRow>(*workbook);
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

    auto bytes = xl::write_sheet_to_memory(XL_FORMAT_XLSX, rows);
    CHECK(bytes.has_value(), "write_sheet_to_memory must succeed");

    auto workbook = xl::Workbook::open_memory(*bytes, XL_FORMAT_XLSX);
    CHECK(workbook.has_value(), "the returned bytes must open");
    auto table = xl::parse_sheet<FullRow>(*workbook);
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
    xl::WriteOptions options{};
    options.sheet_name = "Dados";
    CHECK(xl::write_sheet(path.string(), XL_FORMAT_XLSX, rows, &options).has_value(),
          "writing an empty sheet with a custom name must succeed");
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the written file must open");
        auto name = workbook->sheet_name();
        CHECK(name.has_value(), "the sheet name must be readable");
        CHECK(*name == "Dados", "WriteOptions::sheet_name must reach the file");
    }
    std::filesystem::remove(path);

    const std::filesystem::path csv = temp_path("sheet.csv");
    CHECK(xl::write_sheet(csv.string(), rows).has_value(), "a .csv path must infer XL_FORMAT_CSV");
    std::filesystem::remove(csv);
    return 0;
}

// Writes one header row plus one data row through the raw streaming C ABI (not the C++ wrapper,
// which does not cover it), then reopens the file with xl::Workbook to prove xl_close_write_handle
// actually produced a valid, readable workbook - not just a status code. This is what the earlier
// version of this test skipped: it never opened the file it wrote, so a workbook left without its
// trailing structure (a corrupt XLSX zip) would still have passed.
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
    CHECK(xl_write_date(handle, 20454) == XL_OK, "xl_write_date must succeed"); // 2026-01-01
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
        auto name = workbook->sheet_name();
        CHECK(name.has_value() && *name == "Planilha1", "xl_start_sheet's name must reach the file");
        auto table = xl::parse_sheet<WrittenRow>(*workbook);
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

// Every writer entry point must resolve a bad/closed handle as XL_INVALID_HANDLE, not
// XL_INVALID_ARGUMENT - the same convention xl_close uses on the reader side.
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

// A row/sheet/write out of order must fail as XL_ERROR (not crash, not silently succeed) and must
// leave the handle usable, per the call-order contract documented on xl_writer_handle.
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

// Exercises xl::WriterHandle::open (file-backed) through every write<T> branch, including the
// std::optional<T> null-cell path, then reopens the file to confirm the output matches what the
// raw-C xl_writer_handle test above wrote by hand.
static int test_writer_handle_class_round_trip()
{
    const std::filesystem::path path = temp_path("writer_handle_class.xlsx");
    {
        // Scoped so the destructor closes and releases the handle - including the exclusive file
        // lock xl_open_write_handle takes - before Workbook::open reopens the same path below.
        auto handle = xl::WriterHandle::open(path.string(), XL_FORMAT_XLSX);
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
    }
    {
        auto workbook = xl::Workbook::open(path.string());
        CHECK(workbook.has_value(), "the file WriterHandle produced must open");
        auto table = xl::parse_sheet<FullRow>(*workbook);
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

// Same as test_writer_handle_class_round_trip, but backed by open_memory()/bytes() instead of a
// file.
static int test_writer_handle_class_to_memory_round_trip()
{
    auto handle = xl::WriterHandle::open_memory(XL_FORMAT_XLSX);
    CHECK(handle.has_value(), "WriterHandle::open_memory must succeed");

    CHECK(handle->start_sheet("Dados").has_value(), "start_sheet must succeed");
    CHECK(handle->start_row().has_value(), "start_row must succeed");
    CHECK(handle->write(std::string_view("uma")).has_value(), "write(string_view) must succeed");
    CHECK(handle->write(int64_t{3}).has_value(), "write(int64_t) must succeed");
    CHECK(handle->end_row().has_value(), "end_row must succeed");

    auto bytes = handle->bytes();
    CHECK(bytes.has_value(), "bytes() must succeed");
    CHECK(!bytes->empty(), "bytes() must return non-empty bytes");

    auto workbook = xl::Workbook::open_memory(*bytes, XL_FORMAT_XLSX);
    CHECK(workbook.has_value(), "the bytes bytes() returned must open");
    auto sheet_name = workbook->sheet_name();
    CHECK(sheet_name.has_value() && *sheet_name == "Dados", "start_sheet's name must reach the bytes");

    // bytes() must not have released the handle - unlike end_sheet/start_sheet (bytes() already
    // ended the sheet internally to produce a valid result, so calling those again would rightly
    // fail), a second bytes() call is still valid and must return the same content.
    auto bytes_again = handle->bytes();
    CHECK(bytes_again.has_value(), "a second bytes() call must still succeed");
    CHECK(*bytes_again == *bytes, "a second bytes() call must return the same content");
    return 0;
}

// bytes() on a file-backed handle (opened via open(), not open_memory()) must fail cleanly - the
// same XL_INVALID_ARGUMENT xl_write_handle_bytes itself returns for that case.
static int test_writer_handle_class_bytes_rejects_a_file_backed_handle()
{
    const std::filesystem::path path = temp_path("writer_handle_class_file.xlsx");
    {
        // Scoped for the same reason as test_writer_handle_class_round_trip: the handle holds the
        // path open exclusively until its destructor runs, and std::filesystem::remove below needs
        // that lock released first.
        auto handle = xl::WriterHandle::open(path.string(), XL_FORMAT_XLSX);
        CHECK(handle.has_value(), "WriterHandle::open must succeed");
        CHECK(handle->start_sheet("S").has_value(), "start_sheet must succeed");

        auto bytes = handle->bytes();
        CHECK(!bytes.has_value(), "bytes() on a file-backed handle must fail");
        CHECK(bytes.error().code == XL_INVALID_ARGUMENT,
              "bytes() on a file-backed handle must be XL_INVALID_ARGUMENT");
    }
    std::filesystem::remove(path);
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
    if (test_writer_handle_class_bytes_rejects_a_file_backed_handle() != 0)
    {
        return 1;
    }
    std::printf("OK\n");
    return 0;
}