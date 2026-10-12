#pragma once

#include "excelreader.h"

#include <array>
#include <atomic>
#include <compare>
#include <concepts>
#include <cstdint>
#include <cstring>
#include <exception>
#include <expected>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <ranges>
#include <span>
#include <string>
#include <string_view>
#include <tuple>
#include <type_traits>
#include <utility>
#include <vector>
#include <chrono>

namespace xl
{


    /// @brief A failure reported by any entry point.
    ///
    /// This header throws nothing: every function that can fail returns `std::expected<T, Error>`.
    struct Error
    {
        int32_t code;        ///< One of the `XL_*` status codes.
        std::string message; ///< The native library's description of what went wrong.
    };

    namespace detail
    {

        inline Error make_error(int32_t code)
        {
            int32_t len = 0;
            const uint8_t *ptr = xl_last_error_ptr(&len);
            std::string message = (ptr != nullptr && len > 0)
                                      ? std::string(reinterpret_cast<const char *>(ptr), static_cast<size_t>(len))
                                      : "unknown error";
            return Error{code, std::move(message)};
        }

        template <typename Call>
        std::expected<std::string, Error> fill_string(Call &&call)
        {
            std::string buffer(128, '\0');
            int32_t len = 0;
            int32_t status = call(reinterpret_cast<uint8_t *>(buffer.data()),
                                  static_cast<int32_t>(buffer.size()), &len);
            if (status == XL_BUFFER_TOO_SMALL)
            {
                buffer.assign(static_cast<size_t>(len > 0 ? len : 0), '\0');
                status = call(reinterpret_cast<uint8_t *>(buffer.data()),
                              static_cast<int32_t>(buffer.size()), &len);
            }
            if (status != XL_OK)
            {
                return std::unexpected(make_error(status));
            }
            buffer.resize(static_cast<size_t>(len > 0 ? len : 0));
            return buffer;
        }

    } 


    /// @brief ABI revision of the loaded native library.
    ///
    /// Every open checks it against the `XL_ABI_VERSION` this header was compiled with and fails on a
    /// mismatch.
    inline int32_t abi_version() noexcept { return xl_abi_version(); }

    namespace detail
    {

        inline void warm_up() noexcept
        {
            static constexpr char csv[] = "a,b\n1,2\n3,4\n";
            xl_workbook *workbook = nullptr;
            if (xl_open_memory(reinterpret_cast<const uint8_t *>(csv), static_cast<int32_t>(sizeof(csv) - 1),
                               XL_FORMAT_CSV, nullptr, &workbook) != XL_OK || workbook == nullptr)
            {
                return;
            }
            xl_rows *rows = nullptr;
            if (xl_rows_open(workbook, 0, &rows) == XL_OK && rows != nullptr)
            {
                xl_row row{};
                while (xl_rows_next_view(rows, &row) == XL_OK)
                {
                }
                xl_rows_close(rows);
            }
            xl_close(workbook);
        }

        inline const std::expected<void, Error> &check_abi_version()
        {
            static const std::expected<void, Error> result = []() -> std::expected<void, Error>
            {
                const int32_t loaded = xl_abi_version();
                if (loaded == XL_ABI_VERSION)
                {
                    warm_up();
                    return {};
                }
                return std::unexpected(Error{
                    XL_ERROR,
                    "ExcelReader native library reports ABI version " + std::to_string(loaded) +
                        ", but this header is version " + std::to_string(XL_ABI_VERSION) +
                        ". Update the header and the native library together."});
            }();
            return result;
        }

    } 


    /// @brief One column guessed by Sheet::infer_schema.
    struct InferredColumn
    {
        std::optional<std::string> name; ///< Header text; empty when the column must be matched by `index`.
        int32_t index = 0;                ///< 0-based column position.
        int32_t type = XL_T_STRING;       ///< One of the `XL_T_*` constants.
        bool nullable = false;            ///< Whether the column may hold empty cells.
    };


    /// @brief Reader limits and CSV dialect for the Workbook::open family.
    ///
    /// A field left at 0 or `XL_OPT_DEFAULT` keeps the library default; on/off fields take
    /// `XL_OPT_TRUE` or `XL_OPT_FALSE`. The `max_*` fields bound what a malformed or hostile file can
    /// make the reader allocate: lower them for untrusted input. An out-of-range value fails the open
    /// with `XL_INVALID_ARGUMENT`.
    struct OpenOptions
    {
        int32_t csv_sniff_dialect = XL_OPT_DEFAULT;      ///< Infer delimiter and quote from the start of a CSV. Off by default.
        int32_t csv_delimiter = 0;                       ///< CSV field separator byte. Default `,`.
        int32_t csv_quote = 0;                           ///< CSV quote byte. Default `"`.
        int32_t csv_detect_bom = XL_OPT_DEFAULT;         ///< Skip a leading UTF-8 byte-order mark. On by default.
        int32_t csv_max_cell_bytes = 0;                  ///< For CSV, replaces `max_cell_bytes`.
        int32_t csv_intern_strings = XL_OPT_DEFAULT;     ///< For CSV, replaces `intern_strings`.

        int64_t max_total_decompressed_bytes = 0;        ///< Bytes an XLSX/XLSB may inflate to, or an XLS workbook stream may declare. Default 512 MiB.
        int32_t max_cell_bytes = 0;                      ///< Largest cell value or buffered CSV record. Default 32 MiB.
        int64_t max_shared_string_bytes = 0;             ///< Largest shared-string table. Default 128 MiB.
        int32_t max_zip_entries = 0;                     ///< Most entries an XLSX/XLSB archive may hold. Default 65,536.
        int32_t prefetch_decompression = XL_OPT_DEFAULT; ///< Inflate XLSX/XLSB sheet data on a background thread ahead of parsing. Off by default.
        int32_t intern_strings = XL_OPT_DEFAULT;         ///< Deduplicate repeated text outside the shared-string table. Off by default; pays off only for low-cardinality text.
        int64_t source_cache_bytes = 0;                  ///< Block cache of Workbook::open_source. Default 64 MiB.
        int64_t max_buffered_bytes = 0;                  ///< Most bytes Workbook::open_stream reads before an XLSX, XLSB or XLS opens.
        int32_t source_block_size = 0;                   ///< Block size of Workbook::open_source's cache. Default 4 MiB; negative turns the cache off.

        std::string password_{}; ///< Set through password().

        /// @brief Sets the password that unlocks an encrypted OOXML workbook (.xlsx/.xlsb/.xlsm).
        ///
        /// Open the workbook with `XL_FORMAT_AUTO`: `XL_FORMAT_XLSX` reads the encrypted container as a
        /// plain ZIP and fails before the password is checked.
        /// @param value The password, as UTF-8.
        /// @return `*this`, for chaining.
        OpenOptions &password(std::string_view value)
        {
            password_ = std::string(value);
            return *this;
        }

        /// @brief These options as the C ABI's `xl_open_options`.
        /// @return A struct pointing into this object, valid while it lives unchanged.
        xl_open_options to_c() const noexcept
        {
            xl_open_options opts{};
            opts.struct_size = sizeof(xl_open_options);
            opts.csv_sniff_dialect = csv_sniff_dialect;
            opts.csv_delimiter = csv_delimiter;
            opts.csv_quote = csv_quote;
            opts.csv_detect_bom = csv_detect_bom;
            opts.csv_max_cell_bytes = csv_max_cell_bytes;
            opts.csv_intern_strings = csv_intern_strings;
            opts.max_total_decompressed_bytes = max_total_decompressed_bytes;
            opts.max_cell_bytes = max_cell_bytes;
            opts.max_shared_string_bytes = max_shared_string_bytes;
            opts.max_zip_entries = max_zip_entries;
            opts.prefetch_decompression = prefetch_decompression;
            opts.intern_strings = intern_strings;
            opts.source_cache_bytes = source_cache_bytes;
            opts.max_buffered_bytes = max_buffered_bytes;
            opts.source_block_size = source_block_size;
            opts.password = password_.empty() ? nullptr : reinterpret_cast<const uint8_t *>(password_.data());
            opts.password_len = static_cast<int32_t>(password_.size());
            return opts;
        }
    };


    /// @brief Settings for every writer.
    ///
    /// A field left at 0 or `XL_OPT_DEFAULT` keeps the library default.
    struct WriteOptions
    {
        std::string_view sheet_name{}; ///< Default `Sheet1`; ignored for CSV. Borrowed: must outlive the write.

        int32_t csv_delimiter = 0; ///< CSV field separator byte. Default `,`.
        int32_t csv_quote = 0;     ///< CSV quote byte. Default `"`.

        int32_t date1904 = XL_OPT_DEFAULT;           ///< XLS/XLSB only: use the 1904 date system. Off by default.
        int32_t use_shared_strings = XL_OPT_DEFAULT; ///< XLSX/XLSB only: store repeated text once in a shared-string table. Off by default.

        /// @brief These options as the C ABI's `xl_write_options`.
        /// @return A struct pointing into this object, valid while it lives unchanged.
        xl_write_options to_c() const noexcept
        {
            xl_write_options opts{};
            opts.struct_size = sizeof(xl_write_options);
            opts.sheet_name_len = static_cast<int32_t>(sheet_name.size());
            opts.sheet_name = sheet_name.empty()
                                  ? nullptr
                                  : reinterpret_cast<const uint8_t *>(sheet_name.data());
            opts.csv_delimiter = csv_delimiter;
            opts.csv_quote = csv_quote;
            opts.date1904 = date1904;
            opts.use_shared_strings = use_shared_strings;
            return opts;
        }
    };

    /// @brief Settings for aggregate_csv_file and aggregate_csv_memory.
    ///
    /// A zero field keeps the library default.
    struct CsvParallelOptions
    {
        int32_t degree_of_parallelism = 0;   ///< Worker threads; 0 uses every processor, 1 runs sequentially.
        int32_t header_row = 0;              ///< 1-based header row, not aggregated; 0 means none.
        int32_t delimiter = 0;               ///< Field separator byte. Default `,`.
        int32_t quote = 0;                   ///< Quote byte. Default `"`.
        int32_t detect_bom = XL_OPT_DEFAULT; ///< Skip a leading UTF-8 byte-order mark. On by default.
        int32_t max_cell_bytes = 0;          ///< Largest field or buffered record. Default 32 MiB.

        /// @brief These options as the C ABI's `xl_csv_parallel_options`.
        xl_csv_parallel_options to_c() const noexcept
        {
            xl_csv_parallel_options opts{};
            opts.struct_size = sizeof(xl_csv_parallel_options);
            opts.degree_of_parallelism = degree_of_parallelism;
            opts.header_row = header_row;
            opts.delimiter = delimiter;
            opts.quote = quote;
            opts.detect_bom = detect_bom;
            opts.max_cell_bytes = max_cell_bytes;
            return opts;
        }
    };

    namespace detail
    {

        template <typename Opts, typename Raw>
        inline const Raw *lower_options(const Opts *options, Raw &storage) noexcept
        {
            if (options == nullptr)
            {
                return nullptr;
            }
            storage = options->to_c();
            return &storage;
        }

        constexpr bool ends_with_ci(std::string_view text, std::string_view suffix) noexcept
        {
            if (text.size() < suffix.size())
            {
                return false;
            }
            const std::string_view tail = text.substr(text.size() - suffix.size());
            for (size_t i = 0; i < suffix.size(); ++i)
            {
                const char c = tail[i];
                const char lowered = (c >= 'A' && c <= 'Z') ? static_cast<char>(c - 'A' + 'a') : c;
                if (lowered != suffix[i])
                {
                    return false;
                }
            }
            return true;
        }

    } 

    /// @brief The format a file name's extension names.
    /// @param path File name or path. Only its extension is read: .xlsx, .xlsb, .xls or .csv, in any case.
    /// @return The matching `XL_FORMAT_*` constant, or `XL_FORMAT_AUTO` when none matches, which every
    /// writer rejects.
    constexpr int32_t format_from_path(std::string_view path) noexcept
    {
        const size_t separator = path.find_last_of("/\\");
        const std::string_view name = (separator == std::string_view::npos)
                                          ? path
                                          : path.substr(separator + 1);
        if (detail::ends_with_ci(name, ".xlsx"))
        {
            return XL_FORMAT_XLSX;
        }
        if (detail::ends_with_ci(name, ".xlsb"))
        {
            return XL_FORMAT_XLSB;
        }
        if (detail::ends_with_ci(name, ".xls"))
        {
            return XL_FORMAT_XLS;
        }
        if (detail::ends_with_ci(name, ".csv"))
        {
            return XL_FORMAT_CSV;
        }
        return XL_FORMAT_AUTO;
    }


    /// @brief One column handed to write_columns.
    ///
    /// Its buffers are borrowed, not copied, and must outlive the write. Build it with i64_column,
    /// string_column and the other column functions rather than by hand.
    struct ColumnRef
    {
        std::string_view name{}; 
        int32_t type = XL_T_STRING;
        int64_t length = 0;
        const void *values = nullptr;
        const uint8_t *validity = nullptr; 
        int64_t validity_len = 0;
        const uint8_t *data = nullptr; 
        int64_t data_len = 0;
    };

    namespace detail
    {

        constexpr const uint8_t *validity_pointer(std::span<const uint8_t> validity) noexcept
        {
            return validity.empty() ? nullptr : validity.data();
        }

        template <int32_t Tag, typename E>
        inline constexpr ColumnRef scalar_column(std::string_view name, std::span<const E> values,
                                                 std::span<const uint8_t> validity) noexcept
        {
            return ColumnRef{name, Tag, static_cast<int64_t>(values.size()), values.data(),
                             validity_pointer(validity), static_cast<int64_t>(validity.size()),
                             nullptr, 0};
        }

    } 

    /// @brief An `XL_T_I64` column.
    /// @param name Header text. Either every column of a table has one or none does.
    /// @param values One value per row.
    /// @param validity Optional LSB-first bitmap, one bit per row, set where the row has a value. Empty
    /// means every row has one.
    /// @return A column borrowing `values` and `validity`.
    inline constexpr ColumnRef i64_column(std::string_view name, std::span<const int64_t> values,
                                          std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_I64>(name, values, validity);
    }

    /// @brief An `XL_T_F64` column.
    /// @param name Header text, as for i64_column.
    /// @param values One value per row.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `values` and `validity`.
    inline constexpr ColumnRef f64_column(std::string_view name, std::span<const double> values,
                                          std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_F64>(name, values, validity);
    }

    /// @brief An `XL_T_BOOL` column.
    /// @param name Header text, as for i64_column.
    /// @param values One byte per row, 0 or 1.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `values` and `validity`.
    inline constexpr ColumnRef bool_column(std::string_view name, std::span<const uint8_t> values,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_BOOL>(name, values, validity);
    }

    /// @brief An `XL_T_DATE` column.
    /// @param name Header text, as for i64_column.
    /// @param days_since_epoch Days since 1970-01-01, one per row.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `days_since_epoch` and `validity`.
    inline constexpr ColumnRef date_column(std::string_view name, std::span<const int32_t> days_since_epoch,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_DATE>(name, days_since_epoch, validity);
    }

    /// @brief An `XL_T_TIME` column.
    /// @param name Header text, as for i64_column.
    /// @param micros_since_midnight Microseconds since midnight, one per row.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `micros_since_midnight` and `validity`.
    inline constexpr ColumnRef time_column(std::string_view name, std::span<const int64_t> micros_since_midnight,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_TIME>(name, micros_since_midnight, validity);
    }

    /// @brief An `XL_T_TIMESTAMP` column.
    /// @param name Header text, as for i64_column.
    /// @param micros_since_epoch Microseconds since 1970-01-01T00:00:00Z, one per row.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `micros_since_epoch` and `validity`.
    inline constexpr ColumnRef timestamp_column(std::string_view name, std::span<const int64_t> micros_since_epoch,
                                                std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_TIMESTAMP>(name, micros_since_epoch, validity);
    }

    /// @brief An `XL_T_STRING` column.
    /// @param name Header text, as for i64_column.
    /// @param offsets One entry per row plus one: row `i` is `data[offsets[i], offsets[i + 1])`.
    /// @param data The UTF-8 text of every row, back to back.
    /// @param validity Optional null bitmap, as for i64_column.
    /// @return A column borrowing `offsets`, `data` and `validity`.
    inline constexpr ColumnRef string_column(std::string_view name, std::span<const int32_t> offsets,
                                             std::span<const uint8_t> data,
                                             std::span<const uint8_t> validity = {}) noexcept
    {
        const int64_t rows = offsets.empty() ? 0 : static_cast<int64_t>(offsets.size()) - 1;
        return ColumnRef{name, XL_T_STRING, rows, offsets.data(), detail::validity_pointer(validity),
                         static_cast<int64_t>(validity.size()), data.empty() ? nullptr : data.data(),
                         static_cast<int64_t>(data.size())};
    }

    namespace detail
    {

        inline std::optional<Error> validate_one_write_column(const ColumnRef &column, size_t index,
                                                              int64_t row_count, bool has_header)
        {
            const std::string at = " (column " + std::to_string(index) + ")";
            if (column.length != row_count)
            {
                return Error{XL_INVALID_ARGUMENT,
                             "every column must have the same length; column 0 has " +
                                 std::to_string(row_count) + " rows but this one has " +
                                 std::to_string(column.length) + at};
            }
            if (column.name.empty() == has_header)
            {
                return Error{XL_INVALID_ARGUMENT,
                             "every column must have a name, or none may - xl_write_typed cannot write "
                             "a partial header row" +
                                 at};
            }
            if (column.validity != nullptr && column.validity_len < (row_count + 7) / 8)
            {
                return Error{XL_INVALID_ARGUMENT,
                             "the validity bitmap is " + std::to_string(column.validity_len) +
                                 " bytes, but " + std::to_string(row_count) + " rows need " +
                                 std::to_string((row_count + 7) / 8) + at};
            }
            if (column.type == XL_T_STRING && column.data_len > INT32_MAX)
            {
                return Error{XL_INVALID_ARGUMENT,
                             "the string blob is larger than 2 GiB, which int32 offsets cannot address" + at};
            }
            return std::nullopt;
        }

        inline std::expected<int64_t, Error> validate_write_columns(std::span<const ColumnRef> columns)
        {
            if (columns.empty())
            {
                return std::unexpected(Error{XL_INVALID_ARGUMENT, "write_columns needs at least one column."});
            }
            const int64_t row_count = columns.front().length;
            const bool has_header = !columns.front().name.empty();
            for (size_t i = 0; i < columns.size(); ++i)
            {
                std::optional<Error> problem = validate_one_write_column(columns[i], i, row_count, has_header);
                if (problem.has_value())
                {
                    return std::unexpected(std::move(*problem));
                }
            }
            return row_count;
        }

        inline void fill_write_column(const ColumnRef &column, const uint8_t *&name_slot, int32_t &len_slot,
                                      xl_column_spec &spec, xl_column &raw) noexcept
        {
            name_slot = column.name.empty() ? nullptr : reinterpret_cast<const uint8_t *>(column.name.data());
            len_slot = static_cast<int32_t>(column.name.size());
            spec = xl_column_spec{&name_slot, &len_slot, column.name.empty() ? 0 : 1, 0, column.type, 0};
            raw = xl_column{column.type, column.length, column.values, column.validity, column.data,
                            column.data_len};
        }

    } 


    /// @brief What a cell holds.
    enum class CellType : int32_t
    {
        Empty = XL_CELL_EMPTY,
        String = XL_CELL_STRING,
        Number = XL_CELL_NUMBER,
        Date = XL_CELL_DATE,       ///< A number in a date format; the value is its serial number.
        Bool = XL_CELL_BOOL,
        Formula = XL_CELL_FORMULA, ///< No longer produced: a formula cell reports its cached result's type.
        Error = XL_CELL_ERROR,     ///< An error value such as `#N/A`.
    };

    /// @brief One cell of a row, borrowing the row's memory.
    struct CellView
    {
        int32_t column{};         ///< 0-based column.
        CellType type{};          ///< What the cell holds.
        std::string_view value{}; ///< The text as stored, so a date is its serial number.
    };

    /// @brief The cells of one row: a range of CellView.
    ///
    /// Borrowed from whatever produced it, so it is valid only as long as that RowCursor row,
    /// DecodedRows or aggregation callback says.
    class RowView
    {
    public:
        RowView() = default;

        RowView(std::span<const uint8_t> payload, size_t count) : payload_(payload), count_(count) {}

        RowView(const xl_row_cell *cells, size_t count) : cells_(cells), count_(count) {}

        size_t size() const { return count_; }
        bool empty() const { return count_ == 0; }

        /// @brief The `index`-th cell, counting cells rather than columns. Unchecked.
        /// @param index Position among the row's cells, below size().
        CellView operator[](size_t index) const
        {
            if (cells_ != nullptr)
            {
                const xl_row_cell &raw = cells_[index];
                const size_t length = raw.value_len > 0 ? static_cast<size_t>(raw.value_len) : 0;
                return CellView{raw.column, static_cast<CellType>(raw.type),
                                length == 0 || raw.value == nullptr
                                    ? std::string_view{}
                                    : std::string_view(reinterpret_cast<const char *>(raw.value), length)};
            }
            auto it = begin();
            std::advance(it, static_cast<ptrdiff_t>(index));
            return *it;
        }

        class iterator
        {
        public:
            using iterator_category = std::input_iterator_tag;
            using value_type = CellView;
            using difference_type = ptrdiff_t;

            iterator() = default;
            iterator(const RowView *row, size_t index) : row_(row), index_(index)
            {
                if (row_ != nullptr)
                {
                    next_ = row_->payload_.data();
                }
                load();
            }

            CellView operator*() const { return current_; }
            iterator &operator++()
            {
                ++index_;
                load();
                return *this;
            }
            iterator operator++(int)
            {
                iterator copy = *this;
                ++*this;
                return copy;
            }
            bool operator==(const iterator &other) const { return index_ == other.index_; }

        private:
            void load()
            {
                if (row_ == nullptr || index_ >= row_->count_)
                {
                    return;
                }
                if (row_->cells_ != nullptr)
                {
                    current_ = (*row_)[index_];
                    return;
                }
                const uint8_t *end = row_->payload_.data() + row_->payload_.size();
                int32_t header[3];
                if (static_cast<size_t>(end - next_) < sizeof(header))
                {
                    index_ = row_->count_;
                    return;
                }
                std::memcpy(header, next_, sizeof(header));
                const uint8_t *value = next_ + sizeof(header);
                if (header[2] < 0 || static_cast<size_t>(end - value) < static_cast<size_t>(header[2]))
                {
                    index_ = row_->count_;
                    return;
                }
                current_ = CellView{header[0], static_cast<CellType>(header[1]),
                                    std::string_view(reinterpret_cast<const char *>(value), static_cast<size_t>(header[2]))};
                next_ = value + header[2];
            }

            const RowView *row_{};
            size_t index_{};
            const uint8_t *next_{};
            CellView current_{};
        };

        iterator begin() const { return iterator(this, 0); }
        iterator end() const { return iterator(this, count_); }

    private:
        std::span<const uint8_t> payload_{};
        const xl_row_cell *cells_{};
        size_t count_{};
    };

    /// @brief Every row of a sheet, or every row a cursor had left, decoded in one native call.
    ///
    /// Owns that memory until destruction; the RowView it hands out borrow from it.
    class DecodedRows
    {
    public:
        explicit DecodedRows(xl_rows_decoded raw) : raw_(raw) {}

        DecodedRows(const DecodedRows &) = delete;
        DecodedRows &operator=(const DecodedRows &) = delete;

        DecodedRows(DecodedRows &&other) noexcept : raw_(std::exchange(other.raw_, xl_rows_decoded{})) {}
        DecodedRows &operator=(DecodedRows &&other) noexcept
        {
            if (this != &other)
            {
                release();
                raw_ = std::exchange(other.raw_, xl_rows_decoded{});
            }
            return *this;
        }

        ~DecodedRows() { release(); }

        size_t size() const { return raw_.row_count > 0 ? static_cast<size_t>(raw_.row_count) : 0; }
        bool empty() const { return size() == 0; }

        /// @brief The row at `index`. Unchecked.
        /// @param index 0-based row, below size().
        RowView operator[](size_t index) const
        {
            const xl_row &row = raw_.rows[index];
            return RowView(row.cells, row.cell_count > 0 ? static_cast<size_t>(row.cell_count) : 0);
        }

        class iterator
        {
        public:
            using iterator_category = std::input_iterator_tag;
            using value_type = RowView;
            using difference_type = ptrdiff_t;

            iterator() = default;
            iterator(const DecodedRows *owner, size_t index) : owner_(owner), index_(index) {}

            RowView operator*() const { return (*owner_)[index_]; }
            iterator &operator++()
            {
                ++index_;
                return *this;
            }
            iterator operator++(int)
            {
                iterator copy = *this;
                ++*this;
                return copy;
            }
            bool operator==(const iterator &other) const { return index_ == other.index_; }

        private:
            const DecodedRows *owner_{};
            size_t index_{};
        };

        iterator begin() const { return iterator(this, 0); }
        iterator end() const { return iterator(this, size()); }

    private:
        void release()
        {
            if (raw_.rows != nullptr)
            {
                xl_free_rows(&raw_);
            }
            raw_ = xl_rows_decoded{};
        }

        xl_rows_decoded raw_{};
    };


    /// @brief Reads one sheet row by row.
    ///
    /// Owns a native cursor and closes it on destruction. Use one cursor from one thread at a time;
    /// any number may be open on one workbook.
    class RowCursor
    {
    public:
        /// @brief Takes ownership of a cursor opened through the C ABI.
        /// @param rows Cursor from `xl_rows_open`, closed by this object.
        explicit RowCursor(xl_rows *rows) noexcept : rows_(rows) {}

        RowCursor(const RowCursor &) = delete;
        RowCursor &operator=(const RowCursor &) = delete;

        RowCursor(RowCursor &&other) noexcept : rows_(std::exchange(other.rows_, nullptr)) {}
        RowCursor &operator=(RowCursor &&other) noexcept
        {
            if (this != &other)
            {
                close();
                rows_ = std::exchange(other.rows_, nullptr);
            }
            return *this;
        }

        ~RowCursor() { close(); }

        /// @brief Reads the next row.
        /// @return A view into memory the cursor owns, valid until the next next_row() or the cursor's
        /// destruction. The end of the sheet arrives as an error whose code is `XL_EOF`.
        std::expected<RowView, Error> next_row()
        {
            xl_row row{};
            const int32_t status = xl_rows_next_view(rows_, &row);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            if (row.cell_count < 0)
            {
                return std::unexpected(detail::make_error(XL_ERROR));
            }
            return RowView(row.cells, static_cast<size_t>(row.cell_count));
        }

        /// @brief Every row this cursor has not yet returned, decoded in one native call.
        std::expected<DecodedRows, Error> read_all_decoded()
        {
            xl_rows_decoded raw{};
            const int32_t status = xl_rows_read_all_decoded(rows_, &raw);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return DecodedRows(raw);
        }

        /// @brief The native cursor, for calling the C ABI directly.
        /// @return The cursor, still owned by this object.
        xl_rows *handle() const noexcept { return rows_; }

    private:
        void close() noexcept
        {
            if (rows_ != nullptr)
            {
                xl_rows_close(rows_);
                rows_ = nullptr;
            }
        }

        xl_rows *rows_ = nullptr;
    };

    /// @brief Whether a sheet's tab is shown.
    enum class SheetVisibility : int32_t
    {
        Visible = XL_SHEET_VISIBLE,
        Hidden = XL_SHEET_HIDDEN,          ///< Hidden, but listed in Excel's Unhide dialog.
        VeryHidden = XL_SHEET_VERY_HIDDEN, ///< Hidden and absent from the Unhide dialog; only code can show it.
    };

    /// @brief One sheet of a Workbook: the native workbook handle and an index.
    ///
    /// Holds no resource of its own, so it is cheap to copy and may be handed to another thread; each
    /// read opens what it needs. After its workbook is destroyed, every read on it returns
    /// `XL_INVALID_HANDLE`.
    class Sheet
    {
    public:
        /// @brief A sheet of a workbook opened through the C ABI. Workbook hands these out.
        /// @param handle Workbook handle, not owned.
        /// @param index 0-based sheet position.
        Sheet(xl_workbook *handle, int32_t index) noexcept : handle_(handle), index_(index) {}

        /// @brief 0-based position in the workbook.
        int32_t index() const noexcept { return index_; }

        /// @brief The sheet's name.
        std::expected<std::string, Error> name() const
        {
            return detail::fill_string([this](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name_at(handle_, index_, buffer, capacity, out_len); });
        }

        /// @brief Whether the sheet's tab is shown.
        std::expected<SheetVisibility, Error> visibility() const
        {
            int32_t raw = 0;
            int32_t status = xl_sheet_visibility_at(handle_, index_, &raw);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return static_cast<SheetVisibility>(raw);
        }

        /// @brief A cursor over this sheet, from its first row.
        /// @return A new, independent cursor on every call.
        std::expected<RowCursor, Error> rows() const
        {
            xl_rows *rows = nullptr;
            int32_t status = xl_rows_open(handle_, index_, &rows);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return RowCursor(rows);
        }

        /// @brief The whole sheet from its first row, decoded in one native call.
        std::expected<DecodedRows, Error> read_all_decoded() const
        {
            auto cursor = rows();
            if (!cursor.has_value())
            {
                return std::unexpected(cursor.error());
            }
            return cursor->read_all_decoded();
        }

        /// @brief Guesses each column's name and type from a sample of the sheet.
        /// @param header_row 1-based row holding the headers; 0 when the sheet has none.
        /// @param sample_size Rows after the header to sample.
        /// @param flags 0, or `XL_INFER_PARSE_TEXT` to also type cells that hold text, such as every CSV
        /// field or numbers stored as text.
        /// @return One InferredColumn per column, in order.
        std::expected<std::vector<InferredColumn>, Error> infer_schema(int32_t header_row = 1,
                                                                       int32_t sample_size = 100,
                                                                       int32_t flags = 0) const
        {
            xl_inferred_schema schema{};
            int32_t status = xl_infer_schema(handle_, index_, header_row, sample_size, flags, &schema);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }

            struct SchemaGuard
            {
                xl_inferred_schema *schema;
                ~SchemaGuard() { xl_free_schema(schema); }
            } guard{&schema};

            std::vector<InferredColumn> columns;
            columns.reserve(static_cast<size_t>(schema.column_count > 0 ? schema.column_count : 0));
            for (int32_t i = 0; i < schema.column_count; ++i)
            {
                const xl_column_spec &spec = schema.columns[i];
                InferredColumn column{};
                if (spec.name_count > 0 && spec.names[0] != nullptr && spec.name_lens[0] > 0)
                {
                    column.name = std::string(reinterpret_cast<const char *>(spec.names[0]),
                                              static_cast<size_t>(spec.name_lens[0]));
                }
                column.index = spec.index;
                column.type = spec.type;
                column.nullable = spec.nullable != 0;
                columns.push_back(std::move(column));
            }
            return columns;
        }

        /// @brief The native workbook handle.
        /// @return The handle, not owned: it dies with its Workbook.
        xl_workbook *handle() const noexcept { return handle_; }

    private:
        xl_workbook *handle_ = nullptr;
        int32_t index_ = 0;
    };

    /// @brief Random-access bytes for Workbook::open_source.
    ///
    /// `read_at` is called from several threads at once.
    class Source
    {
    public:
        virtual ~Source() = default;
        /// @brief Total length in bytes.
        virtual std::uint64_t size() const = 0;
        /// @brief Reads the bytes at an offset.
        /// @param offset Position of the first byte.
        /// @param buffer Where to put them.
        /// @return How many bytes were written, or an error message. An error, or an exception, fails
        /// the read that needed the bytes.
        virtual std::expected<std::size_t, std::string> read_at(std::uint64_t offset, std::span<std::byte> buffer) const = 0;
    };

    /// @brief Sequential bytes for Workbook::open_stream.
    ///
    /// `read` is never called concurrently.
    class InputStream
    {
    public:
        virtual ~InputStream() = default;
        /// @brief Reads the next bytes.
        /// @param buffer Where to put up to `buffer.size()` bytes.
        /// @return How many bytes were written, 0 at the end of the stream, or an error message. An error,
        /// or an exception, fails the open or read that needed the bytes.
        virtual std::expected<std::size_t, std::string> read(std::span<std::byte> buffer) = 0;
    };

    namespace detail
    {
        inline int64_t report_source_error(std::string_view message) noexcept
        {
            xl_set_source_error(reinterpret_cast<const uint8_t *>(message.data()), static_cast<int32_t>(message.size()));
            return -1;
        }

        template <typename Call>
        inline int64_t guarded_read(Call &&call) noexcept
        {
            try
            {
                auto result = call();
                return result.has_value() ? static_cast<int64_t>(*result) : report_source_error(result.error());
            }
            catch (const std::exception &e)
            {
                return report_source_error(e.what());
            }
            catch (...)
            {
                return report_source_error("the source threw an exception");
            }
        }

        inline int64_t source_read_at(void *user_data, int64_t offset, uint8_t *buffer, int64_t length) noexcept
        {
            return guarded_read([&]
                                { return static_cast<const Source *>(user_data)->read_at(
                                      static_cast<std::uint64_t>(offset),
                                      std::span<std::byte>(reinterpret_cast<std::byte *>(buffer), static_cast<std::size_t>(length))); });
        }

        inline int64_t stream_read(void *user_data, uint8_t *buffer, int64_t length) noexcept
        {
            return guarded_read([&]
                                { return static_cast<InputStream *>(user_data)->read(
                                      std::span<std::byte>(reinterpret_cast<std::byte *>(buffer), static_cast<std::size_t>(length))); });
        }

        inline void source_release(void *user_data) noexcept
        {
            delete static_cast<Source *>(user_data);
        }

        inline void stream_release(void *user_data) noexcept
        {
            delete static_cast<InputStream *>(user_data);
        }
    }

    /// @brief An open workbook, closed on destruction.
    ///
    /// Safe to use from several threads at once: every read goes through a Sheet and opens its own
    /// cursor.
    class Workbook
    {
    public:
        Workbook(const Workbook &) = delete;
        Workbook &operator=(const Workbook &) = delete;

        Workbook(Workbook &&other) noexcept : handle_(std::exchange(other.handle_, nullptr)) {}
        Workbook &operator=(Workbook &&other) noexcept
        {
            if (this != &other)
            {
                close();
                handle_ = std::exchange(other.handle_, nullptr);
            }
            return *this;
        }

        ~Workbook() { close(); }

        /// @brief Opens the workbook at a path.
        /// @param path File to open.
        /// @param format An `XL_FORMAT_*` constant. `XL_FORMAT_AUTO` detects XLSX, XLSB and XLS from the
        /// file's signature; CSV has none and needs `XL_FORMAT_CSV`.
        /// @param options Reader limits and CSV dialect, or null for the defaults.
        /// @return The workbook. Fails if the native library's ABI does not match this header.
        static std::expected<Workbook, Error> open(std::string_view path, int32_t format = XL_FORMAT_AUTO,
                                                   const OpenOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            xl_workbook *handle = nullptr;
            xl_open_options c_options{};
            const xl_open_options *c_options_ptr = detail::lower_options(options, c_options);
            int32_t status = xl_open_file(reinterpret_cast<const uint8_t *>(path.data()),
                                          static_cast<int32_t>(path.size()), format,
                                          c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

        /// @brief Opens a workbook from bytes in memory.
        /// @param data The file's bytes. The native library copies them, so they need not outlive the call.
        /// @param format As for Workbook::open.
        /// @param options As for Workbook::open.
        static std::expected<Workbook, Error> open_memory(std::span<const uint8_t> data, int32_t format = XL_FORMAT_AUTO,
                                                          const OpenOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            xl_workbook *handle = nullptr;
            xl_open_options c_options{};
            const xl_open_options *c_options_ptr = detail::lower_options(options, c_options);
            int32_t status = xl_open_memory(data.data(), static_cast<int32_t>(data.size()), format,
                                            c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

        /// @brief Opens a workbook over random-access bytes.
        /// @param source Owned by the library from here on. Read from several threads, so every sheet can
        /// still be read in parallel.
        /// @param format As for Workbook::open.
        /// @param options As for Workbook::open; `source_block_size` and `source_cache_bytes` tune the
        /// block cache.
        static std::expected<Workbook, Error> open_source(std::unique_ptr<Source> source, int32_t format = XL_FORMAT_AUTO,
                                                          const OpenOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            if (source == nullptr)
            {
                return std::unexpected(Error{XL_INVALID_ARGUMENT, "source is null"});
            }
            const std::uint64_t size = source->size();
            if (size > static_cast<std::uint64_t>(std::numeric_limits<int64_t>::max()))
            {
                return std::unexpected(Error{XL_INVALID_ARGUMENT, "source size exceeds INT64_MAX"});
            }
            xl_source raw{};
            raw.struct_size = sizeof(xl_source);
            raw.length = static_cast<int64_t>(size);
            raw.read_at = &detail::source_read_at;
            raw.release = &detail::source_release;
            raw.user_data = source.release();
            xl_workbook *handle = nullptr;
            xl_open_options c_options{};
            int32_t status = xl_open_source(&raw, format, detail::lower_options(options, c_options), &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

        /// @brief Opens a workbook over a stream that cannot seek.
        ///
        /// A CSV is read as it arrives, and again from the start until 16 MiB of it were read; an XLSX,
        /// XLSB or XLS is read whole first.
        /// @param stream Owned by the library from here on.
        /// @param format As for Workbook::open.
        /// @param options As for Workbook::open; `max_buffered_bytes` caps what an XLSX, XLSB or XLS may
        /// buffer.
        static std::expected<Workbook, Error> open_stream(std::unique_ptr<InputStream> stream, int32_t format = XL_FORMAT_AUTO,
                                                          const OpenOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            if (stream == nullptr)
            {
                return std::unexpected(Error{XL_INVALID_ARGUMENT, "stream is null"});
            }
            xl_stream raw{};
            raw.struct_size = sizeof(xl_stream);
            raw.read = &detail::stream_read;
            raw.release = &detail::stream_release;
            raw.user_data = stream.release();
            xl_workbook *handle = nullptr;
            xl_open_options c_options{};
            int32_t status = xl_open_stream(&raw, format, detail::lower_options(options, c_options), &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }


        /// @brief Number of sheets.
        std::expected<int32_t, Error> sheet_count() const
        {
            int32_t count = 0;
            int32_t status = xl_sheet_count(handle_, &count);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return count;
        }

        /// @brief Name of the sheet at `index`.
        /// @param index 0-based sheet position.
        std::expected<std::string, Error> sheet_name_at(int32_t index) const
        {
            return detail::fill_string([this, index](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name_at(handle_, index, buffer, capacity, out_len); });
        }

        /// @brief Every sheet's name, in workbook order.
        std::expected<std::vector<std::string>, Error> sheet_names() const
        {
            auto count = sheet_count();
            if (!count.has_value())
            {
                return std::unexpected(count.error());
            }
            std::vector<std::string> names;
            names.reserve(static_cast<size_t>(*count > 0 ? *count : 0));
            for (int32_t i = 0; i < *count; ++i)
            {
                auto name = sheet_name_at(i);
                if (!name.has_value())
                {
                    return std::unexpected(name.error());
                }
                names.push_back(std::move(*name));
            }
            return names;
        }

        /// @brief The sheet at `index`.
        ///
        /// Makes no native call: an index outside the workbook is reported by the first read on the
        /// returned sheet.
        /// @param index 0-based sheet position.
        Sheet sheet(int32_t index) const noexcept { return Sheet(handle_, index); }

        /// @brief Every sheet, in workbook order. Opens nothing.
        std::expected<std::vector<Sheet>, Error> sheets() const
        {
            auto count = sheet_count();
            if (!count.has_value())
            {
                return std::unexpected(count.error());
            }
            std::vector<Sheet> result;
            result.reserve(static_cast<size_t>(*count > 0 ? *count : 0));
            for (int32_t i = 0; i < *count; ++i)
            {
                result.emplace_back(handle_, i);
            }
            return result;
        }

        /// @brief The sheet called `name`, compared without regard to case.
        /// @param name Sheet name.
        /// @return The sheet, or empty when no sheet matches.
        std::expected<std::optional<Sheet>, Error> sheet_by_name(std::string_view name) const
        {
            int32_t index = -1;
            int32_t status = xl_sheet_index(handle_, reinterpret_cast<const uint8_t *>(name.data()),
                                            static_cast<int32_t>(name.size()), &index);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            if (index < 0)
            {
                return std::optional<Sheet>{};
            }
            return std::optional<Sheet>{Sheet(handle_, index)};
        }

        /// @brief Whether dates count from 1904 instead of 1900.
        ///
        /// A date cell's serial number depends on it.
        std::expected<bool, Error> is_date1904() const
        {
            int32_t flag = 0;
            int32_t status = xl_is_date1904(handle_, &flag);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return flag != 0;
        }

        /// @brief The native handle, for calling the C ABI directly.
        /// @return The handle, still owned by this object.
        xl_workbook *handle() const noexcept { return handle_; }

    private:
        explicit Workbook(xl_workbook *handle) noexcept : handle_(handle) {}

        void close() noexcept
        {
            if (handle_ != nullptr)
            {
                xl_close(handle_);
                handle_ = nullptr;
            }
        }

        xl_workbook *handle_ = nullptr;
    };


    /// @brief The `XL_T_*` column type `T` reads and writes as, in `value`.
    ///
    /// Defined for std::string and std::string_view, integers, floating point, bool, the std::chrono
    /// date and time types, and std::optional of any of them.
    /// @tparam T A field type.
    template <typename T>
    struct XlType; 

    template <>
    struct XlType<std::string>
    {
        static constexpr int32_t value = XL_T_STRING;
    };

    template <>
    struct XlType<std::string_view>
    {
        static constexpr int32_t value = XL_T_STRING;
    };

    template <typename T>
        requires std::is_integral_v<T>
    struct XlType<T>
    {
        static constexpr int32_t value = XL_T_I64;
    };

    template <typename T>
        requires std::is_floating_point_v<T>
    struct XlType<T>
    {
        static constexpr int32_t value = XL_T_F64;
    };

    template <>
    struct XlType<bool>
    {
        static constexpr int32_t value = XL_T_BOOL;
    };

    template <>
    struct XlType<std::chrono::system_clock::time_point>
    {
        static constexpr int32_t value = XL_T_TIMESTAMP;
    };

    template <>
    struct XlType<std::chrono::year_month_day>
    {
        static constexpr int32_t value = XL_T_DATE;
    };

    template <>
    struct XlType<std::chrono::sys_days>
    {
        static constexpr int32_t value = XL_T_DATE;
    };

    template <>
    struct XlType<std::chrono::microseconds>
    {
        static constexpr int32_t value = XL_T_TIME;
    };

    template <>
    struct XlType<std::chrono::hh_mm_ss<std::chrono::microseconds>>
    {
        static constexpr int32_t value = XL_T_TIME;
    };

    template <typename T>
    struct XlType<std::optional<T>>
    {
        static constexpr int32_t value = XlType<T>::value;
    };

    namespace detail
    {

        template <typename T>
        struct IsOptional : std::false_type
        {
            using Inner = T;
        };

        template <typename T>
        struct IsOptional<std::optional<T>> : std::true_type
        {
            using Inner = T;
        };

        template <typename T>
        inline constexpr bool is_optional_v = IsOptional<T>::value;

        template <typename T>
        using unwrap_optional_t = typename IsOptional<T>::Inner;

    } 


    /// @brief A member bound to a column by header name. Built by make_field.
    /// @tparam Class The row type.
    /// @tparam T The member's type.
    /// @tparam N How many header names the binding accepts.
    template <typename Class, typename T, std::size_t N = 1>
    struct FieldBinding
    {
        std::array<const char *, N> column_names;
        T Class::*member;
        using FieldType = T;
    };

    /// @brief Binds a member to the column with one header.
    /// @param name Header text.
    /// @param member The member to fill or write.
    /// @return The binding, for ExcelMapper::get_bindings.
    template <typename Class, typename T>
    constexpr FieldBinding<Class, T> make_field(const char *name, T Class::*member)
    {
        return {{name}, member};
    }

    /// @brief Binds a member to the column headed by any of several names.
    /// @param names Accepted header texts.
    /// @param member The member to fill or write.
    /// @return The binding, for ExcelMapper::get_bindings.
    template <typename Class, typename T, std::size_t N>
    constexpr FieldBinding<Class, T, N> make_field(const char *(&&names)[N], T Class::*member)
    {
        FieldBinding<Class, T, N> result{};
        for (std::size_t i = 0; i < N; ++i)
        {
            result.column_names[i] = names[i];
        }
        result.member = member;
        return result;
    }

    /// @brief Maps a row type to columns.
    ///
    /// Specialize it with a `static constexpr auto get_bindings()` that returns a std::tuple of
    /// make_field results; parse_sheet, typed_reader, parse_arrow and write_sheet all read the same
    /// mapping.
    /// @tparam T The row type.
    template <typename T>
    struct ExcelMapper;

    namespace detail
    {
        template <typename Class, typename T, std::size_t N>
        xl_column_spec build_one_spec(const FieldBinding<Class, T, N> &binding, std::vector<int32_t> &name_lens_storage)
        {
            name_lens_storage.resize(N);
            for (std::size_t i = 0; i < N; ++i)
            {
                name_lens_storage[i] = static_cast<int32_t>(std::strlen(binding.column_names[i]));
            }
            return xl_column_spec{
                reinterpret_cast<const uint8_t *const *>(binding.column_names.data()),
                name_lens_storage.data(),
                static_cast<int32_t>(N),
                0, 
                XlType<T>::value,
                1 
            };
        }
        template <typename Tuple, std::size_t... Is>
        std::array<xl_column_spec, sizeof...(Is)> build_specs(const Tuple &bindings, std::index_sequence<Is...>, std::array<std::vector<int32_t>, sizeof...(Is)> &name_lens_storage)
        {
            return {build_one_spec(std::get<Is>(bindings), name_lens_storage[Is])...};
        }

        inline bool is_valid(const xl_column &col, int64_t row)
        {
            if (col.validity == nullptr)
            {
                return true;
            }
            int64_t byte_idx = row / 8;
            int64_t bit_idx = row % 8;
            return (col.validity[byte_idx] & (1 << bit_idx)) != 0;
        }

        template <typename Class, typename T, std::size_t N>
        void assign_field(Class &instance, const xl_column &col, int64_t row, const FieldBinding<Class, T, N> &binding)
        {
            if (!is_valid(col, row))
            {
                return; 
            }
            if constexpr (detail::is_optional_v<T>)
            {
                using Inner = detail::unwrap_optional_t<T>;
                struct Holder
                {
                    Inner value{};
                };
                Holder holder{};
                const FieldBinding<Holder, Inner, N> inner_binding{binding.column_names, &Holder::value};
                assign_field(holder, col, row, inner_binding);
                instance.*(binding.member) = std::move(holder.value);
            }
            else if constexpr (std::is_same_v<T, std::string> || std::is_same_v<T, std::string_view>)
            {
                const int32_t *offsets = static_cast<const int32_t *>(col.values);
                int32_t start = offsets[row];
                int32_t end = offsets[row + 1];
                if (end > start && col.data != nullptr)
                {
                    const char *str_data = reinterpret_cast<const char *>(col.data) + start;
                    instance.*(binding.member) = T(str_data, static_cast<size_t>(end - start));
                }
            }
            else if constexpr (std::is_same_v<T, bool>)
            {
                instance.*(binding.member) = (static_cast<const uint8_t *>(col.values)[row] != 0);
            }
            else if constexpr (std::is_integral_v<T>)
            {
                instance.*(binding.member) = T(static_cast<const int64_t *>(col.values)[row]);
            }
            else if constexpr (std::is_floating_point_v<T>)
            {
                instance.*(binding.member) = T(static_cast<const double *>(col.values)[row]);
            }
            else if constexpr (std::is_same_v<T, std::chrono::sys_days>)
            {
                int32_t days = static_cast<const int32_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::sys_days{std::chrono::days{days}};
            }
            else if constexpr (std::is_same_v<T, std::chrono::year_month_day>)
            {
                int32_t days = static_cast<const int32_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::year_month_day{std::chrono::sys_days{std::chrono::days{days}}};
            }
            else if constexpr (std::is_same_v<T, std::chrono::microseconds>)
            {
                int64_t micros = static_cast<const int64_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::microseconds{micros};
            }
            else if constexpr (std::is_same_v<T, std::chrono::hh_mm_ss<std::chrono::microseconds>>)
            {
                int64_t micros = static_cast<const int64_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::hh_mm_ss<std::chrono::microseconds>{std::chrono::microseconds{micros}};
            }
            else if constexpr (std::is_same_v<T, std::chrono::system_clock::time_point>)
            {
                int64_t micros = static_cast<const int64_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::time_point_cast<std::chrono::system_clock::duration>(
                    std::chrono::sys_time<std::chrono::microseconds>{std::chrono::microseconds{micros}});
            }
        }

        template <typename T, typename Tuple, std::size_t... Is>
        inline void populate_instance(T &instance, const xl_table &table, int64_t row, const Tuple &bindings, std::index_sequence<Is...>)
        {
            (..., assign_field(instance, table.columns[Is], row, std::get<Is>(bindings)));
        }

        template <typename T>
        inline T row_at(const xl_table &table, int64_t row)
        {
            T instance{};
            static constexpr auto bindings = ExcelMapper<T>::get_bindings();
            static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
            populate_instance(instance, table, row, bindings, std::make_index_sequence<num_fields>{});
            return instance;
        }

    } 


    /// @brief The rows of a typed parse, held in native columnar memory.
    ///
    /// Each row converts to `T` on access. Owns that memory; a random-access range of `T`.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    template <typename T>
    class TableView
    {
    public:
        TableView(const TableView &) = delete;
        TableView &operator=(const TableView &) = delete;

        TableView(TableView &&other) noexcept : table_(std::exchange(other.table_, xl_table{})) {}
        TableView &operator=(TableView &&other) noexcept
        {
            if (this != &other)
            {
                xl_free_table(&table_);
                table_ = std::exchange(other.table_, xl_table{});
            }
            return *this;
        }

        ~TableView() { xl_free_table(&table_); } 

        class iterator
        {
        public:
            using iterator_category = std::random_access_iterator_tag;
            using iterator_concept = std::random_access_iterator_tag;
            using value_type = T;
            using difference_type = std::ptrdiff_t;
            using pointer = void;
            using reference = T;

            iterator() = default;

            T operator*() const { return detail::row_at<T>(*table_, row_); }

            T operator[](difference_type n) const { return *(*this + n); }

            iterator &operator++()
            {
                ++row_;
                return *this;
            }

            iterator operator++(int)
            {
                iterator tmp = *this;
                ++row_;
                return tmp;
            }

            iterator &operator--()
            {
                --row_;
                return *this;
            }

            iterator operator--(int)
            {
                iterator tmp = *this;
                --row_;
                return tmp;
            }

            iterator &operator+=(difference_type n)
            {
                row_ += n;
                return *this;
            }

            iterator &operator-=(difference_type n)
            {
                row_ -= n;
                return *this;
            }

            friend iterator operator+(iterator it, difference_type n)
            {
                it += n;
                return it;
            }

            friend iterator operator+(difference_type n, iterator it)
            {
                it += n;
                return it;
            }

            friend iterator operator-(iterator it, difference_type n)
            {
                it -= n;
                return it;
            }

            friend difference_type operator-(const iterator &lhs, const iterator &rhs) { return lhs.row_ - rhs.row_; }

            friend bool operator==(const iterator &lhs, const iterator &rhs) { return lhs.row_ == rhs.row_; }
            friend std::strong_ordering operator<=>(const iterator &lhs, const iterator &rhs) { return lhs.row_ <=> rhs.row_; }

        private:
            friend class TableView;
            iterator(const xl_table *table, int64_t row) : table_(table), row_(row) {}

            const xl_table *table_ = nullptr;
            int64_t row_ = 0;
        };

        static_assert(std::random_access_iterator<iterator>);

        iterator begin() const { return iterator(&table_, 0); }
        iterator end() const { return iterator(&table_, table_.row_count); }

        int64_t size() const noexcept { return table_.row_count; }
        bool empty() const noexcept { return table_.row_count == 0; }

        /// @brief Row `row` as `T`. Unchecked; at() is the checked read.
        /// @param row 0-based row, below size().
        T operator[](int64_t row) const { return detail::row_at<T>(table_, row); }

        /// @brief Row `row` as `T`, bounds-checked.
        /// @param row 0-based row.
        /// @return The row, or empty outside `[0, size())`.
        std::optional<T> at(int64_t row) const
        {
            if (row < 0 || row >= table_.row_count)
            {
                return std::nullopt;
            }
            return (*this)[row];
        }

        /// @brief Every row converted to `T`.
        std::vector<T> to_vector() const
        {
            std::vector<T> result;
            result.reserve(static_cast<size_t>(size()));
            for (T item : *this)
            {
                result.push_back(std::move(item));
            }
            return result;
        }

        /// @brief Takes ownership of a table produced by the C ABI.
        /// @param table Table from `xl_parse_typed` or `xl_typed_reader_next`, freed by this object.
        static TableView from_raw(xl_table table) { return TableView(table); }

    private:
        explicit TableView(xl_table table) noexcept : table_(table) {}

        xl_table table_{};
    };


    /// @brief Reads a whole sheet as `T`.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    /// @param sheet The sheet to read.
    /// @param header_row 1-based row holding the headers; 0 when the sheet has none.
    /// @param degree_of_parallelism Threads to split an XLSX, XLSB or CSV across: 0 uses every
    /// processor, 1 runs sequentially. The result is the same either way.
    /// @return Every row after the header.
    template <typename T>
    std::expected<TableView<T>, Error> parse_sheet(Sheet sheet, int32_t header_row = 1,
                                                   int32_t degree_of_parallelism = 1)
    {
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);
        std::span<const xl_column_spec> specs(specs_array);

        xl_table table{};
        int32_t status = xl_parse_typed(sheet.handle(), sheet.index(), specs.data(),
                                        static_cast<int32_t>(specs.size()), header_row,
                                        degree_of_parallelism, &table);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return TableView<T>::from_raw(table);
    }


    /// @brief parse_sheet delivered a batch at a time.
    ///
    /// An input range of `std::expected<TableView<T>, Error>` that stops after the first error. Each
    /// batch owns its memory and outlives the reader. Use from one thread at a time.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    template <typename T>
    class TypedReader
    {
    public:
        TypedReader(const TypedReader &) = delete;
        TypedReader &operator=(const TypedReader &) = delete;

        TypedReader(TypedReader &&other) noexcept : reader_(std::exchange(other.reader_, nullptr)) {}

        TypedReader &operator=(TypedReader &&other) noexcept
        {
            if (this != &other)
            {
                xl_typed_reader_close(reader_);
                reader_ = std::exchange(other.reader_, nullptr);
            }
            return *this;
        }

        ~TypedReader() { xl_typed_reader_close(reader_); } 

        /// @brief Reads the next batch.
        /// @return The batch, or empty at the end of the sheet.
        std::expected<std::optional<TableView<T>>, Error> next()
        {
            xl_table table{};
            int32_t status = xl_typed_reader_next(reader_, &table);
            if (status == XL_EOF)
            {
                return std::optional<TableView<T>>{};
            }
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return std::optional<TableView<T>>(TableView<T>::from_raw(table));
        }

        class iterator
        {
        public:
            using iterator_category = std::input_iterator_tag;
            using iterator_concept = std::input_iterator_tag;
            using value_type = std::expected<TableView<T>, Error>;
            using difference_type = std::ptrdiff_t;

            iterator() = default;

            value_type &operator*() { return *current_; }

            iterator &operator++()
            {
                advance();
                return *this;
            }

            void operator++(int) { advance(); }

            friend bool operator==(const iterator &it, std::default_sentinel_t) noexcept
            {
                return !it.current_.has_value();
            }

        private:
            friend class TypedReader;

            explicit iterator(TypedReader *reader) : reader_(reader) { advance(); }

            void advance()
            {
                if (reader_ == nullptr)
                {
                    current_.reset();
                    return;
                }
                current_.reset();
                std::expected<std::optional<TableView<T>>, Error> batch = reader_->next();
                if (!batch.has_value())
                {
                    current_ = value_type(std::unexpect, std::move(batch.error()));
                    reader_ = nullptr;
                    return;
                }
                if (!batch->has_value())
                {
                    current_.reset();
                    reader_ = nullptr;
                    return;
                }
                current_ = value_type(std::move(**batch));
            }

            TypedReader *reader_ = nullptr;
            std::optional<value_type> current_{};
        };

        iterator begin() { return iterator(this); }
        std::default_sentinel_t end() const noexcept { return {}; }

        /// @brief Takes ownership of a reader opened through the C ABI.
        /// @param reader Reader from `xl_typed_reader_open`, closed by this object.
        static TypedReader from_raw(xl_typed_reader *reader) { return TypedReader(reader); }

    private:
        explicit TypedReader(xl_typed_reader *reader) noexcept : reader_(reader) {}

        xl_typed_reader *reader_ = nullptr;
    };

    /// @brief Opens a TypedReader over a sheet.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    /// @param sheet The sheet to read.
    /// @param header_row As for parse_sheet.
    /// @param batch_size Rows per batch: 0 makes one batch of the whole sheet; negative is an error.
    template <typename T>
    std::expected<TypedReader<T>, Error> typed_reader(Sheet sheet, int32_t header_row = 1,
                                                      int64_t batch_size = 10000)
    {
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);

        xl_typed_reader *reader = nullptr;
        int32_t status = xl_typed_reader_open(sheet.handle(), sheet.index(), specs_array.data(),
                                              static_cast<int32_t>(specs_array.size()), header_row,
                                              batch_size, &reader);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return TypedReader<T>::from_raw(reader);
    }

    /// @brief Writes columns as one sheet, copying nothing.
    /// @param path File to create.
    /// @param format An `XL_FORMAT_*` constant. `XL_FORMAT_AUTO` is rejected: a file being created has
    /// no signature to detect.
    /// @param columns Columns of equal length, borrowed for the call. Either every column has a name,
    /// written as a header row, or none does.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    inline std::expected<void, Error> write_columns(std::string_view path, int32_t format,
                                                    std::span<const ColumnRef> columns,
                                                    const WriteOptions *options = nullptr)
    {
        const std::expected<void, Error> &abi = detail::check_abi_version();
        if (!abi.has_value())
        {
            return std::unexpected(abi.error());
        }
        std::expected<int64_t, Error> row_count = detail::validate_write_columns(columns);
        if (!row_count.has_value())
        {
            return std::unexpected(std::move(row_count.error()));
        }

        const size_t count = columns.size();
        std::vector<const uint8_t *> name_slots(count);
        std::vector<int32_t> name_lens(count);
        std::vector<xl_column_spec> specs(count);
        std::vector<xl_column> raw_columns(count);
        for (size_t i = 0; i < count; ++i)
        {
            detail::fill_write_column(columns[i], name_slots[i], name_lens[i], specs[i], raw_columns[i]);
        }

        xl_table table{static_cast<int32_t>(count), *row_count, raw_columns.data()};
        xl_write_options raw_options{};
        const xl_write_options *options_pointer = detail::lower_options(options, raw_options);

        const int32_t status = xl_write_typed(reinterpret_cast<const uint8_t *>(path.data()),
                                              static_cast<int32_t>(path.size()), format, specs.data(),
                                              &table, options_pointer);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return {};
    }

    /// @brief write_columns with the format taken from the file's extension.
    /// @param path File to create; its extension picks the format.
    /// @param columns As for the overload that takes a format.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    inline std::expected<void, Error> write_columns(std::string_view path, std::span<const ColumnRef> columns,
                                                    const WriteOptions *options = nullptr)
    {
        return write_columns(path, format_from_path(path), columns, options);
    }

    /// @brief write_columns into memory.
    /// @param format As for write_columns. There is no default: no file name to infer it from.
    /// @param columns As for write_columns.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    /// @return The file's bytes.
    inline std::expected<std::vector<uint8_t>, Error> write_columns_to_memory(int32_t format,
                                                                              std::span<const ColumnRef> columns,
                                                                              const WriteOptions *options = nullptr)
    {
        const std::expected<void, Error> &abi = detail::check_abi_version();
        if (!abi.has_value())
        {
            return std::unexpected(abi.error());
        }
        std::expected<int64_t, Error> row_count = detail::validate_write_columns(columns);
        if (!row_count.has_value())
        {
            return std::unexpected(std::move(row_count.error()));
        }

        const size_t count = columns.size();
        std::vector<const uint8_t *> name_slots(count);
        std::vector<int32_t> name_lens(count);
        std::vector<xl_column_spec> specs(count);
        std::vector<xl_column> raw_columns(count);
        for (size_t i = 0; i < count; ++i)
        {
            detail::fill_write_column(columns[i], name_slots[i], name_lens[i], specs[i], raw_columns[i]);
        }

        xl_table table{static_cast<int32_t>(count), *row_count, raw_columns.data()};
        xl_write_options raw_options{};
        const xl_write_options *options_pointer = detail::lower_options(options, raw_options);

        xl_buffer buffer{};
        const int32_t status = xl_write_typed_to_memory(format, specs.data(), &table, options_pointer, &buffer);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        struct BufferGuard
        {
            xl_buffer *buffer;
            ~BufferGuard() { xl_free_buffer(buffer); }
        } guard{&buffer};
        return std::vector<uint8_t>(buffer.data, buffer.data + buffer.len);
    }


    /// @brief Encrypts a finished XLSX or XLSB.
    ///
    /// The encryption parameters are fixed at Excel's own defaults; the result opens in Excel, and in
    /// Workbook::open given the same password.
    /// @param package_path The plaintext package. Read twice, so it must already be a finished file.
    /// @param destination_path Where to write the encrypted file.
    /// @param password The password, as UTF-8.
    inline std::expected<void, Error> encrypt_package(std::string_view package_path,
                                                       std::string_view destination_path,
                                                       std::string_view password)
    {
        const std::expected<void, Error> &abi = detail::check_abi_version();
        if (!abi.has_value())
        {
            return std::unexpected(abi.error());
        }
        const int32_t status = xl_encrypt_package(
            reinterpret_cast<const uint8_t *>(package_path.data()), static_cast<int32_t>(package_path.size()),
            reinterpret_cast<const uint8_t *>(destination_path.data()), static_cast<int32_t>(destination_path.size()),
            reinterpret_cast<const uint8_t *>(password.data()), static_cast<int32_t>(password.size()));
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return {};
    }


    namespace detail
    {

        struct StringBuffer
        {
            std::vector<int32_t> offsets{0};
            std::vector<uint8_t> data{};
            bool overflowed = false;

            void reserve(size_t rows)
            {
                offsets.reserve(rows + 1);
            }

            void push(std::string_view value)
            {
                if (data.size() + value.size() > static_cast<size_t>(INT32_MAX))
                {
                    overflowed = true;
                    offsets.push_back(offsets.back());
                    return;
                }
                const uint8_t *bytes = reinterpret_cast<const uint8_t *>(value.data());
                data.insert(data.end(), bytes, bytes + value.size());
                offsets.push_back(static_cast<int32_t>(data.size()));
            }
        };

        template <int32_t Type>
        struct ColumnStorage;

        template <>
        struct ColumnStorage<XL_T_STRING>
        {
            using type = StringBuffer;
        };
        template <>
        struct ColumnStorage<XL_T_I64>
        {
            using type = std::vector<int64_t>;
        };
        template <>
        struct ColumnStorage<XL_T_F64>
        {
            using type = std::vector<double>;
        };
        template <>
        struct ColumnStorage<XL_T_BOOL>
        {
            using type = std::vector<uint8_t>;
        };
        template <>
        struct ColumnStorage<XL_T_DATE>
        {
            using type = std::vector<int32_t>;
        };
        template <>
        struct ColumnStorage<XL_T_TIME>
        {
            using type = std::vector<int64_t>;
        };
        template <>
        struct ColumnStorage<XL_T_TIMESTAMP>
        {
            using type = std::vector<int64_t>;
        };

        template <typename T>
        struct ColumnBuilder
        {
            using Field = unwrap_optional_t<T>;
            static constexpr bool nullable = is_optional_v<T>;
            static constexpr int32_t column_type = XlType<Field>::value;

            typename ColumnStorage<column_type>::type storage{};
            std::vector<uint8_t> validity{};
            int64_t rows = 0;

            void reserve(size_t count)
            {
                storage.reserve(count);
                if constexpr (nullable)
                {
                    validity.reserve((count + 7) / 8);
                }
            }

            void push(const T &value)
            {
                if constexpr (nullable)
                {
                    validity.resize(static_cast<size_t>((rows + 8) / 8), 0);
                    if (value.has_value())
                    {
                        validity[static_cast<size_t>(rows / 8)] |=
                            static_cast<uint8_t>(1u << static_cast<unsigned>(rows % 8));
                        append(*value);
                    }
                    else
                    {
                        append_placeholder();
                    }
                }
                else
                {
                    append(value);
                }
                ++rows;
            }

            bool overflowed() const
            {
                if constexpr (column_type == XL_T_STRING)
                {
                    return storage.overflowed;
                }
                else
                {
                    return false;
                }
            }

            ColumnRef to_ref(std::string_view name) const
            {
                if constexpr (column_type == XL_T_STRING)
                {
                    return string_column(name, storage.offsets, storage.data, validity);
                }
                else if constexpr (column_type == XL_T_I64)
                {
                    return i64_column(name, storage, validity);
                }
                else if constexpr (column_type == XL_T_F64)
                {
                    return f64_column(name, storage, validity);
                }
                else if constexpr (column_type == XL_T_BOOL)
                {
                    return bool_column(name, storage, validity);
                }
                else if constexpr (column_type == XL_T_DATE)
                {
                    return date_column(name, storage, validity);
                }
                else if constexpr (column_type == XL_T_TIME)
                {
                    return time_column(name, storage, validity);
                }
                else
                {
                    return timestamp_column(name, storage, validity);
                }
            }

        private:
            void append_placeholder()
            {
                if constexpr (column_type == XL_T_STRING)
                {
                    storage.push(std::string_view{});
                }
                else
                {
                    storage.push_back({});
                }
            }

            void append(const Field &value)
            {
                if constexpr (std::is_same_v<Field, std::string> || std::is_same_v<Field, std::string_view>)
                {
                    storage.push(std::string_view(value));
                }
                else if constexpr (std::is_same_v<Field, bool>)
                {
                    storage.push_back(static_cast<uint8_t>(value ? 1 : 0));
                }
                else if constexpr (std::is_integral_v<Field>)
                {
                    storage.push_back(static_cast<int64_t>(value));
                }
                else if constexpr (std::is_floating_point_v<Field>)
                {
                    storage.push_back(static_cast<double>(value));
                }
                else if constexpr (std::is_same_v<Field, std::chrono::sys_days>)
                {
                    storage.push_back(static_cast<int32_t>(value.time_since_epoch().count()));
                }
                else if constexpr (std::is_same_v<Field, std::chrono::year_month_day>)
                {
                    storage.push_back(
                        static_cast<int32_t>(std::chrono::sys_days{value}.time_since_epoch().count()));
                }
                else if constexpr (std::is_same_v<Field, std::chrono::microseconds>)
                {
                    storage.push_back(value.count());
                }
                else if constexpr (std::is_same_v<Field, std::chrono::hh_mm_ss<std::chrono::microseconds>>)
                {
                    storage.push_back(value.to_duration().count());
                }
                else if constexpr (std::is_same_v<Field, std::chrono::system_clock::time_point>)
                {
                    storage.push_back(std::chrono::time_point_cast<std::chrono::microseconds>(value)
                                          .time_since_epoch()
                                          .count());
                }
            }
        };

        template <typename Tuple, typename Indices>
        struct BuildersFor;
        template <typename Tuple, std::size_t... Is>
        struct BuildersFor<Tuple, std::index_sequence<Is...>>
        {
            using type = std::tuple<ColumnBuilder<typename std::tuple_element_t<Is, Tuple>::FieldType>...>;
        };

        template <typename Builders, typename T, typename Tuple, std::size_t... Is>
        void push_row(Builders &builders, const T &row, const Tuple &bindings, std::index_sequence<Is...>)
        {
            (..., std::get<Is>(builders).push(row.*(std::get<Is>(bindings).member)));
        }

        template <typename Builders, std::size_t... Is>
        void reserve_all(Builders &builders, size_t count, std::index_sequence<Is...>)
        {
            (..., std::get<Is>(builders).reserve(count));
        }

        template <typename Builders, std::size_t... Is>
        bool any_overflowed(const Builders &builders, std::index_sequence<Is...>)
        {
            return (... || std::get<Is>(builders).overflowed());
        }

        template <typename Builders, typename Tuple, std::size_t... Is>
        std::array<ColumnRef, sizeof...(Is)> to_refs(const Builders &builders, const Tuple &bindings,
                                                     std::index_sequence<Is...>)
        {
            return {std::get<Is>(builders).to_ref(std::string_view(std::get<Is>(bindings).column_names[0]))...};
        }

    } 

    /// @brief Writes rows as one sheet.
    /// @param path File to create.
    /// @param format As for write_columns.
    /// @param rows Input range of a row type with an ExcelMapper, the mapping parse_sheet reads with.
    /// An empty std::optional field writes a blank cell.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    template <std::ranges::input_range R>
    std::expected<void, Error> write_sheet(std::string_view path, int32_t format, R &&rows,
                                           const WriteOptions *options = nullptr)
    {
        using T = std::remove_cvref_t<std::ranges::range_value_t<R>>;
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t field_count = std::tuple_size_v<decltype(bindings)>;
        static constexpr auto indices = std::make_index_sequence<field_count>{};
        typename detail::BuildersFor<decltype(bindings), std::remove_cvref_t<decltype(indices)>>::type builders{};
        if constexpr (std::ranges::sized_range<R>)
        {
            detail::reserve_all(builders, static_cast<size_t>(std::ranges::size(rows)), indices);
        }
        for (const auto &row : rows)
        {
            detail::push_row(builders, row, bindings, indices);
        }

        if (detail::any_overflowed(builders, indices))
        {
            return std::unexpected(Error{XL_INVALID_ARGUMENT,
                                         "a string column exceeds 2 GiB, which int32 offsets cannot address."});
        }

        // refs borrows pointers into builders, so both must stay in the scope of the write call.
        const std::array<ColumnRef, field_count> refs = detail::to_refs(builders, bindings, indices);
        return write_columns(path, format, refs, options);
    }

    /// @brief write_sheet with the format taken from the file's extension.
    /// @param path File to create; its extension picks the format.
    /// @param rows As for the overload that takes a format.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    template <std::ranges::input_range R>
    std::expected<void, Error> write_sheet(std::string_view path, R &&rows,
                                           const WriteOptions *options = nullptr)
    {
        return write_sheet(path, format_from_path(path), std::forward<R>(rows), options);
    }

    /// @brief write_sheet into memory.
    /// @param format As for write_columns. There is no default: no file name to infer it from.
    /// @param rows As for write_sheet.
    /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
    /// @return The file's bytes.
    template <std::ranges::input_range R>
    std::expected<std::vector<uint8_t>, Error> write_sheet_to_memory(int32_t format, R &&rows,
                                                                     const WriteOptions *options = nullptr)
    {
        using T = std::remove_cvref_t<std::ranges::range_value_t<R>>;
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t field_count = std::tuple_size_v<decltype(bindings)>;
        static constexpr auto indices = std::make_index_sequence<field_count>{};
        typename detail::BuildersFor<decltype(bindings), std::remove_cvref_t<decltype(indices)>>::type builders{};
        if constexpr (std::ranges::sized_range<R>)
        {
            detail::reserve_all(builders, static_cast<size_t>(std::ranges::size(rows)), indices);
        }
        for (const auto &row : rows)
        {
            detail::push_row(builders, row, bindings, indices);
        }

        if (detail::any_overflowed(builders, indices))
        {
            return std::unexpected(Error{XL_INVALID_ARGUMENT,
                                         "a string column exceeds 2 GiB, which int32 offsets cannot address."});
        }

        // refs borrows pointers into builders, so both must stay in the scope of the write call.
        const std::array<ColumnRef, field_count> refs = detail::to_refs(builders, bindings, indices);
        return write_columns_to_memory(format, refs, options);
    }


    /// @brief Writes a workbook as calls arrive instead of building it in memory first.
    ///
    /// Per sheet: start_sheet, then for each row start_row, one write per cell from left to right, and
    /// end_row; then end_sheet. A call out of order returns an error rather than corrupting the
    /// output. Closed on destruction.
    class WriterHandle
    {
    public:
        WriterHandle(const WriterHandle &) = delete;
        WriterHandle &operator=(const WriterHandle &) = delete;

        WriterHandle(WriterHandle &&other) noexcept : handle_(std::exchange(other.handle_, nullptr)) {}
        WriterHandle &operator=(WriterHandle &&other) noexcept
        {
            if (this != &other)
            {
                close();
                handle_ = std::exchange(other.handle_, nullptr);
            }
            return *this;
        }

        ~WriterHandle() { close(); }

        /// @brief Creates a file to write into.
        /// @param path File to create.
        /// @param format An `XL_FORMAT_*` constant other than `XL_FORMAT_AUTO`.
        /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
        static std::expected<WriterHandle, Error> open(std::string_view path, int32_t format,
                                                       const WriteOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            xl_writer_handle *handle = nullptr;
            xl_write_options c_options{};
            const xl_write_options *c_options_ptr = detail::lower_options(options, c_options);
            int32_t status = xl_open_write_handle(reinterpret_cast<const uint8_t *>(path.data()),
                                                  static_cast<int32_t>(path.size()), format,
                                                  c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return WriterHandle(handle);
        }

        /// @brief Writes into memory instead of a file; read the result with bytes().
        /// @param format An `XL_FORMAT_*` constant other than `XL_FORMAT_AUTO`.
        /// @param options Sheet name, CSV dialect and format toggles, or null for the defaults.
        static std::expected<WriterHandle, Error> open_memory(int32_t format,
                                                              const WriteOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            xl_writer_handle *handle = nullptr;
            xl_write_options c_options{};
            const xl_write_options *c_options_ptr = detail::lower_options(options, c_options);
            int32_t status = xl_open_write_handle_to_memory(format, c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return WriterHandle(handle);
        }

        /// @brief Starts a sheet.
        /// @param name The sheet's name.
        std::expected<void, Error> start_sheet(std::string_view name)
        {
            return status_result(xl_start_sheet(handle_, reinterpret_cast<const uint8_t *>(name.data()),
                                                static_cast<int32_t>(name.size())));
        }

        /// @brief Starts a row.
        std::expected<void, Error> start_row() { return status_result(xl_start_row(handle_)); }

        /// @brief Ends the current row.
        std::expected<void, Error> end_row() { return status_result(xl_end_row(handle_)); }

        /// @brief Ends the current sheet.
        std::expected<void, Error> end_sheet() { return status_result(xl_end_sheet(handle_)); }

        /// @brief Writes the next cell.
        /// @param value Any type XlType covers; an empty std::optional writes a blank cell.
        template <typename T>
        std::expected<void, Error> write(const T &value)
        {
            using Value = std::remove_cvref_t<T>;
            if constexpr (detail::is_optional_v<Value>)
            {
                if (!value.has_value())
                {
                    return write_null(XlType<Value>::value);
                }
                return write(*value);
            }
            else if constexpr (std::is_same_v<Value, std::string> || std::is_same_v<Value, std::string_view>)
            {
                const std::string_view text(value);
                return status_result(xl_write_string(handle_, reinterpret_cast<const uint8_t *>(text.data()),
                                                      static_cast<int32_t>(text.size())));
            }
            else if constexpr (std::is_same_v<Value, bool>)
            {
                return status_result(xl_write_bool(handle_, value ? 1 : 0));
            }
            else if constexpr (std::is_integral_v<Value>)
            {
                return status_result(xl_write_int64(handle_, static_cast<int64_t>(value)));
            }
            else if constexpr (std::is_floating_point_v<Value>)
            {
                return status_result(xl_write_float64(handle_, static_cast<double>(value)));
            }
            else if constexpr (std::is_same_v<Value, std::chrono::sys_days>)
            {
                return status_result(
                    xl_write_date(handle_, static_cast<int32_t>(value.time_since_epoch().count())));
            }
            else if constexpr (std::is_same_v<Value, std::chrono::year_month_day>)
            {
                return status_result(xl_write_date(
                    handle_, static_cast<int32_t>(std::chrono::sys_days{value}.time_since_epoch().count())));
            }
            else if constexpr (std::is_same_v<Value, std::chrono::microseconds>)
            {
                return status_result(xl_write_time(handle_, value.count()));
            }
            else if constexpr (std::is_same_v<Value, std::chrono::hh_mm_ss<std::chrono::microseconds>>)
            {
                return status_result(xl_write_time(handle_, value.to_duration().count()));
            }
            else if constexpr (std::is_same_v<Value, std::chrono::system_clock::time_point>)
            {
                return status_result(xl_write_timestamp(
                    handle_,
                    std::chrono::time_point_cast<std::chrono::microseconds>(value).time_since_epoch().count()));
            }
            else
            {
                static_assert(sizeof(Value) == 0, "unsupported type for xl::WriterHandle::write");
            }
        }

        /// @brief Writes a blank cell.
        /// @param type The column's `XL_T_*` constant.
        std::expected<void, Error> write_null(int32_t type) { return status_result(xl_write_null(handle_, type)); }

        /// @brief Everything written so far, for a handle from open_memory.
        ///
        /// Finishes the workbook's trailing structure but leaves the handle open.
        /// @return The file's bytes.
        std::expected<std::vector<uint8_t>, Error> bytes()
        {
            xl_buffer buffer{};
            const int32_t status = xl_write_handle_bytes(handle_, &buffer);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            struct BufferGuard
            {
                xl_buffer *buffer;
                ~BufferGuard() { xl_free_buffer(buffer); }
            } guard{&buffer};
            return std::vector<uint8_t>(buffer.data, buffer.data + buffer.len);
        }

        /// @brief The native handle, for calling the C ABI directly.
        /// @return The handle, still owned by this object.
        xl_writer_handle *handle() const noexcept { return handle_; }

    private:
        explicit WriterHandle(xl_writer_handle *handle) noexcept : handle_(handle) {}

        static std::expected<void, Error> status_result(int32_t status)
        {
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return {};
        }

        void close() noexcept
        {
            if (handle_ != nullptr)
            {
                xl_close_write_handle(handle_);
                handle_ = nullptr;
            }
        }

        xl_writer_handle *handle_ = nullptr;
    };


    /// @brief A fold over CSV records.
    ///
    /// `Acc` must be move-constructible; the wrapper returns the surviving
    /// accumulator by moving out of it. `combine` must fold `next` into `*this` without destroying
    /// it - the library owns `next`'s lifetime and frees it afterwards.
    ///
    /// `accumulate` runs concurrently on worker threads, one `Acc` per partition, never on the
    /// calling thread. `combine` is effectively single-threaded but is not pinned to any particular
    /// thread and also never runs on the calling thread. `~Acc()` does run on the calling thread,
    /// once the run has finished, and must not throw - `std::move_constructible` subsumes
    /// `std::destructible`, which is `is_nothrow_destructible_v`, so a throwing destructor fails
    /// this concept rather than reaching the shims.
    ///
    /// The `RowView` handed to `accumulate` borrows the library's cell buffer and is valid only for
    /// the duration of that one call; copy anything you keep.
    ///
    /// All callbacks must be noexcept in effect: an exception escaping into the library is undefined
    /// behavior, so this wrapper catches everything and reports it as an error. See
    /// `excelreader.h`'s `xl_csv_aggregate_file` contract.
    /// @tparam Acc The accumulator type.
    template <typename Acc>
    concept CsvAccumulator =
        std::move_constructible<Acc> && requires(Acc acc, Acc &next, RowView row) {
            { acc.accumulate(row) } -> std::same_as<int32_t>;
            { acc.combine(next) } -> std::same_as<int32_t>;
        };

    namespace detail
    {
        /// Status a shim returns for "a callback threw". Positive, so it can never collide with a
        /// library status (every one of those is <= 0). It CAN collide with a caller's own
        /// `INT32_MAX`, so the entry points key on `threw` rather than on the status alone.
        inline constexpr int32_t kCsvCallbackException = (std::numeric_limits<int32_t>::max)();

        /// What the four shims reach through `user_data`. The C ABI shares one `user_data` across
        /// every worker thread and requires it to be read-only or internally synchronized, so the
        /// seed callable is held by pointer-to-const and the "a callback threw" flag is atomic.
        template <typename Acc, typename Seed>
        struct CsvContext
        {
            const Seed *seed;
            std::atomic<bool> threw{false};
        };

        template <typename Acc, typename Seed>
        int32_t csv_seed(void **out_state, void *user_data) noexcept
        {
            auto *context = static_cast<CsvContext<Acc, Seed> *>(user_data);
            *out_state = nullptr;
            try
            {
                *out_state = new Acc((*context->seed)());
                return XL_OK;
            }
            catch (...)
            {
                context->threw.store(true, std::memory_order_relaxed);
                return kCsvCallbackException;
            }
        }

        template <typename Acc, typename Seed>
        int32_t csv_accumulate(void *state, const xl_row *row, void *user_data) noexcept
        {
            auto *context = static_cast<CsvContext<Acc, Seed> *>(user_data);
            try
            {
                const size_t count = row->cell_count > 0 ? static_cast<size_t>(row->cell_count) : 0;
                return static_cast<Acc *>(state)->accumulate(RowView(row->cells, count));
            }
            catch (...)
            {
                context->threw.store(true, std::memory_order_relaxed);
                return kCsvCallbackException;
            }
        }

        template <typename Acc, typename Seed>
        int32_t csv_combine(void *acc, void *next, void *user_data) noexcept
        {
            auto *context = static_cast<CsvContext<Acc, Seed> *>(user_data);
            try
            {
                return static_cast<Acc *>(acc)->combine(*static_cast<Acc *>(next));
            }
            catch (...)
            {
                context->threw.store(true, std::memory_order_relaxed);
                return kCsvCallbackException;
            }
        }

        template <typename Acc, typename Seed>
        void csv_free_state(void *state, void *) noexcept
        {
            delete static_cast<Acc *>(state);
        }

        template <typename Acc, typename Seed>
        xl_csv_aggregation csv_aggregation(CsvContext<Acc, Seed> &context) noexcept
        {
            xl_csv_aggregation agg{};
            agg.struct_size = static_cast<int32_t>(sizeof(xl_csv_aggregation));
            agg.seed = &csv_seed<Acc, Seed>;
            agg.accumulate = &csv_accumulate<Acc, Seed>;
            agg.combine = &csv_combine<Acc, Seed>;
            agg.free_state = &csv_free_state<Acc, Seed>;
            agg.user_data = &context;
            return agg;
        }

        /// Turns the run's outcome into an `expected`, taking ownership of the surviving state.
        ///
        /// A recorded exception wins over the status on EVERY path, including `XL_OK`: an exception
        /// is a bug in the caller's accumulator, not a data condition, so whether it surfaces must
        /// not depend on how the library happened to chunk the source - a throw on a partition that
        /// was later re-read and discarded would otherwise vanish. Keying on the flag rather than on
        /// the status is also what keeps `kCsvCallbackException` distinguishable from a caller that
        /// legitimately returned `INT32_MAX`.
        template <typename Acc, typename Seed>
        std::expected<Acc, Error> csv_finish(const CsvContext<Acc, Seed> &context, int32_t status, void *state)
        {
            const bool threw = context.threw.load(std::memory_order_relaxed);

            std::unique_ptr<Acc> owned(static_cast<Acc *>(state));

            if (threw)
            {
                if (status < XL_OK)
                {
                    return std::unexpected(
                        Error{status, "an accumulator callback threw: " + make_error(status).message});
                }
                if (status == XL_OK)
                {
                    return std::unexpected(Error{kCsvCallbackException, "an accumulator callback threw"});
                }
                return std::unexpected(Error{status, "an accumulator callback threw"});
            }
            if (status != XL_OK)
            {
                if (status > 0)
                {
                    return std::unexpected(
                        Error{status, "an aggregation callback returned status " + std::to_string(status)});
                }
                return std::unexpected(make_error(status));
            }
            if (owned == nullptr)
            {
                return std::unexpected(
                    Error{XL_ERROR, "native reported success but wrote no aggregation state"});
            }
            return std::move(*owned);
        }
    }

    /// @brief Folds the CSV file at `path` into one `Acc` across several threads.
    ///
    /// `seed` is called at least once per partition, on worker threads, and may be called again for
    /// a partition that has to be re-read - a seeded accumulator can therefore be destroyed without
    /// ever being combined. The call blocks until the run finishes, so `seed` may capture locals
    /// freely. It must be const-callable: the C ABI shares one `user_data` across every worker and
    /// requires it to be read-only or internally synchronized, so the shims invoke it through a
    /// reference to const. A capturing `mutable` lambda therefore fails this constraint cleanly
    /// instead of hard-erroring inside a shim. The `remove_cvref_t` is load-bearing - `Seed` deduces
    /// to an lvalue reference for a named callable, and `const Seed &` would collapse the `const`
    /// away.
    /// @tparam Acc The accumulator; see CsvAccumulator.
    /// @param path CSV file to read.
    /// @param seed Makes a fresh `Acc` for a partition.
    /// @param options Threads, header row and CSV dialect, or null for the defaults.
    /// @return The combined accumulator.
    template <CsvAccumulator Acc, typename Seed>
        requires std::invocable<const std::remove_cvref_t<Seed> &>
    std::expected<Acc, Error> aggregate_csv_file(std::string_view path, Seed &&seed,
                                                 const CsvParallelOptions *options = nullptr)
    {
        if (const auto &abi = detail::check_abi_version(); !abi.has_value())
        {
            return std::unexpected(abi.error());
        }

        using SeedType = std::remove_cvref_t<Seed>;
        detail::CsvContext<Acc, SeedType> context{std::addressof(seed), {}};
        const xl_csv_aggregation agg = detail::csv_aggregation<Acc, SeedType>(context);
        xl_csv_parallel_options c_options{};
        const xl_csv_parallel_options *c_options_ptr = detail::lower_options(options, c_options);
        void *state = nullptr;
        const int32_t status = xl_csv_aggregate_file(
            reinterpret_cast<const uint8_t *>(path.data()), static_cast<int32_t>(path.size()),
            &agg, c_options_ptr, &state);
        return detail::csv_finish<Acc, SeedType>(context, status, state);
    }

    /// @brief Folds the CSV file at `path` into one `Acc`, seeding each partition with `Acc()`.
    /// @tparam Acc The accumulator; see CsvAccumulator.
    /// @param path CSV file to read.
    /// @param options Threads, header row and CSV dialect, or null for the defaults.
    /// @return The combined accumulator.
    template <CsvAccumulator Acc>
        requires std::default_initializable<Acc>
    std::expected<Acc, Error> aggregate_csv_file(std::string_view path,
                                                 const CsvParallelOptions *options = nullptr)
    {
        auto seed = [] { return Acc(); };
        return aggregate_csv_file<Acc>(path, seed, options);
    }

    /// @brief Folds an in-memory CSV buffer into one `Acc`.
    ///
    /// Otherwise identical to aggregate_csv_file.
    /// @tparam Acc The accumulator; see CsvAccumulator.
    /// @param data The CSV bytes. Not copied, so they must stay valid for the call.
    /// @param seed As for aggregate_csv_file.
    /// @param options Threads, header row and CSV dialect, or null for the defaults.
    /// @return The combined accumulator.
    template <CsvAccumulator Acc, typename Seed>
        requires std::invocable<const std::remove_cvref_t<Seed> &>
    std::expected<Acc, Error> aggregate_csv_memory(std::span<const uint8_t> data, Seed &&seed,
                                                   const CsvParallelOptions *options = nullptr)
    {
        if (const auto &abi = detail::check_abi_version(); !abi.has_value())
        {
            return std::unexpected(abi.error());
        }

        using SeedType = std::remove_cvref_t<Seed>;
        detail::CsvContext<Acc, SeedType> context{std::addressof(seed), {}};
        const xl_csv_aggregation agg = detail::csv_aggregation<Acc, SeedType>(context);
        xl_csv_parallel_options c_options{};
        const xl_csv_parallel_options *c_options_ptr = detail::lower_options(options, c_options);
        void *state = nullptr;
        const int32_t status = xl_csv_aggregate_memory(
            data.data(), static_cast<int32_t>(data.size()), &agg, c_options_ptr, &state);
        return detail::csv_finish<Acc, SeedType>(context, status, state);
    }

    /// @brief Folds an in-memory CSV buffer into one `Acc`, seeding each partition with `Acc()`.
    /// @tparam Acc The accumulator; see CsvAccumulator.
    /// @param data The CSV bytes. Not copied, so they must stay valid for the call.
    /// @param options Threads, header row and CSV dialect, or null for the defaults.
    /// @return The combined accumulator.
    template <CsvAccumulator Acc>
        requires std::default_initializable<Acc>
    std::expected<Acc, Error> aggregate_csv_memory(std::span<const uint8_t> data,
                                                   const CsvParallelOptions *options = nullptr)
    {
        auto seed = [] { return Acc(); };
        return aggregate_csv_memory<Acc>(data, seed, options);
    }
}
