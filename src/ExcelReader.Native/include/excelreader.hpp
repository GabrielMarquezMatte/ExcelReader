#pragma once

#include "excelreader.h"

#include <array>
#include <compare>
#include <concepts>
#include <cstdint>
#include <cstring>
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


    struct Error
    {
        int32_t code;
        std::string message;
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


    inline int32_t abi_version() noexcept { return xl_abi_version(); }

    namespace detail
    {

        inline const std::expected<void, Error> &check_abi_version()
        {
            static const std::expected<void, Error> result = []() -> std::expected<void, Error>
            {
                const int32_t loaded = xl_abi_version();
                if (loaded == XL_ABI_VERSION)
                {
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


    struct InferredColumn
    {
        std::optional<std::string> name;
        int32_t index = 0;
        int32_t type = XL_T_STRING; 
        bool nullable = false;
    };


    struct OpenOptions
    {
        int32_t csv_sniff_dialect = XL_OPT_DEFAULT;
        int32_t csv_delimiter = 0;
        int32_t csv_quote = 0;
        int32_t csv_detect_bom = XL_OPT_DEFAULT;
        int32_t csv_max_cell_bytes = 0;
        int32_t csv_intern_strings = XL_OPT_DEFAULT;

        int64_t max_total_decompressed_bytes = 0;
        int32_t max_cell_bytes = 0;
        int64_t max_shared_string_bytes = 0;
        int32_t max_zip_entries = 0;
        int32_t prefetch_decompression = XL_OPT_DEFAULT;
        int32_t intern_strings = XL_OPT_DEFAULT;

        std::string password_{};

        OpenOptions &password(std::string_view value)
        {
            password_ = std::string(value);
            return *this;
        }

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
            opts.password = password_.empty() ? nullptr : reinterpret_cast<const uint8_t *>(password_.data());
            opts.password_len = static_cast<int32_t>(password_.size());
            return opts;
        }
    };


    struct WriteOptions
    {
        std::string_view sheet_name{}; 

        int32_t csv_delimiter = 0;
        int32_t csv_quote = 0;

        int32_t date1904 = XL_OPT_DEFAULT;           
        int32_t use_shared_strings = XL_OPT_DEFAULT; 

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

    struct CsvParallelOptions
    {
        int32_t degree_of_parallelism = 0;
        int32_t header_row = 0;
        int32_t delimiter = 0;
        int32_t quote = 0;
        int32_t detect_bom = XL_OPT_DEFAULT;
        int32_t max_cell_bytes = 0;

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

    inline constexpr ColumnRef i64_column(std::string_view name, std::span<const int64_t> values,
                                          std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_I64>(name, values, validity);
    }

    inline constexpr ColumnRef f64_column(std::string_view name, std::span<const double> values,
                                          std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_F64>(name, values, validity);
    }

    inline constexpr ColumnRef bool_column(std::string_view name, std::span<const uint8_t> values,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_BOOL>(name, values, validity);
    }

    inline constexpr ColumnRef date_column(std::string_view name, std::span<const int32_t> days_since_epoch,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_DATE>(name, days_since_epoch, validity);
    }

    inline constexpr ColumnRef time_column(std::string_view name, std::span<const int64_t> micros_since_midnight,
                                           std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_TIME>(name, micros_since_midnight, validity);
    }

    inline constexpr ColumnRef timestamp_column(std::string_view name, std::span<const int64_t> micros_since_epoch,
                                                std::span<const uint8_t> validity = {}) noexcept
    {
        return detail::scalar_column<XL_T_TIMESTAMP>(name, micros_since_epoch, validity);
    }

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


    enum class CellType : int32_t
    {
        Empty = XL_CELL_EMPTY,
        String = XL_CELL_STRING,
        Number = XL_CELL_NUMBER,
        Date = XL_CELL_DATE,
        Bool = XL_CELL_BOOL,
        Formula = XL_CELL_FORMULA,
        Error = XL_CELL_ERROR,
    };

    struct CellView
    {
        int32_t column{};
        CellType type{};
        std::string_view value{};
    };

    namespace detail
    {
        inline std::optional<std::pair<CellView, size_t>> decode_cell(std::span<const uint8_t> blob, size_t offset)
        {
            const auto read_i32 = [&](size_t at) -> std::optional<int32_t> {
                if (at + 4 > blob.size())
                {
                    return std::nullopt;
                }
                int32_t value = 0;
                std::memcpy(&value, blob.data() + at, sizeof(value));
                return value;
            };

            const auto column = read_i32(offset);
            const auto type = read_i32(offset + 4);
            const auto length = read_i32(offset + 8);
            if (!column || !type || !length || *length < 0)
            {
                return std::nullopt;
            }
            const size_t start = offset + 12;
            const size_t end = start + static_cast<size_t>(*length);
            if (end > blob.size())
            {
                return std::nullopt;
            }
            CellView cell{*column, static_cast<CellType>(*type),
                          std::string_view(reinterpret_cast<const char *>(blob.data() + start),
                                           static_cast<size_t>(*length))};
            return std::make_pair(cell, end);
        }
    }

    class RowView
    {
    public:
        RowView() = default;

        RowView(std::span<const uint8_t> payload, size_t count) : payload_(payload), count_(count) {}

        RowView(const xl_row_cell *cells, size_t count) : cells_(cells), count_(count) {}

        size_t size() const { return count_; }
        bool empty() const { return count_ == 0; }

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
            iterator(const RowView *row, size_t index) : row_(row), index_(index) { load(); }

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
                auto decoded = detail::decode_cell(row_->payload_, offset_);
                if (!decoded)
                {
                    index_ = row_->count_;   
                    return;
                }
                current_ = decoded->first;
                offset_ = decoded->second;
            }

            const RowView *row_{};
            size_t index_{};
            size_t offset_{};
            CellView current_{};
        };

        iterator begin() const { return iterator(this, 0); }
        iterator end() const { return iterator(this, count_); }

    private:
        std::span<const uint8_t> payload_{};
        const xl_row_cell *cells_{};
        size_t count_{};
    };

    class RowCursor
    {
    public:
        explicit RowCursor(xl_workbook *handle) : handle_(handle), buffer_(kInitialRowBuffer) {}

        RowCursor(const RowCursor &) = delete;
        RowCursor &operator=(const RowCursor &) = delete;
        RowCursor(RowCursor &&) noexcept = default;
        RowCursor &operator=(RowCursor &&) noexcept = default;

        std::expected<RowView, Error> next_row()
        {
            while (true)
            {
                int32_t written = 0;
                const auto capacity = static_cast<int32_t>(buffer_.size());
                const int32_t status = xl_next_row(handle_, buffer_.data(), capacity, &written);

                if (status == XL_OK)
                {
                    const size_t length = written > 0 ? static_cast<size_t>(written) : 0;
                    if (length < 4)
                    {
                        return std::unexpected(detail::make_error(XL_ERROR));
                    }
                    int32_t count = 0;
                    std::memcpy(&count, buffer_.data(), sizeof(count));
                    if (count < 0)
                    {
                        return std::unexpected(detail::make_error(XL_ERROR));
                    }
                    return RowView(std::span<const uint8_t>(buffer_.data() + 4, length - 4),
                                   static_cast<size_t>(count));
                }
                if (status == XL_BUFFER_TOO_SMALL)
                {
                    const size_t needed = written > 0 ? static_cast<size_t>(written) : buffer_.size() * 2;
                    if (needed <= buffer_.size())
                    {
                        return std::unexpected(detail::make_error(XL_ERROR));
                    }
                    buffer_.resize(needed);
                    continue;
                }
                return std::unexpected(detail::make_error(status));   
            }
        }

    private:
        static constexpr size_t kInitialRowBuffer = 64 * 1024;

        xl_workbook *handle_{};
        std::vector<uint8_t> buffer_;
    };

    class DecodedRows
    {
    public:
        explicit DecodedRows(xl_rows raw) : raw_(raw) {}

        DecodedRows(const DecodedRows &) = delete;
        DecodedRows &operator=(const DecodedRows &) = delete;

        DecodedRows(DecodedRows &&other) noexcept : raw_(std::exchange(other.raw_, xl_rows{})) {}
        DecodedRows &operator=(DecodedRows &&other) noexcept
        {
            if (this != &other)
            {
                release();
                raw_ = std::exchange(other.raw_, xl_rows{});
            }
            return *this;
        }

        ~DecodedRows() { release(); }

        size_t size() const { return raw_.row_count > 0 ? static_cast<size_t>(raw_.row_count) : 0; }
        bool empty() const { return size() == 0; }

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
            raw_ = xl_rows{};
        }

        xl_rows raw_{};
    };


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
            int32_t status = xl_open_file_ex(reinterpret_cast<const uint8_t *>(path.data()),
                                             static_cast<int32_t>(path.size()), format,
                                             c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

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
            int32_t status = xl_open_memory_ex(data.data(), static_cast<int32_t>(data.size()), format,
                                               c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }


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

        std::expected<std::string, Error> sheet_name() const
        {
            return detail::fill_string([this](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name(handle_, buffer, capacity, out_len); });
        }

        std::expected<std::string, Error> sheet_name_at(int32_t index) const
        {
            return detail::fill_string([this, index](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name_at(handle_, index, buffer, capacity, out_len); });
        }

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

        std::expected<void, Error> move_to_sheet(int32_t index)
        {
            int32_t status = xl_move_to_sheet(handle_, index);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return {};
        }

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

        RowCursor rows() { return RowCursor(handle_); }

        std::expected<DecodedRows, Error> read_all_decoded()
        {
            xl_rows raw{};
            const int32_t status = xl_read_all_decoded(handle_, &raw);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return DecodedRows(raw);
        }


        std::expected<std::vector<InferredColumn>, Error> infer_schema(int32_t header_row = 1,
                                                                       int32_t sample_size = 100) const
        {
            xl_inferred_schema schema{};
            int32_t status = xl_infer_schema(handle_, header_row, sample_size, &schema);
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


    template <typename Class, typename T, std::size_t N = 1>
    struct FieldBinding
    {
        std::array<const char *, N> column_names;
        T Class::*member;
        using FieldType = T;
    };

    template <typename Class, typename T>
    constexpr FieldBinding<Class, T> make_field(const char *name, T Class::*member)
    {
        return {{name}, member};
    }

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

        T operator[](int64_t row) const { return detail::row_at<T>(table_, row); }

        std::optional<T> at(int64_t row) const
        {
            if (row < 0 || row >= table_.row_count)
            {
                return std::nullopt;
            }
            return (*this)[row];
        }

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

        static TableView from_raw(xl_table table) { return TableView(table); }

    private:
        explicit TableView(xl_table table) noexcept : table_(table) {}

        xl_table table_{};
    };


    template <typename T>
    std::expected<TableView<T>, Error> parse_sheet(Workbook &workbook, int32_t header_row = 1)
    {
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);
        std::span<const xl_column_spec> specs(specs_array);

        xl_table table{};
        int32_t status = xl_parse_typed(workbook.handle(), specs.data(), static_cast<int32_t>(specs.size()), header_row, &table);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return TableView<T>::from_raw(table);
    }


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

        static TypedReader from_raw(xl_typed_reader *reader) { return TypedReader(reader); }

    private:
        explicit TypedReader(xl_typed_reader *reader) noexcept : reader_(reader) {}

        xl_typed_reader *reader_ = nullptr;
    };

    template <typename T>
    std::expected<TypedReader<T>, Error> typed_reader(Workbook &workbook, int32_t header_row = 1,
                                                      int64_t batch_size = 10000)
    {
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);

        xl_typed_reader *reader = nullptr;
        int32_t status = xl_typed_reader_open(workbook.handle(), specs_array.data(),
                                              static_cast<int32_t>(specs_array.size()), header_row,
                                              batch_size, &reader);
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return TypedReader<T>::from_raw(reader);
    }

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

    inline std::expected<void, Error> write_columns(std::string_view path, std::span<const ColumnRef> columns,
                                                    const WriteOptions *options = nullptr)
    {
        return write_columns(path, format_from_path(path), columns, options);
    }

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

    // NOTE for write_sheet_to_memory below: this body is intentionally NOT factored into a shared
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

        const std::array<ColumnRef, field_count> refs = detail::to_refs(builders, bindings, indices);
        return write_columns(path, format, refs, options);
    }

    template <std::ranges::input_range R>
    std::expected<void, Error> write_sheet(std::string_view path, R &&rows,
                                           const WriteOptions *options = nullptr)
    {
        return write_sheet(path, format_from_path(path), std::forward<R>(rows), options);
    }

    // In-memory equivalent of write_sheet: same transpose (see the NOTE on write_sheet above for why
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

        const std::array<ColumnRef, field_count> refs = detail::to_refs(builders, bindings, indices);
        return write_columns_to_memory(format, refs, options);
    }


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

        std::expected<void, Error> start_sheet(std::string_view name)
        {
            return status_result(xl_start_sheet(handle_, reinterpret_cast<const uint8_t *>(name.data()),
                                                static_cast<int32_t>(name.size())));
        }

        std::expected<void, Error> start_row() { return status_result(xl_start_row(handle_)); }

        std::expected<void, Error> end_row() { return status_result(xl_end_row(handle_)); }

        std::expected<void, Error> end_sheet() { return status_result(xl_end_sheet(handle_)); }

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

        std::expected<void, Error> write_null(int32_t type) { return status_result(xl_write_null(handle_, type)); }

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


    /// A fold over CSV records. `Acc` must be default-constructible and move-constructible; `seed`
    /// builds a fresh one with `Acc()` and the wrapper returns it by moving out of the surviving
    /// state. `combine` must fold `next` into `*this` without destroying it - the library owns
    /// `next`'s lifetime and frees it afterwards.
    ///
    /// `accumulate` runs concurrently on worker threads, one `Acc` per partition, never on the
    /// calling thread. `combine` is effectively single-threaded but is not pinned to any particular
    /// thread, per `excelreader.h`'s `xl_csv_aggregate_file` contract. Both must be noexcept in
    /// effect: an exception escaping into the library is undefined behavior, so this wrapper catches
    /// everything and reports a reserved sentinel status that this header maps back to a descriptive
    /// error instead of `XL_ERROR`.
    template <typename Acc>
    concept CsvAccumulator =
        std::default_initializable<Acc> && std::move_constructible<Acc> &&
        requires(Acc acc, Acc &next, const xl_row *row) {
            { acc.accumulate(row) } -> std::same_as<int32_t>;
            { acc.combine(next) } -> std::same_as<int32_t>;
        };

    namespace detail
    {
        /// Reserved status a shim below returns to signal "a user callback threw", never a real
        /// accumulate/combine/seed result. Positive, so it can never collide with a library status
        /// (every one of those is <= 0) or a caller's own error code space in the same way `XL_ERROR`
        /// would - `aggregate_csv_file`/`aggregate_csv_memory` recognize it and report a descriptive
        /// error instead of the generic "unknown error" that `XL_ERROR` would produce here, since
        /// `xl_last_error` is never populated for a callback exception. Mirrors the Rust wrapper's
        /// `PANIC_STATUS`.
        inline constexpr int32_t kCsvCallbackException = (std::numeric_limits<int32_t>::max)();

        template <typename Acc>
        int32_t csv_seed(void **out_state, void *) noexcept
        {
            try
            {
                *out_state = new Acc();
                return XL_OK;
            }
            catch (...)
            {
                return kCsvCallbackException;
            }
        }

        template <typename Acc>
        int32_t csv_accumulate(void *state, const xl_row *row, void *) noexcept
        {
            try
            {
                return static_cast<Acc *>(state)->accumulate(row);
            }
            catch (...)
            {
                return kCsvCallbackException;
            }
        }

        template <typename Acc>
        int32_t csv_combine(void *acc, void *next, void *) noexcept
        {
            try
            {
                return static_cast<Acc *>(acc)->combine(*static_cast<Acc *>(next));
            }
            catch (...)
            {
                return kCsvCallbackException;
            }
        }

        template <typename Acc>
        void csv_free_state(void *state, void *) noexcept
        {
            delete static_cast<Acc *>(state);
        }

        template <typename Acc>
        xl_csv_aggregation csv_aggregation() noexcept
        {
            xl_csv_aggregation agg{};
            agg.struct_size = static_cast<int32_t>(sizeof(xl_csv_aggregation));
            agg.seed = &csv_seed<Acc>;
            agg.accumulate = &csv_accumulate<Acc>;
            agg.combine = &csv_combine<Acc>;
            agg.free_state = &csv_free_state<Acc>;
            agg.user_data = nullptr;
            return agg;
        }
    }

    template <CsvAccumulator Acc>
    std::expected<Acc, Error> aggregate_csv_file(std::string_view path,
                                                 const CsvParallelOptions *options = nullptr)
    {
        if (const auto &abi = detail::check_abi_version(); !abi.has_value())
        {
            return std::unexpected(abi.error());
        }

        const xl_csv_aggregation agg = detail::csv_aggregation<Acc>();
        xl_csv_parallel_options c_options{};
        const xl_csv_parallel_options *c_options_ptr = detail::lower_options(options, c_options);
        void *state = nullptr;
        const int32_t status = xl_csv_aggregate_file(
            reinterpret_cast<const uint8_t *>(path.data()), static_cast<int32_t>(path.size()),
            &agg, c_options_ptr, &state);
        if (status == detail::kCsvCallbackException)
        {
            return std::unexpected(Error{status, "an accumulator callback threw"});
        }
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }

        std::unique_ptr<Acc> owned(static_cast<Acc *>(state));
        return std::move(*owned);
    }

    template <CsvAccumulator Acc>
    std::expected<Acc, Error> aggregate_csv_memory(std::span<const uint8_t> data,
                                                   const CsvParallelOptions *options = nullptr)
    {
        if (const auto &abi = detail::check_abi_version(); !abi.has_value())
        {
            return std::unexpected(abi.error());
        }

        const xl_csv_aggregation agg = detail::csv_aggregation<Acc>();
        xl_csv_parallel_options c_options{};
        const xl_csv_parallel_options *c_options_ptr = detail::lower_options(options, c_options);
        void *state = nullptr;
        const int32_t status = xl_csv_aggregate_memory(data.data(), static_cast<int32_t>(data.size()),
                                                        &agg, c_options_ptr, &state);
        if (status == detail::kCsvCallbackException)
        {
            return std::unexpected(Error{status, "an accumulator callback threw"});
        }
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }

        std::unique_ptr<Acc> owned(static_cast<Acc *>(state));
        return std::move(*owned);
    }
}
