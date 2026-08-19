/* Header-only C++ wrapper around excelreader.h (the C ABI).
 *
 * Scope of this first pass: opening a workbook and schema-driven typed table parsing
 * (xl_parse_typed) only - no writing, no row-by-row decoded reads.
 *
 * Design constraints, matching the native library's own perf/memory posture:
 *   - No exceptions anywhere in this header. Every fallible operation returns
 *     std::expected<T, xl::Error>.
 *   - xl::parse_sheet<T> does NOT materialize a std::vector<T>. It returns a
 *     xl::TableView<T> that owns the raw xl_table (freed via xl_free_table) and builds
 *     one T per dereference of its iterator - the only allocation is the one
 *     xl_parse_typed itself already makes for the columnar buffers. Call
 *     TableView<T>::to_vector() (or the parse_sheet_vector<T> convenience) if you
 *     actually want a materialized vector.
 *   - std::string_view fields are zero-copy views into the xl_table's own string blob:
 *     valid ONLY as long as the owning TableView<T> is alive. Use std::string for a
 *     field that needs to outlive the view (e.g. after to_vector()).
 */
#pragma once

#include "excelreader.h"

#include <array>
#include <compare>
#include <cstdint>
#include <cstring>
#include <expected>
#include <iterator>
#include <optional>
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

    // ---- Errors --------------------------------------------------------------------------------

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

        // Shared two-pass buffer dance for the xl_* functions that write a UTF-8 name into a caller
        // buffer and report the required capacity through XL_BUFFER_TOO_SMALL.
        template <typename Call>
        std::expected<std::string, Error> fill_string(Call &&call)
        {
            // One sized attempt first: Excel caps sheet names at 31 characters, so 128 bytes clears
            // even the 4-byte-per-character worst case and the retry never runs in practice.
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

    } // namespace detail

    // ---- ABI guard -------------------------------------------------------------------------------

    // Revision of the loaded shared library, to compare against XL_ABI_VERSION - the revision this
    // header was compiled against.
    inline int32_t abi_version() noexcept { return xl_abi_version(); }

    namespace detail
    {

        // The native binary is resolved at build time from a GitHub release asset (see
        // cmake/FetchNativeLib.cmake) or from EXCELREADER_NATIVE_LIB, so it can easily be built from
        // a different ABI revision than this header. Every struct below is laid out against
        // XL_ABI_VERSION, so proceeding past a mismatch would mean reading native memory through the
        // wrong layout. Both Workbook constructors gate on this.
        //
        // The result is cached in a function-local static: it cannot change for the lifetime of the
        // process, its initialization is thread-safe since C++11, and every open() would otherwise
        // pay an FFI call for it.
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

    } // namespace detail

    // ---- Inferred schema ---------------------------------------------------------------------------

    // One column guessed by Workbook::infer_schema.
    struct InferredColumn
    {
        // Header text, or nullopt when the column must be resolved by `index` instead (no header
        // row was requested, the header cell was blank, or the column never appeared in it).
        std::optional<std::string> name;
        int32_t index = 0;
        int32_t type = XL_T_STRING; // XL_T_*
        bool nullable = false;
    };

    // ---- Open options ----------------------------------------------------------------------------

    // C++ mirror of xl_open_options: same fields, same meaning (XL_OPT_DEFAULT/FALSE/TRUE for every
    // tri-state field, 0 for "use the library default" on every numeric field), but with default
    // member initializers so a caller only sets the fields they actually want to override - no
    // memset, no struct_size bookkeeping (to_c() fills it in).
    struct OpenOptions
    {
        // CSV only (format == XL_FORMAT_CSV); ignored for every other format.
        int32_t csv_sniff_dialect = XL_OPT_DEFAULT;
        int32_t csv_delimiter = 0;
        int32_t csv_quote = 0;
        int32_t csv_detect_bom = XL_OPT_DEFAULT;
        int32_t csv_max_cell_bytes = 0;
        int32_t csv_intern_strings = XL_OPT_DEFAULT;

        // XLS/XLSX/XLSB only; ignored for CSV.
        int64_t max_total_decompressed_bytes = 0;
        int32_t max_cell_bytes = 0;
        int64_t max_shared_string_bytes = 0;
        int32_t max_zip_entries = 0;
        int32_t prefetch_decompression = XL_OPT_DEFAULT;
        int32_t intern_strings = XL_OPT_DEFAULT;

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
            return opts;
        }
    };

    // ---- Workbook (RAII) ------------------------------------------------------------------------

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
            const xl_open_options *c_options_ptr = nullptr;
            if (options != nullptr)
            {
                c_options = options->to_c();
                c_options_ptr = &c_options;
            }
            int32_t status = xl_open_file_ex(reinterpret_cast<const uint8_t *>(path.data()),
                                             static_cast<int32_t>(path.size()), format,
                                             c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

        // In-memory equivalent of open(): `data` is copied by the native library, so it need not
        // outlive this call.
        static std::expected<Workbook, Error> open_memory(std::span<const uint8_t> data, int32_t format = XL_FORMAT_AUTO,
                                                          const OpenOptions *options = nullptr)
        {
            if (const auto &abi = detail::check_abi_version(); !abi.has_value())
            {
                return std::unexpected(abi.error());
            }
            xl_workbook *handle = nullptr;
            xl_open_options c_options{};
            const xl_open_options *c_options_ptr = nullptr;
            if (options != nullptr)
            {
                c_options = options->to_c();
                c_options_ptr = &c_options;
            }
            int32_t status = xl_open_memory_ex(data.data(), static_cast<int32_t>(data.size()), format,
                                               c_options_ptr, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
        }

        // ---- Sheet navigation --------------------------------------------------------------------

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

        // Name of the currently selected sheet.
        std::expected<std::string, Error> sheet_name() const
        {
            return detail::fill_string([this](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name(handle_, buffer, capacity, out_len); });
        }

        // Name of the sheet at `index`, without changing the current sheet or disturbing row
        // enumeration.
        std::expected<std::string, Error> sheet_name_at(int32_t index) const
        {
            return detail::fill_string([this, index](uint8_t *buffer, int32_t capacity, int32_t *out_len)
                                       { return xl_sheet_name_at(handle_, index, buffer, capacity, out_len); });
        }

        // Every sheet name, in workbook order.
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

        // Selects the sheet at `index`, resetting row enumeration to its first row. Non-const: it
        // moves the cursor every subsequent read on this handle shares.
        std::expected<void, Error> move_to_sheet(int32_t index)
        {
            int32_t status = xl_move_to_sheet(handle_, index);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return {};
        }

        // Whether the workbook uses the 1904 date system - needed to interpret raw Excel serials.
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

        // ---- Schema inference --------------------------------------------------------------------

        // Guesses a parse_sheet schema by sampling the current sheet. `header_row` has the same
        // meaning as in parse_sheet (0 = no header); `sample_size` bounds how many rows after the
        // header are inspected. A guess over a sample, not a guarantee - always check it fits before
        // trusting it against the full sheet. Const: the native call samples independently of the
        // shared row cursor and never disturbs it.
        std::expected<std::vector<InferredColumn>, Error> infer_schema(int32_t header_row = 1,
                                                                       int32_t sample_size = 100) const
        {
            xl_inferred_schema schema{};
            int32_t status = xl_infer_schema(handle_, header_row, sample_size, &schema);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }

            // The schema is native-owned from here. This guard returns it on every exit path,
            // including the one where the vector's own allocation throws - the only way out of this
            // function that is not a plain return.
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
                // A guessed name is exactly name_len bytes with no NUL terminator, and is NULL
                // whenever the column had no usable header cell.
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

    // ---- Type traits: map a C++ field type to its XL_T_* column type -----------------------------

    template <typename T>
    struct XlType; // no default: a field type not specialized below is a compile error, not a silent bug.

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

    // XL_T_TIME's native width is microseconds since midnight (see excelreader.h) - std::chrono::
    // microseconds is the primary field type for it, matching that exactly with no conversion.
    // hh_mm_ss<microseconds> is also supported, for callers who want hours/minutes/seconds broken
    // out rather than a raw duration; its precision must match XL_T_TIME's for the same reason.
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

    // ---- Struct <-> column bindings ---------------------------------------------------------------

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

    // Users specialize this for each struct they want to parse into.
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
                0, // index is ignored: resolved by name
                XlType<T>::value,
                1 // nullable = 1 (safe default)
            };
        }
        template <typename Tuple, std::size_t... Is>
        std::array<xl_column_spec, sizeof...(Is)> build_specs(const Tuple &bindings, std::index_sequence<Is...>, std::array<std::vector<int32_t>, sizeof...(Is)> &name_lens_storage)
        {
            return {build_one_spec(std::get<Is>(bindings), name_lens_storage[Is])...};
        }

        // Whether row `row` is non-null in `col`; columns with no null values have validity == nullptr.
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
                return; // leave the struct member default-initialized
            }

            if constexpr (std::is_same_v<T, std::string> || std::is_same_v<T, std::string_view>)
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
            else if constexpr (std::is_integral_v<T>)
            {
                instance.*(binding.member) = T(static_cast<const int64_t *>(col.values)[row]);
            }
            else if constexpr (std::is_floating_point_v<T>)
            {
                instance.*(binding.member) = T(static_cast<const double *>(col.values)[row]);
            }
            else if constexpr (std::is_same_v<T, bool>)
            {
                instance.*(binding.member) = (static_cast<const uint8_t *>(col.values)[row] != 0);
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
                // system_clock::time_point's own Duration is implementation-defined (nanoseconds on
                // libstdc++/MSVC) - time_point_cast converts the microseconds XL_T_TIMESTAMP provides
                // into whatever that is. sys_time<microseconds> is the time_point<system_clock,
                // microseconds> alias; constructing through it (rather than time_point's raw Duration
                // constructor) keeps the "microseconds since epoch" meaning explicit at the call site.
                int64_t micros = static_cast<const int64_t *>(col.values)[row];
                instance.*(binding.member) = std::chrono::time_point_cast<std::chrono::system_clock::duration>(
                    std::chrono::sys_time<std::chrono::microseconds>{std::chrono::microseconds{micros}});
            }
        }

        template <typename T, typename Tuple, std::size_t... Is>
        void populate_instance(T &instance, const xl_table &table, int64_t row, const Tuple &bindings, std::index_sequence<Is...>)
        {
            (..., assign_field(instance, table.columns[Is], row, std::get<Is>(bindings)));
        }

    } // namespace detail

    // ---- TableView<T>: a lazy, non-owning-of-T view over a parsed xl_table ------------------------

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

        ~TableView() { xl_free_table(&table_); } // safe on a zeroed table

        // Random access: every row is independently addressable in the columnar buffers, so
        // dereferencing has no traversal-order dependency and the full random-access surface
        // applies. operator*() still returns T by value (same as any forward/bidirectional
        // iterator here) - only algorithms that need a real lvalue reference (e.g. std::sort)
        // are unavailable; read-only traversal, std::advance/distance, and indexed access all work.
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

            T operator*() const
            {
                T instance{};
                static constexpr auto bindings = ExcelMapper<T>::get_bindings();
                static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
                detail::populate_instance(instance, *table_, row_, bindings, std::make_index_sequence<num_fields>{});
                return instance;
            }

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

        // Random access into the view itself, without going through an iterator.
        //
        // Unchecked, exactly like std::vector::operator[]: a `row` outside [0, size()) reads past
        // the columnar buffers and returns whatever sits after the allocation. Use at() below unless
        // the caller has already established the bound.
        T operator[](int64_t row) const
        {
            T instance{};
            static constexpr auto bindings = ExcelMapper<T>::get_bindings();
            static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
            detail::populate_instance(instance, table_, row, bindings, std::make_index_sequence<num_fields>{});
            return instance;
        }

        // Bounds-checked counterpart to operator[]. Returns nullopt rather than throwing, since this
        // header is exception-free by design (std::vector::at's out_of_range is not an option here).
        std::optional<T> at(int64_t row) const
        {
            if (row < 0 || row >= table_.row_count)
            {
                return std::nullopt;
            }
            return (*this)[row];
        }

        // Opt-in materialization for callers who want an owned std::vector<T> instead of the
        // lazy view (e.g. because they need it to outlive the view, or need random access).
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

        // Internal: constructed only by parse_sheet, which owns the xl_parse_typed call.
        static TableView from_raw(xl_table table) { return TableView(table); }

    private:
        explicit TableView(xl_table table) noexcept : table_(table) {}

        xl_table table_{};
    };

    // ---- Entry points -------------------------------------------------------------------------

    // Schema-driven, zero-vector-allocation parse: only the xl_parse_typed columnar buffers are
    // allocated. Iterate the returned TableView directly, or call .to_vector() to materialize.
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

    // Convenience for callers who want a materialized std::vector<T> up front.
    template <typename T>
    std::expected<std::vector<T>, Error> parse_sheet_vector(Workbook &workbook, int32_t header_row = 1)
    {
        auto view = parse_sheet<T>(workbook, header_row);
        if (!view.has_value())
        {
            return std::unexpected(std::move(view.error()));
        }
        return view->to_vector();
    }

} // namespace xl
