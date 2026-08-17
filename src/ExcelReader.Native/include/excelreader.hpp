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
#include <span>
#include <string>
#include <string_view>
#include <tuple>
#include <type_traits>
#include <utility>
#include <vector>

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
                                      : std::string("unknown error");
            return Error{code, std::move(message)};
        }

    } // namespace detail

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

        static std::expected<Workbook, Error> open(std::string_view path, int32_t format = XL_FORMAT_AUTO)
        {
            xl_workbook *handle = nullptr;
            int32_t status = xl_open_file(reinterpret_cast<const uint8_t *>(path.data()),
                                          static_cast<int32_t>(path.size()), format, &handle);
            if (status != XL_OK)
            {
                return std::unexpected(detail::make_error(status));
            }
            return Workbook(handle);
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

    template <>
    struct XlType<int64_t>
    {
        static constexpr int32_t value = XL_T_I64;
    };

    template <>
    struct XlType<double>
    {
        static constexpr int32_t value = XL_T_F64;
    };

    template <>
    struct XlType<bool>
    {
        static constexpr int32_t value = XL_T_BOOL;
    };

    // ---- Struct <-> column bindings ---------------------------------------------------------------

    template <typename Class, typename T>
    struct FieldBinding
    {
        const char *column_name;
        T Class::*member;
        using FieldType = T;
    };

    template <typename Class, typename T>
    constexpr FieldBinding<Class, T> make_field(const char *name, T Class::*member)
    {
        return {name, member};
    }

    // Users specialize this for each struct they want to parse into.
    template <typename T>
    struct ExcelMapper;

    namespace detail
    {

        template <typename Tuple, std::size_t... Is>
        std::array<xl_column_spec, sizeof...(Is)> build_specs(const Tuple &bindings, std::index_sequence<Is...>)
        {
            return {xl_column_spec{
                reinterpret_cast<const uint8_t *>(std::get<Is>(bindings).column_name),
                static_cast<int32_t>(std::strlen(std::get<Is>(bindings).column_name)),
                0, // index is ignored: resolved by name
                XlType<typename std::tuple_element_t<Is, Tuple>::FieldType>::value,
                1 // nullable = 1 (safe default)
            }...};
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

        template <typename Class, typename T>
        void assign_field(Class &instance, const xl_column &col, int64_t row, const FieldBinding<Class, T> &binding)
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
            else if constexpr (std::is_same_v<T, int64_t>)
            {
                instance.*(binding.member) = static_cast<const int64_t *>(col.values)[row];
            }
            else if constexpr (std::is_same_v<T, double>)
            {
                instance.*(binding.member) = static_cast<const double *>(col.values)[row];
            }
            else if constexpr (std::is_same_v<T, bool>)
            {
                instance.*(binding.member) = (static_cast<const uint8_t *>(col.values)[row] != 0);
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
        T operator[](int64_t row) const
        {
            T instance{};
            static constexpr auto bindings = ExcelMapper<T>::get_bindings();
            static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
            detail::populate_instance(instance, table_, row, bindings, std::make_index_sequence<num_fields>{});
            return instance;
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

        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{});
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
