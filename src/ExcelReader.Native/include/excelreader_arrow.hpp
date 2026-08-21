/* Optional Arrow C Data Interface export for the C++ binding. Deliberately a separate header from
 * excelreader.hpp (mirroring the excelreader.h / excelreader_arrow.h split): a caller who never
 * wants Arrow never pulls these declarations in.
 *
 * This header does NOT depend on the Apache Arrow C++ library. The Arrow C Data Interface is a
 * fixed, versioned ABI - handing back the raw ArrowArray/ArrowSchema pair lets the caller feed it
 * into whichever Arrow implementation they already link, instead of forcing one on them. */
#ifndef XL_EXCELREADER_ARROW_HPP
#define XL_EXCELREADER_ARROW_HPP

#include "excelreader.hpp"
#include "excelreader_arrow.h"

#include <expected>
#include <utility>

namespace xl
{
    // Owns one ArrowArray/ArrowSchema pair produced by parse_arrow, releasing both on destruction.
    //
    // Ownership note: xl_parse_arrow's results are released through their OWN release callbacks,
    // never through xl_free_table - the native side already freed its intermediate table before
    // returning. Releasing the schema and array is independent; both are done here.
    struct ArrowTable
    {
        ArrowArray array{};
        ArrowSchema schema{};

        ArrowTable() = default;

        ArrowTable(const ArrowTable &) = delete;
        ArrowTable &operator=(const ArrowTable &) = delete;

        // A released-or-moved-from ArrowArray/ArrowSchema is defined by the Arrow spec as one whose
        // `release` member is null, so zeroing the source is exactly what "moved-from" means here.
        ArrowTable(ArrowTable &&other) noexcept
            : array(std::exchange(other.array, ArrowArray{})),
              schema(std::exchange(other.schema, ArrowSchema{}))
        {
        }

        ArrowTable &operator=(ArrowTable &&other) noexcept
        {
            if (this != &other)
            {
                release();
                array = std::exchange(other.array, ArrowArray{});
                schema = std::exchange(other.schema, ArrowSchema{});
            }
            return *this;
        }

        ~ArrowTable() { release(); }

        // Releases both structures early. Idempotent: the Arrow spec requires a release callback to
        // null its own struct's `release` member, so a second call is a no-op.
        void release() noexcept
        {
            if (array.release != nullptr)
            {
                array.release(&array);
            }
            if (schema.release != nullptr)
            {
                schema.release(&schema);
            }
        }
    };

    // Same schema-driven parse as xl::parse_sheet<T>, returned as one top-level Arrow struct
    // array/schema instead of a TableView<T>. `header_row` has the same meaning as in parse_sheet
    // (0 = no header). Consumes the workbook's shared row cursor, hence Workbook&.
    template <typename T>
    std::expected<ArrowTable, Error> parse_arrow(Workbook &workbook, int32_t header_row = 1)
    {
        // The next four lines are xl::parse_sheet<T>'s own spec-building block
        // (cpp/include/xl/excelreader.hpp:1101-1106), copied verbatim: both entry points take an
        // identical xl_column_spec array, built from the same ExcelMapper<T>::get_bindings().
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);
        std::span<const xl_column_spec> specs(specs_array);

        ArrowTable result;
        int32_t status = xl_parse_arrow(workbook.handle(), specs.data(),
                                        static_cast<int32_t>(specs.size()), header_row,
                                        &result.array, &result.schema);
        if (status != XL_OK)
        {
            // On failure the ABI leaves both outputs untouched (still zeroed), so ~ArrowTable is a
            // no-op and there is nothing to release here.
            return std::unexpected(detail::make_error(status));
        }
        return result;
    }
}

#endif /* XL_EXCELREADER_ARROW_HPP */
