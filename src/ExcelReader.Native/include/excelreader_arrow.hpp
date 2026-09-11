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

    // One owned ArrowSchema. Separate from ArrowTable because a stream hands schemas and arrays out
    // independently - get_schema allocates a fresh one on every call, and each batch is its own
    // allocation - so they cannot share one guard.
    struct ArrowSchemaGuard
    {
        ArrowSchema schema{};

        ArrowSchemaGuard() = default;

        ArrowSchemaGuard(const ArrowSchemaGuard &) = delete;
        ArrowSchemaGuard &operator=(const ArrowSchemaGuard &) = delete;

        ArrowSchemaGuard(ArrowSchemaGuard &&other) noexcept
            : schema(std::exchange(other.schema, ArrowSchema{}))
        {
        }

        ArrowSchemaGuard &operator=(ArrowSchemaGuard &&other) noexcept
        {
            if (this != &other)
            {
                release();
                schema = std::exchange(other.schema, ArrowSchema{});
            }
            return *this;
        }

        ~ArrowSchemaGuard() { release(); }

        void release() noexcept
        {
            if (schema.release != nullptr)
            {
                schema.release(&schema);
            }
        }
    };

    // One owned ArrowArray - a single batch from an ArrowStream.
    struct ArrowArrayGuard
    {
        ArrowArray array{};

        ArrowArrayGuard() = default;

        ArrowArrayGuard(const ArrowArrayGuard &) = delete;
        ArrowArrayGuard &operator=(const ArrowArrayGuard &) = delete;

        ArrowArrayGuard(ArrowArrayGuard &&other) noexcept
            : array(std::exchange(other.array, ArrowArray{}))
        {
        }

        ArrowArrayGuard &operator=(ArrowArrayGuard &&other) noexcept
        {
            if (this != &other)
            {
                release();
                array = std::exchange(other.array, ArrowArray{});
            }
            return *this;
        }

        ~ArrowArrayGuard() { release(); }

        void release() noexcept
        {
            if (array.release != nullptr)
            {
                array.release(&array);
            }
        }
    };

    // parse_arrow delivered a batch at a time, over the Arrow C stream interface. Owns the stream
    // and releases it on destruction, which is what closes the underlying read.
    //
    // Borrows the workbook under the same rules as xl::TypedReader: one chunked read per workbook,
    // and any other read on it invalidates this stream permanently.
    class ArrowStream
    {
    public:
        ArrowStream(const ArrowStream &) = delete;
        ArrowStream &operator=(const ArrowStream &) = delete;

        ArrowStream(ArrowStream &&other) noexcept
            : stream_(std::exchange(other.stream_, ArrowArrayStream{}))
        {
        }

        ArrowStream &operator=(ArrowStream &&other) noexcept
        {
            if (this != &other)
            {
                release();
                stream_ = std::exchange(other.stream_, ArrowArrayStream{});
            }
            return *this;
        }

        ~ArrowStream() { release(); }

        // A FRESH schema on every call - each returned guard owns its own allocation, so polling
        // this without keeping the guards would leak.
        std::expected<ArrowSchemaGuard, Error> schema()
        {
            ArrowSchemaGuard guard;
            int rc = stream_.get_schema(&stream_, &guard.schema);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            return guard;
        }

        // The next batch, or an empty optional at end of stream. Arrow signals end of stream by
        // returning success with a RELEASED array, which is what the null release check reads.
        std::expected<std::optional<ArrowArrayGuard>, Error> next()
        {
            ArrowArrayGuard guard;
            int rc = stream_.get_next(&stream_, &guard.array);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            if (guard.array.release == nullptr)
            {
                return std::optional<ArrowArrayGuard>{};
            }
            return std::optional<ArrowArrayGuard>(std::move(guard));
        }

        // Internal: constructed only by arrow_stream, which owns the xl_parse_arrow_stream call.
        static ArrowStream from_raw(ArrowArrayStream stream) { return ArrowStream(stream); }

    private:
        explicit ArrowStream(ArrowArrayStream stream) noexcept : stream_(stream) {}

        void release() noexcept
        {
            if (stream_.release != nullptr)
            {
                stream_.release(&stream_);
            }
        }

        // get_next/get_schema are errno-style, not XL_*; the reason lives in get_last_error and is
        // only valid until the next call on this stream, so it is copied into the Error here.
        Error stream_error(int rc)
        {
            const char *message =
                stream_.get_last_error != nullptr ? stream_.get_last_error(&stream_) : nullptr;
            if (message != nullptr)
            {
                return Error{XL_ERROR, std::string(message)};
            }
            return Error{XL_ERROR, "Arrow stream failed with errno " + std::to_string(rc)};
        }

        ArrowArrayStream stream_{};
    };

    // parse_arrow, delivered a batch at a time. `header_row` and `batch_size` mean exactly what
    // they do in xl::typed_reader.
    template <typename T>
    std::expected<ArrowStream, Error> arrow_stream(Workbook &workbook, int32_t header_row = 1,
                                                   int64_t batch_size = 10000)
    {
        static constexpr auto bindings = ExcelMapper<T>::get_bindings();
        static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;
        std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
        std::array<xl_column_spec, num_fields> specs_array =
            detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);

        ArrowArrayStream stream{};
        int32_t status = xl_parse_arrow_stream(workbook.handle(), specs_array.data(),
                                               static_cast<int32_t>(specs_array.size()), header_row,
                                               batch_size, &stream);
        if (status != XL_OK)
        {
            // Every failure path zeroes *out_stream, so there is nothing to release here.
            return std::unexpected(detail::make_error(status));
        }
        return ArrowStream::from_raw(stream);
    }
}

#endif /* XL_EXCELREADER_ARROW_HPP */
