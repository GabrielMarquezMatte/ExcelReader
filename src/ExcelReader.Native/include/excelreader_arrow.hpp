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
    struct ArrowTable
    {
        ArrowArray array{};
        ArrowSchema schema{};

        ArrowTable() = default;

        ArrowTable(const ArrowTable &) = delete;
        ArrowTable &operator=(const ArrowTable &) = delete;

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

    template <typename T>
    std::expected<ArrowTable, Error> parse_arrow(Workbook &workbook, int32_t header_row = 1)
    {
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
            return std::unexpected(detail::make_error(status));
        }
        return result;
    }

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

        std::expected<ArrowSchemaGuard, Error> schema()
        {
            if (stream_.release == nullptr)
            {
                return std::unexpected(Error{XL_ERROR, "the Arrow stream is released or moved-from"});
            }
            ArrowSchemaGuard guard;
            int rc = stream_.get_schema(&stream_, &guard.schema);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            return guard;
        }

        std::expected<std::optional<ArrowArrayGuard>, Error> next()
        {
            if (stream_.release == nullptr)
            {
                return std::unexpected(Error{XL_ERROR, "the Arrow stream is released or moved-from"});
            }
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
            return std::unexpected(detail::make_error(status));
        }
        return ArrowStream::from_raw(stream);
    }
}

#endif /* XL_EXCELREADER_ARROW_HPP */
