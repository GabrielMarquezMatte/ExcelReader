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
    /// @brief A sheet or one batch of it, exported through the Arrow C Data Interface: one struct
    /// array and its schema.
    ///
    /// Releases both on destruction; move them out to hand them to an Arrow library.
    /// ArrowStream::schema() returns one whose `array` is empty.
    struct ArrowTable
    {
        ArrowArray array{};   ///< One struct array holding every column.
        ArrowSchema schema{}; ///< The struct type of `array`.

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

        /// @brief Releases both now. Safe to call again.
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

    /// @brief parse into Arrow instead of `T`.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    /// @param sheet The sheet to read.
    /// @param options Header row and parallelism.
    template <typename T>
    std::expected<ArrowTable, Error> to_arrow(Sheet sheet, const ReadOptions &options = {})
    {
        ArrowTable result;
        const int32_t status = detail::with_specs<T>(
            sheet, [&](xl_workbook *handle, int32_t index, const xl_column_spec *specs, int32_t count)
            { return xl_parse_arrow(handle, index, specs, count, options.header_row, options.parallelism,
                                    &result.array, &result.schema); });
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return result;
    }

    /// @brief parse_batches's Arrow counterpart: record batches pulled one at a time.
    ///
    /// Use from one thread at a time.
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

        /// @brief The schema every batch shares, in a table whose `array` is empty.
        std::expected<ArrowTable, Error> schema()
        {
            if (stream_.release == nullptr)
            {
                return std::unexpected(Error{XL_ERROR, "the Arrow stream is released or moved-from"});
            }
            ArrowTable table;
            int rc = stream_.get_schema(&stream_, &table.schema);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            return table;
        }

        /// @brief Reads the next batch.
        /// @return The batch with its own copy of the schema, or empty at the end of the sheet.
        std::expected<std::optional<ArrowTable>, Error> next()
        {
            if (stream_.release == nullptr)
            {
                return std::unexpected(Error{XL_ERROR, "the Arrow stream is released or moved-from"});
            }
            ArrowTable table;
            int rc = stream_.get_next(&stream_, &table.array);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            if (table.array.release == nullptr)
            {
                return std::optional<ArrowTable>{};
            }
            rc = stream_.get_schema(&stream_, &table.schema);
            if (rc != 0)
            {
                return std::unexpected(stream_error(rc));
            }
            return std::optional<ArrowTable>(std::move(table));
        }

    private:
        friend struct detail::Access;

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

    /// @brief Opens an ArrowStream over a sheet.
    /// @tparam T Row type, mapped by `ExcelMapper<T>`.
    /// @param sheet The sheet to read.
    /// @param options Header row and batch size.
    template <typename T>
    std::expected<ArrowStream, Error> to_arrow_stream(Sheet sheet, const ReadOptions &options = {})
    {
        if (std::optional<Error> problem = detail::check_batch_size(options.batch_size))
        {
            return std::unexpected(std::move(*problem));
        }
        ArrowArrayStream stream{};
        const int32_t status = detail::with_specs<T>(
            sheet, [&](xl_workbook *handle, int32_t index, const xl_column_spec *specs, int32_t count)
            { return xl_parse_arrow_stream(handle, index, specs, count, options.header_row, options.batch_size, &stream); });
        if (status != XL_OK)
        {
            return std::unexpected(detail::make_error(status));
        }
        return detail::Access::make<ArrowStream>(stream);
    }
}

#endif /* XL_EXCELREADER_ARROW_HPP */
