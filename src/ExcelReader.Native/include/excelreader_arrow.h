/* Optional Arrow C Data Interface export for ExcelReader. Separate from excelreader.h so a caller
 * that doesn't want Arrow never pulls these declarations in.
 *
 * struct ArrowSchema / struct ArrowArray below are the Arrow C Data Interface's standard structs,
 * reproduced verbatim from the Apache Arrow specification (Apache License 2.0) - they are a fixed,
 * versioned ABI shared across every Arrow producer/consumer, not an ExcelReader invention. */
#ifndef EXCELREADER_ARROW_H
#define EXCELREADER_ARROW_H

#include "excelreader.h"

#ifdef __cplusplus
extern "C" {
#endif

#ifndef ARROW_C_DATA_INTERFACE
#define ARROW_C_DATA_INTERFACE

#define ARROW_FLAG_DICTIONARY_ORDERED 1
#define ARROW_FLAG_NULLABLE 2
#define ARROW_FLAG_MAP_KEYS_SORTED 4

struct ArrowSchema {
    /* Array type description */
    const char* format;
    const char* name;
    const char* metadata;
    int64_t flags;
    int64_t n_children;
    struct ArrowSchema** children;
    struct ArrowSchema* dictionary;

    /* Release callback */
    void (*release)(struct ArrowSchema*);
    /* Opaque producer-specific data */
    void* private_data;
};

struct ArrowArray {
    /* Array data description */
    int64_t length;
    int64_t null_count;
    int64_t offset;
    int64_t n_buffers;
    int64_t n_children;
    const void** buffers;
    struct ArrowArray** children;
    struct ArrowArray* dictionary;

    /* Release callback */
    void (*release)(struct ArrowArray*);
    /* Opaque producer-specific data */
    void* private_data;
};

#endif /* ARROW_C_DATA_INTERFACE */

/* Guarded separately from the data interface above, exactly as Arrow's own arrow/c/abi.h does: the
 * stream ABI is a distinct opt-in, and a consumer can legitimately have defined
 * ARROW_C_DATA_INTERFACE from a data-only subset (nanoarrow, for one) without ever declaring
 * ArrowArrayStream. Folding this into that guard would give such a consumer xl_parse_arrow_stream
 * declared against an undeclared struct, and would redefine the struct for anyone including
 * arrow/c/abi.h after this header. */
#ifndef ARROW_C_STREAM_INTERFACE
#define ARROW_C_STREAM_INTERFACE

struct ArrowArrayStream {
    int (*get_schema)(struct ArrowArrayStream*, struct ArrowSchema* out);
    int (*get_next)(struct ArrowArrayStream*, struct ArrowArray* out);
    const char* (*get_last_error)(struct ArrowArrayStream*);
    void (*release)(struct ArrowArrayStream*);
    void* private_data;
};

#endif /* ARROW_C_STREAM_INTERFACE */

/* Same schema/column semantics as xl_parse_typed (see excelreader.h) - `header_row`, `nullable`,
 * name-vs-index resolution, XL_T_* dispatch, and every XL_INVALID_ARGUMENT/XL_ERROR case are
 * identical. The whole table is exported as ONE top-level Arrow struct array (format "+s"): its
 * length is the row count, and each xl_column_spec becomes one child array/child schema, in order.
 *
 * XL_T_* to Arrow format code:
 *     XL_T_STRING     "u"     (utf8, int32 offsets)
 *     XL_T_I64        "l"     (int64)
 *     XL_T_F64        "g"     (float64)
 *     XL_T_BOOL       "b"     (bit-packed boolean - NOT the same byte-per-value layout xl_column
 *                              uses for XL_T_BOOL; this function repacks it)
 *     XL_T_DATE       "tdD"   (date32, days)
 *     XL_T_TIME       "ttu"   (time64, microseconds)
 *     XL_T_TIMESTAMP  "tsu:"  (timestamp, microseconds, no timezone)
 *
 * `out_array`/`out_schema` are CALLER-OWNED storage (stack or heap) that this function fills in
 * place - it does not allocate the top-level struct itself, matching the Arrow C Data Interface's
 * usual producer contract. Every buffer, child, and string this function points them at IS
 * heap-allocated by ExcelReader and is released by calling `out_array->release(out_array)` and
 * `out_schema->release(out_schema)` yourself (the standard Arrow consumer contract) - there is no
 * separate xl_free_* for these; do not call xl_free_table on anything reached through them.
 *
 * Only read the out params when the call returns XL_OK. On failure they are either zeroed or left
 * untouched, depending on how early the failure was caught, and in neither case do they describe a
 * result - a zeroed struct has a NULL `release`, so releasing one is a no-op rather than a crash,
 * but nothing about it is meaningful.
 *
 * `spec_count` and each spec's `name_len` are bounded by XL_MAX_COLUMN_SPECS and
 * XL_MAX_COLUMN_NAME_BYTES (see excelreader.h); anything past either is XL_INVALID_ARGUMENT. */
int32_t xl_parse_arrow(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                       int32_t header_row, struct ArrowArray* out_array, struct ArrowSchema* out_schema);

/* Batched counterpart to xl_parse_arrow: the same schema-driven read, delivered as an Arrow C
 * stream so peak memory is one batch rather than one sheet. max_rows is the batch size in rows;
 * 0 means unbounded, negative is XL_INVALID_ARGUMENT.
 *
 * On XL_OK the caller owns *out_stream and MUST eventually call out_stream->release, which is what
 * closes the underlying read - there is no xl_free_* for it. Driving the stream follows the Arrow
 * contract, not this ABI's: get_next returns 0 on success, and signals end of stream by returning 0
 * with a RELEASED array (out->release == NULL). A non-zero return is errno-style; get_last_error
 * then yields a non-empty message valid until the next call on the same stream.
 *
 * Unlike xl_parse_arrow's out params, *out_stream is ALWAYS written: every failure path zeroes it,
 * so a caller that ignores the return code and calls out_stream->release still finds a NULL release
 * (a no-op) rather than an uninitialized function pointer. get_next and get_schema give their own
 * out params the same guarantee. The only case where *out_stream is untouched is out_stream == NULL
 * itself, which returns XL_INVALID_ARGUMENT.
 *
 * The stream borrows the workbook, with the same rules as xl_typed_reader_open. */
int32_t xl_parse_arrow_stream(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                              int32_t header_row, int64_t max_rows,
                              struct ArrowArrayStream* out_stream);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_ARROW_H */
