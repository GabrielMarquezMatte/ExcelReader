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

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_ARROW_H */
