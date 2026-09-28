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
    const char* format;
    const char* name;
    const char* metadata;
    int64_t flags;
    int64_t n_children;
    struct ArrowSchema** children;
    struct ArrowSchema* dictionary;

    void (*release)(struct ArrowSchema*);
    void* private_data;
};

struct ArrowArray {
    int64_t length;
    int64_t null_count;
    int64_t offset;
    int64_t n_buffers;
    int64_t n_children;
    const void** buffers;
    struct ArrowArray** children;
    struct ArrowArray* dictionary;

    void (*release)(struct ArrowArray*);
    void* private_data;
};

#endif /* ARROW_C_DATA_INTERFACE */

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

int32_t xl_parse_arrow(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                       int32_t header_row, struct ArrowArray* out_array, struct ArrowSchema* out_schema);

/* xl_parse_arrow with a thread count; same contract as xl_parse_typed_ex. */
int32_t xl_parse_arrow_ex(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                          int32_t header_row, int32_t degree_of_parallelism,
                          struct ArrowArray* out_array, struct ArrowSchema* out_schema);

int32_t xl_parse_arrow_stream(xl_workbook* handle, const xl_column_spec* specs, int32_t spec_count,
                              int32_t header_row, int64_t max_rows,
                              struct ArrowArrayStream* out_stream);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_ARROW_H */
