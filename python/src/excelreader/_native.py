"""ctypes binding to the ExcelReader NativeAOT shared library.

Everything here mirrors src/ExcelReader.Native/include/excelreader.h. If you change one, change both.
"""

from __future__ import annotations

import ctypes
import os
import platform
from collections.abc import Sequence
from functools import lru_cache
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    # Import-time only: types.py imports nothing from this module, but keeping the dependency out of
    # the runtime path keeps this file loadable on its own, the way the rest of it already is.
    from excelreader.types import OpenOptions, WriteOptions

XL_OK = 0
XL_EOF = -1
XL_BUFFER_TOO_SMALL = -2
XL_INVALID_HANDLE = -3
XL_INVALID_ARGUMENT = -4
XL_ERROR = -5
XL_STATUS_PASSWORD_REQUIRED = -6
XL_STATUS_PASSWORD_INCORRECT = -7

# Bumped on any change to a struct layout, a status code, or the meaning of an existing function;
# adding a new function does not bump it. Mirrors XL_ABI_VERSION in include/excelreader.h.
XL_ABI_VERSION = 4

XL_FORMAT_AUTO = 0
XL_FORMAT_XLS = 1
XL_FORMAT_XLSX = 2
XL_FORMAT_XLSB = 3
XL_FORMAT_CSV = 4
# The one mapping from a public format name to its XL_FORMAT_* value. reader.py and writer.py both
# derive their tables from these rather than restating them.
FORMATS = {
    "auto": XL_FORMAT_AUTO,
    "xls": XL_FORMAT_XLS,
    "xlsx": XL_FORMAT_XLSX,
    "xlsb": XL_FORMAT_XLSB,
    "csv": XL_FORMAT_CSV,
}
# xl_write_typed rejects XL_FORMAT_AUTO - a file being created has no signature bytes to sniff - so
# the write side gets the same table minus that entry.
WRITE_FORMATS = {name: value for name, value in FORMATS.items() if name != "auto"}

# Every boolean-shaped NativeOpenOptions field uses one of these three states, never a plain 0/1 -
# several of them default to true, so a bare 0 would be ambiguous between "off" and "use the library
# default". Mirrors XL_OPT_* in include/excelreader.h.
XL_OPT_DEFAULT = 0
XL_OPT_FALSE = 1
XL_OPT_TRUE = 2

# xl_column_spec.type / xl_column.type. Mirrors XL_T_* in include/excelreader.h.
XL_T_STRING = 0
XL_T_I64 = 1
XL_T_F64 = 2
XL_T_BOOL = 3
XL_T_DATE = 4
XL_T_TIME = 5
XL_T_TIMESTAMP = 6


class NativeRowCell(ctypes.Structure):
    _fields_ = [
        ("column", ctypes.c_int32),
        ("type", ctypes.c_int32),
        ("value_len", ctypes.c_int32),
        ("value", ctypes.POINTER(ctypes.c_uint8)),
    ]


class NativeRow(ctypes.Structure):
    _fields_ = [
        ("cell_count", ctypes.c_int32),
        ("cells", ctypes.POINTER(NativeRowCell)),
    ]


class NativeColumnSpec(ctypes.Structure):
    """Mirrors xl_column_spec. `names`/`name_lens`/`name_count` may be left NULL/NULL/0 to resolve by
    `index` instead. Build one with `column_spec_by_name`/`column_spec_by_names`, never directly —
    the `names`/`name_lens` pointers must stay alive as long as the struct is in use, and those
    helpers keep the backing buffers alive via ctypes' `_objects` mechanism."""

    _fields_ = [
        ("names", ctypes.POINTER(ctypes.POINTER(ctypes.c_uint8))),
        ("name_lens", ctypes.POINTER(ctypes.c_int32)),
        ("name_count", ctypes.c_int32),
        ("index", ctypes.c_int32),
        ("type", ctypes.c_int32),
        ("nullable", ctypes.c_int32),
    ]


def column_spec_by_names(names: Sequence[str], type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    encoded = [name.encode("utf-8") for name in names]
    # One ctypes buffer per name (kept alive by `spec._objects` through the pointer array below),
    # plus the pointer array and length array themselves.
    buffers = [ctypes.create_string_buffer(e, len(e)) for e in encoded]
    name_ptrs = (ctypes.POINTER(ctypes.c_uint8) * len(buffers))(
        *(ctypes.cast(b, ctypes.POINTER(ctypes.c_uint8)) for b in buffers)
    )
    name_lens = (ctypes.c_int32 * len(encoded))(*(len(e) for e in encoded))
    spec = NativeColumnSpec(
        names=ctypes.cast(name_ptrs, ctypes.POINTER(ctypes.POINTER(ctypes.c_uint8))),
        name_lens=name_lens,
        name_count=len(encoded),
        index=0,
        type=type_,
        nullable=int(nullable),
    )
    # ctypes only auto-keeps-alive objects assigned directly to a field; `name_ptrs`/`name_lens`/
    # `buffers` were only cast/wrapped, so pin them explicitly on the returned struct.
    spec._name_storage = (buffers, name_ptrs, name_lens)  # type: ignore[attr-defined]
    return spec


def column_spec_by_name(name: str, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    return column_spec_by_names([name], type_, nullable=nullable)


def column_spec_by_index(index: int, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    return NativeColumnSpec(
        names=None, name_lens=None, name_count=0, index=index, type=type_, nullable=int(nullable)
    )


class NativeInferredColumnSpec(NativeColumnSpec):
    """The OUTPUT direction of the same xl_column_spec (xl_infer_schema fills these in).

    Layout is inherited, not restated: it is literally the same C struct, and a second copy of the
    field list is a second place for it to drift from the header. Only the ownership rules differ,
    which is why this carries its own name.

    Always carries `name_count` 0 or 1 — inference never guesses more than one candidate name per
    column. Unlike `column_spec_by_name`'s buffers, `names[0]` here is a raw pointer with no
    guaranteed NUL terminator, so decode it with `ctypes.string_at(names[0], name_lens[0])`.
    """


class NativeInferredSchema(ctypes.Structure):
    """Mirrors xl_inferred_schema. `columns` is a native-owned array of `column_count` values, freed
    (along with each column's own `name`) by `xl_free_schema`."""

    _fields_ = [
        ("columns", ctypes.POINTER(NativeInferredColumnSpec)),
        ("column_count", ctypes.c_int32),
    ]


class NativeColumn(ctypes.Structure):
    """Mirrors xl_column. `data`/`data_len` are only meaningful for XL_T_STRING columns; `values` is
    always the ONE allocation this column owns directly (see xl_column's doc comment in the header —
    `data` is an interior pointer into `values` for strings, never a separate allocation)."""

    _fields_ = [
        ("type", ctypes.c_int32),
        ("length", ctypes.c_int64),
        ("values", ctypes.c_void_p),
        ("validity", ctypes.POINTER(ctypes.c_uint8)),
        ("data", ctypes.POINTER(ctypes.c_uint8)),
        ("data_len", ctypes.c_int64),
    ]


class NativeTable(ctypes.Structure):
    _fields_ = [
        ("column_count", ctypes.c_int32),
        ("row_count", ctypes.c_int64),
        ("columns", ctypes.POINTER(NativeColumn)),
    ]


# Mirrors the Arrow C Data Interface's struct ArrowSchema/struct ArrowArray (see
# excelreader_arrow.h) — a fixed, versioned ABI shared across every Arrow producer/consumer, not an
# ExcelReader invention. Defined here (not left to pyarrow) so xl_parse_arrow is usable via plain
# ctypes even without pyarrow installed; pyarrow.Array._import_from_c consumes the same layout for a
# zero-copy handoff when it IS installed.
class ArrowSchema(ctypes.Structure):
    pass


class ArrowArray(ctypes.Structure):
    pass


ArrowSchema._fields_ = [
    ("format", ctypes.c_char_p),
    ("name", ctypes.c_char_p),
    ("metadata", ctypes.c_char_p),
    ("flags", ctypes.c_int64),
    ("n_children", ctypes.c_int64),
    ("children", ctypes.POINTER(ctypes.POINTER(ArrowSchema))),
    ("dictionary", ctypes.POINTER(ArrowSchema)),
    ("release", ctypes.c_void_p),
    ("private_data", ctypes.c_void_p),
]

ArrowArray._fields_ = [
    ("length", ctypes.c_int64),
    ("null_count", ctypes.c_int64),
    ("offset", ctypes.c_int64),
    ("n_buffers", ctypes.c_int64),
    ("n_children", ctypes.c_int64),
    ("buffers", ctypes.POINTER(ctypes.c_void_p)),
    ("children", ctypes.POINTER(ctypes.POINTER(ArrowArray))),
    ("dictionary", ctypes.POINTER(ArrowArray)),
    ("release", ctypes.c_void_p),
    ("private_data", ctypes.c_void_p),
]

# ARROW_FLAG_NULLABLE from the Arrow C Data Interface spec.
ARROW_FLAG_NULLABLE = 2


class NativeRows(ctypes.Structure):
    _fields_ = [
        ("row_count", ctypes.c_int32),
        ("rows", ctypes.POINTER(NativeRow)),
    ]


class NativeBuffer(ctypes.Structure):
    """Mirrors xl_buffer."""

    _fields_ = [
        ("data", ctypes.POINTER(ctypes.c_uint8)),
        ("len", ctypes.c_int64),
    ]


class NativeOpenOptions(ctypes.Structure):
    """Mirrors xl_open_options. `struct_size` is set for you by `default_open_options()`; every other
    numeric field is 0 (use the library default), and every XL_OPT_* field starts at XL_OPT_DEFAULT."""

    _fields_ = [
        ("struct_size", ctypes.c_int32),
        ("csv_sniff_dialect", ctypes.c_int32),
        ("csv_delimiter", ctypes.c_int32),
        ("csv_quote", ctypes.c_int32),
        ("csv_detect_bom", ctypes.c_int32),
        ("csv_max_cell_bytes", ctypes.c_int32),
        ("csv_intern_strings", ctypes.c_int32),
        ("max_total_decompressed_bytes", ctypes.c_int64),
        ("max_cell_bytes", ctypes.c_int32),
        ("max_shared_string_bytes", ctypes.c_int64),
        ("max_zip_entries", ctypes.c_int32),
        ("prefetch_decompression", ctypes.c_int32),
        ("intern_strings", ctypes.c_int32),
        ("password", ctypes.POINTER(ctypes.c_uint8)),
        ("password_len", ctypes.c_int32),
    ]


def default_open_options() -> NativeOpenOptions:
    """A NativeOpenOptions with every field at its "use the library default" value and struct_size
    filled in — start from this rather than constructing NativeOpenOptions() directly, since the raw
    zero-value for struct_size is never valid."""
    options = NativeOpenOptions()
    options.struct_size = ctypes.sizeof(NativeOpenOptions)
    return options


def _opt_state(value: bool | None) -> int:
    if value is None:
        return XL_OPT_DEFAULT
    return XL_OPT_TRUE if value else XL_OPT_FALSE


def _opt_number(value: int | None) -> int:
    # Explicit None test rather than `value or 0`: the latter would also collapse a caller's
    # deliberate 0, and a limit silently turning into "use the default" is exactly the kind of
    # quiet semantic change the reader's own rules forbid.
    return 0 if value is None else value


def to_native_open_options(options: OpenOptions) -> NativeOpenOptions:
    """Converts a public `OpenOptions` into the raw ABI struct.

    This is where the public API's single "None means default" convention splits back into the two
    the ABI uses: 0 for an unset number, XL_OPT_DEFAULT for an unset boolean. Note the numeric
    fields cannot express a deliberate 0 — the ABI spends that value on "default" — but no field
    here has a meaningful 0 (a zero delimiter, or a zero-byte cell limit, is not a setting).

    Values are passed through unvalidated on purpose: the native side owns the real bounds and
    reports a rejection through xl_last_error, so checking them here too would give those bounds a
    second place to drift from.
    """
    raw = default_open_options()
    raw.csv_sniff_dialect = _opt_state(options.csv_sniff_dialect)
    raw.csv_detect_bom = _opt_state(options.csv_detect_bom)
    raw.csv_intern_strings = _opt_state(options.csv_intern_strings)
    raw.prefetch_decompression = _opt_state(options.prefetch_decompression)
    raw.intern_strings = _opt_state(options.intern_strings)
    raw.csv_delimiter = _opt_number(options.csv_delimiter)
    raw.csv_quote = _opt_number(options.csv_quote)
    raw.csv_max_cell_bytes = _opt_number(options.csv_max_cell_bytes)
    raw.max_total_decompressed_bytes = _opt_number(options.max_total_decompressed_bytes)
    raw.max_cell_bytes = _opt_number(options.max_cell_bytes)
    raw.max_shared_string_bytes = _opt_number(options.max_shared_string_bytes)
    raw.max_zip_entries = _opt_number(options.max_zip_entries)
    return raw


class NativeWriteOptions(ctypes.Structure):
    """Mirrors xl_write_options. `struct_size` is set for you by `default_write_options()`; every other
    numeric field is 0 (use the library default), and every XL_OPT_* field starts at XL_OPT_DEFAULT."""

    _fields_ = [
        ("struct_size", ctypes.c_int32),
        ("sheet_name_len", ctypes.c_int32),
        ("sheet_name", ctypes.c_char_p),
        ("csv_delimiter", ctypes.c_int32),
        ("csv_quote", ctypes.c_int32),
        ("date1904", ctypes.c_int32),
        ("use_shared_strings", ctypes.c_int32),
    ]


def default_write_options() -> NativeWriteOptions:
    """A NativeWriteOptions with every field at its "use the library default" value and struct_size
    filled in — start from this rather than constructing NativeWriteOptions() directly, since the raw
    zero-value for struct_size is never valid."""
    options = NativeWriteOptions()
    options.struct_size = ctypes.sizeof(NativeWriteOptions)
    return options


def to_native_write_options(options: WriteOptions) -> NativeWriteOptions:
    """Converts a public `WriteOptions` into the raw ABI struct.

    Same split as `to_native_open_options()`: the public API's single "None means default" convention
    becomes 0 for an unset number and XL_OPT_DEFAULT for an unset boolean. Values are passed through
    unvalidated on purpose — the native side owns the real bounds and reports a rejection through
    xl_last_error, so checking them here too would give those bounds a second place to drift from.
    """
    raw = default_write_options()
    if options.sheet_name is not None:
        encoded = options.sheet_name.encode("utf-8")
        raw.sheet_name = encoded
        raw.sheet_name_len = len(encoded)
        # No separate keepalive is needed, but NOT because anything is copied: assigning a bytes
        # object to a c_char_p field stores a POINTER into that object's buffer and files the object
        # in the structure's `_objects` dict, so `raw` itself keeps `encoded` alive. Drop `raw` (or
        # rebuild the field from a temporary) and the pointer dangles.
    raw.csv_delimiter = _opt_number(options.csv_delimiter)
    raw.csv_quote = _opt_number(options.csv_quote)
    raw.date1904 = _opt_state(options.date1904)
    raw.use_shared_strings = _opt_state(options.use_shared_strings)
    return raw


_LIB_NAMES = {
    "Windows": "ExcelReader.Native.dll",
    "Linux": "ExcelReader.Native.so",
    "Darwin": "ExcelReader.Native.dylib",
}

# NativeAOT emits no "lib" prefix, so the filename is the assembly name on every platform.
def library_filename() -> str:
    try:
        return _LIB_NAMES[platform.system()]
    except KeyError:
        raise RuntimeError(f"unsupported platform: {platform.system()}") from None


def _candidate_paths() -> list[Path]:
    override = os.environ.get("EXCELREADER_NATIVE_LIB")
    if override:
        return [Path(override)]
    return [Path(__file__).resolve().parent / "_lib" / library_filename()]


@lru_cache(maxsize=1)
def load_library() -> ctypes.CDLL:
    for path in _candidate_paths():
        if not path.exists():
            continue
        lib = _bind(ctypes.CDLL(str(path)))
        version = lib.xl_abi_version()
        if version != XL_ABI_VERSION:
            raise RuntimeError(
                f"{path} is ABI version {version}, but this package expects {XL_ABI_VERSION}. "
                f"Rebuild the native library with python python/scripts/build_native.py."
            )
        return lib
    raise RuntimeError(
        f"{library_filename()} not found. Build it with:\n"
        f"    python python/scripts/build_native.py\n"
        f"or point EXCELREADER_NATIVE_LIB at an existing binary."
    )


def _bind(lib: ctypes.CDLL) -> ctypes.CDLL:
    c_int = ctypes.c_int32
    p_int = ctypes.POINTER(ctypes.c_int32)
    p_void = ctypes.c_void_p
    pp_void = ctypes.POINTER(ctypes.c_void_p)
    p_bytes = ctypes.c_char_p

    p_open_options = ctypes.POINTER(NativeOpenOptions)

    lib.xl_open_file.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_file.restype = c_int
    lib.xl_open_file_ex.argtypes = [p_bytes, c_int, c_int, p_open_options, pp_void]
    lib.xl_open_file_ex.restype = c_int
    lib.xl_open_memory.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_memory.restype = c_int
    lib.xl_open_memory_ex.argtypes = [p_bytes, c_int, c_int, p_open_options, pp_void]
    lib.xl_open_memory_ex.restype = c_int
    lib.xl_close.argtypes = [p_void]
    lib.xl_close.restype = c_int
    lib.xl_sheet_count.argtypes = [p_void, p_int]
    lib.xl_sheet_count.restype = c_int
    # xl_sheet_name / xl_next_row / xl_last_error write INTO their buffer argument.
    # Callers MUST pass ctypes.create_string_buffer(n), never a bytes literal — bytes
    # objects are immutable/interned in CPython, and letting native code write into
    # one is undefined behavior.
    lib.xl_sheet_name.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_sheet_name.restype = c_int
    lib.xl_sheet_name_at.argtypes = [p_void, c_int, p_bytes, c_int, p_int]
    lib.xl_sheet_name_at.restype = c_int
    lib.xl_move_to_sheet.argtypes = [p_void, c_int]
    lib.xl_move_to_sheet.restype = c_int
    lib.xl_is_date1904.argtypes = [p_void, p_int]
    lib.xl_is_date1904.restype = c_int
    lib.xl_next_row.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_next_row.restype = c_int
    lib.xl_read_all_blob.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_read_all_blob.restype = c_int
    lib.xl_last_error.argtypes = [p_bytes, c_int, p_int]
    lib.xl_last_error.restype = c_int
    lib.xl_last_error_ptr.argtypes = [p_int]
    lib.xl_last_error_ptr.restype = ctypes.POINTER(ctypes.c_uint8)
    lib.xl_read_all_decoded.argtypes = [p_void, ctypes.POINTER(NativeRows)]
    lib.xl_read_all_decoded.restype = c_int
    lib.xl_free_rows.argtypes = [ctypes.POINTER(NativeRows)]
    lib.xl_free_rows.restype = None
    lib.xl_parse_typed.argtypes = [p_void, ctypes.POINTER(NativeColumnSpec), c_int, c_int, ctypes.POINTER(NativeTable)]
    lib.xl_parse_typed.restype = c_int
    lib.xl_free_table.argtypes = [ctypes.POINTER(NativeTable)]
    lib.xl_free_table.restype = None
    lib.xl_infer_schema.argtypes = [p_void, c_int, c_int, ctypes.POINTER(NativeInferredSchema)]
    lib.xl_infer_schema.restype = c_int
    lib.xl_free_schema.argtypes = [ctypes.POINTER(NativeInferredSchema)]
    lib.xl_free_schema.restype = None
    lib.xl_parse_arrow.argtypes = [p_void, ctypes.POINTER(NativeColumnSpec), c_int, c_int, ctypes.POINTER(ArrowArray), ctypes.POINTER(ArrowSchema)]
    lib.xl_parse_arrow.restype = c_int
    lib.xl_write_typed.argtypes = [
        p_bytes, c_int, c_int,
        ctypes.POINTER(NativeColumnSpec), ctypes.POINTER(NativeTable), ctypes.POINTER(NativeWriteOptions),
    ]
    lib.xl_write_typed.restype = c_int
    lib.xl_write_typed_to_memory.argtypes = [
        ctypes.c_int32, ctypes.POINTER(NativeColumnSpec), ctypes.POINTER(NativeTable),
        ctypes.POINTER(NativeWriteOptions), ctypes.POINTER(NativeBuffer),
    ]
    lib.xl_write_typed_to_memory.restype = ctypes.c_int32
    lib.xl_encrypt_package.argtypes = [p_bytes, c_int, p_bytes, c_int, p_bytes, c_int]
    lib.xl_encrypt_package.restype = c_int
    lib.xl_abi_version.argtypes = []
    lib.xl_abi_version.restype = c_int

    lib.xl_open_write_handle.argtypes = [
        ctypes.POINTER(ctypes.c_uint8), ctypes.c_int32, ctypes.c_int32,
        ctypes.POINTER(NativeWriteOptions), ctypes.POINTER(ctypes.c_void_p),
    ]
    lib.xl_open_write_handle.restype = ctypes.c_int32

    lib.xl_open_write_handle_to_memory.argtypes = [
        ctypes.c_int32, ctypes.POINTER(NativeWriteOptions), ctypes.POINTER(ctypes.c_void_p),
    ]
    lib.xl_open_write_handle_to_memory.restype = ctypes.c_int32

    lib.xl_start_sheet.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint8), ctypes.c_int32]
    lib.xl_start_sheet.restype = ctypes.c_int32

    lib.xl_start_row.argtypes = [ctypes.c_void_p]
    lib.xl_start_row.restype = ctypes.c_int32

    lib.xl_write_string.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint8), ctypes.c_int32]
    lib.xl_write_string.restype = ctypes.c_int32

    lib.xl_write_int64.argtypes = [ctypes.c_void_p, ctypes.c_int64]
    lib.xl_write_int64.restype = ctypes.c_int32

    lib.xl_write_float64.argtypes = [ctypes.c_void_p, ctypes.c_double]
    lib.xl_write_float64.restype = ctypes.c_int32

    lib.xl_write_bool.argtypes = [ctypes.c_void_p, ctypes.c_int32]
    lib.xl_write_bool.restype = ctypes.c_int32

    lib.xl_write_date.argtypes = [ctypes.c_void_p, ctypes.c_int32]
    lib.xl_write_date.restype = ctypes.c_int32

    lib.xl_write_time.argtypes = [ctypes.c_void_p, ctypes.c_int64]
    lib.xl_write_time.restype = ctypes.c_int32

    lib.xl_write_timestamp.argtypes = [ctypes.c_void_p, ctypes.c_int64]
    lib.xl_write_timestamp.restype = ctypes.c_int32

    lib.xl_write_null.argtypes = [ctypes.c_void_p, ctypes.c_int32]
    lib.xl_write_null.restype = ctypes.c_int32

    lib.xl_end_row.argtypes = [ctypes.c_void_p]
    lib.xl_end_row.restype = ctypes.c_int32

    lib.xl_end_sheet.argtypes = [ctypes.c_void_p]
    lib.xl_end_sheet.restype = ctypes.c_int32

    lib.xl_close_write_handle.argtypes = [ctypes.c_void_p]
    lib.xl_close_write_handle.restype = ctypes.c_int32

    lib.xl_write_handle_bytes.argtypes = [ctypes.c_void_p, ctypes.POINTER(NativeBuffer)]
    lib.xl_write_handle_bytes.restype = ctypes.c_int32

    lib.xl_free_buffer.argtypes = [ctypes.POINTER(NativeBuffer)]
    lib.xl_free_buffer.restype = None

    return lib
