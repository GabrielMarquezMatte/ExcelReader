"""ctypes binding to the ExcelReader NativeAOT shared library.

Everything here mirrors src/ExcelReader.Native/include/excelreader.h. If you change one, change both.
"""

from __future__ import annotations

import ctypes
import os
import platform
from functools import lru_cache
from pathlib import Path

XL_OK = 0
XL_EOF = -1
XL_BUFFER_TOO_SMALL = -2
XL_INVALID_HANDLE = -3
XL_INVALID_ARGUMENT = -4
XL_ERROR = -5

# Bumped on any change to a struct layout, a status code, or the meaning of an existing function;
# adding a new function does not bump it. Mirrors XL_ABI_VERSION in include/excelreader.h.
XL_ABI_VERSION = 1

XL_FORMAT_AUTO = 0
XL_FORMAT_XLS = 1
XL_FORMAT_XLSX = 2
XL_FORMAT_XLSB = 3
XL_FORMAT_CSV = 4

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
    """Mirrors xl_column_spec. `name`/`name_len` may be left NULL/0 to resolve by `index` instead."""

    _fields_ = [
        ("name", ctypes.c_char_p),
        ("name_len", ctypes.c_int32),
        ("index", ctypes.c_int32),
        ("type", ctypes.c_int32),
        ("nullable", ctypes.c_int32),
    ]


def column_spec_by_name(name: str, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    encoded = name.encode("utf-8")
    return NativeColumnSpec(name=encoded, name_len=len(encoded), index=0, type=type_, nullable=int(nullable))


def column_spec_by_index(index: int, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    return NativeColumnSpec(name=None, name_len=0, index=index, type=type_, nullable=int(nullable))


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


class NativeRows(ctypes.Structure):
    _fields_ = [
        ("row_count", ctypes.c_int32),
        ("rows", ctypes.POINTER(NativeRow)),
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
    ]


def default_open_options() -> NativeOpenOptions:
    """A NativeOpenOptions with every field at its "use the library default" value and struct_size
    filled in — start from this rather than constructing NativeOpenOptions() directly, since the raw
    zero-value for struct_size is never valid."""
    options = NativeOpenOptions()
    options.struct_size = ctypes.sizeof(NativeOpenOptions)
    return options


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
    lib.xl_next_row_decoded.argtypes = [p_void, ctypes.POINTER(NativeRow)]
    lib.xl_next_row_decoded.restype = c_int
    lib.xl_free_row.argtypes = [ctypes.POINTER(NativeRow)]
    lib.xl_free_row.restype = None
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
    lib.xl_abi_version.argtypes = []
    lib.xl_abi_version.restype = c_int
    return lib
