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

XL_FORMAT_AUTO = 0
XL_FORMAT_XLS = 1
XL_FORMAT_XLSX = 2
XL_FORMAT_XLSB = 3
XL_FORMAT_CSV = 4

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


class NativeRows(ctypes.Structure):
    _fields_ = [
        ("row_count", ctypes.c_int32),
        ("rows", ctypes.POINTER(NativeRow)),
    ]


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
        if path.exists():
            return _bind(ctypes.CDLL(str(path)))
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

    lib.xl_open_file.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_file.restype = c_int
    lib.xl_open_memory.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_memory.restype = c_int
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
    lib.xl_move_to_sheet.argtypes = [p_void, c_int]
    lib.xl_move_to_sheet.restype = c_int
    lib.xl_is_date1904.argtypes = [p_void, p_int]
    lib.xl_is_date1904.restype = c_int
    lib.xl_next_row.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_next_row.restype = c_int
    lib.xl_last_error.argtypes = [p_bytes, c_int, p_int]
    lib.xl_last_error.restype = c_int
    lib.xl_read_all_decoded.argtypes = [p_void, ctypes.POINTER(NativeRows)]
    lib.xl_read_all_decoded.restype = c_int
    lib.xl_free_rows.argtypes = [ctypes.POINTER(NativeRows)]
    lib.xl_free_rows.restype = None
    return lib
