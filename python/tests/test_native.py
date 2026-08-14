import ctypes
import re
from pathlib import Path

from excelreader import _native

REPO_ROOT = Path(__file__).resolve().parents[2]
HEADER_PATH = REPO_ROOT / "src" / "ExcelReader.Native" / "include" / "excelreader.h"


def test_library_loads():
    lib = _native.load_library()
    assert isinstance(lib, ctypes.CDLL)


def test_library_is_memoized():
    assert _native.load_library() is _native.load_library()


def test_exported_functions_are_present():
    lib = _native.load_library()
    for name in (
        "xl_open_file",
        "xl_open_memory",
        "xl_close",
        "xl_sheet_count",
        "xl_sheet_name",
        "xl_sheet_name_at",
        "xl_move_to_sheet",
        "xl_is_date1904",
        "xl_next_row",
        "xl_read_all_blob",
        "xl_next_row_decoded",
        "xl_free_row",
        "xl_read_all_decoded",
        "xl_free_rows",
        "xl_last_error",
        "xl_last_error_ptr",
        "xl_abi_version",
    ):
        assert hasattr(lib, name), name


def test_status_constants_match_the_abi():
    assert (_native.XL_OK, _native.XL_EOF, _native.XL_BUFFER_TOO_SMALL) == (0, -1, -2)
    assert (_native.XL_INVALID_HANDLE, _native.XL_INVALID_ARGUMENT, _native.XL_ERROR) == (-3, -4, -5)


def test_last_error_two_call_form_still_works():
    # reader.py uses xl_last_error_ptr exclusively now, but xl_last_error (the ask-the-size-then-copy
    # form) is still part of the public ABI for callers who prefer to own the buffer — exercise its
    # ctypes binding directly so a signature mismatch here doesn't go unnoticed.
    lib = _native.load_library()
    path = str(REPO_ROOT / "nonexistent-workbook-for-error-test.xlsx").encode("utf-8")
    handle = ctypes.c_void_p()
    status = lib.xl_open_file(path, len(path), _native.XL_FORMAT_XLSX, ctypes.byref(handle))
    assert status == _native.XL_ERROR

    length = ctypes.c_int32()
    buffer = ctypes.create_string_buffer(4)
    assert lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_BUFFER_TOO_SMALL
    assert length.value > 0

    buffer = ctypes.create_string_buffer(length.value)
    assert lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_OK
    assert buffer.raw[: length.value]


def test_abi_version_matches_the_header():
    header = HEADER_PATH.read_text(encoding="utf-8")
    match = re.search(r"#define XL_ABI_VERSION (\d+)", header)
    assert match, "XL_ABI_VERSION not found in excelreader.h"
    header_version = int(match.group(1))

    assert _native.XL_ABI_VERSION == header_version
    assert _native.load_library().xl_abi_version() == header_version
