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
        "xl_open_file_ex",
        "xl_open_memory",
        "xl_open_memory_ex",
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
        "xl_parse_typed",
        "xl_free_table",
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


def test_open_file_ex_with_default_options_opens_like_open_file(xlsx_path):
    # Exercises the NativeOpenOptions ctypes.Structure layout end-to-end against the real library —
    # a field-order/size mismatch with the C# NativeOpenOptionsRaw would show up here as
    # XL_INVALID_ARGUMENT (struct_size mismatch) or a crash, not merely a missing symbol.
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    options = _native.default_open_options()
    handle = ctypes.c_void_p()

    status = lib.xl_open_file_ex(path, len(path), _native.XL_FORMAT_XLSX, ctypes.byref(options), ctypes.byref(handle))

    assert status == _native.XL_OK
    assert handle.value
    assert lib.xl_close(handle) == _native.XL_OK


def test_open_file_ex_with_null_options_pointer_behaves_like_open_file(xlsx_path):
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    handle = ctypes.c_void_p()

    status = lib.xl_open_file_ex(path, len(path), _native.XL_FORMAT_XLSX, None, ctypes.byref(handle))

    assert status == _native.XL_OK
    assert lib.xl_close(handle) == _native.XL_OK


def test_open_file_ex_applies_a_csv_delimiter_override(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "semicolons.csv"
    csv_file.write_text("name;qty\nwidget;7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")

    options = _native.default_open_options()
    options.csv_delimiter = ord(";")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file_ex(path, len(path), _native.XL_FORMAT_CSV, ctypes.byref(options), ctypes.byref(handle)) == _native.XL_OK

    buffer = ctypes.create_string_buffer(4096)
    written = ctypes.c_int32()
    assert lib.xl_next_row(handle, buffer, len(buffer), ctypes.byref(written)) == _native.XL_OK
    assert b"name" in buffer.raw[: written.value]
    assert b"qty" in buffer.raw[: written.value]
    lib.xl_close(handle)


def test_open_file_ex_rejects_an_unrecognized_struct_size(xlsx_path):
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    options = _native.default_open_options()
    options.struct_size = 1
    handle = ctypes.c_void_p()

    status = lib.xl_open_file_ex(path, len(path), _native.XL_FORMAT_XLSX, ctypes.byref(options), ctypes.byref(handle))

    assert status == _native.XL_INVALID_ARGUMENT


def test_parse_typed_returns_typed_columns_by_name(tmp_path):
    # Exercises the NativeColumnSpec/NativeColumn/NativeTable ctypes.Structure layouts end-to-end
    # against the real library, the same way test_open_file_ex_with_default_options_opens_like_open_file
    # does for NativeOpenOptions — a field mismatch here would surface as garbage values or a crash,
    # not a clean assertion failure, so this is the test that actually proves the layouts match.
    lib = _native.load_library()
    csv_file = tmp_path / "typed.csv"
    csv_file.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, ctypes.byref(handle)) == _native.XL_OK

    specs = (_native.NativeColumnSpec * 2)(
        _native.column_spec_by_name("name", _native.XL_T_STRING),
        _native.column_spec_by_name("qty", _native.XL_T_I64),
    )
    table = _native.NativeTable()
    status = lib.xl_parse_typed(handle, specs, len(specs), 1, ctypes.byref(table))
    assert status == _native.XL_OK
    assert table.row_count == 2
    assert table.column_count == 2

    qty_column = table.columns[1]
    qty_values = ctypes.cast(qty_column.values, ctypes.POINTER(ctypes.c_int64))
    assert [qty_values[i] for i in range(2)] == [3, 7]

    name_column = table.columns[0]
    offsets = ctypes.cast(name_column.values, ctypes.POINTER(ctypes.c_int32))
    data = ctypes.string_at(name_column.data, name_column.data_len)
    names = [data[offsets[i] : offsets[i + 1]].decode("utf-8") for i in range(2)]
    assert names == ["widget", "gadget"]

    lib.xl_free_table(ctypes.byref(table))
    lib.xl_close(handle)


def test_abi_version_matches_the_header():
    header = HEADER_PATH.read_text(encoding="utf-8")
    match = re.search(r"#define XL_ABI_VERSION (\d+)", header)
    assert match, "XL_ABI_VERSION not found in excelreader.h"
    header_version = int(match.group(1))

    assert _native.XL_ABI_VERSION == header_version
    assert _native.load_library().xl_abi_version() == header_version
