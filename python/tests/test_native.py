import pytest
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
        "xl_open_source",
        "xl_open_stream",
        "xl_set_source_error",
        "xl_close",
        "xl_sheet_count",
        "xl_sheet_name_at",
        "xl_sheet_visibility_at",
        "xl_sheet_index",
        "xl_is_date1904",
        "xl_rows_open",
        "xl_rows_close",
        "xl_rows_next",
        "xl_rows_read_all_blob",
        "xl_rows_read_all_decoded",
        "xl_free_rows",
        "xl_parse_typed",
        "xl_infer_schema",
        "xl_free_table",
        "xl_parse_arrow",
        "xl_typed_reader_open",
        "xl_typed_reader_next",
        "xl_typed_reader_close",
        "xl_parse_arrow_stream",
        "xl_write_typed",
        "xl_abort_write_handle",
        "xl_last_error",
        "xl_last_error_ptr",
        "xl_abi_version",
    ):
        assert hasattr(lib, name), name


def test_status_constants_match_the_abi():
    assert (_native.XL_OK, _native.XL_EOF, _native.XL_BUFFER_TOO_SMALL) == (0, -1, -2)
    assert (_native.XL_INVALID_HANDLE, _native.XL_INVALID_ARGUMENT, _native.XL_ERROR) == (-3, -4, -5)


def test_last_error_two_call_form_still_works():
    lib = _native.load_library()
    path = str(REPO_ROOT / "nonexistent-workbook-for-error-test.xlsx").encode("utf-8")
    handle = ctypes.c_void_p()
    status = lib.xl_open_file(path, len(path), _native.XL_FORMAT_XLSX, None, ctypes.byref(handle))
    assert status == _native.XL_ERROR

    length = ctypes.c_int32()
    buffer = ctypes.create_string_buffer(4)
    assert lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_BUFFER_TOO_SMALL
    assert length.value > 0

    buffer = ctypes.create_string_buffer(length.value)
    assert lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_OK
    assert buffer.raw[: length.value]


def test_open_file_with_default_options_opens(xlsx_path):
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    options = _native.default_open_options()
    handle = ctypes.c_void_p()

    status = lib.xl_open_file(path, len(path), _native.XL_FORMAT_XLSX, ctypes.byref(options), ctypes.byref(handle))

    assert status == _native.XL_OK
    assert handle.value
    assert lib.xl_close(handle) == _native.XL_OK


def test_open_file_with_a_null_options_pointer_opens(xlsx_path):
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    handle = ctypes.c_void_p()

    status = lib.xl_open_file(path, len(path), _native.XL_FORMAT_XLSX, None, ctypes.byref(handle))

    assert status == _native.XL_OK
    assert lib.xl_close(handle) == _native.XL_OK


def test_open_file_applies_a_csv_delimiter_override(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "semicolons.csv"
    csv_file.write_text("name;qty\nwidget;7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")

    options = _native.default_open_options()
    options.csv_delimiter = ord(";")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, ctypes.byref(options), ctypes.byref(handle)) == _native.XL_OK

    cursor = ctypes.c_void_p()
    assert lib.xl_rows_open(handle, 0, ctypes.byref(cursor)) == _native.XL_OK
    buffer = ctypes.create_string_buffer(4096)
    written = ctypes.c_int32()
    assert lib.xl_rows_next(cursor, buffer, len(buffer), ctypes.byref(written)) == _native.XL_OK
    assert b"name" in buffer.raw[: written.value]
    assert b"qty" in buffer.raw[: written.value]
    lib.xl_rows_close(cursor)
    lib.xl_close(handle)


def test_open_file_rejects_an_unrecognized_struct_size(xlsx_path):
    lib = _native.load_library()
    path = str(xlsx_path).encode("utf-8")
    options = _native.default_open_options()
    options.struct_size = 1
    handle = ctypes.c_void_p()

    status = lib.xl_open_file(path, len(path), _native.XL_FORMAT_XLSX, ctypes.byref(options), ctypes.byref(handle))

    assert status == _native.XL_INVALID_ARGUMENT


def test_parse_typed_returns_typed_columns_by_name(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "typed.csv"
    csv_file.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, None, ctypes.byref(handle)) == _native.XL_OK

    specs = (_native.NativeColumnSpec * 2)(
        _native.column_spec_by_name("name", _native.XL_T_STRING),
        _native.column_spec_by_name("qty", _native.XL_T_I64),
    )
    table = _native.NativeTable()
    status = lib.xl_parse_typed(handle, 0, specs, len(specs), 1, 1, ctypes.byref(table))
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

def test_parse_typed_resolves_the_first_alias_present_in_the_header_row(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "typed.csv"
    csv_file.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, None, ctypes.byref(handle)) == _native.XL_OK

    specs = (_native.NativeColumnSpec * 2)(
        _native.column_spec_by_names(["does-not-exist", "name"], _native.XL_T_STRING),
        _native.column_spec_by_name("qty", _native.XL_T_I64),
    )
    table = _native.NativeTable()
    status = lib.xl_parse_typed(handle, 0, specs, len(specs), 1, 1, ctypes.byref(table))
    assert status == _native.XL_OK
    assert table.row_count == 2

    lib.xl_free_table(ctypes.byref(table))
    lib.xl_close(handle)


def test_parse_arrow_returns_a_struct_array_with_a_matching_schema(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "typed.csv"
    csv_file.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, None, ctypes.byref(handle)) == _native.XL_OK

    specs = (_native.NativeColumnSpec * 2)(
        _native.column_spec_by_name("name", _native.XL_T_STRING),
        _native.column_spec_by_name("qty", _native.XL_T_I64),
    )
    array = _native.ArrowArray()
    schema = _native.ArrowSchema()
    status = lib.xl_parse_arrow(handle, 0, specs, len(specs), 1, 1, ctypes.byref(array), ctypes.byref(schema))
    assert status == _native.XL_OK

    assert schema.format == b"+s"
    assert schema.n_children == 2
    name_schema = schema.children[0].contents
    qty_schema = schema.children[1].contents
    assert name_schema.format == b"u"
    assert name_schema.name == b"name"
    assert qty_schema.format == b"l"

    assert array.length == 2
    assert array.n_children == 2
    name_array = array.children[0].contents
    qty_array = array.children[1].contents

    offsets = ctypes.cast(qty_array.buffers[0], ctypes.c_void_p)  
    assert not offsets.value
    qty_values = ctypes.cast(qty_array.buffers[1], ctypes.POINTER(ctypes.c_int64))
    assert [qty_values[i] for i in range(2)] == [3, 7]

    name_offsets = ctypes.cast(name_array.buffers[1], ctypes.POINTER(ctypes.c_int32))
    name_data = ctypes.string_at(name_array.buffers[2], name_offsets[2])
    names = [name_data[name_offsets[i] : name_offsets[i + 1]].decode("utf-8") for i in range(2)]
    assert names == ["widget", "gadget"]

    array_release = ctypes.CFUNCTYPE(None, ctypes.POINTER(_native.ArrowArray))(array.release)
    schema_release = ctypes.CFUNCTYPE(None, ctypes.POINTER(_native.ArrowSchema))(schema.release)
    array_release(ctypes.byref(array))
    schema_release(ctypes.byref(schema))
    assert not array.release
    assert not schema.release
    array_release(ctypes.byref(array))
    schema_release(ctypes.byref(schema))

    lib.xl_close(handle)


def test_write_options_default_to_a_serial_write():
    from excelreader import WriteOptions

    assert ctypes.sizeof(_native.NativeWriteOptions) == 40
    raw = _native.to_native_write_options(WriteOptions())
    assert (raw.degree_of_parallelism, raw.batch_rows) == (1, 0)
    raw = _native.to_native_write_options(WriteOptions(degree_of_parallelism=0, batch_rows=2))
    assert (raw.degree_of_parallelism, raw.batch_rows) == (0, 2)


def test_abi_version_matches_the_header():
    header = HEADER_PATH.read_text(encoding="utf-8")
    match = re.search(r"#define XL_ABI_VERSION (\d+)", header)
    assert match, "XL_ABI_VERSION not found in excelreader.h"
    header_version = int(match.group(1))

    assert _native.XL_ABI_VERSION == header_version
    assert _native.load_library().xl_abi_version() == header_version


class _Oversized:
    def __len__(self):
        return 2**31


def test_int32_len_rejects_a_length_the_abi_would_wrap():
    assert _native.int32_len(b"abc") == 3
    with pytest.raises(ValueError, match="int32"):
        _native.int32_len(_Oversized())


def test_cell_style_structs_match_the_header_layout():
    style = _native.NativeCellStyle
    assert ctypes.sizeof(style) == 104
    assert ctypes.sizeof(_native.NativeBorder) == 8
    offsets = {name: getattr(style, name).offset for name, _ in style._fields_}
    assert offsets == {
        "struct_size": 0, "number_format_len": 4, "number_format": 8, "font_name": 16, "font_name_len": 24,
        "font_flags": 28, "font_size": 32, "font_color": 40, "fill_color": 44, "horizontal_alignment": 48,
        "vertical_alignment": 52, "wrap_text": 56, "indent": 60, "rotation": 64,
        "left": 68, "right": 76, "top": 84, "bottom": 92,
    }
    options = _native.NativeSheetOptions
    assert ctypes.sizeof(options) == 64
    assert [options.column_styles.offset, options.column_widths.offset, options.header_style.offset] == [40, 48, 56]


def test_default_cell_style_marshals_every_color_as_unset():
    from excelreader import CellStyle

    raw = _native.to_native_cell_style(CellStyle())
    assert raw.struct_size == 104
    assert (raw.font_color, raw.fill_color) == (-1, -1)
    assert [edge.color for edge in (raw.left, raw.right, raw.top, raw.bottom)] == [-1] * 4
