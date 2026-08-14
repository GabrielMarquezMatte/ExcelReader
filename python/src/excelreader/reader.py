"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import struct
from array import array
from collections.abc import Iterator
from pathlib import Path

from typing_extensions import Self

from excelreader import _native
from excelreader.types import Cell, CellType, ColumnarSheet, ExcelReaderError

try:
    import numpy as _numpy
except ImportError:
    _numpy = None  # NumPy is an optional extra (pip install excelreader-native[numpy]) — see
    # read_all_columnar()/_to_columnar_arrays() below, the only two places that consult it.

_FORMATS = {
    "auto": _native.XL_FORMAT_AUTO,
    "xls": _native.XL_FORMAT_XLS,
    "xlsx": _native.XL_FORMAT_XLSX,
    "xlsb": _native.XL_FORMAT_XLSB,
    "csv": _native.XL_FORMAT_CSV,
}

_CELL_HEADER = struct.Struct("<iii")
_INITIAL_ROW_BUFFER = 64 * 1024
_INITIAL_ALL_ROWS_BUFFER = 1024 * 1024


def _last_error() -> str:
    # xl_last_error_ptr borrows a pointer straight into the native side's thread-local error buffer —
    # no ask-the-size-then-copy round trip. The pointer is only valid until the next ExcelReader call
    # on this thread, so it must be decoded immediately, before any other native call — which is
    # exactly how the sole caller below uses it.
    lib = _native.load_library()
    length = ctypes.c_int32()
    pointer = lib.xl_last_error_ptr(ctypes.byref(length))
    if not pointer:
        return ""
    return ctypes.string_at(pointer, length.value).decode("utf-8", errors="replace")


def _check(status: int) -> None:
    if status == _native.XL_OK:
        return
    if status == _native.XL_INVALID_HANDLE:
        raise ExcelReaderError("workbook is closed or the handle is invalid")
    if status == _native.XL_INVALID_ARGUMENT:
        raise ExcelReaderError("invalid argument passed to the native library")
    raise ExcelReaderError(_last_error() or f"native call failed with status {status}")


def _resolve_format(name: str | None, path: Path | None) -> int:
    if name is not None:
        try:
            return _FORMATS[name.lower()]
        except KeyError:
            raise ValueError(f"unknown format {name!r}; expected one of {sorted(_FORMATS)}") from None
    # The signature sniffer covers XLS/XLSX/XLSB but CSV has no signature, so the extension is the
    # only hint available for it.
    if path is not None and path.suffix.lower() == ".csv":
        return _native.XL_FORMAT_CSV
    return _native.XL_FORMAT_AUTO


class Workbook:
    """A read cursor over one workbook. Not thread-safe; use one instance per thread."""

    def __init__(self, handle: ctypes.c_void_p) -> None:
        self._lib = _native.load_library()
        self._handle: ctypes.c_void_p | None = handle

    def _require_handle(self) -> ctypes.c_void_p:
        if self._handle is None:
            raise ExcelReaderError("workbook is closed")
        return self._handle

    @property
    def sheet_count(self) -> int:
        count = ctypes.c_int32()
        _check(self._lib.xl_sheet_count(self._require_handle(), ctypes.byref(count)))
        return count.value

    @property
    def sheet_name(self) -> str:
        handle = self._require_handle()
        length = ctypes.c_int32()
        buffer = ctypes.create_string_buffer(256)
        status = self._lib.xl_sheet_name(handle, buffer, len(buffer), ctypes.byref(length))
        if status == _native.XL_BUFFER_TOO_SMALL:
            buffer = ctypes.create_string_buffer(length.value)
            status = self._lib.xl_sheet_name(handle, buffer, len(buffer), ctypes.byref(length))
        _check(status)
        return buffer.raw[: length.value].decode("utf-8")

    def sheet_name_at(self, index: int) -> str:
        """Name of the sheet at `index`, without changing the current sheet or disturbing row enumeration."""
        handle = self._require_handle()
        length = ctypes.c_int32()
        buffer = ctypes.create_string_buffer(256)
        status = self._lib.xl_sheet_name_at(handle, index, buffer, len(buffer), ctypes.byref(length))
        if status == _native.XL_BUFFER_TOO_SMALL:
            buffer = ctypes.create_string_buffer(length.value)
            status = self._lib.xl_sheet_name_at(handle, index, buffer, len(buffer), ctypes.byref(length))
        _check(status)
        return buffer.raw[: length.value].decode("utf-8")

    @property
    def sheet_names(self) -> list[str]:
        """Every sheet's name, in order. Does not change the current sheet or disturb row enumeration."""
        return [self.sheet_name_at(index) for index in range(self.sheet_count)]

    def sheets(self) -> Iterator[tuple[int, str]]:
        """(index, name) for every sheet, in order. Does not change the current sheet or disturb row enumeration."""
        for index in range(self.sheet_count):
            yield index, self.sheet_name_at(index)

    @property
    def is_date1904(self) -> bool:
        flag = ctypes.c_int32()
        _check(self._lib.xl_is_date1904(self._require_handle(), ctypes.byref(flag)))
        return flag.value != 0

    def move_to_sheet(self, index: int) -> None:
        """Selects a sheet and restarts row enumeration from its first row."""
        _check(self._lib.xl_move_to_sheet(self._require_handle(), index))

    def rows(self) -> Iterator[list[Cell]]:
        written = ctypes.c_int32()
        capacity = _INITIAL_ROW_BUFFER
        buffer = ctypes.create_string_buffer(capacity)
        while True:
            handle = self._require_handle()
            status = self._lib.xl_next_row(handle, buffer, capacity, ctypes.byref(written))
            if status == _native.XL_EOF:
                return
            if status == _native.XL_BUFFER_TOO_SMALL:
                # The native side holds the row until it fits, so growing loses nothing.
                capacity = written.value
                buffer = ctypes.create_string_buffer(capacity)
                continue
            _check(status)
            yield _decode_row(buffer.raw, written.value)

    def read_all(self) -> list[list[Cell]]:
        """Materializes every remaining row of the current sheet in one native call.

        Builds one `Cell`/`str` object per cell — for a large sheet, `read_all_columnar()` is
        substantially faster because it never does that.
        """
        handle = self._require_handle()
        rows = _native.NativeRows()
        _check(self._lib.xl_read_all_decoded(handle, ctypes.byref(rows)))
        try:
            return [_decode_native_row(rows.rows[index]) for index in range(rows.row_count)]
        finally:
            self._lib.xl_free_rows(ctypes.byref(rows))

    def read_all_columnar(self) -> ColumnarSheet:
        """Materializes every remaining row of the current sheet as parallel flat arrays.

        The fast path for large sheets: unlike `read_all()`, this never constructs a `Cell`/`str`
        per cell. Use `decode_cell()` to materialize one cell on demand.
        """
        written = ctypes.c_int32()
        capacity = _INITIAL_ALL_ROWS_BUFFER
        buffer = ctypes.create_string_buffer(capacity)
        handle = self._require_handle()
        status = self._lib.xl_read_all_blob(handle, buffer, capacity, ctypes.byref(written))
        if status == _native.XL_BUFFER_TOO_SMALL:
            # xl_read_all_blob guarantees the accumulated result is held, not lost, on a too-small
            # buffer — growing and retrying costs one copy, not a re-read.
            capacity = written.value
            buffer = ctypes.create_string_buffer(capacity)
            status = self._lib.xl_read_all_blob(handle, buffer, capacity, ctypes.byref(written))
        _check(status)
        return _decode_columnar(buffer.raw, written.value)

    def close(self) -> None:
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        _check(self._lib.xl_close(handle))

    def __enter__(self) -> Self:
        return self

    def __exit__(self, *_exc_info: object) -> None:
        self.close()

    def __del__(self) -> None:
        # Backstop only, not a substitute for explicit close()/`with`: if a Workbook is dropped
        # without one, this still releases the native handle and the file lock it holds. During
        # interpreter shutdown or GC, module globals (_native, ctypes) may already be partially torn
        # down, so a finalizer must never let an exception escape — swallow anything broadly here,
        # which is the standard, accepted exception to "never bare-except" for __del__ specifically.
        try:
            self.close()
        except Exception:  # noqa: BLE001, S110
            pass


def _decode_row(blob: bytes, length: int) -> list[Cell]:
    count = struct.unpack_from("<i", blob, 0)[0]
    cells: list[Cell] = []
    offset = 4
    for _ in range(count):
        column, cell_type, value_length = _CELL_HEADER.unpack_from(blob, offset)
        offset += _CELL_HEADER.size
        value = blob[offset : offset + value_length].decode("utf-8")
        offset += value_length
        cells.append(Cell(column=column, type=CellType(cell_type), value=value))
    if offset != length:
        raise ExcelReaderError(f"row blob is malformed: consumed {offset} of {length} bytes")
    return cells


def _decode_native_row(row: _native.NativeRow) -> list[Cell]:
    cells: list[Cell] = []
    for index in range(row.cell_count):
        cell = row.cells[index]
        value = ctypes.string_at(cell.value, cell.value_len).decode("utf-8")
        cells.append(Cell(column=cell.column, type=CellType(cell.type), value=value))
    return cells


def _decode_columnar(blob: bytes, length: int) -> ColumnarSheet:
    # xl_read_all_blob layout: int32 row_count, then row_count * {int32 row_length, row blob}, where
    # a row blob is int32 cell_count, then cell_count * {int32 column, int32 type, int32 value_len,
    # uint8 value[value_len]} (the same per-row shape xl_next_row/_decode_row already use).
    #
    # This loop only ever appends plain ints to array('i') and slices bytes — no Cell/str object is
    # constructed per cell, which is the entire point of this method over read_all(). A fully
    # vectorized (no-Python-loop) parse was considered and skipped: the values are variable-length
    # and interleaved with their headers, so finding cell boundaries needs a pass over the blob
    # regardless — int-list-append is already fast, the object construction was what was
    # expensive. (ponytail: if this loop measurably dominates for very wide/tall sheets, revisit with
    # a vectorized header scan; not attempted without a measurement showing it's needed.)
    row_offsets = array("i", [0])
    columns = array("i")
    types = array("i")
    value_offsets = array("i", [0])
    values = bytearray()

    row_count = struct.unpack_from("<i", blob, 0)[0]
    offset = 4
    cell_index = 0
    for _ in range(row_count):
        (row_length,) = struct.unpack_from("<i", blob, offset)
        offset += 4
        row_end = offset + row_length

        (cell_count,) = struct.unpack_from("<i", blob, offset)
        cell_offset = offset + 4
        for _ in range(cell_count):
            column, cell_type, value_length = _CELL_HEADER.unpack_from(blob, cell_offset)
            cell_offset += _CELL_HEADER.size
            values += blob[cell_offset : cell_offset + value_length]
            cell_offset += value_length
            columns.append(column)
            types.append(cell_type)
            value_offsets.append(len(values))
            cell_index += 1

        if cell_offset != row_end:
            raise ExcelReaderError(f"row blob is malformed: consumed {cell_offset - offset} of {row_length} bytes")
        row_offsets.append(cell_index)
        offset = row_end

    if offset != length:
        raise ExcelReaderError(f"all-rows blob is malformed: consumed {offset} of {length} bytes")

    return ColumnarSheet(
        row_offsets=_to_columnar_array(row_offsets),
        columns=_to_columnar_array(columns),
        types=_to_columnar_array(types),
        value_offsets=_to_columnar_array(value_offsets),
        values=bytes(values),
    )


def _to_columnar_array(values: array) -> object:
    # NumPy is optional; array('i') already gives every ColumnarSheet field a shared int32-ish
    # len()/index/slice interface either way, so callers don't need to branch on which one they got.
    if _numpy is None:
        return values
    return _numpy.frombuffer(values, dtype=_numpy.int32)


def decode_cell(sheet: ColumnarSheet, index: int) -> Cell:
    """Materializes the `Cell` at flat cell index `index` in `sheet`, decoding only that one value."""
    start, end = int(sheet.value_offsets[index]), int(sheet.value_offsets[index + 1])
    value = bytes(sheet.values[start:end]).decode("utf-8")
    return Cell(column=int(sheet.columns[index]), type=CellType(int(sheet.types[index])), value=value)


def open_workbook(path: str | Path, format: str | None = None) -> Workbook:
    """Opens a workbook from disk. `format` is one of auto/xls/xlsx/xlsb/csv; None infers it."""
    resolved = Path(path)
    encoded = str(resolved).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_file(encoded, len(encoded), _resolve_format(format, resolved), ctypes.byref(handle)))
    return Workbook(handle)


def open_bytes(data: bytes, format: str | None = None) -> Workbook:
    """Opens a workbook from an in-memory buffer. The native side copies `data` immediately."""
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_memory(data, len(data), _resolve_format(format, None), ctypes.byref(handle)))
    return Workbook(handle)
