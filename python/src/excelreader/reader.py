"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import struct
from pathlib import Path
from typing import Iterator, Optional, Union

from excelreader import _native
from excelreader.types import Cell, CellType, ExcelReaderError

_FORMATS = {
    "auto": _native.XL_FORMAT_AUTO,
    "xls": _native.XL_FORMAT_XLS,
    "xlsx": _native.XL_FORMAT_XLSX,
    "xlsb": _native.XL_FORMAT_XLSB,
    "csv": _native.XL_FORMAT_CSV,
}

_CELL_HEADER = struct.Struct("<iii")
_INITIAL_ROW_BUFFER = 64 * 1024


def _last_error() -> str:
    lib = _native.load_library()
    length = ctypes.c_int32()
    buffer = ctypes.create_string_buffer(1024)
    if lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_BUFFER_TOO_SMALL:
        buffer = ctypes.create_string_buffer(length.value)
        lib.xl_last_error(buffer, len(buffer), ctypes.byref(length))
    return buffer.raw[: length.value].decode("utf-8", errors="replace")


def _check(status: int) -> None:
    if status == _native.XL_OK:
        return
    if status == _native.XL_INVALID_HANDLE:
        raise ExcelReaderError("workbook is closed or the handle is invalid")
    if status == _native.XL_INVALID_ARGUMENT:
        raise ExcelReaderError("invalid argument passed to the native library")
    raise ExcelReaderError(_last_error() or f"native call failed with status {status}")


def _resolve_format(name: Optional[str], path: Optional[Path]) -> int:
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
        self._handle: Optional[ctypes.c_void_p] = handle

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

    def close(self) -> None:
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        _check(self._lib.xl_close(handle))

    def __enter__(self) -> "Workbook":
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
        except Exception:
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


def open_workbook(path: Union[str, Path], format: Optional[str] = None) -> Workbook:
    """Opens a workbook from disk. `format` is one of auto/xls/xlsx/xlsb/csv; None infers it."""
    resolved = Path(path)
    encoded = str(resolved).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_file(encoded, len(encoded), _resolve_format(format, resolved), ctypes.byref(handle)))
    return Workbook(handle)


def open_bytes(data: bytes, format: Optional[str] = None) -> Workbook:
    """Opens a workbook from an in-memory buffer. The native side copies `data` immediately."""
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_memory(data, len(data), _resolve_format(format, None), ctypes.byref(handle)))
    return Workbook(handle)
