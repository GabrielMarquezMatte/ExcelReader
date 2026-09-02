"""The streaming writer: one sheet and one row open at a time, written as each call arrives.

Mirrors the xl_writer_handle call order documented in
src/ExcelReader.Native/include/excelreader.h. `write_workbook()` in writer.py is the columnar
alternative, which builds the whole table first.
"""

from __future__ import annotations

import ctypes
import datetime
from pathlib import Path
from typing import Any, Sequence

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import ColumnType, WriteOptions

_EPOCH_DATE = datetime.date(1970, 1, 1)
_EPOCH_DATETIME = datetime.datetime(1970, 1, 1)


def _encode(value: str) -> tuple[Any, int]:
    raw = value.encode("utf-8")
    return (ctypes.cast(ctypes.c_char_p(raw), ctypes.POINTER(ctypes.c_uint8)), len(raw))


class SheetWriter:
    """A streaming workbook writer. Use as a context manager; closing finalizes the file.

    The required call order is `start_sheet` -> `start_row` -> one write per cell -> `end_row` ->
    `end_sheet`. Calling out of order raises `ExcelReaderError`, and the writer stays usable: fix
    the order and continue, or close to discard.
    """

    __slots__ = ("_handle", "_lib", "_in_memory")

    def __init__(self, handle: ctypes.c_void_p, in_memory: bool) -> None:
        self._handle: ctypes.c_void_p | None = handle
        self._lib = _native.load_library()
        self._in_memory = in_memory

    def _require_handle(self) -> ctypes.c_void_p:
        if self._handle is None:
            raise ValueError("this SheetWriter is closed")
        return self._handle

    def start_sheet(self, name: str) -> None:
        pointer, length = _encode(name)
        _check(self._lib.xl_start_sheet(self._require_handle(), pointer, length))

    def start_row(self) -> None:
        _check(self._lib.xl_start_row(self._require_handle()))

    def end_row(self) -> None:
        _check(self._lib.xl_end_row(self._require_handle()))

    def end_sheet(self) -> None:
        _check(self._lib.xl_end_sheet(self._require_handle()))

    def write_str(self, value: str | None) -> None:
        """Writes the next cell as text, or a blank cell for None."""
        handle = self._require_handle()
        if value is None:
            _check(self._lib.xl_write_string(handle, None, 0))
            return
        pointer, length = _encode(value)
        _check(self._lib.xl_write_string(handle, pointer, length))

    def write_i64(self, value: int | None) -> None:
        if value is None:
            self.write_null(ColumnType.I64)
            return
        _check(self._lib.xl_write_int64(self._require_handle(), value))

    def write_f64(self, value: float | None) -> None:
        if value is None:
            self.write_null(ColumnType.F64)
            return
        _check(self._lib.xl_write_float64(self._require_handle(), value))

    def write_bool(self, value: bool | None) -> None:
        if value is None:
            self.write_null(ColumnType.BOOL)
            return
        _check(self._lib.xl_write_bool(self._require_handle(), 1 if value else 0))

    def write_date(self, value: datetime.date | None) -> None:
        if value is None:
            self.write_null(ColumnType.DATE)
            return
        _check(self._lib.xl_write_date(self._require_handle(), (value - _EPOCH_DATE).days))

    def write_time(self, value: datetime.time | None) -> None:
        if value is None:
            self.write_null(ColumnType.TIME)
            return
        micros = (
            (value.hour * 3600 + value.minute * 60 + value.second) * 1_000_000 + value.microsecond
        )
        _check(self._lib.xl_write_time(self._require_handle(), micros))

    def write_timestamp(self, value: datetime.datetime | None) -> None:
        if value is None:
            self.write_null(ColumnType.TIMESTAMP)
            return
        # A tz-aware value is converted to UTC first, because the ABI's wire format is
        # microseconds since the Unix epoch in UTC.
        if value.tzinfo is not None:
            value = value.astimezone(datetime.timezone.utc).replace(tzinfo=None)
        delta = value - _EPOCH_DATETIME
        micros = (delta.days * 86_400 + delta.seconds) * 1_000_000 + delta.microseconds
        _check(self._lib.xl_write_timestamp(self._require_handle(), micros))

    def write_null(self, column_type: ColumnType) -> None:
        """Writes a blank cell typed as `column_type`."""
        _check(self._lib.xl_write_null(self._require_handle(), int(column_type)))

    def close(self) -> None:
        """Finalizes and releases the writer. Safe to call more than once."""
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        # xl_close_write_handle always releases the handle, including on error, so the field is
        # cleared before the call rather than after.
        _check(self._lib.xl_close_write_handle(handle))

    def __enter__(self) -> "SheetWriter":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()


def open_writer(
    path: str | Path,
    format: str | None = None,
    options: WriteOptions | None = None,
) -> SheetWriter:
    """Opens a streaming writer for `path`, truncating any existing file.

    `format` is one of xlsx/xlsb/xls/csv; None infers it from the path's extension.
    """
    from excelreader.writer import _resolve_write_format

    resolved = Path(path)
    format_id = _resolve_write_format(format, resolved)
    raw_options = _native.to_native_write_options(options or WriteOptions())

    encoded = str(resolved).encode("utf-8")
    pointer = ctypes.cast(ctypes.c_char_p(encoded), ctypes.POINTER(ctypes.c_uint8))
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(
        lib.xl_open_write_handle(
            pointer, len(encoded), format_id, ctypes.byref(raw_options), ctypes.byref(handle)
        )
    )
    return SheetWriter(handle, in_memory=False)
