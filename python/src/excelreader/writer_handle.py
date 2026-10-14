"""The streaming writer: one sheet and one row open at a time, written as each call arrives.

Mirrors the xl_writer_handle call order documented in
src/ExcelReader.Native/include/excelreader.h. `write_workbook()` in writer.py is the columnar
alternative, which builds the whole table first.
"""

from __future__ import annotations

import ctypes
import datetime
import os
from pathlib import Path
from typing import IO, Any, Sequence, Union

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import ColumnType, WriteOptions

_EPOCH_DATE = datetime.date(1970, 1, 1)
_EPOCH_DATETIME = datetime.datetime(1970, 1, 1)
_INT64_MIN = -(2**63)
_INT64_MAX = 2**63 - 1


def _encode(value: str) -> tuple[Any, int]:
    raw = value.encode("utf-8")
    return (ctypes.cast(ctypes.c_char_p(raw), ctypes.POINTER(ctypes.c_uint8)), len(raw))


def _is_writable(value: Any) -> bool:
    return value is None or isinstance(
        value, (bool, int, float, str, datetime.datetime, datetime.date, datetime.time)
    )


def _is_typed_blank(value: Any) -> bool:
    return type(value) is tuple and len(value) == 2 and value[0] is None and isinstance(value[1], ColumnType)


class SheetWriter:
    """A streaming workbook writer. Use as a context manager; closing finalizes the workbook, and for
    a destination with `write()` hands it the finished bytes.

    The required call order is `start_sheet` -> `write_row` per row -> `end_sheet`. Calling out of
    order raises `ExcelReaderError`, and the writer stays usable: fix the order and continue, or
    close to discard.
    """

    __slots__ = ("_handle", "_lib", "_destination")

    def __init__(self, handle: ctypes.c_void_p, destination: IO[bytes] | None) -> None:
        self._handle: ctypes.c_void_p | None = handle
        self._lib = _native.load_library()
        self._destination = destination

    def _require_handle(self) -> ctypes.c_void_p:
        if self._handle is None:
            raise ValueError("this SheetWriter is closed")
        return self._handle

    def start_sheet(self, name: str) -> None:
        pointer, length = _encode(name)
        _check(self._lib.xl_start_sheet(self._require_handle(), pointer, length))

    def end_sheet(self) -> None:
        _check(self._lib.xl_end_sheet(self._require_handle()))

    def write_row(self, values: Sequence[Any]) -> None:
        """Writes one row, one cell per value, its type inferred from the value.

        `str` is text, `bool` a boolean, `int` a 64-bit integer, `float` a number, `datetime` a
        timestamp (an aware one converted to UTC), `date` a date and `time` a time of day. `None` is a
        blank text cell; the tuple `(None, ColumnType.X)` is a blank cell typed as X.

        Raises, before writing anything, `TypeError` for a value whose type has no mapping and
        `OverflowError` for an `int` outside the signed 64-bit range.
        """
        for position, value in enumerate(values):
            if not _is_writable(value) and not _is_typed_blank(value):
                raise TypeError(
                    f"write_row cannot infer a cell type for {type(value).__name__} "
                    f"at position {position}"
                )
            # ctypes would keep the low 64 bits of a larger value without complaint.
            if type(value) is int and not _INT64_MIN <= value <= _INT64_MAX:
                raise OverflowError(f"{value} at position {position} does not fit a 64-bit integer cell")

        handle = self._require_handle()
        _check(self._lib.xl_start_row(handle))
        for value in values:
            _check(self._write_cell(handle, value))
        _check(self._lib.xl_end_row(handle))

    def _write_cell(self, handle: ctypes.c_void_p, value: Any) -> int:
        lib = self._lib
        if value is None:
            return lib.xl_write_string(handle, None, 0)
        if isinstance(value, tuple):
            return lib.xl_write_null(handle, int(value[1]))
        if isinstance(value, bool):
            return lib.xl_write_bool(handle, 1 if value else 0)
        if isinstance(value, int):
            return lib.xl_write_int64(handle, value)
        if isinstance(value, float):
            return lib.xl_write_float64(handle, value)
        if isinstance(value, datetime.datetime):
            if value.tzinfo is not None:
                value = value.astimezone(datetime.timezone.utc).replace(tzinfo=None)
            delta = value - _EPOCH_DATETIME
            return lib.xl_write_timestamp(handle, (delta.days * 86_400 + delta.seconds) * 1_000_000 + delta.microseconds)
        if isinstance(value, datetime.date):
            return lib.xl_write_date(handle, (value - _EPOCH_DATE).days)
        if isinstance(value, datetime.time):
            return lib.xl_write_time(
                handle, (value.hour * 3600 + value.minute * 60 + value.second) * 1_000_000 + value.microsecond
            )
        pointer, length = _encode(value)
        return lib.xl_write_string(handle, pointer, length)

    def _finished_bytes(self, handle: ctypes.c_void_p) -> bytes:
        buffer = _native.NativeBuffer()
        _check(self._lib.xl_write_handle_bytes(handle, ctypes.byref(buffer)))
        try:
            if not buffer.data or buffer.len <= 0:
                return b""
            return ctypes.string_at(buffer.data, buffer.len)
        finally:
            self._lib.xl_free_buffer(ctypes.byref(buffer))

    def close(self) -> None:
        """Finalizes and releases the writer. Safe to call more than once.

        For a destination with `write()`, the finished workbook is written to it in one call, even
        when sheets or rows were left open.
        """
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        destination, self._destination = self._destination, None
        try:
            if destination is not None:
                destination.write(self._finished_bytes(handle))
        finally:
            _check(self._lib.xl_close_write_handle(handle))

    def __enter__(self) -> "SheetWriter":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:  # noqa: BLE001, S110
            pass


def open_writer(
    destination: Union[str, os.PathLike, IO[bytes]],
    format: str | None = None,
    options: WriteOptions | None = None,
) -> SheetWriter:
    """Opens a streaming writer for `destination`: a path, truncating any existing file, or a binary
    object with `write()` such as `io.BytesIO` or an open file.

    `format` is one of xlsx/xlsb/xls/csv. For a path, None infers it from the extension; an object
    with `write()` needs it. Such an object receives the whole workbook in one `write()` call when the
    writer closes, and is not closed itself.
    """
    from excelreader.writer import _resolve_write_format

    raw_options = _native.to_native_write_options(options or WriteOptions())
    handle = ctypes.c_void_p()
    lib = _native.load_library()

    if isinstance(destination, (str, os.PathLike)):
        resolved = Path(destination)
        format_id = _resolve_write_format(format, resolved)
        encoded = str(resolved).encode("utf-8")
        pointer = ctypes.cast(ctypes.c_char_p(encoded), ctypes.POINTER(ctypes.c_uint8))
        _check(
            lib.xl_open_write_handle(
                pointer, len(encoded), format_id, ctypes.byref(raw_options), ctypes.byref(handle)
            )
        )
        return SheetWriter(handle, None)

    if not hasattr(destination, "write"):
        raise TypeError(
            f"destination must be a path or an object with write(); got {type(destination).__name__}"
        )
    if format is None:
        raise ValueError(
            "format= is required when writing to an object with write(); "
            f"one of {sorted(_native.WRITE_FORMATS)}"
        )
    try:
        format_id = _native.WRITE_FORMATS[format.lower()]
    except KeyError:
        raise ValueError(f"unknown format {format!r}; expected one of {sorted(_native.WRITE_FORMATS)}") from None
    _check(lib.xl_open_write_handle_to_memory(format_id, ctypes.byref(raw_options), ctypes.byref(handle)))
    return SheetWriter(handle, destination)
