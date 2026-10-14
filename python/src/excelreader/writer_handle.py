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


class SheetWriter:
    """A streaming workbook writer. Use as a context manager; closing finalizes the workbook, and for
    a destination with `write()` hands it the finished bytes.

    The required call order is `start_sheet` -> `start_row` -> one write per cell -> `end_row` ->
    `end_sheet`. Calling out of order raises `ExcelReaderError`, and the writer stays usable: fix
    the order and continue, or close to discard.
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
        """Writes the next cell as a whole number. Raises `OverflowError` outside the signed 64-bit range."""
        if value is None:
            self.write_null(ColumnType.I64)
            return
        # ctypes would keep the low 64 bits of a larger value without complaint.
        if not _INT64_MIN <= value <= _INT64_MAX:
            raise OverflowError(f"{value} does not fit a 64-bit integer cell; write it with write_f64 or write_str")
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
        if value.tzinfo is not None:
            value = value.astimezone(datetime.timezone.utc).replace(tzinfo=None)
        delta = value - _EPOCH_DATETIME
        micros = (delta.days * 86_400 + delta.seconds) * 1_000_000 + delta.microseconds
        _check(self._lib.xl_write_timestamp(self._require_handle(), micros))

    def write_null(self, column_type: ColumnType) -> None:
        """Writes a blank cell typed as `column_type`."""
        _check(self._lib.xl_write_null(self._require_handle(), int(column_type)))

    def write_row(self, values: Sequence[Any]) -> None:
        """`start_row()`, one inferred write per value, `end_row()`.

        The type of each value picks the write method. `None` becomes a blank *string* cell — it is
        the one case inference cannot resolve; use `write_null(ColumnType.X)` for a typed blank.

        Raises, before writing anything, `TypeError` for a value whose type has no mapping and
        `OverflowError` for an `int` outside the signed 64-bit range.
        """
        for position, value in enumerate(values):
            if not _is_writable(value):
                raise TypeError(
                    f"write_row cannot infer a cell type for {type(value).__name__} "
                    f"at position {position}"
                )
            if type(value) is int and not _INT64_MIN <= value <= _INT64_MAX:
                raise OverflowError(f"{value} at position {position} does not fit a 64-bit integer cell")

        self.start_row()
        for value in values:
            self._write_inferred(value)
        self.end_row()

    def _write_inferred(self, value: Any) -> None:
        if value is None:
            self.write_str(None)
        elif isinstance(value, bool):
            self.write_bool(value)
        elif isinstance(value, int):
            self.write_i64(value)
        elif isinstance(value, float):
            self.write_f64(value)
        elif isinstance(value, datetime.datetime):
            self.write_timestamp(value)
        elif isinstance(value, datetime.date):
            self.write_date(value)
        elif isinstance(value, datetime.time):
            self.write_time(value)
        else:
            self.write_str(value)

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
