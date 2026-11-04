"""The streaming writer: one sheet and one row open at a time, written as each call arrives.

Mirrors the xl_writer_handle call order documented in
src/ExcelReader.Native/include/excelreader.h. `write_sheet()` is the columnar
alternative, which builds the whole table first.
"""

from __future__ import annotations

import ctypes
import datetime
import os
from pathlib import Path
from typing import IO, Any, Mapping, Sequence, Union

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import CellStyle, ColumnType, TableOptions, WriteOptions
from excelreader.writer import _marshal, _typed_table, _write_all

_EPOCH_DATE = datetime.date(1970, 1, 1)
_EPOCH_DATETIME = datetime.datetime(1970, 1, 1)
_INT64_MIN = -(2**63)
_INT64_MAX = 2**63 - 1


def _encode(value: str) -> tuple[Any, int]:
    raw = value.encode("utf-8")
    return (ctypes.cast(ctypes.c_char_p(raw), ctypes.POINTER(ctypes.c_uint8)), _native.int32_len(raw))


def _is_writable(value: Any) -> bool:
    return value is None or isinstance(
        value, (bool, int, float, str, datetime.datetime, datetime.date, datetime.time)
    )


def _is_typed_blank(value: Any) -> bool:
    return type(value) is tuple and len(value) == 2 and value[0] is None and isinstance(value[1], ColumnType)


class SheetWriter:
    """A streaming workbook writer. Use as a context manager; closing finalizes the workbook, and for
    a destination with `write()` hands it the finished bytes.

    Only `close()`, or leaving a `with` block normally, commits. Leaving it by an exception, `abort()`,
    or a writer collected without being closed discards what was written: an existing file at the
    path is left as it was, and a destination with `write()` gets nothing.

    The required call order is `start_sheet` -> `write_row` per row -> `end_sheet`, or one
    `write_sheet` per whole table. Calling out of order raises `ExcelReaderError`, and the writer
    stays usable: fix the order and continue, or close to discard.
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

    def write_sheet(
        self,
        name: str,
        data: Any,
        types: list[ColumnType] | None = None,
        *,
        table: TableOptions | None = None,
        column_styles: Mapping[str, CellStyle] | None = None,
        header_style: CellStyle | None = None,
        column_widths: Mapping[str, float] | None = None,
    ) -> None:
        """Writes `data` as a whole sheet named `name`, in one native call instead of one per cell.

        `data` and `types` follow `write_workbook`'s rules: a `TypedTable` with one `ColumnType` per
        column, or Arrow-compatible data (pyarrow, pandas, polars) with `types` None. The column names
        become the header row. With `table`, the sheet is written as an Excel table (xlsx/xlsb only).
        Parallelism and the other options come from `open_writer`; `WriteOptions.sheet_name` does not
        apply. Raises `ExcelReaderError` while a sheet started with `start_sheet` is open.

        `column_styles` and `column_widths` (in characters) are keyed by column name; a column left
        out keeps its default, and a name that is not a column raises `ValueError`. A column style
        reaches the data cells only, never the header row, and one without a `number_format` keeps a
        DATE, TIME or TIMESTAMP column's format. `header_style` styles the header row.
        """
        typed, types = _typed_table(data, types)
        unknown = [key for mapping in (column_styles, column_widths) for key in mapping or () if key not in typed.names]
        if unknown:
            raise ValueError(f"unknown column name {unknown[0]!r}; the columns are {typed.names}")
        specs, native_table, _, keepalive = _marshal(typed, types, None)
        pointer, length = _encode(name)
        options = _native.to_native_sheet_options(
            table,
            None if column_styles is None else [column_styles.get(column) for column in typed.names],
            header_style,
            None if column_widths is None else [column_widths.get(column, 0.0) for column in typed.names],
        )
        _check(
            self._lib.xl_write_sheet(
                self._require_handle(), pointer, length, specs, ctypes.byref(native_table),
                None if options is None else ctypes.byref(options),
            )
        )
        del keepalive

    def add_style(self, style: CellStyle) -> int:
        """Registers `style` and returns its id; an equal style returns the same id, and id 0 is the
        general style, which needs no registering. Raises `ExcelReaderError` for an invalid value or
        a format limit, leaving the writer usable.
        """
        raw = _native.to_native_cell_style(style)
        style_id = ctypes.c_int32()
        _check(self._lib.xl_add_style(self._require_handle(), ctypes.byref(raw), ctypes.byref(style_id)))
        return style_id.value

    def set_column_style(self, column: int, style_id: int) -> None:
        """Styles every cell of the 0-based `column` of the open sheet; call it before the sheet's first row."""
        _check(self._lib.xl_set_column_style(self._require_handle(), column, style_id))

    def set_column_width(self, column: int, width: float) -> None:
        """Sets the 0-based `column`'s width in characters (0-255); call it before the sheet's first row."""
        _check(self._lib.xl_set_column_width(self._require_handle(), column, width))

    def write_row(
        self,
        values: Sequence[Any],
        *,
        style_id: int = 0,
        height: float = 0,
        cell_styles: Sequence[int | None] | None = None,
    ) -> None:
        """Writes one row, one cell per value, its type inferred from the value.

        `style_id` styles every cell of the row, beating column styles, and `height` is the row's
        height in points (0 for the default). A non-None `cell_styles[i]` styles value `i` alone,
        over the row and column styles; its length must equal `len(values)`, else `ValueError`.

        `str` is text, `bool` a boolean, `int` a 64-bit integer, `float` a number, `datetime` a
        timestamp (an aware one converted to UTC), `date` a date and `time` a time of day. `None` is a
        blank text cell; the tuple `(None, ColumnType.X)` is a blank cell typed as X.

        Raises, before writing anything, `ValueError` for a `cell_styles` of the wrong length,
        `TypeError` for a value whose type has no mapping and `OverflowError` for an `int` outside the
        signed 64-bit range.
        """
        if cell_styles is not None and len(cell_styles) != len(values):
            raise ValueError(f"cell_styles has {len(cell_styles)} entries but values has {len(values)}")
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
        _check(self._lib.xl_start_row_styled(handle, style_id, height))
        try:
            for position, value in enumerate(values):
                if cell_styles is not None and cell_styles[position] is not None:
                    _check(self._lib.xl_style_next_cell(handle, cell_styles[position]))
                _check(self._write_cell(handle, value))
        except BaseException:
            self._lib.xl_end_row(handle)
            raise
        _check(self._lib.xl_end_row(handle))

    def _write_cell(self, handle: ctypes.c_void_p, value: Any) -> int:
        lib = self._lib
        kind = type(value)
        if kind is str:
            raw = value.encode("utf-8")
            return lib.xl_write_string(handle, raw, _native.int32_len(raw))
        if kind is int:
            return lib.xl_write_int64(handle, value)
        if kind is float:
            return lib.xl_write_float64(handle, value)
        if kind is bool:
            return lib.xl_write_bool(handle, 1 if value else 0)
        if value is None:
            return lib.xl_write_string(handle, None, 0)
        if kind is tuple:
            return lib.xl_write_null(handle, int(value[1]))
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
        if isinstance(value, bool):
            return lib.xl_write_bool(handle, 1 if value else 0)
        if isinstance(value, int):
            return lib.xl_write_int64(handle, value)
        if isinstance(value, float):
            return lib.xl_write_float64(handle, value)
        raw = value.encode("utf-8")
        return lib.xl_write_string(handle, raw, _native.int32_len(raw))

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

        For a destination with `write()`, the finished workbook is written to it, even when sheets
        or rows were left open; a `write()` that takes only part of it is called again for the rest.
        """
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        destination, self._destination = self._destination, None
        try:
            if destination is not None:
                _write_all(destination, self._finished_bytes(handle))
        finally:
            _check(self._lib.xl_close_write_handle(handle))

    def abort(self) -> None:
        """Discards everything written and releases the writer, leaving an existing file at the path
        as it was and handing a destination with `write()` nothing. Safe to call more than once, and
        after `close()`.
        """
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        self._destination = None
        self._lib.xl_abort_write_handle(handle)

    def __enter__(self) -> "SheetWriter":
        return self

    def __exit__(self, exc_type: object, *exc: object) -> None:
        if exc_type is not None:
            self.abort()
            return
        self.close()

    def __del__(self) -> None:
        try:
            self.abort()
        except Exception:  # noqa: BLE001, S110
            pass


def open_writer(
    destination: Union[str, os.PathLike, IO[bytes]],
    format: str | None = None,
    *,
    options: WriteOptions | None = None,
) -> SheetWriter:
    """Opens a streaming writer for `destination`: a path, whose existing file is replaced when the
    writer closes, or a binary object with `write()` such as `io.BytesIO` or an open file.

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
                pointer, _native.int32_len(encoded), format_id, ctypes.byref(raw_options), ctypes.byref(handle)
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
