"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import struct
from array import array
from collections.abc import Iterator, Sequence
from pathlib import Path

from typing_extensions import Self

from excelreader import _native
from excelreader.types import (
    Cell,
    CellType,
    ColumnarSheet,
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    OpenOptions,
    StringColumn,
    TypedTable,
)

try:
    import numpy as _numpy
except ImportError:
    _numpy = None  # NumPy is an optional extra (pip install excelreader-native[numpy]) — see
    # _buffer_to_array()/_to_columnar_array() below, the only two places that consult it.

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

    def _fill_buffer(self, fn: object, *args: object, initial: int) -> tuple[bytes, int]:
        # Every buffer-returning export shares one convention: on XL_BUFFER_TOO_SMALL it reports the
        # size it needs through the out-length and holds the result, so growing and retrying once is
        # always enough and never re-reads.
        length = ctypes.c_int32()
        buffer = ctypes.create_string_buffer(initial)
        status = fn(*args, buffer, len(buffer), ctypes.byref(length))
        if status == _native.XL_BUFFER_TOO_SMALL:
            buffer = ctypes.create_string_buffer(length.value)
            status = fn(*args, buffer, len(buffer), ctypes.byref(length))
        _check(status)
        return buffer.raw, length.value

    @property
    def sheet_name(self) -> str:
        raw, length = self._fill_buffer(self._lib.xl_sheet_name, self._require_handle(), initial=256)
        return raw[:length].decode("utf-8")

    def sheet_name_at(self, index: int) -> str:
        """Name of the sheet at `index`, without changing the current sheet or disturbing row enumeration."""
        raw, length = self._fill_buffer(self._lib.xl_sheet_name_at, self._require_handle(), index, initial=256)
        return raw[:length].decode("utf-8")

    @property
    def sheet_names(self) -> list[str]:
        """Every sheet's name, in order. Does not change the current sheet or disturb row enumeration."""
        return [name for _, name in self.sheets()]

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
        raw, written = self._fill_buffer(
            self._lib.xl_read_all_blob, self._require_handle(), initial=_INITIAL_ALL_ROWS_BUFFER
        )
        return _decode_columnar(raw, written)

    def parse_typed(self, schema: Sequence[ColumnSpec], header_row: int = 1) -> TypedTable:
        """Reads the whole current sheet into typed columns, converting values on the native side.

        By far the fastest path in this library: unlike every other read method, no cell is ever
        formatted to text, so a large sheet costs a fraction of `read_all_columnar()`. It reads the
        sheet from its first row regardless of how far `rows()` has advanced, and does not disturb
        that cursor.

        `header_row` is a 1-based row number whose values name the columns (that row is skipped and
        never yielded as data); 0 means the sheet has no header, in which case every spec must
        resolve by `index`.
        """
        handle = self._require_handle()
        specs = _build_specs(schema)
        table = _native.NativeTable()
        _check(self._lib.xl_parse_typed(handle, specs, len(specs), header_row, ctypes.byref(table)))
        try:
            return _decode_table(schema, table)
        finally:
            # Every column is copied out above, so the native table is dead the moment we return —
            # nothing this method hands back points into it.
            self._lib.xl_free_table(ctypes.byref(table))

    def to_arrow(self, schema: Sequence[ColumnSpec], header_row: int = 1) -> object:
        """The same read as `parse_typed()`, handed to pyarrow as one `StructArray`, zero-copy.

        Requires pyarrow. The returned array owns the native buffers through the Arrow C Data
        Interface's release callback, so it stays valid after this workbook is closed. Wrap it with
        `pyarrow.RecordBatch.from_struct_array()` for a column-named batch.
        """
        try:
            import pyarrow
        except ImportError:
            raise ImportError(
                "to_arrow() requires pyarrow — install it with `pip install pyarrow`, or use "
                "parse_typed(), which returns the same data with no third-party dependency."
            ) from None

        handle = self._require_handle()
        specs = _build_specs(schema)
        array = _native.ArrowArray()
        arrow_schema = _native.ArrowSchema()
        _check(self._lib.xl_parse_arrow(handle, specs, len(specs), header_row, ctypes.byref(array), ctypes.byref(arrow_schema)))
        # _import_from_c takes ownership of both structs' release callbacks — releasing either one
        # here as well would be a double free. On failure _check raised above and the native side
        # exported nothing, so there is no leak on that path either.
        return pyarrow.Array._import_from_c(ctypes.addressof(array), ctypes.addressof(arrow_schema))

    def infer_schema(self, header_row: int = 1, sample_size: int = 100) -> list[ColumnSpec]:
        """Guesses a `parse_typed()`/`to_arrow()` schema by sampling this sheet's cells.

        Reads `header_row` for column names (0 means no header — every returned spec resolves by
        `index` instead) and up to `sample_size` rows after it, guessing each column's type from
        Excel's own per-cell type tag — no text sniffing. A column with a real mix of kinds, only
        formula/error results, or nothing sampled falls back to `ColumnType.STRING`; `nullable` is set
        when any sampled row left the column empty.

        This is a guess over a sample, not a guarantee — a column that looks like `ColumnType.I64` in
        the sample can still hold a fractional value further down the sheet, which `parse_typed()`
        would then reject unless the spec is `nullable`. Reads from the sheet's first row regardless of
        how far `rows()` has advanced, and does not disturb that cursor — same as `parse_typed()`.
        """
        handle = self._require_handle()
        schema = _native.NativeInferredSchema()
        _check(self._lib.xl_infer_schema(handle, header_row, sample_size, ctypes.byref(schema)))
        try:
            return _decode_inferred_schema(schema)
        finally:
            self._lib.xl_free_schema(ctypes.byref(schema))

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


def _decode_inferred_schema(schema: _native.NativeInferredSchema) -> list[ColumnSpec]:
    specs: list[ColumnSpec] = []
    for index in range(schema.column_count):
        raw = schema.columns[index]
        name = ctypes.string_at(raw.name, raw.name_len).decode("utf-8") if raw.name else None
        specs.append(ColumnSpec(ColumnType(raw.type), name=name, index=raw.index, nullable=bool(raw.nullable)))
    return specs


def _build_specs(schema: Sequence[ColumnSpec]) -> ctypes.Array:
    if not schema:
        raise ValueError("schema must name at least one column")
    specs = (_native.NativeColumnSpec * len(schema))()
    for index, spec in enumerate(schema):
        if spec.name is None:
            specs[index] = _native.column_spec_by_index(spec.index, int(spec.type), nullable=spec.nullable)
        else:
            specs[index] = _native.column_spec_by_name(spec.name, int(spec.type), nullable=spec.nullable)
    return specs


# ColumnType -> (ctypes element type, array.array typecode, NumPy dtype name). STRING is absent on
# purpose: it is the one type whose `values` buffer is offsets rather than data, handled separately.
_COLUMN_BUFFERS = {
    ColumnType.I64: (ctypes.c_int64, "q", "int64"),
    ColumnType.F64: (ctypes.c_double, "d", "float64"),
    ColumnType.BOOL: (ctypes.c_uint8, "b", "int8"),
    ColumnType.DATE: (ctypes.c_int32, "i", "int32"),
    ColumnType.TIME: (ctypes.c_int64, "q", "int64"),
    ColumnType.TIMESTAMP: (ctypes.c_int64, "q", "int64"),
}


def _decode_table(schema: Sequence[ColumnSpec], table: _native.NativeTable) -> TypedTable:
    row_count = int(table.row_count)
    names: list[str] = []
    columns: list[object] = []
    validity: list[bytes | None] = []

    for index, spec in enumerate(schema):
        column = table.columns[index]
        names.append(spec.name if spec.name is not None else str(spec.index))
        if ColumnType(column.type) is ColumnType.STRING:
            columns.append(_decode_string_column(column, row_count))
        else:
            columns.append(_decode_value_column(column, row_count))
        validity.append(_decode_validity(column, row_count))

    return TypedTable(row_count=row_count, names=names, columns=columns, validity=validity)


def _buffer_to_array(raw: bytes, typecode: str, dtype: str) -> object:
    # NumPy is optional, so every buffer decoded here lands as either an ndarray or an array.array;
    # both give callers the same len()/index/slice interface, so nothing downstream has to branch.
    # frombuffer over `raw` is a view on that bytes object, which keeps it alive — the native block
    # it was copied from is already out of the picture by then.
    if _numpy is not None:
        return _numpy.frombuffer(raw, dtype=dtype)
    values = array(typecode)
    values.frombytes(raw)
    return values


def _decode_value_column(column: _native.NativeColumn, row_count: int) -> object:
    element, typecode, dtype = _COLUMN_BUFFERS[ColumnType(column.type)]
    raw = ctypes.string_at(column.values, row_count * ctypes.sizeof(element))
    return _buffer_to_array(raw, typecode, dtype)


def _decode_string_column(column: _native.NativeColumn, row_count: int) -> StringColumn:
    # xl_column's STRING layout: `values` is (row_count + 1) int32 offsets, and `data` points just
    # past them INTO THE SAME allocation (see excelreader.h) — one copy each, not one per row.
    offsets_bytes = ctypes.string_at(column.values, (row_count + 1) * ctypes.sizeof(ctypes.c_int32))
    data = ctypes.string_at(column.data, int(column.data_len)) if column.data_len else b""
    return StringColumn(_buffer_to_array(offsets_bytes, "i", "int32"), data)


def _decode_validity(column: _native.NativeColumn, row_count: int) -> bytes | None:
    # A NULL validity pointer is the native side's "this column has no nulls" signal, not an error.
    if not column.validity:
        return None
    return ctypes.string_at(column.validity, (row_count + 7) // 8)


def _to_columnar_array(values: array) -> object:
    # The already-an-array('i') counterpart to _buffer_to_array: with NumPy absent the accumulator is
    # already the right type, so it is handed back untouched rather than round-tripped through bytes.
    if _numpy is None:
        return values
    return _numpy.frombuffer(values, dtype=_numpy.int32)


def decode_cell(sheet: ColumnarSheet, index: int) -> Cell:
    """Materializes the `Cell` at flat cell index `index` in `sheet`, decoding only that one value."""
    start, end = int(sheet.value_offsets[index]), int(sheet.value_offsets[index + 1])
    value = bytes(sheet.values[start:end]).decode("utf-8")
    return Cell(column=int(sheet.columns[index]), type=CellType(int(sheet.types[index])), value=value)


def _raw_options(options: OpenOptions | None) -> object:
    # A NULL options pointer means "every library default" on the native side — identical to calling
    # the non-_ex entry point — so both paths can go through xl_open_*_ex and there is only one call
    # site to keep correct.
    if options is None:
        return None
    return ctypes.byref(_native.to_native_open_options(options))


def open_workbook(path: str | Path, format: str | None = None, *, options: OpenOptions | None = None) -> Workbook:
    """Opens a workbook from disk. `format` is one of auto/xls/xlsx/xlsb/csv; None infers it.

    `options` overrides reader limits and CSV dialect settings; see `OpenOptions`.
    """
    resolved = Path(path)
    encoded = str(resolved).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(
        lib.xl_open_file_ex(
            encoded, len(encoded), _resolve_format(format, resolved), _raw_options(options), ctypes.byref(handle)
        )
    )
    return Workbook(handle)


def open_bytes(data: bytes, format: str | None = None, *, options: OpenOptions | None = None) -> Workbook:
    """Opens a workbook from an in-memory buffer. The native side copies `data` immediately.

    `options` overrides reader limits and CSV dialect settings; see `OpenOptions`.
    """
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(
        lib.xl_open_memory_ex(
            data, len(data), _resolve_format(format, None), _raw_options(options), ctypes.byref(handle)
        )
    )
    return Workbook(handle)
