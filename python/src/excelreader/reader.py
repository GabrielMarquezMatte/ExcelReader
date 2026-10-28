"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import io
import operator
import os
import struct
import threading
from array import array
from collections.abc import Iterator, Sequence
from pathlib import Path
from typing import IO, TYPE_CHECKING, SupportsIndex, overload

if TYPE_CHECKING:
    from typing_extensions import Self

    from excelreader.sources import Source

from excelreader import _native
from excelreader.arrow_stream import ArrowStream
from excelreader.types import (
    Cell,
    CellType,
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    OpenOptions,
    PasswordIncorrectError,
    PasswordRequiredError,
    SheetVisibility,
    StringColumn,
    TypedTable,
)

try:
    import numpy as _numpy
except ImportError:
    _numpy = None  

_FORMATS = _native.FORMATS

_CELL_HEADER = struct.Struct("<iii")
_CELL_TYPES = tuple(CellType)
_INITIAL_ROW_BUFFER = 64 * 1024


def _last_error() -> str:
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
        raise ExcelReaderError(_last_error() or "workbook is closed or the handle is invalid")
    if status == _native.XL_INVALID_ARGUMENT:
        raise ExcelReaderError(_last_error() or "invalid argument passed to the native library")
    if status == _native.XL_STATUS_PASSWORD_REQUIRED:
        raise PasswordRequiredError(_last_error() or "the workbook is encrypted and no password was supplied")
    if status == _native.XL_STATUS_PASSWORD_INCORRECT:
        raise PasswordIncorrectError(_last_error() or "the supplied password did not match the workbook's verifier")
    raise ExcelReaderError(_last_error() or f"native call failed with status {status}")


def _resolve_format(name: str | None, path: Path | None) -> int:
    if name is not None:
        try:
            return _FORMATS[name.lower()]
        except KeyError:
            raise ValueError(f"unknown format {name!r}; expected one of {sorted(_FORMATS)}") from None
    if path is not None and path.suffix.lower() == ".csv":
        return _native.XL_FORMAT_CSV
    return _native.XL_FORMAT_AUTO


def _fill_buffer(fn: object, *args: object, initial: int) -> tuple[bytes, int]:
    length = ctypes.c_int32()
    buffer = ctypes.create_string_buffer(initial)
    status = fn(*args, buffer, len(buffer), ctypes.byref(length))
    if status == _native.XL_BUFFER_TOO_SMALL:
        buffer = ctypes.create_string_buffer(length.value)
        status = fn(*args, buffer, len(buffer), ctypes.byref(length))
    _check(status)
    return buffer.raw, length.value


class Workbook:
    """An open workbook. Read it through `sheets`.

    Safe to share between threads: every read goes through a `Sheet` and opens its own native
    cursor, so reads on one workbook never disturb one another.
    """

    def __init__(self, handle: ctypes.c_void_p) -> None:
        self._close_lock = threading.Lock()
        self._lib = _native.load_library()
        self._handle: ctypes.c_void_p | None = handle

    def _require_handle(self) -> ctypes.c_void_p:
        if self._handle is None:
            raise ExcelReaderError("workbook is closed")
        return self._handle

    def _sheet_count(self) -> int:
        count = ctypes.c_int32()
        _check(self._lib.xl_sheet_count(self._require_handle(), ctypes.byref(count)))
        return count.value

    def _sheet_name_at(self, index: int) -> str:
        raw, length = _fill_buffer(self._lib.xl_sheet_name_at, self._require_handle(), index, initial=256)
        return raw[:length].decode("utf-8")

    @property
    def sheets(self) -> Sheets:
        """The workbook's sheets: `len()`, iteration, and lookup by index, slice or name.

        `[sheet.name for sheet in workbook.sheets]` lists every sheet's name.
        """
        return Sheets(self)

    @property
    def tables(self) -> Tables:
        """The workbook's Excel tables: `len()`, iteration, and lookup by index, slice or name
        (ignoring case). Each reads like a `Sheet`, bounded to the table's range."""
        return Tables(self)

    @property
    def is_date1904(self) -> bool:
        flag = ctypes.c_int32()
        _check(self._lib.xl_is_date1904(self._require_handle(), ctypes.byref(flag)))
        return flag.value != 0

    def close(self) -> None:
        """Closes the workbook. A generator that has already started, and any Arrow stream already opened, keep working."""
        with self._close_lock:
            handle, self._handle = self._handle, None
        if handle is None:
            return
        _check(self._lib.xl_close(handle))

    def __enter__(self) -> Self:
        return self

    def __exit__(self, *_exc_info: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:  # noqa: BLE001, S110
            pass


class Sheets(Sequence["Sheet"]):
    """A workbook's sheets, in order. Listing them opens nothing."""

    def __init__(self, workbook: Workbook) -> None:
        self._workbook = workbook

    def __len__(self) -> int:
        return self._workbook._sheet_count()

    @overload
    def __getitem__(self, key: SupportsIndex | str) -> Sheet: ...
    @overload
    def __getitem__(self, key: slice) -> list[Sheet]: ...
    def __getitem__(self, key: SupportsIndex | str | slice) -> Sheet | list[Sheet]:  # type: ignore[override]
        workbook = self._workbook
        if isinstance(key, str):
            encoded = key.encode("utf-8")
            index = ctypes.c_int32()
            _check(workbook._lib.xl_sheet_index(workbook._require_handle(), encoded, _native.int32_len(encoded), ctypes.byref(index)))
            if index.value < 0:
                raise KeyError(key)
            return Sheet(workbook, index.value)
        count = len(self)
        if isinstance(key, slice):
            return [Sheet(workbook, index) for index in range(*key.indices(count))]
        index = operator.index(key)
        if index < 0:
            index += count
        if not 0 <= index < count:
            raise IndexError(f"sheet index {key} is out of range; the workbook has {count} sheets")
        return Sheet(workbook, index)


class Sheet:
    """One sheet of a `Workbook`: the workbook and an index. Holds no native resource.

    Every read opens its own cursor, so several reads may be in flight on one sheet or on several,
    on this thread or on others.
    """

    def __init__(self, workbook: Workbook, index: int) -> None:
        self._workbook = workbook
        self._lib = workbook._lib
        self._index = index

    def __repr__(self) -> str:
        return f"Sheet(index={self._index})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Sheet):
            return NotImplemented
        return self._workbook is other._workbook and self._index == other._index

    def __hash__(self) -> int:
        return hash((id(self._workbook), self._index))

    @property
    def index(self) -> int:
        """Zero-based position in the workbook."""
        return self._index

    @property
    def name(self) -> str:
        return self._workbook._sheet_name_at(self._index)

    @property
    def visibility(self) -> SheetVisibility:
        value = ctypes.c_int32()
        _check(
            self._lib.xl_sheet_visibility_at(self._workbook._require_handle(), self._index, ctypes.byref(value))
        )
        return SheetVisibility(value.value)

    def _open_rows(self) -> ctypes.c_void_p:
        cursor = ctypes.c_void_p()
        _check(self._lib.xl_rows_open(self._workbook._require_handle(), self._index, ctypes.byref(cursor)))
        return cursor

    def _close_rows(self, cursor: ctypes.c_void_p) -> None:
        self._lib.xl_rows_close(cursor)

    def rows(self) -> Iterator[list[Cell]]:
        """This sheet's rows, from the first. Each call is an independent pass over the sheet."""
        date1904 = self._workbook.is_date1904
        cursor = self._open_rows()
        try:
            written = ctypes.c_int32()
            capacity = _INITIAL_ROW_BUFFER
            buffer = ctypes.create_string_buffer(capacity)
            view = memoryview(buffer)
            while True:
                status = self._lib.xl_rows_next(cursor, buffer, capacity, ctypes.byref(written))
                if status == _native.XL_EOF:
                    return
                if status == _native.XL_BUFFER_TOO_SMALL:
                    capacity = written.value
                    buffer = ctypes.create_string_buffer(capacity)
                    view = memoryview(buffer)
                    continue
                _check(status)
                yield _decode_row(view, written.value, date1904)
        finally:
            self._close_rows(cursor)

    def parse(self, schema: Sequence[ColumnSpec], *, header_row: int = 1, parallelism: int = 1) -> TypedTable:
        """Reads this whole sheet into typed columns, converting values on the native side.

        By far the fastest path in this library: unlike `rows()`, no cell is ever formatted to text
        or turned into a Python object.

        `header_row` is a 1-based row number whose values name the columns (that row is skipped and
        never yielded as data); 0 means the sheet has no header, in which case every spec must
        resolve by `index`.

        `parallelism` is the thread count for a CSV: 1 (the default) reads on one thread, 0 uses every
        core, n up to n threads. Any other format, or a CSV too small to split, is read on one thread.
        The table is identical either way; a parallel read holds every partition's columns until they
        are merged, so its peak memory is higher.
        """
        handle = self._workbook._require_handle()
        specs = _build_specs(schema)
        table = _native.NativeTable()
        _check(self._lib.xl_parse_typed(handle, self._index, specs, len(specs), header_row, parallelism, ctypes.byref(table)))
        try:
            return _decode_table(schema, table)
        finally:
            self._lib.xl_free_table(ctypes.byref(table))

    def parse_batches(
        self, schema: Sequence[ColumnSpec], *, header_row: int = 1, batch_size: int = 10000
    ) -> Iterator[TypedTable]:
        """`parse()` a batch at a time, yielding one `TypedTable` per batch.

        `batch_size` is rows per batch; 0 means one unbounded batch, exactly what `parse()`
        does. A generator that has already started owns its position in the sheet and keeps working after
        the workbook is closed.

        Being a generator, it opens nothing until the first iteration, so a bad `batch_size` is not
        reported until then; `to_arrow_stream()` raises immediately instead.
        """
        handle = self._workbook._require_handle()
        specs = _build_specs(schema)
        reader = ctypes.c_void_p()
        _check(
            self._lib.xl_typed_reader_open(
                handle, self._index, specs, len(specs), header_row, batch_size, ctypes.byref(reader)
            )
        )
        try:
            while True:
                table = _native.NativeTable()
                status = self._lib.xl_typed_reader_next(reader, ctypes.byref(table))
                if status == _native.XL_EOF:
                    return
                _check(status)
                try:
                    yield _decode_table(schema, table)
                finally:
                    self._lib.xl_free_table(ctypes.byref(table))
        finally:
            self._lib.xl_typed_reader_close(reader)

    def to_arrow(self, schema: Sequence[ColumnSpec], *, header_row: int = 1, parallelism: int = 1) -> object:
        """The same read as `parse()`, handed to pyarrow as one column-named `RecordBatch`, zero-copy.

        Requires pyarrow. The batch owns the native buffers through the Arrow C Data Interface's
        release callback, so it stays valid after this workbook is closed. `parallelism` works as in
        `parse()`. To stream instead, see `to_arrow_stream()`.
        """
        pyarrow = _import_pyarrow("to_arrow()", "parse()")
        handle = self._workbook._require_handle()
        specs = _build_specs(schema)
        array = _native.ArrowArray()
        arrow_schema = _native.ArrowSchema()
        _check(self._lib.xl_parse_arrow(handle, self._index, specs, len(specs), header_row, parallelism, ctypes.byref(array), ctypes.byref(arrow_schema)))
        return pyarrow.RecordBatch._import_from_c(ctypes.addressof(array), ctypes.addressof(arrow_schema))

    def _open_arrow_stream(
        self, schema: Sequence[ColumnSpec], header_row: int, batch_size: int
    ) -> _native.ArrowArrayStream:
        handle = self._workbook._require_handle()
        specs = _build_specs(schema)
        stream = _native.ArrowArrayStream()
        _check(
            self._lib.xl_parse_arrow_stream(
                handle, self._index, specs, len(specs), header_row, batch_size, ctypes.byref(stream)
            )
        )
        return stream

    def to_arrow_stream(
        self, schema: Sequence[ColumnSpec], *, header_row: int = 1, batch_size: int = 10000
    ) -> ArrowStream:
        """The same read as `to_arrow()`, streamed `batch_size` rows at a time as a single-use Arrow
        PyCapsule stream. Peak memory is one batch rather than one sheet.

        Needs no third-party package to create; pass it to any consumer of `__arrow_c_stream__`:
        `polars.DataFrame(stream)`, or `pyarrow.RecordBatchReader.from_stream(stream)` for a
        `RecordBatchReader`. The stream stays valid after this workbook is closed.
        """
        return ArrowStream(self._open_arrow_stream(schema, header_row, batch_size))

    def to_pandas(
        self, schema: Sequence[ColumnSpec], *, header_row: int = 1, batch_size: int = 10000, parallelism: int = 1
    ) -> object:
        """Same read as `to_arrow()`, materialized as a `pandas.DataFrame`. Requires pyarrow and pandas.

        Streams `batch_size` rows at a time, then concatenates them, so peak memory is not bounded
        by the batch size; `self_destruct` frees each Arrow chunk as pandas takes it, so the sheet
        is never held twice. With `parallelism` other than 1 the sheet is parsed whole, as in
        `parse()`, instead of streamed.
        """
        pyarrow = _import_pyarrow("to_pandas()", "parse()")
        if parallelism != 1:
            data = self.to_arrow(schema, header_row=header_row, parallelism=parallelism)
        else:
            stream = self._open_arrow_stream(schema, header_row, batch_size)
            data = pyarrow.RecordBatchReader._import_from_c(ctypes.addressof(stream)).read_all()
        return data.to_pandas(self_destruct=True, split_blocks=True)

    def to_polars(
        self, schema: Sequence[ColumnSpec], *, header_row: int = 1, batch_size: int = 10000, parallelism: int = 1
    ) -> object:
        """Same read as `to_arrow()`, materialized as a `polars.DataFrame`, zero-copy.

        Streams internally through `to_arrow_stream()` — polars consumes the stream a batch at a
        time, so unlike `to_pandas()` the whole sheet is never resident as Arrow buffers. Requires
        polars only, not pyarrow. With `parallelism` other than 1 the sheet is parsed whole, as in
        `parse()`, instead of streamed, and that path does require pyarrow.
        """
        try:
            import polars
        except ImportError:
            raise ImportError(
                "to_polars() requires polars — install it with `pip install polars`, or use "
                "to_arrow_stream(), which returns the same data with no polars dependency."
            ) from None

        if parallelism != 1:
            return polars.from_arrow(self.to_arrow(schema, header_row=header_row, parallelism=parallelism))
        return polars.DataFrame(self.to_arrow_stream(schema, header_row=header_row, batch_size=batch_size))

    def infer_schema(self, *, header_row: int = 1, sample_size: int = 100, parse_text: bool = False) -> list[ColumnSpec]:
        """Guesses a `parse()`/`to_arrow()` schema by sampling this sheet's cells.

        Reads `header_row` for column names (0 means no header — every returned spec resolves by
        `index` instead) and up to `sample_size` rows after it, guessing each column's type from
        Excel's own per-cell type tag. A column with a real mix of kinds, only
        formula/error results, or nothing sampled falls back to `ColumnType.STRING`; `nullable` is set
        when any sampled row left the column empty.

        `parse_text=True` also types cells that hold text — every CSV field, or numbers stored as text:
        integers, decimals, `true`/`false` and ISO dates or date-times, when the text has exactly that
        shape. Codes with a leading zero (`00123`), scientific notation (`12E4`), padded or comma-decimal
        numbers and non-ISO dates stay `ColumnType.STRING`.

        This is a guess over a sample, not a guarantee — a column that looks like `ColumnType.I64` in
        the sample can still hold a fractional value further down the sheet, which `parse()`
        would then reject unless the spec is `nullable`.
        """
        handle = self._workbook._require_handle()
        schema = _native.NativeInferredSchema()
        flags = _native.XL_INFER_PARSE_TEXT if parse_text else 0
        _check(self._lib.xl_infer_schema(handle, self._index, header_row, sample_size, flags, ctypes.byref(schema)))
        try:
            return _decode_inferred_schema(schema)
        finally:
            self._lib.xl_free_schema(ctypes.byref(schema))


class Tables(Sequence["Table"]):
    """A workbook's Excel tables, in workbook order. Listing them opens nothing."""

    def __init__(self, workbook: Workbook) -> None:
        self._workbook = workbook

    def __len__(self) -> int:
        workbook = self._workbook
        count = ctypes.c_int32()
        _check(workbook._lib.xl_table_count(workbook._require_handle(), ctypes.byref(count)))
        return count.value

    @overload
    def __getitem__(self, key: SupportsIndex | str) -> Table: ...
    @overload
    def __getitem__(self, key: slice) -> list[Table]: ...
    def __getitem__(self, key: SupportsIndex | str | slice) -> Table | list[Table]:  # type: ignore[override]
        workbook = self._workbook
        if isinstance(key, str):
            encoded = key.encode("utf-8")
            index = ctypes.c_int32()
            _check(workbook._lib.xl_table_index(workbook._require_handle(), encoded, _native.int32_len(encoded), ctypes.byref(index)))
            if index.value < 0:
                raise KeyError(key)
            return Table(workbook, index.value)
        count = len(self)
        if isinstance(key, slice):
            return [Table(workbook, index) for index in range(*key.indices(count))]
        index = operator.index(key)
        if index < 0:
            index += count
        if not 0 <= index < count:
            raise IndexError(f"table index {key} is out of range; the workbook has {count} tables")
        return Table(workbook, index)


class Table(Sheet):
    """One Excel table of a `Workbook`. Every `Sheet` read works on it and reads only the table's
    range, header row first; cells keep the column indexes of the sheet holding it.

    Rows and columns are zero-based and include the header and totals rows.
    """

    def __init__(self, workbook: Workbook, index: int) -> None:
        super().__init__(workbook, _native.table_sheet(index))
        self._table = index

    def __repr__(self) -> str:
        return f"Table(index={self._table})"

    @property
    def index(self) -> int:
        """Zero-based position among the workbook's tables."""
        return self._table

    @property
    def name(self) -> str:
        raw, length = _fill_buffer(self._lib.xl_table_name_at, self._workbook._require_handle(), self._table, initial=256)
        return raw[:length].decode("utf-8")

    @property
    def sheet(self) -> Sheet:
        """The sheet that holds the table."""
        return Sheet(self._workbook, self._info().sheet)

    @property
    def first_row(self) -> int:
        return self._info().first_row

    @property
    def last_row(self) -> int:
        return self._info().last_row

    @property
    def first_column(self) -> int:
        return self._info().first_column

    @property
    def last_column(self) -> int:
        return self._info().last_column

    @property
    def header_row_count(self) -> int:
        """1, or 0 when the table's header row is turned off."""
        return self._info().header_row_count

    @property
    def totals_row_count(self) -> int:
        """1 when the table shows a totals row, otherwise 0."""
        return self._info().totals_row_count

    def _info(self) -> _native.NativeTableInfo:
        info = _native.NativeTableInfo()
        _check(self._lib.xl_table_info_at(self._workbook._require_handle(), self._table, ctypes.byref(info)))
        return info


def _decode_row(blob: memoryview, length: int, date1904: bool) -> list[Cell]:
    count = struct.unpack_from("<i", blob, 0)[0]
    cells: list[Cell] = []
    offset = 4
    for _ in range(count):
        column, cell_type, value_length = _CELL_HEADER.unpack_from(blob, offset)
        offset += _CELL_HEADER.size
        value = str(blob[offset : offset + value_length], "utf-8")
        offset += value_length
        cells.append(Cell(column=column, type=_CELL_TYPES[cell_type], value=value, date1904=date1904))
    if offset != length:
        raise ExcelReaderError(f"row blob is malformed: consumed {offset} of {length} bytes")
    return cells


def _import_pyarrow(method: str, alternative: str) -> object:
    try:
        import pyarrow
    except ImportError:
        raise ImportError(
            f"{method} requires pyarrow — install it with `pip install pyarrow`, or use "
            f"{alternative}, which returns the same data with no third-party dependency."
        ) from None
    return pyarrow


def _decode_inferred_schema(schema: _native.NativeInferredSchema) -> list[ColumnSpec]:
    specs: list[ColumnSpec] = []
    for index in range(schema.column_count):
        raw = schema.columns[index]
        name = None
        if raw.name_count > 0:
            name = ctypes.string_at(raw.names[0], raw.name_lens[0]).decode("utf-8")
        specs.append(ColumnSpec(ColumnType(raw.type), name=name, index=raw.index, nullable=bool(raw.nullable)))
    return specs


def _build_specs(schema: Sequence[ColumnSpec]) -> ctypes.Array:
    if not schema:
        raise ValueError("schema must name at least one column")
    specs = (_native.NativeColumnSpec * len(schema))()
    for index, spec in enumerate(schema):
        if spec.name is None:
            specs[index] = _native.column_spec_by_index(spec.index, int(spec.type), nullable=spec.nullable)
            continue
        if isinstance(spec.name, str):
            specs[index] = _native.column_spec_by_name(spec.name, int(spec.type), nullable=spec.nullable)
            continue
        specs[index] = _native.column_spec_by_names(spec.name, int(spec.type), nullable=spec.nullable)
    return specs


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
    offsets_bytes = ctypes.string_at(column.values, (row_count + 1) * ctypes.sizeof(ctypes.c_int32))
    data = ctypes.string_at(column.data, int(column.data_len)) if column.data_len else b""
    return StringColumn(_buffer_to_array(offsets_bytes, "i", "int32"), data)


def _decode_validity(column: _native.NativeColumn, row_count: int) -> bytes | None:
    if not column.validity:
        return None
    return ctypes.string_at(column.validity, (row_count + 7) // 8)


def _raw_options(
    options: OpenOptions | None, password: str | None
) -> tuple[object, _native.NativeOpenOptions | None, ctypes.Array[ctypes.c_char] | None]:
    """Builds the argument to pass as `xl_open_*`'s `options` parameter, plus the two Python
    objects that back it and must outlive the native call.

    A NULL options pointer means "every library default" on the native side, which is what neither
    `options` nor `password` being given produces.

    Returns `(raw_options_arg, raw, password_buffer)`. `raw` must be kept alive (referenced) until
    after the native call returns, since `raw_options_arg` points into it; `password_buffer` must be
    kept alive the same way, since `raw.password` points into IT. The C ABI copies the password bytes
    immediately, but the buffer must not be collected before it does.
    """
    if options is None and password is None:
        return None, None, None
    raw = _native.to_native_open_options(options) if options is not None else _native.default_open_options()
    password_buffer = None
    if password is not None:
        encoded = password.encode("utf-8")
        password_buffer = ctypes.create_string_buffer(encoded, len(encoded))
        raw.password = ctypes.cast(password_buffer, ctypes.POINTER(ctypes.c_uint8))
        raw.password_len = _native.int32_len(encoded)
    return ctypes.byref(raw), raw, password_buffer


def open_workbook(
    source: str | os.PathLike[str] | bytes | bytearray | memoryview | IO[bytes] | Source,
    format: str | None = None,
    *,
    options: OpenOptions | None = None,
    password: str | None = None,
) -> Workbook:
    """Opens a workbook from a path, bytes in memory, a `Source`, or a readable stream.

    `source` picks how the workbook is read:

    - a `str` or path-like: the file on disk.
    - `bytes`, `bytearray`, `memoryview` or `io.BytesIO`: a copy the native side takes immediately.
      A `BytesIO` is read whole, whatever its position.
    - a `Source` (`size` and `read_at`): fetched in blocks on demand, and every sheet can be read on
      its own thread. The library keeps a reference to it until the workbook and everything read
      from it are closed. `read_at` runs with the GIL held, so a local file is better opened by path.
    - any other object with `readinto()` or `read()`, such as an open file, a pipe or
      `sys.stdin.buffer`: read as a stream that cannot seek. A CSV is read as it arrives, and can be
      read again from the start (`infer_schema()`, then a read) until 16 MiB of it were read; an
      XLSX, XLSB or XLS is read whole before the workbook opens. The library never closes it.

    `format` is one of auto/xls/xlsx/xlsb/csv; None infers it, from the extension for a path and from
    the content otherwise. The content tells XLS, XLSX and XLSB apart, and a stream that is none of
    them reads as CSV; a CSV in memory or behind a `Source` needs `format="csv"`. `options` overrides reader limits and CSV dialect settings; see
    `OpenOptions`. `password` unlocks an encrypted OOXML workbook (.xlsx/.xlsb); omitting it for one
    raises `PasswordRequiredError`, and a wrong one raises `PasswordIncorrectError`.
    """
    if isinstance(source, (str, os.PathLike)):
        return _open_path(Path(source), format, options, password)
    if isinstance(source, io.BytesIO):
        source = source.getvalue()
    if isinstance(source, (bytes, bytearray, memoryview)):
        return _open_memory(source, format, options, password)

    from excelreader import sources

    if hasattr(source, "read_at"):
        return sources._open_source(source, format, options, password)
    if hasattr(source, "readinto") or hasattr(source, "read"):
        return sources._open_stream(source, format, options, password)
    raise TypeError(
        "open_workbook() needs a path, bytes, io.BytesIO, a Source, or an object with readinto() or read();"
        f" got {type(source).__name__}"
    )


def _open_path(path: Path, format: str | None, options: OpenOptions | None, password: str | None) -> Workbook:
    encoded = str(path).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    _check(lib.xl_open_file(encoded, _native.int32_len(encoded), _resolve_format(format, path), raw_options, ctypes.byref(handle)))
    return Workbook(handle)


def _open_memory(
    data: bytes | bytearray | memoryview, format: str | None, options: OpenOptions | None, password: str | None
) -> Workbook:
    if not isinstance(data, bytes):
        data = bytes(data)
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    _check(lib.xl_open_memory(data, _native.int32_len(data), _resolve_format(format, None), raw_options, ctypes.byref(handle)))
    return Workbook(handle)
