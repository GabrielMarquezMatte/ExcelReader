"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import struct
from array import array
from collections.abc import Iterator, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
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
    PasswordIncorrectError,
    PasswordRequiredError,
    StringColumn,
    TypedTable,
)

try:
    import numpy as _numpy
except ImportError:
    _numpy = None  

_FORMATS = _native.FORMATS

_CELL_HEADER = struct.Struct("<iii")
_INITIAL_ROW_BUFFER = 64 * 1024
_INITIAL_ALL_ROWS_BUFFER = 1024 * 1024


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
        raise ExcelReaderError("workbook is closed or the handle is invalid")
    if status == _native.XL_INVALID_ARGUMENT:
        raise ExcelReaderError("invalid argument passed to the native library")
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
            self._lib.xl_free_table(ctypes.byref(table))

    def iter_parse_typed(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> Iterator[TypedTable]:
        """`parse_typed()` a batch at a time, yielding one `TypedTable` per batch.

        `batch_size` is rows per batch; 0 means one unbounded batch, exactly what `parse_typed()`
        does. Any other read on this workbook — `parse_typed()`, `rows()`, `move_to_sheet()`,
        `close()` — invalidates the generator, whose next batch then raises. Finish it, or
        `.close()` it, first.

        Being a generator, it opens nothing until the first iteration, so a bad `batch_size` is not
        reported until then; `to_record_batch_reader()` raises immediately instead.
        """
        handle = self._require_handle()
        specs = _build_specs(schema)
        reader = ctypes.c_void_p()
        _check(
            self._lib.xl_typed_reader_open(
                handle, specs, len(specs), header_row, batch_size, ctypes.byref(reader)
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
        return pyarrow.Array._import_from_c(ctypes.addressof(array), ctypes.addressof(arrow_schema))

    def to_record_batch(self, schema: Sequence[ColumnSpec], header_row: int = 1) -> object:
        """Same read as `to_arrow()`, wrapped as a column-named `pyarrow.RecordBatch`.

        Requires pyarrow — see `to_arrow()`.
        """
        array = self.to_arrow(schema, header_row=header_row)
        import pyarrow

        return pyarrow.RecordBatch.from_struct_array(array)

    def to_record_batch_reader(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> object:
        """The same read as `to_arrow()`, as a streaming `pyarrow.RecordBatchReader`.

        Peak memory is one batch rather than one sheet. Requires pyarrow.

        The stream borrows this workbook's row cursor under the same rules as `iter_parse_typed()`:
        one chunked read at a time, and any other read on this workbook invalidates the stream.
        pyarrow owns the stream once it is imported — exhausting or dropping the reader is what
        closes the underlying read.
        """
        try:
            import pyarrow
        except ImportError:
            raise ImportError(
                "to_record_batch_reader() requires pyarrow — install it with `pip install "
                "pyarrow`, or use iter_parse_typed(), which streams the same data with no "
                "third-party dependency."
            ) from None

        handle = self._require_handle()
        specs = _build_specs(schema)
        stream = _native.ArrowArrayStream()
        _check(
            self._lib.xl_parse_arrow_stream(
                handle, specs, len(specs), header_row, batch_size, ctypes.byref(stream)
            )
        )
        return pyarrow.RecordBatchReader._import_from_c(ctypes.addressof(stream))

    def iter_pandas(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> Iterator[object]:
        """One `pandas.DataFrame` per batch, streamed. Requires pyarrow and pandas.

        A generator, so a missing dependency is reported at the first iteration, not at the call.
        """
        try:
            import pyarrow  # noqa: F401
        except ImportError:
            raise ImportError(
                "iter_pandas() requires pyarrow — install it with `pip install pyarrow`, or use "
                "iter_parse_typed(), which streams the same data with no third-party dependency."
            ) from None

        reader = self.to_record_batch_reader(schema, header_row=header_row, batch_size=batch_size)
        for batch in reader:
            yield batch.to_pandas(self_destruct=True, split_blocks=True)

    def iter_polars(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> Iterator[object]:
        """One `polars.DataFrame` per batch, streamed. Requires pyarrow and polars.

        A generator, so a missing dependency is reported at the first iteration, not at the call.
        """
        try:
            import polars
        except ImportError:
            raise ImportError(
                "iter_polars() requires polars — install it with `pip install polars`, or use "
                "to_record_batch_reader(), which streams the same data with no polars dependency."
            ) from None

        reader = self.to_record_batch_reader(schema, header_row=header_row, batch_size=batch_size)
        for batch in reader:
            yield polars.from_arrow(batch, rechunk=False)

    def to_pandas(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> object:
        """Same read as `to_arrow()`, materialized as a `pandas.DataFrame`.

        `read_all()` concatenates every batch, so this does NOT bound peak memory the way
        `iter_pandas()` does. What it buys is the conversion: `self_destruct` frees each Arrow
        chunk as pandas takes it, so the sheet is never held twice. Requires pyarrow and pandas.
        """
        reader = self.to_record_batch_reader(schema, header_row=header_row, batch_size=batch_size)
        return reader.read_all().to_pandas(self_destruct=True, split_blocks=True)

    def to_polars(
        self, schema: Sequence[ColumnSpec], header_row: int = 1, batch_size: int = 10000
    ) -> object:
        """Same read as `to_arrow()`, materialized as a `polars.DataFrame`, zero-copy.

        Streams internally — polars consumes the reader a batch at a time, so unlike `to_pandas()`
        the whole sheet is never resident as Arrow buffers. Requires pyarrow and polars.
        """
        try:
            import polars
        except ImportError:
            raise ImportError(
                "to_polars() requires polars — install it with `pip install polars`, or use "
                "to_arrow()/to_record_batch(), which return the same data with no polars dependency."
            ) from None

        reader = self.to_record_batch_reader(schema, header_row=header_row, batch_size=batch_size)
        return polars.from_arrow(reader, rechunk=False)

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
    # expensive. (ponytail: if this loop measurably dominates for very wide/tall sheets, revisit with
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


def _to_columnar_array(values: array) -> object:
    if _numpy is None:
        return values
    return _numpy.frombuffer(values, dtype=_numpy.int32)


def decode_cell(sheet: ColumnarSheet, index: int) -> Cell:
    """Materializes the `Cell` at flat cell index `index` in `sheet`, decoding only that one value."""
    start, end = int(sheet.value_offsets[index]), int(sheet.value_offsets[index + 1])
    value = bytes(sheet.values[start:end]).decode("utf-8")
    return Cell(column=int(sheet.columns[index]), type=CellType(int(sheet.types[index])), value=value)


def _raw_options(
    options: OpenOptions | None, password: str | None
) -> tuple[object, _native.NativeOpenOptions | None, ctypes.Array[ctypes.c_char] | None]:
    """Builds the argument to pass as `xl_open_*_ex`'s `options` parameter, plus the two Python
    objects that back it and must outlive the native call.

    A NULL options pointer means "every library default" on the native side — identical to calling
    the non-_ex entry point — so both paths can go through xl_open_*_ex and there is only one call
    site to keep correct; that NULL case only applies when neither `options` nor `password` was given.

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
        raw.password_len = len(encoded)
    return ctypes.byref(raw), raw, password_buffer


def open_workbook(
    path: str | Path, format: str | None = None, *, options: OpenOptions | None = None, password: str | None = None
) -> Workbook:
    """Opens a workbook from disk. `format` is one of auto/xls/xlsx/xlsb/csv; None infers it.

    `options` overrides reader limits and CSV dialect settings; see `OpenOptions`. `password` unlocks
    an encrypted OOXML workbook (.xlsx/.xlsb); omitting it for one raises `PasswordRequiredError`, and
    a wrong one raises `PasswordIncorrectError`. An explicit xlsx/xlsb `format` works for an encrypted
    file too, the same as leaving `format` unset — both routes decrypt correctly given the right
    `password`.
    """
    resolved = Path(path)
    encoded = str(resolved).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    _check(
        lib.xl_open_file_ex(
            encoded, len(encoded), _resolve_format(format, resolved), raw_options, ctypes.byref(handle)
        )
    )
    return Workbook(handle)


def open_bytes(
    data: bytes, format: str | None = None, *, options: OpenOptions | None = None, password: str | None = None
) -> Workbook:
    """Opens a workbook from an in-memory buffer. The native side copies `data` immediately.

    `options` overrides reader limits and CSV dialect settings; see `OpenOptions`. `password` unlocks
    an encrypted OOXML workbook — see `open_workbook()` for the full contract.
    """
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    _check(
        lib.xl_open_memory_ex(
            data, len(data), _resolve_format(format, None), raw_options, ctypes.byref(handle)
        )
    )
    return Workbook(handle)
