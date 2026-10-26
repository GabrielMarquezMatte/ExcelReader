"""Writing a columnar table to a workbook file or a binary stream. Read cursors live in reader.py; this module has none.

Everything here funnels into the single `xl_write_typed` export. Each column's buffer is copied once
into a ctypes block that this module owns for the duration of the call: `xl_write_typed` borrows
whatever it is handed and never frees it, so the pointers it reads must outlive the call and must not
be invalidated by anything Python does to the original object meanwhile. Budget for one extra copy of
the data being written.
"""

from __future__ import annotations

import ctypes
import datetime
import io
import os
import sys
from array import array
from pathlib import Path
from typing import IO, Any, Union

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import ColumnType, StringColumn, TypedTable, WriteOptions

_FORMAT_NAMES = _native.WRITE_FORMATS
_WRITE_FORMATS = {f".{name}": value for name, value in _FORMAT_NAMES.items()}


def _resolve_write_format(name: str | None, path: Path) -> int:
    if name is not None:
        try:
            return _FORMAT_NAMES[name.lower()]
        except KeyError:
            raise ValueError(f"unknown format {name!r}; expected one of {sorted(_FORMAT_NAMES)}") from None
    try:
        return _WRITE_FORMATS[path.suffix.lower()]
    except KeyError:
        raise ValueError(
            f"cannot infer a write format from {path.suffix!r}; pass format= explicitly "
            f"(one of {sorted(_FORMAT_NAMES)})"
        ) from None


_SIGNED_FORMATS = frozenset("bhilqn")
_ELEMENTS = {
    ColumnType.I64: (_SIGNED_FORMATS, 8),
    ColumnType.TIME: (_SIGNED_FORMATS, 8),
    ColumnType.TIMESTAMP: (_SIGNED_FORMATS, 8),
    ColumnType.F64: (frozenset("d"), 8),
    ColumnType.BOOL: (frozenset("bB?"), 1),
    ColumnType.DATE: (_SIGNED_FORMATS, 4),
}


def _checked_view(values: Any, formats: frozenset[str], itemsize: int, count: int, what: str) -> memoryview:
    """The native side trusts `length` and the column type, and cannot see how big a Python
    allocation really is, so a buffer of the wrong shape must be refused here, before any pointer
    to it exists.
    """
    view = memoryview(getattr(values, "data", values))
    code = view.format.lstrip("@=")
    if view.ndim != 1 or view.itemsize != itemsize or code not in formats:
        raise TypeError(
            f"{what} must be a 1-D buffer of {itemsize}-byte elements with format one of {sorted(formats)}; "
            f"got format {view.format!r} with {view.itemsize}-byte elements"
        )
    if len(view) != count:
        raise ValueError(f"{what} has {len(view)} elements but {count} are needed")
    return view


def _copy(view: memoryview) -> Any:
    """COPIES the bytes into a ctypes block rather than borrowing them: the copy is owned by this
    module and cannot be moved or freed out from under the native side while the call runs. The
    caller must keep the returned object alive for as long as the pointer is in use.
    """
    blob = view.tobytes()
    return (ctypes.c_char * len(blob)).from_buffer_copy(blob)


def _byte_pointer(buffer: Any) -> Any:
    # ctypes.cast(buffer, ...) puts buffer in a reference cycle, so the copy would outlive `keepalive` until a gc pass.
    return ctypes.cast(ctypes.addressof(buffer), ctypes.POINTER(ctypes.c_uint8))


def _native_column(column: Any, column_type: ColumnType, validity: bytes | None, row_count: int, keepalive: list[Any]) -> _native.NativeColumn:
    if isinstance(column, StringColumn):
        offsets = _copy(_checked_view(column.offsets, _SIGNED_FORMATS, 4, row_count + 1, "a STRING column's offsets"))
        data_view = memoryview(column.data)
        data_len = data_view.nbytes
        data = _copy(data_view)
        keepalive += [offsets, data]
        return _native.NativeColumn(
            type=int(ColumnType.STRING),
            length=row_count,
            values=ctypes.addressof(offsets),
            validity=_validity_pointer(validity, row_count, keepalive),
            data=_byte_pointer(data) if data_len else None,
            data_len=data_len,
        )

    if column_type is ColumnType.STRING:
        raise TypeError("a STRING column must be a StringColumn")
    formats, itemsize = _ELEMENTS[column_type]
    values = _copy(_checked_view(column, formats, itemsize, row_count, f"a {column_type.name} column"))
    keepalive.append(values)
    return _native.NativeColumn(
        type=int(column_type),
        length=row_count,
        values=ctypes.addressof(values),
        validity=_validity_pointer(validity, row_count, keepalive),
        data=None,
        data_len=0,
    )


def _validity_pointer(validity: bytes | None, row_count: int, keepalive: list[Any]) -> Any:
    if validity is None:
        return None
    if len(validity) < (row_count + 7) // 8:
        raise ValueError(
            f"the validity bitmap is {len(validity)} bytes, but {row_count} rows need {(row_count + 7) // 8}"
        )
    buffer = (ctypes.c_uint8 * len(validity)).from_buffer_copy(validity)
    keepalive.append(buffer)
    return _byte_pointer(buffer)


def _marshal(
    table: TypedTable,
    types: list[ColumnType],
    options: WriteOptions | None,
) -> tuple[Any, Any, Any, list[Any]]:
    """Lowers a TypedTable into the native structs xl_write_typed* expect.

    Returns the specs array, the NativeTable, the NativeWriteOptions, and a keepalive list whose
    contents must outlive the native call — dropping it early frees buffers the call is still
    reading.
    """
    if len(types) != len(table.columns):
        raise ValueError(
            f"types must give one ColumnType per column in table.columns; got {len(types)} for {len(table.columns)}"
        )

    keepalive: list[Any] = []
    columns = (_native.NativeColumn * len(table.columns))()
    specs = (_native.NativeColumnSpec * len(table.columns))()
    for index, column in enumerate(table.columns):
        columns[index] = _native_column(column, types[index], table.validity[index], table.row_count, keepalive)
        specs[index] = _native.column_spec_by_name(table.names[index], int(types[index]))

    native_table = _native.NativeTable(
        column_count=len(table.columns), row_count=table.row_count, columns=columns
    )
    raw_options = _native.to_native_write_options(options or WriteOptions())
    return specs, native_table, raw_options, keepalive


Destination = Union[str, os.PathLike, IO[bytes]]


def write_workbook(
    destination: Destination,
    data: Any,
    types: list[ColumnType] | None = None,
    *,
    format: str | None = None,
    options: WriteOptions | None = None,
) -> None:
    """Writes `data` as a single sheet to `destination`: a path, or a binary object with `write()`
    such as `io.BytesIO` or an open file.

    `data` is one of:

    - a `TypedTable`, what `Sheet.parse()` returns, so reading a sheet and writing it back out
      is a round-trip through the same buffers. `types` then gives each column's `ColumnType` and is
      required: a `TypedTable` carries raw buffers, and nothing about a raw buffer's element size
      alone distinguishes, say, I64 from TIME — both are 8-byte-per-row arrays.
    - a `pyarrow.RecordBatch` or `pyarrow.Table`, a `pandas.DataFrame` (its index is not written),
      or any other table exposing `__arrow_c_stream__`, such as a `polars.DataFrame`. Each column's
      type comes from its Arrow type, so `types` must be None. Requires pyarrow.

    `format` is one of xlsx/xlsb/xls/csv. For a path, None infers it from the extension; an object
    with `write()` has no extension, so it needs `format`. The whole workbook is built in memory and
    handed to `write()` at the object's current position, again for the rest whenever `write()`
    returns that it took only part; the object is not closed.

    The column names become the header row. `options` sets the sheet name and CSV dialect; see
    `WriteOptions`. For several sheets, use `SheetWriter.write_sheet`.
    """
    table, types = _typed_table(data, types)
    if isinstance(destination, (str, os.PathLike)):
        _write_file(Path(destination), table, types, format, options)
        return
    if not hasattr(destination, "write"):
        raise TypeError(
            f"destination must be a path or an object with write(); got {type(destination).__name__}"
        )
    if format is None:
        raise ValueError(
            f"format= is required when writing to an object with write(); one of {sorted(_FORMAT_NAMES)}"
        )
    _write_all(destination, _write_memory(table, types, format, options))


def _typed_table(data: Any, types: list[ColumnType] | None) -> tuple[TypedTable, list[ColumnType]]:
    if isinstance(data, TypedTable):
        if types is None:
            raise TypeError("types is required when data is a TypedTable: one ColumnType per column")
        return data, types
    if types is not None:
        raise TypeError("types must be None unless data is a TypedTable; Arrow data carries its own types")
    return _table_from_arrow(_arrow_table(data))


def _write_all(destination: IO[bytes], data: bytes) -> None:
    """Calls `write()` until it has taken all of `data`. A `write()` that returns None is taken to
    have written everything, as files that predate the `int` return convention do, except on a raw
    stream, where None means the write would block and nothing was written.
    """
    view = memoryview(data)
    offset = 0
    while offset < len(data):
        written = destination.write(data if offset == 0 else view[offset:])
        if written is None and isinstance(destination, io.RawIOBase):
            raise BlockingIOError(
                f"write() would block with {len(data) - offset} bytes left to write; use a blocking stream"
            )
        if written is None:
            return
        if not 0 < written <= len(data) - offset:
            raise OSError(
                f"write() made no progress: it returned {written} with {len(data) - offset} bytes left to write"
            )
        offset += written


def _write_file(
    path: Path, table: TypedTable, types: list[ColumnType], format: str | None, options: WriteOptions | None
) -> None:
    format_id = _resolve_write_format(format, path)
    specs, native_table, raw_options, keepalive = _marshal(table, types, options)
    encoded = str(path).encode("utf-8")
    lib = _native.load_library()
    _check(
        lib.xl_write_typed(
            encoded, len(encoded), format_id, specs, ctypes.byref(native_table), ctypes.byref(raw_options)
        )
    )
    del keepalive


def _write_memory(
    table: TypedTable, types: list[ColumnType], format: str, options: WriteOptions | None
) -> bytes:
    try:
        format_id = _FORMAT_NAMES[format.lower()]
    except KeyError:
        raise ValueError(f"unknown format {format!r}; expected one of {sorted(_FORMAT_NAMES)}") from None
    specs, native_table, raw_options, keepalive = _marshal(table, types, options)
    buffer = _native.NativeBuffer()
    lib = _native.load_library()
    _check(
        lib.xl_write_typed_to_memory(
            format_id, specs, ctypes.byref(native_table), ctypes.byref(raw_options), ctypes.byref(buffer)
        )
    )
    try:
        if not buffer.data or buffer.len <= 0:
            return b""
        return ctypes.string_at(buffer.data, buffer.len)
    finally:
        lib.xl_free_buffer(ctypes.byref(buffer))
        del keepalive


_ARROW_TYPES = {
    "string": ColumnType.STRING,
    "large_string": ColumnType.STRING,
    "string_view": ColumnType.STRING,
    "int64": ColumnType.I64,
    "int32": ColumnType.I64,
    "double": ColumnType.F64,
    "float": ColumnType.F64,
    "bool": ColumnType.BOOL,
    "date32[day]": ColumnType.DATE,
    "time64[us]": ColumnType.TIME,
    "timestamp[us]": ColumnType.TIMESTAMP,
}

_ARRAY_TYPECODES = {
    ColumnType.I64: "q",
    ColumnType.TIME: "q",
    ColumnType.TIMESTAMP: "q",
    ColumnType.F64: "d",
    ColumnType.BOOL: "b",
    ColumnType.DATE: "i",
}

_EPOCH_DATE = datetime.date(1970, 1, 1)
_EPOCH_DATETIME = datetime.datetime(1970, 1, 1, tzinfo=datetime.timezone.utc)


def _arrow_column_type(field: Any) -> ColumnType:
    key = str(field.type)
    try:
        return _ARROW_TYPES[key]
    except KeyError:
        raise ValueError(
            f"column {field.name!r} has Arrow type {key!r}, which xl_write_typed cannot write; "
            f"cast it to one of {sorted({str(t) for t in _ARROW_TYPES})} first"
        ) from None


def _column_from_pylist(values: list[Any], column_type: ColumnType, row_count: int) -> tuple[Any, bytes | None]:
    """Builds one column's native buffer layout from a plain Python list, as `to_pylist()` returns it.

    Returns `(column, validity)`: `column` is a `StringColumn` for `ColumnType.STRING` (UTF-8 data plus
    an `array("i")` of `row_count + 1` offsets), or an `array.array` of the typecode matching
    `column_type` otherwise. `validity` is an Arrow-style LSB-first bitmap — `(row_count + 7) // 8`
    bytes, bit `r` set when row `r` is valid — or `None` when no value in `values` was ever `None`,
    matching `TypedTable.validity`'s "None means no nulls" convention. A `None` value contributes a
    zero/empty placeholder to `column` and a cleared bit to the bitmap.
    """
    validity = bytearray((row_count + 7) // 8)
    has_null = False

    if column_type is ColumnType.STRING:
        offsets = array("i", [0])
        chunks: list[bytes] = []
        total = 0
        for row, value in enumerate(values):
            if value is None:
                has_null = True
            else:
                validity[row // 8] |= 1 << (row % 8)
                encoded = value.encode("utf-8")
                chunks.append(encoded)
                total += len(encoded)
            offsets.append(total)
        string_column: Any = StringColumn(offsets, b"".join(chunks))
        return string_column, (bytes(validity) if has_null else None)

    data_array = array(_ARRAY_TYPECODES[column_type])
    for row, value in enumerate(values):
        if value is None:
            has_null = True
            data_array.append(0)
            continue
        validity[row // 8] |= 1 << (row % 8)
        if column_type is ColumnType.DATE:
            data_array.append((value - _EPOCH_DATE).days)
        elif column_type is ColumnType.TIME:
            data_array.append(
                value.hour * 3_600_000_000
                + value.minute * 60_000_000
                + value.second * 1_000_000
                + value.microsecond
            )
        elif column_type is ColumnType.TIMESTAMP:
            delta = value.replace(tzinfo=datetime.timezone.utc) - _EPOCH_DATETIME
            data_array.append(delta.days * 86_400_000_000 + delta.seconds * 1_000_000 + delta.microseconds)
        elif column_type is ColumnType.BOOL:
            data_array.append(1 if value else 0)
        else:
            data_array.append(value)

    return data_array, (bytes(validity) if has_null else None)


def _table_from_arrow(batch: Any) -> tuple[TypedTable, list[ColumnType]]:
    """Converts through Python lists rather than borrowing Arrow's buffers directly: Arrow's null
    representation, offset conventions and chunking are the producer's choice, and reproducing all of
    them faithfully is exactly the parser this package deliberately does not have. The conversion is
    one pass and keeps the native side reading only layouts it produced itself.
    """
    import pyarrow

    if isinstance(batch, pyarrow.Table):
        batches = batch.combine_chunks().to_batches()
        batch = batches[0] if batches else pyarrow.RecordBatch.from_pylist([], schema=batch.schema)

    types = [_arrow_column_type(field) for field in batch.schema]
    columns: list[Any] = []
    validity: list[bytes | None] = []
    for index, column in enumerate(batch.columns):
        built, mask = _column_from_pylist(column.to_pylist(), types[index], batch.num_rows)
        columns.append(built)
        validity.append(mask)

    table = TypedTable(
        row_count=batch.num_rows,
        names=list(batch.schema.names),
        columns=columns,
        validity=validity,
    )
    return table, types


def _arrow_table(data: Any) -> Any:
    try:
        import pyarrow
    except ImportError:
        raise ImportError(
            "write_workbook() requires pyarrow to write Arrow or DataFrame data — install it with "
            "`pip install pyarrow`, or pass a TypedTable and its types."
        ) from None

    if isinstance(data, (pyarrow.RecordBatch, pyarrow.Table)):
        return data
    pandas = sys.modules.get("pandas")
    if pandas is not None and isinstance(data, pandas.DataFrame):
        return pyarrow.Table.from_pandas(data, preserve_index=False)
    if hasattr(data, "__arrow_c_stream__"):
        return pyarrow.RecordBatchReader._import_from_c_capsule(data.__arrow_c_stream__()).read_all()
    raise TypeError(
        "data must be a TypedTable, a pyarrow RecordBatch or Table, a pandas DataFrame, or an object "
        f"with __arrow_c_stream__; got {type(data).__name__}"
    )


def encrypt_package(
    package: Union[str, os.PathLike, bytes, bytearray, memoryview, io.BytesIO],
    destination: Destination,
    password: str,
) -> None:
    """Wraps a finished plaintext XLSX/XLSB package in an agile-encrypted (ECMA-376 4.4) CFB
    container, the inverse of passing `password=` to `open_workbook`.

    `package` is a path to a complete file, or its bytes (`bytes`, `bytearray`, `memoryview`, or an
    `io.BytesIO` read whole). `destination` is a path, or a binary object with `write()` that
    receives the encrypted workbook, as `write_workbook` hands it over, and is not closed. Encryption parameters are fixed
    at Excel's own defaults; there are no options.
    """
    password_encoded = password.encode("utf-8")
    lib = _native.load_library()
    if isinstance(package, (str, os.PathLike)) and isinstance(destination, (str, os.PathLike)):
        package_encoded = str(Path(package)).encode("utf-8")
        destination_encoded = str(Path(destination)).encode("utf-8")
        _check(
            lib.xl_encrypt_package(
                package_encoded, len(package_encoded),
                destination_encoded, len(destination_encoded),
                password_encoded, len(password_encoded),
            )
        )
        return

    if not isinstance(destination, (str, os.PathLike)) and not hasattr(destination, "write"):
        raise TypeError(
            f"destination must be a path or an object with write(); got {type(destination).__name__}"
        )
    if isinstance(package, (str, os.PathLike)):
        data = Path(package).read_bytes()
    elif isinstance(package, io.BytesIO):
        data = package.getvalue()
    elif isinstance(package, (bytes, bytearray, memoryview)):
        data = bytes(package)
    else:
        raise TypeError(f"package must be a path or bytes; got {type(package).__name__}")

    buffer = _native.NativeBuffer()
    _check(
        lib.xl_encrypt_package_to_memory(
            data, len(data), password_encoded, len(password_encoded), ctypes.byref(buffer)
        )
    )
    try:
        encrypted = ctypes.string_at(buffer.data, buffer.len)
    finally:
        lib.xl_free_buffer(ctypes.byref(buffer))
    if isinstance(destination, (str, os.PathLike)):
        Path(destination).write_bytes(encrypted)
    else:
        _write_all(destination, encrypted)
