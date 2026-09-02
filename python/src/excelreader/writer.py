"""Writing a columnar table to a workbook file. Read cursors live in reader.py; this module has none.

Everything here funnels into the single `xl_write_typed` export. Each column's buffer is copied once
into a ctypes block that this module owns for the duration of the call: `xl_write_typed` borrows
whatever it is handed and never frees it, so the pointers it reads must outlive the call and must not
be invalidated by anything Python does to the original object meanwhile. Budget for one extra copy of
the data being written.
"""

from __future__ import annotations

import ctypes
import datetime
from array import array
from pathlib import Path
from typing import Any

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import ColumnType, StringColumn, TypedTable, WriteOptions

_FORMAT_NAMES = _native.WRITE_FORMATS
# The by-extension table is the by-name one with a leading dot, so a new format is added in exactly
# one place (_native.FORMATS) and reaches both lookups.
_WRITE_FORMATS = {f".{name}": value for name, value in _FORMAT_NAMES.items()}


def _resolve_write_format(name: str | None, path: Path) -> int:
    # XL_FORMAT_AUTO is not accepted by xl_write_typed — a file being created has no signature bytes
    # to sniff — so the extension is the only inference available, and an unknown one is an error
    # rather than a silent default.
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


def _buffer_of(values: Any) -> tuple[Any, int]:
    """Returns (ctypes-addressable buffer, byte length) for one column's values.

    NumPy arrays, `array.array` and `bytes` all expose the buffer protocol, so one `memoryview` reads
    every input shape the same way. The bytes are then COPIED into a ctypes block rather than
    borrowed: the copy is owned by this module and cannot be moved or freed out from under the native
    side while the call runs. The caller must keep the returned object alive for as long as the
    pointer is in use.
    """
    if isinstance(values, (bytes, bytearray)):
        raw: Any = values
    elif isinstance(values, array):
        raw = values
    else:
        raw = getattr(values, "data", values)  # numpy ndarray -> memoryview
    blob = memoryview(raw).tobytes()
    return (ctypes.c_char * len(blob)).from_buffer_copy(blob), len(blob)


def _native_column(column: Any, column_type: ColumnType, validity: bytes | None, row_count: int, keepalive: list[Any]) -> _native.NativeColumn:
    if isinstance(column, StringColumn):
        offsets, _ = _buffer_of(column.offsets)
        data, data_len = _buffer_of(column.data)
        keepalive += [offsets, data]
        return _native.NativeColumn(
            type=int(ColumnType.STRING),
            length=row_count,
            values=ctypes.cast(offsets, ctypes.c_void_p),
            validity=_validity_pointer(validity, keepalive),
            data=ctypes.cast(data, ctypes.POINTER(ctypes.c_uint8)) if data_len else None,
            data_len=data_len,
        )

    values, _ = _buffer_of(column)
    keepalive.append(values)
    return _native.NativeColumn(
        type=int(column_type),
        length=row_count,
        values=ctypes.cast(values, ctypes.c_void_p),
        validity=_validity_pointer(validity, keepalive),
        data=None,
        data_len=0,
    )


def _validity_pointer(validity: bytes | None, keepalive: list[Any]) -> Any:
    if validity is None:
        return None
    buffer = (ctypes.c_uint8 * len(validity)).from_buffer_copy(validity)
    keepalive.append(buffer)
    return ctypes.cast(buffer, ctypes.POINTER(ctypes.c_uint8))


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


def write_workbook(
    path: str | Path,
    table: TypedTable,
    types: list[ColumnType],
    *,
    format: str | None = None,
    options: WriteOptions | None = None,
) -> None:
    """Writes `table` to `path` as a single sheet.

    `table` is what `Workbook.parse_typed()` returns, so reading a sheet and writing it back out is a
    round-trip through the same buffers. `types` gives each column's `ColumnType` and is required
    positionally: a `TypedTable` carries raw buffers (an `array`/`StringColumn`/NumPy array), and
    nothing about a raw buffer's element size alone distinguishes, say, I64 from TIME — both are
    8-byte-per-row arrays. The type tag has to travel with the call, not be guessed from the data.
    `format` is one of xlsx/xlsb/xls/csv; None infers it from the path's extension.

    `table.names` becomes the header row. `options` sets the sheet name and CSV dialect; see
    `WriteOptions`.
    """
    resolved = Path(path)
    format_id = _resolve_write_format(format, resolved)
    specs, native_table, raw_options, keepalive = _marshal(table, types, options)

    encoded = str(resolved).encode("utf-8")
    lib = _native.load_library()
    _check(
        lib.xl_write_typed(
            encoded, len(encoded), format_id, specs, ctypes.byref(native_table), ctypes.byref(raw_options)
        )
    )
    del keepalive


def write_workbook_to_bytes(
    table: TypedTable,
    types: list[ColumnType],
    *,
    format: str,
    options: WriteOptions | None = None,
) -> bytes:
    """The in-memory twin of `write_workbook()`.

    `format` is required and is one of xlsx/xlsb/xls/csv — unlike `write_workbook`, there is no path
    to infer it from.
    """
    if format not in _native.WRITE_FORMATS:
        raise ValueError(f"format must be one of {sorted(_native.WRITE_FORMATS)}; got {format!r}")

    specs, native_table, raw_options, keepalive = _marshal(table, types, options)

    buffer = _native.NativeBuffer()
    lib = _native.load_library()
    _check(
        lib.xl_write_typed_to_memory(
            _native.WRITE_FORMATS[format],
            specs,
            ctypes.byref(native_table),
            ctypes.byref(raw_options),
            ctypes.byref(buffer),
        )
    )
    try:
        if not buffer.data or buffer.len <= 0:
            return b""
        return ctypes.string_at(buffer.data, buffer.len)
    finally:
        lib.xl_free_buffer(ctypes.byref(buffer))
        del keepalive


# Arrow type id -> the ColumnType whose buffer layout matches it exactly. Only the seven types
# xl_write_typed accepts are listed; anything else is rejected by name rather than coerced, so a
# caller learns what happened instead of getting a silently stringified column.
_ARROW_TYPES = {
    "string": ColumnType.STRING,
    "large_string": ColumnType.STRING,
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
            # timedelta arithmetic, not datetime.timestamp(): that returns a float whose ULP is
            # already ~0.24 us at 2024-era magnitudes, so int() truncation would silently shift a
            # microsecond-precision value by one microsecond.
            delta = value.replace(tzinfo=datetime.timezone.utc) - _EPOCH_DATETIME
            data_array.append(delta.days * 86_400_000_000 + delta.seconds * 1_000_000 + delta.microseconds)
        elif column_type is ColumnType.BOOL:
            data_array.append(1 if value else 0)
        else:
            data_array.append(value)

    return data_array, (bytes(validity) if has_null else None)


def write_arrow(path: str | Path, batch: Any, **kwargs: Any) -> None:
    """Writes a `pyarrow.RecordBatch` (or `Table`) to `path`.

    Converts through Python lists rather than borrowing Arrow's buffers directly: Arrow's null
    representation, offset conventions and chunking are the producer's choice, and reproducing all of
    them faithfully is exactly the parser this package deliberately does not have. The conversion is
    one pass and keeps the native side reading only layouts it produced itself.
    """
    import pyarrow

    if isinstance(batch, pyarrow.Table):
        batch = batch.combine_chunks().to_batches()[0]

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
    write_workbook(path, table, types, **kwargs)


def write_pandas(path: str | Path, df: Any, **kwargs: Any) -> None:
    """Writes a `pandas.DataFrame` to `path`. Requires pyarrow and pandas."""
    import pyarrow

    write_arrow(path, pyarrow.RecordBatch.from_pandas(df, preserve_index=False), **kwargs)


def write_polars(path: str | Path, df: Any, **kwargs: Any) -> None:
    """Writes a `polars.DataFrame` to `path`. Requires pyarrow and polars."""
    write_arrow(path, df.to_arrow(), **kwargs)


def encrypt_package(package_path: str | Path, destination_path: str | Path, password: str) -> None:
    """Wraps a finished plaintext XLSX/XLSB package in an agile-encrypted (ECMA-376 4.4) CFB
    container — the inverse of passing `password=` to `open_workbook`/`open_bytes`. `package_path`
    is read twice, so it must already be a complete file (write it with `write_workbook` first).
    Encryption parameters are fixed at Excel's own defaults; there are no options.
    """
    package_encoded = str(Path(package_path)).encode("utf-8")
    destination_encoded = str(Path(destination_path)).encode("utf-8")
    password_encoded = password.encode("utf-8")
    lib = _native.load_library()
    _check(
        lib.xl_encrypt_package(
            package_encoded, len(package_encoded),
            destination_encoded, len(destination_encoded),
            password_encoded, len(password_encoded),
        )
    )
