"""Writing a columnar table to a workbook file. Read cursors live in reader.py; this module has none.

Everything here funnels into the single `xl_write_typed` export: it borrows the buffers a `TypedTable`
already holds, so nothing is re-encoded on the way out.
"""

from __future__ import annotations

import ctypes
from array import array
from pathlib import Path
from typing import Any

from excelreader import _native
from excelreader.reader import _check
from excelreader.types import ColumnType, StringColumn, TypedTable, WriteOptions

_WRITE_FORMATS = {
    ".xlsx": _native.XL_FORMAT_XLSX,
    ".xlsb": _native.XL_FORMAT_XLSB,
    ".xls": _native.XL_FORMAT_XLS,
    ".csv": _native.XL_FORMAT_CSV,
}

_FORMAT_NAMES = {
    "xlsx": _native.XL_FORMAT_XLSX,
    "xlsb": _native.XL_FORMAT_XLSB,
    "xls": _native.XL_FORMAT_XLS,
    "csv": _native.XL_FORMAT_CSV,
}


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

    NumPy arrays, `array.array` and `bytes` all expose the buffer protocol, so
    `ctypes.c_char.from_buffer` borrows them without copying. The caller must keep the returned
    object alive for as long as the pointer is in use — the native side copies nothing.
    """
    if isinstance(values, (bytes, bytearray)):
        raw: Any = values
    elif isinstance(values, array):
        raw = values
    else:
        raw = getattr(values, "data", values)  # numpy ndarray -> memoryview
    buffer = (ctypes.c_char * len(memoryview(raw).tobytes())).from_buffer_copy(memoryview(raw).tobytes())
    return buffer, len(buffer)


def _native_column(column: Any, column_type: ColumnType, validity: bytes | None, row_count: int, keepalive: list[Any]) -> _native.NativeColumn:
    if isinstance(column, StringColumn):
        offsets, offsets_len = _buffer_of(column.offsets)
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

    encoded = str(resolved).encode("utf-8")
    lib = _native.load_library()
    _check(
        lib.xl_write_typed(
            encoded, len(encoded), format_id, specs, ctypes.byref(native_table), ctypes.byref(raw_options)
        )
    )
