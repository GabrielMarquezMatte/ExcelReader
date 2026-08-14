"""Value types crossing the FFI boundary. Mirrors ExcelReader.Core.Enums.CellType."""

from __future__ import annotations

import datetime
from enum import IntEnum
from typing import Any, NamedTuple


class CellType(IntEnum):
    EMPTY = 0
    STRING = 1
    NUMBER = 2
    DATE = 3
    BOOL = 4
    FORMULA = 5
    ERROR = 6


class Cell(NamedTuple):
    """One cell. `value` is the raw text as stored, so DATE cells are Excel serial numbers."""

    column: int
    type: CellType
    value: str

    def as_date(self, date1904: bool = False) -> datetime.date | None:
        """Converts a DATE cell's serial value to a date. Returns None for any other cell type.

        `date1904` should come from `Workbook.is_date1904` for the workbook this cell came from.
        """
        if self.type is not CellType.DATE:
            return None
        epoch = datetime.date(1904, 1, 1) if date1904 else datetime.date(1899, 12, 30)
        return epoch + datetime.timedelta(days=int(float(self.value)))


class ColumnarSheet(NamedTuple):
    """Every cell of a sheet as parallel flat arrays, produced by `Workbook.read_all_columnar()`.

    The fast path for large sheets: unlike `Workbook.read_all()`, decoding never constructs one
    `Cell`/`str` object per cell — only `int` values fill the index arrays below, and `values` holds
    every cell's raw UTF-8 bytes concatenated. Use `decode_cell()` to materialize a single `Cell` on
    demand, or slice `values` directly for bulk work.

    Each array is `array.array('i')` normally, or a NumPy `int32` `ndarray` when NumPy is installed
    (see `reader.py`) — both support the same `len()`/indexing/slicing used here.

    - `row_offsets[i] : row_offsets[i + 1]` are the cell indices belonging to row `i`, indexing into
      `columns`/`types`/`value_offsets`. Length `row_count + 1`.
    - `columns[j]` / `types[j]` are cell `j`'s column index and `CellType` value. Length `cell_count`.
    - `value_offsets[j] : value_offsets[j + 1]` is cell `j`'s byte slice into `values`. Length
      `cell_count + 1`.
    """

    row_offsets: Any
    columns: Any
    types: Any
    value_offsets: Any
    values: bytes


class ColumnType(IntEnum):
    """A column's requested type in a `Workbook.parse_typed()`/`to_arrow()` schema.

    Mirrors XL_T_* in src/ExcelReader.Native/include/excelreader.h — the values are ABI, do not
    renumber them. The buffer layout each one produces is documented on `TypedTable.columns`.
    """

    STRING = 0
    I64 = 1
    F64 = 2
    BOOL = 3
    DATE = 4
    TIME = 5
    TIMESTAMP = 6


class ColumnSpec(NamedTuple):
    """One column to read, and the type to convert it to.

    Leave `name` as None to resolve the column by `index` instead. `nullable` decides what a failed
    conversion means: False makes it an error that aborts the whole read, True records a null in the
    column's validity bitmap and keeps going.
    """

    type: ColumnType
    name: str | None = None
    index: int = 0
    nullable: bool = False


class StringColumn:
    """A `ColumnType.STRING` column, held as UTF-8 bytes plus offsets rather than one `str` per row.

    Decoding is deferred to `column[i]`, so reading a million-row string column costs no Python
    objects until something actually asks for a value — the same trade `ColumnarSheet` makes.
    """

    __slots__ = ("_data", "_offsets")

    def __init__(self, offsets: Any, data: bytes) -> None:
        self._offsets = offsets
        self._data = data

    @property
    def offsets(self) -> Any:
        """int32 offsets into `data`; `offsets[i]:offsets[i + 1]` is row `i`. Length `len(self) + 1`."""
        return self._offsets

    @property
    def data(self) -> bytes:
        """Every row's UTF-8 bytes, concatenated."""
        return self._data

    def __len__(self) -> int:
        return len(self._offsets) - 1

    def __getitem__(self, index: int) -> str:
        count = len(self)
        if index < 0:
            index += count
        if not 0 <= index < count:
            raise IndexError(f"index {index} is out of range for a column of {count} rows")
        return self._data[int(self._offsets[index]) : int(self._offsets[index + 1])].decode("utf-8")

    def __iter__(self) -> Any:
        return (self[index] for index in range(len(self)))

    def __repr__(self) -> str:
        return f"StringColumn({len(self)} rows, {len(self._data)} bytes)"


class TypedTable(NamedTuple):
    """The result of `Workbook.parse_typed()`: one flat buffer per column, already type-converted.

    - `names[i]` is column `i`'s name as requested, or its index as a string when the spec resolved
      by index.
    - `columns[i]` is a `StringColumn` for `ColumnType.STRING`, and otherwise an `array.array` (or a
      NumPy array when NumPy is installed) of exactly `row_count` values: `'q'`/int64 for I64, TIME
      (microseconds since midnight) and TIMESTAMP (microseconds since 1970-01-01T00:00:00Z), `'d'`
      for F64, `'i'`/int32 for DATE (days since 1970-01-01), `'b'` 0/1 per row for BOOL.
    - `validity[i]` is None when column `i` has no nulls, otherwise an Arrow-style bit-packed bitmap
      where bit `r` (least-significant-bit first) is set when row `r` holds a real value. Only a
      column whose spec set `nullable=True` can ever have one.
    """

    row_count: int
    names: list[str]
    columns: list[Any]
    validity: list[bytes | None]


class ExcelReaderError(Exception):
    """Raised when the native library reports a failure."""
