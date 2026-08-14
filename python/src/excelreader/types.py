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


class ExcelReaderError(Exception):
    """Raised when the native library reports a failure."""
