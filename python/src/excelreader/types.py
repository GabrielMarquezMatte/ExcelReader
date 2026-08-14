"""Value types crossing the FFI boundary. Mirrors ExcelReader.Core.Enums.CellType."""

from __future__ import annotations

import datetime
from enum import IntEnum
from typing import NamedTuple


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


class ExcelReaderError(Exception):
    """Raised when the native library reports a failure."""
