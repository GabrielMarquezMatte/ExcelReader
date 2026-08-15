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


class OpenOptions(NamedTuple):
    """Reader limits and dialect settings for `open_workbook()`/`open_bytes()`.

    Every field is None for "use the library default", so you only set what you actually want to
    override. That single convention replaces two the C ABI uses internally: numeric fields there
    treat 0 as "default" (making 0 unsettable), and boolean fields carry a third XL_OPT_DEFAULT state
    because several of them default to true. Neither leaks up here.

    The `max_*` fields are resource limits, not tuning knobs — they bound what a malformed or hostile
    file can make the reader allocate. Lower them when parsing untrusted uploads.

    `csv_delimiter` and `csv_quote` are byte values, not strings: pass `ord(';')`, not `';'`.

    Validation happens on the native side, which owns the real limits; an out-of-range value raises
    `ExcelReaderError` with the reason. Fields mirror xl_open_options in
    src/ExcelReader.Native/include/excelreader.h.
    """

    csv_sniff_dialect: bool | None = None
    csv_delimiter: int | None = None
    csv_quote: int | None = None
    csv_detect_bom: bool | None = None
    csv_max_cell_bytes: int | None = None
    csv_intern_strings: bool | None = None
    max_total_decompressed_bytes: int | None = None
    max_cell_bytes: int | None = None
    max_shared_string_bytes: int | None = None
    max_zip_entries: int | None = None
    prefetch_decompression: bool | None = None
    intern_strings: bool | None = None


class WriteOptions(NamedTuple):
    """Sheet name and dialect settings for `write_workbook()`.

    Every field is None for "use the library default", the same convention `OpenOptions` uses.

    `csv_delimiter` and `csv_quote` are byte values, not strings: pass `ord(';')`, not `';'`. They
    apply to CSV output only. `date1904` applies to xls/xlsb only; `use_shared_strings` to
    xlsx/xlsb only, where it shrinks files with many repeated strings at the cost of a string table.

    Validation happens on the native side; an out-of-range value or an invalid sheet name raises
    `ExcelReaderError`. Fields mirror xl_write_options in
    src/ExcelReader.Native/include/excelreader.h.
    """

    sheet_name: str | None = None
    csv_delimiter: int | None = None
    csv_quote: int | None = None
    date1904: bool | None = None
    use_shared_strings: bool | None = None


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
