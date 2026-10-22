"""Value types crossing the FFI boundary. Mirrors ExcelReader.Core.Reader.CellType."""

from __future__ import annotations

import datetime
from dataclasses import KW_ONLY, dataclass
from enum import IntEnum
from typing import Any

from excelreader import _native


class CellType(IntEnum):
    EMPTY = 0
    STRING = 1
    NUMBER = 2
    DATE = 3
    BOOL = 4
    FORMULA = 5
    ERROR = 6


class SheetVisibility(IntEnum):
    """How a sheet is shown in the application's tab bar.

    Mirrors XL_SHEET_* in src/ExcelReader.Native/include/excelreader.h — the values are ABI, do not
    renumber them.
    """

    VISIBLE = _native.XL_SHEET_VISIBLE
    HIDDEN = _native.XL_SHEET_HIDDEN
    VERY_HIDDEN = _native.XL_SHEET_VERY_HIDDEN


# One past 9999-12-31, the last day Excel shows as a date.
_MAX_SERIAL = 2_958_466


@dataclass(slots=True, kw_only=True)
class Cell:
    """One cell. `value` is the raw text as stored, so DATE cells are Excel serial numbers;
    `date1904` is the date system of the workbook it came from.

    Unlike the other value types it is not frozen: `rows()` builds one per cell, and a frozen
    dataclass's `__init__` costs twice as much."""

    column: int
    type: CellType
    value: str
    date1904: bool = False

    def as_date(self) -> datetime.date | None:
        """Converts a DATE cell's serial value to a date. Returns None for any other cell type.

        Raises `ValueError` for a serial outside what Excel can show as a date.
        """
        if self.type is not CellType.DATE:
            return None
        serial = float(self.value)
        if not 0 <= serial < _MAX_SERIAL:
            raise ValueError(f"{self.value} is not an Excel date serial")
        days = int(serial)
        if self.date1904:
            return datetime.date(1904, 1, 1) + datetime.timedelta(days=days)
        # Excel counts a 29 February 1900 that never existed, so serials before it sit one day later
        # against this epoch than those after. Serial 60, that day itself, reads as 28 February.
        if days < 60:
            days += 1
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=days)


class ColumnType(IntEnum):
    """A column's requested type in a `Sheet.parse()`/`to_arrow()` schema.

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


@dataclass(frozen=True, slots=True)
class ColumnSpec:
    """One column to read, and the type to convert it to.

    Leave `name` as None to resolve the column by `index` instead. `nullable` decides what a failed
    conversion means: False makes it an error that aborts the whole read, True records a null in the
    column's validity bitmap and keeps going.
    """

    type: ColumnType
    _: KW_ONLY
    name: str | None = None
    index: int = 0
    nullable: bool = False


@dataclass(frozen=True, slots=True, kw_only=True)
class OpenOptions:
    """Reader limits and dialect settings for `open_workbook()`.

    Every field is None for "use the library default", so you only set what you actually want to
    override. That single convention replaces two the C ABI uses internally: numeric fields there
    treat 0 as "default" (making 0 unsettable), and boolean fields carry a third XL_OPT_DEFAULT state
    because several of them default to true. Neither leaks up here.

    The `max_*` fields are resource limits, not tuning knobs — they bound what a malformed or hostile
    file can make the reader allocate. Lower them when parsing untrusted uploads. `max_cell_bytes`
    and `intern_strings` apply to every format, a CSV included.

    `csv_delimiter` and `csv_quote` are byte values, not strings: pass `ord(';')`, not `';'`.

    `source_block_size` and `source_cache_bytes` tune the block cache of a `Source` passed to
    `open_workbook()` (4 MiB blocks in 64 MiB by default; a negative block size turns the cache off).
    `max_buffered_bytes` caps how much `open_workbook()` reads from a stream before an XLSX, XLSB or
    XLS opens.

    `verify_zip_checksums` (on by default) checks each ZIP entry of an XLSX or XLSB against its
    CRC-32, so a damaged file raises instead of returning wrong cells. Set it to `False` to read a
    file whose producer wrote wrong checksums.

    Validation happens on the native side, which owns the real limits; an out-of-range value raises
    `ExcelReaderError` with the reason. Fields mirror xl_open_options in
    src/ExcelReader.Native/include/excelreader.h.
    """

    csv_sniff_dialect: bool | None = None
    csv_delimiter: int | None = None
    csv_quote: int | None = None
    csv_detect_bom: bool | None = None
    max_total_decompressed_bytes: int | None = None
    max_cell_bytes: int | None = None
    max_shared_string_bytes: int | None = None
    max_zip_entries: int | None = None
    prefetch_decompression: bool | None = None
    intern_strings: bool | None = None
    source_block_size: int | None = None
    source_cache_bytes: int | None = None
    max_buffered_bytes: int | None = None
    verify_zip_checksums: bool | None = None


@dataclass(frozen=True, slots=True, kw_only=True)
class WriteOptions:
    """Sheet name and dialect settings for `write_workbook()`.

    Every field is None for "use the library default", the same convention `OpenOptions` uses.

    `csv_delimiter` and `csv_quote` are byte values, not strings: pass `ord(';')`, not `';'`. They
    apply to CSV output only. `date1904` applies to xls/xlsb only; `use_shared_strings` to
    xlsx/xlsb only, where it shrinks files with many repeated strings at the cost of a string table.

    `degree_of_parallelism` (xlsx/xlsb only) formats and compresses batches of `batch_rows` rows on
    several threads: 0 uses every processor, 1 writes serially, and None is serial. A `batch_rows`
    of None keeps the library default of 8192. With shared strings on, two parallel runs give the
    same content in files that differ byte for byte.

    Validation happens on the native side; an out-of-range value or an invalid sheet name raises
    `ExcelReaderError`. Fields mirror xl_write_options in
    src/ExcelReader.Native/include/excelreader.h.
    """

    sheet_name: str | None = None
    csv_delimiter: int | None = None
    csv_quote: int | None = None
    date1904: bool | None = None
    use_shared_strings: bool | None = None
    degree_of_parallelism: int | None = None
    batch_rows: int | None = None


@dataclass(frozen=True, slots=True)
class TableOptions:
    """Writes a sheet as an Excel table ("Format as Table"), for `SheetWriter.write_sheet`. xlsx/xlsb only.

    `name` must be unique in the workbook, ignoring case, and valid as an Excel name. `style` is a
    built-in table style such as "TableStyleLight9", or None for no style. The flags turn on the
    style's banded rows and columns and its first- and last-column emphasis.
    """

    name: str
    _: KW_ONLY
    style: str | None = "TableStyleMedium2"
    row_stripes: bool = True
    column_stripes: bool = False
    first_column: bool = False
    last_column: bool = False


class StringColumn:
    """A `ColumnType.STRING` column, held as UTF-8 bytes plus offsets rather than one `str` per row.

    Decoding is deferred to `column[i]`, so reading a million-row string column costs no Python
    objects until something actually asks for a value.
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


@dataclass(frozen=True, slots=True, kw_only=True)
class TypedTable:
    """The result of `Sheet.parse()`: one flat buffer per column, already type-converted.

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


class PasswordRequiredError(ExcelReaderError):
    """Raised by `open_workbook()` when the workbook is encrypted and no `password`
    was supplied. Maps XL_STATUS_PASSWORD_REQUIRED from the native ABI."""


class PasswordIncorrectError(ExcelReaderError):
    """Raised by `open_workbook()` when the supplied `password` did not match the
    workbook's verifier. Maps XL_STATUS_PASSWORD_INCORRECT from the native ABI."""
