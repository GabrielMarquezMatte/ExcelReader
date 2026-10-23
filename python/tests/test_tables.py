from __future__ import annotations

import io
import zipfile
from array import array

import pytest

from excelreader import (
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    StringColumn,
    TableOptions,
    TypedTable,
    WriteOptions,
    open_workbook,
    open_writer,
)

_SCHEMA = [ColumnSpec(ColumnType.STRING, name="name"), ColumnSpec(ColumnType.I64, name="qty")]
_TYPES = [spec.type for spec in _SCHEMA]


def _table(names: list[str], quantities: list[int]) -> TypedTable:
    data = "".join(names).encode()
    offsets = array("i", [0])
    for name in names:
        offsets.append(offsets[-1] + len(name.encode()))
    return TypedTable(
        row_count=len(names),
        names=["name", "qty"],
        columns=[StringColumn(offsets, data), array("q", quantities)],
        validity=[None, None],
    )


def _workbook_bytes(extension: str, options: WriteOptions | None = None) -> bytes:
    buffer = io.BytesIO()
    with open_writer(buffer, extension, options=options) as writer:
        writer.start_sheet("Notas")
        writer.write_row(["antes"])
        writer.end_sheet()
        writer.write_sheet("Vendas", _table(["widget", "gadget"], [3, 7]), _TYPES, table=TableOptions("Sales"))
        writer.write_sheet(
            "Estoque",
            _table(["a", "b", "c"], [10, 20, 30]),
            _TYPES,
            table=TableOptions("Stock", style=None, row_stripes=False),
        )
    return buffer.getvalue()


@pytest.mark.parametrize("extension", ["xlsx", "xlsb"])
@pytest.mark.parametrize("degree", [1, 4])
def test_tables_written_by_write_sheet_read_back(extension, degree):
    payload = _workbook_bytes(extension, WriteOptions(degree_of_parallelism=degree, batch_rows=1))

    with open_workbook(payload, format=extension) as workbook:
        tables = workbook.tables
        assert len(tables) == 2
        assert [table.name for table in tables] == ["Sales", "Stock"]

        sales = tables["sales"]
        assert sales == tables[0]
        assert sales.index == 0
        assert sales.sheet.name == "Vendas"
        assert (sales.first_row, sales.last_row, sales.first_column, sales.last_column) == (0, 2, 0, 1)
        assert (sales.header_row_count, sales.totals_row_count) == (1, 0)
        assert [[cell.value for cell in row] for row in sales.rows()] == [["name", "qty"], ["widget", "3"], ["gadget", "7"]]

        stock = tables[-1].parse(_SCHEMA)
        assert list(stock.columns[0]) == ["a", "b", "c"]
        assert list(stock.columns[1]) == [10, 20, 30]

        with pytest.raises(KeyError):
            tables["missing"]
        with pytest.raises(IndexError):
            tables[2]


def test_table_options_reach_the_table_part():
    with zipfile.ZipFile(io.BytesIO(_workbook_bytes("xlsx"))) as package:
        parts = {name: package.read(name).decode() for name in package.namelist() if name.startswith("xl/tables/")}

    sales = next(xml for xml in parts.values() if 'name="Sales"' in xml)
    stock = next(xml for xml in parts.values() if 'name="Stock"' in xml)
    assert '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1"' in sales
    assert '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0"' in stock


def test_a_workbook_without_tables_lists_none():
    buffer = io.BytesIO()
    with open_writer(buffer, "xlsx") as writer:
        writer.write_sheet("Vendas", _table(["widget"], [3]), _TYPES)
    with open_workbook(buffer.getvalue(), format="xlsx") as workbook:
        assert len(workbook.tables) == 0


@pytest.mark.parametrize("extension", ["xls", "csv"])
def test_a_table_is_refused_outside_xlsx_and_xlsb(extension):
    with open_writer(io.BytesIO(), extension) as writer:
        with pytest.raises(ExcelReaderError, match="XLSX or XLSB"):
            writer.write_sheet("Vendas", _table(["widget"], [3]), _TYPES, table=TableOptions("Sales"))
        writer.write_sheet("Vendas", _table(["widget"], [3]), _TYPES)
