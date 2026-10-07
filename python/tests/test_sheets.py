import gc

import pytest
from excelreader import ExcelReaderError, Sheet, SheetVisibility, open_bytes, open_writer_to_memory


def _two_sheets() -> bytes:
    with open_writer_to_memory("xlsx") as writer:
        for name, values in (("First", ["a1", "a2"]), ("Second", ["b1", "b2", "b3"])):
            writer.start_sheet(name)
            for value in values:
                writer.write_row([value])
            writer.end_sheet()
        return writer.bytes()


def _first_column(sheet: Sheet) -> list[str]:
    return [row[0].value for row in sheet.rows()]


def test_a_sheet_reads_its_own_rows():
    with open_bytes(_two_sheets()) as workbook:
        assert _first_column(workbook.sheets[1]) == ["b1", "b2", "b3"]
        assert _first_column(workbook.sheets[0]) == ["a1", "a2"]


def test_sheets_is_a_sequence_of_sheets():
    with open_bytes(_two_sheets()) as workbook:
        sheets = workbook.sheets

        assert len(sheets) == 2
        assert [sheet.name for sheet in sheets] == ["First", "Second"]
        assert sheets[-1].index == 1
        assert [sheet.index for sheet in sheets[0:2]] == [0, 1]
        assert sheets[0].visibility is SheetVisibility.VISIBLE


def test_sheets_looks_a_sheet_up_by_name_without_regard_to_case():
    with open_bytes(_two_sheets()) as workbook:
        assert workbook.sheets["Second"].index == 1
        assert workbook.sheets["sEcOnD"].index == 1


@pytest.mark.parametrize(("key", "error"), [(5, IndexError), (-3, IndexError), ("nope", KeyError), (1.0, TypeError)])
def test_sheets_rejects_a_key_that_names_no_sheet(key, error):
    with open_bytes(_two_sheets()) as workbook, pytest.raises(error):
        workbook.sheets[key]


def test_two_cursors_on_one_sheet_each_return_every_row():
    with open_bytes(_two_sheets()) as workbook:
        sheet = workbook.sheets[1]
        one = sheet.rows()
        two = sheet.rows()

        assert next(one)[0].value == "b1"
        assert next(two)[0].value == "b1"
        assert [row[0].value for row in one] == ["b2", "b3"]
        assert [row[0].value for row in two] == ["b2", "b3"]


def test_an_abandoned_rows_generator_closes_its_cursor(monkeypatch):
    with open_bytes(_two_sheets()) as workbook:
        sheet = workbook.sheets[1]
        closed = []
        real_close = workbook._lib.xl_rows_close
        monkeypatch.setattr(
            type(sheet), "_close_rows", lambda self, cursor: closed.append(real_close(cursor)), raising=True
        )

        rows = sheet.rows()
        next(rows)
        rows.close()
        del rows
        gc.collect()

        assert closed == [0]


def test_a_sheet_of_a_closed_workbook_raises():
    workbook = open_bytes(_two_sheets())
    sheet = workbook.sheets[0]
    workbook.close()

    with pytest.raises(ExcelReaderError):
        list(sheet.rows())
    with pytest.raises(ExcelReaderError):
        sheet.name


def test_rows_opened_before_close_read_to_the_end_after_it():
    workbook = open_bytes(_two_sheets())
    rows = workbook.sheets[1].rows()
    assert next(rows)[0].value == "b1"

    workbook.close()

    assert [row[0].value for row in rows] == ["b2", "b3"]


def test_read_all_and_parse_typed_read_the_named_sheet():
    from excelreader import ColumnSpec, ColumnType

    with open_bytes(_two_sheets()) as workbook:
        second = workbook.sheets[1]
        schema = [ColumnSpec(ColumnType.STRING, index=0, nullable=True)]

        assert len(second.read_all()) == 3
        assert len(second.read_all_columnar().row_offsets) == 4
        assert second.parse_typed(schema, header_row=0).row_count == 3
        assert sum(table.row_count for table in second.iter_parse_typed(schema, header_row=0, batch_size=2)) == 3
        assert len(second.infer_schema(header_row=0)) == 1
