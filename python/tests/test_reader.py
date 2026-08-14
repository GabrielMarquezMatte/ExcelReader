import gc
import shutil

import pytest
from excelreader import Cell, CellType, ExcelReaderError, open_bytes, open_workbook


def test_reads_the_xlsx_header_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        header = next(workbook.rows())

    assert len(header) == 18
    assert header[0] == Cell(column=0, type=CellType.STRING, value="Coluna1")
    assert header[17].value == "Coluna18"


def test_reads_every_xlsx_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = list(workbook.rows())

    assert len(rows) == 101


def test_reports_xlsx_cell_types(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        next(rows)
        first_data_row = next(rows)

    assert first_data_row[0].type is CellType.STRING
    assert first_data_row[1].type is CellType.DATE
    assert first_data_row[2].type is CellType.NUMBER


def test_reads_csv_when_the_extension_is_csv(csv_path):
    with open_workbook(csv_path) as workbook:
        rows = list(workbook.rows())

    assert [cell.value for cell in rows[0]] == ["name", "qty"]
    assert [cell.value for cell in rows[2]] == ["gadget", "9"]


def test_reads_xlsb(xlsb_path):
    with open_workbook(xlsb_path) as workbook:
        first = next(workbook.rows())

    assert len(first) > 0


def test_reads_xls(xls_path):
    with open_workbook(xls_path) as workbook:
        rows = workbook.rows()
        # 65K rows / 11 MB: reading a handful proves the path works without the wall-clock cost.
        sampled = [next(rows) for _ in range(10)]

    assert all(len(row) > 0 for row in sampled)


def test_reads_from_bytes(xlsx_path):
    with open_bytes(xlsx_path.read_bytes()) as workbook:
        header = next(workbook.rows())

    assert header[0].value == "Coluna1"


def test_exposes_sheet_metadata(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        assert workbook.sheet_count >= 1
        assert workbook.sheet_name
        assert workbook.is_date1904 is False


def test_move_to_sheet_restarts_enumeration(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        first = next(workbook.rows())
        workbook.move_to_sheet(0)
        again = next(workbook.rows())

    assert first == again


def test_missing_file_raises(tmp_path):
    with pytest.raises(ExcelReaderError):
        open_workbook(tmp_path / "nope.xlsx")


def test_use_after_close_raises(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()

    with pytest.raises(ExcelReaderError):
        workbook.sheet_count  # noqa: B018


def test_close_is_idempotent(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()
    workbook.close()


def test_rows_iterator_raises_after_close_mid_iteration(xlsx_path):
    workbook = open_workbook(xlsx_path)
    rows = workbook.rows()
    next(rows)
    workbook.close()

    with pytest.raises(ExcelReaderError):
        next(rows)


def test_dropping_a_workbook_without_close_still_releases_the_file(xlsx_path, tmp_path):
    # Work on a copy — never touch the real fixture, and a copy also lets us assert the OS-level
    # file lock is actually gone by deleting it afterwards.
    copy_path = tmp_path / "dropped.xlsx"
    shutil.copyfile(xlsx_path, copy_path)

    workbook = open_workbook(copy_path)
    next(workbook.rows())
    del workbook
    gc.collect()

    # If __del__ didn't close the native handle, the native side still holds the file open and this
    # raises PermissionError on Windows (the platform where an open-file lock is actually enforced).
    copy_path.unlink()


def test_as_date_converts_the_serial_value(xlsx_path):
    from datetime import date

    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        next(rows)
        first_data_row = next(rows)
        date1904 = workbook.is_date1904

    assert first_data_row[1].as_date(date1904) == date(2026, 1, 1)


def test_as_date_returns_none_for_non_date_cells(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        next(rows)
        first_data_row = next(rows)

    assert first_data_row[0].as_date() is None
