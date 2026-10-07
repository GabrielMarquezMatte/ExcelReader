import gc
import threading
from concurrent.futures import ThreadPoolExecutor

import pytest
from excelreader import ExcelReaderError, Sheet, SheetVisibility, Workbook, open_bytes, open_writer_to_memory


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


_PARALLEL_SHEETS = 6
_PARALLEL_ROWS = 500


def _many_sheets() -> bytes:
    with open_writer_to_memory("xlsx") as writer:
        for sheet in range(_PARALLEL_SHEETS):
            writer.start_sheet(f"sheet{sheet}")
            for row in range(_PARALLEL_ROWS):
                writer.write_row([f"s{sheet}-r{row}", sheet * 1_000_000 + row])
            writer.end_sheet()
        return writer.bytes()


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


def test_every_sheet_read_on_its_own_thread_matches_a_sequential_read():
    with open_bytes(_many_sheets()) as workbook:
        sheets = list(workbook.sheets)
        sequential = [_first_column(sheet) for sheet in sheets]
        assert sequential[3][0] == "s3-r0"
        assert len(sequential[3]) == _PARALLEL_ROWS

        with ThreadPoolExecutor(max_workers=_PARALLEL_SHEETS * 2) as pool:
            parallel = list(pool.map(_first_column, sheets + sheets))

    for position, rows in enumerate(parallel):
        assert rows == sequential[position % _PARALLEL_SHEETS]


def test_typed_parses_of_every_sheet_run_on_separate_threads():
    from excelreader import ColumnSpec, ColumnType

    schema = [ColumnSpec(ColumnType.I64, index=1)]
    with open_bytes(_many_sheets()) as workbook, ThreadPoolExecutor(max_workers=_PARALLEL_SHEETS) as pool:
        tables = list(pool.map(lambda sheet: sheet.parse_typed(schema, header_row=0), workbook.sheets))

    for sheet, table in enumerate(tables):
        assert table.row_count == _PARALLEL_ROWS
        assert int(table.columns[0][0]) == sheet * 1_000_000


def test_a_record_batch_reader_outlives_its_workbook():
    pyarrow = pytest.importorskip("pyarrow")
    from excelreader import ColumnSpec, ColumnType

    schema = [ColumnSpec(ColumnType.I64, index=1)]
    with open_bytes(_many_sheets()) as workbook:
        reader = workbook.sheets[2].to_record_batch_reader(schema, header_row=0, batch_size=100)

    table = reader.read_all()

    assert isinstance(table, pyarrow.Table)
    assert table.num_rows == _PARALLEL_ROWS
    assert table.column(0)[0].as_py() == 2_000_000


def test_sheets_compare_by_workbook_and_index():
    data = _two_sheets()
    with open_bytes(data) as workbook, open_bytes(data) as other:
        sheets = workbook.sheets
        assert sheets.index(sheets[1]) == 1
        assert sheets[0] in sheets
        assert sheets[0] == sheets[0]
        assert sheets[0] != sheets[1]
        assert sheets[0] != other.sheets[0]
        assert len({sheets[0], sheets[0], sheets[1]}) == 2


def test_a_sheet_with_no_rows_reads_empty():
    with open_writer_to_memory("xlsx") as writer:
        writer.start_sheet("Empty")
        writer.end_sheet()
        data = writer.bytes()

    with open_bytes(data) as workbook:
        sheet = workbook.sheets[0]
        assert sheet.read_all() == []
        columnar = sheet.read_all_columnar()
        assert list(columnar.row_offsets) == [0]
        assert columnar.values == b""
        assert list(sheet.rows()) == []


class _SlowCheckWorkbook(Workbook):
    """Parks each thread at its first read of the handle until two threads have read it."""

    _gate = threading.Barrier(2)
    _seen = threading.local()

    @property
    def _handle(self):
        value = self._value
        if not getattr(self._seen, "done", False):
            self._seen.done = True
            try:
                self._gate.wait(timeout=0.5)
            except threading.BrokenBarrierError:
                pass
        return value

    @_handle.setter
    def _handle(self, value):
        self._value = value


def test_closing_from_two_threads_at_once_raises_nowhere():
    source = open_bytes(_two_sheets())
    workbook = _SlowCheckWorkbook(source._handle)
    source._handle = None
    errors: list[BaseException] = []

    def close() -> None:
        try:
            workbook.close()
        except BaseException as error:  # noqa: BLE001
            errors.append(error)

    threads = [threading.Thread(target=close) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert errors == []
