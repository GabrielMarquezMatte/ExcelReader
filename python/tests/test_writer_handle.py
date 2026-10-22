from __future__ import annotations

import datetime

import pytest

from excelreader import CellType, ColumnSpec, ColumnType, open_workbook, open_writer


def test_round_trips_typed_cells(tmp_path):
    path = tmp_path / "out.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(["name", "qty"])
        writer.write_row(["widget", 7])
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert rows[0] == ["name", "qty"]
    assert rows[1][0] == "widget"
    assert rows[1][1] == "7"


@pytest.mark.xfail(
    strict=True,
    reason=(
        "native reader bug: XlsxReader.Enumerator.cs ParseCellSpan skips self-closing <c/> elements "
        "instead of recording them as blank cells, so a trailing typed blank cell is "
        "missing from the decoded row (len 7, not 8). Pre-existing in the native core, out of scope "
        "for Task 6 (Python bindings only) - remove this marker once the native reader is fixed."
    ),
)
def test_writes_every_scalar_type(tmp_path):
    path = tmp_path / "types.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(
            [
                "s",
                1,
                1.5,
                True,
                datetime.date(2020, 3, 4),
                datetime.time(13, 45, 6),
                datetime.datetime(2020, 3, 4, 13, 45, 6),
                (None, ColumnType.I64),
            ]
        )
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert rows[0][0] == "s"
    assert rows[0][1] == "1"
    assert len(rows[0]) == 8


def test_none_is_a_blank_cell(tmp_path):
    path = tmp_path / "blank.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row([None, "after"])
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert rows[0][-1] == "after"


def test_a_typed_blank_reads_back_as_a_null_of_that_type(tmp_path):
    path = tmp_path / "typed-blank.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(["qty"])
        writer.write_row([(None, ColumnType.I64)])
        writer.write_row([5])
        writer.end_sheet()

    with open_workbook(path) as workbook:
        assert workbook.sheets[0].infer_schema() == [ColumnSpec(ColumnType.I64, name="qty", nullable=True)]
        table = workbook.sheets[0].parse([ColumnSpec(ColumnType.I64, name="qty", nullable=True)])

    assert table.row_count == 2
    assert table.validity[0][0] & 0b11 == 0b10
    assert int(table.columns[0][1]) == 5


@pytest.mark.parametrize("value", [(1, ColumnType.I64), (None, 1), (None,), [None, ColumnType.I64]])
def test_write_row_rejects_a_malformed_typed_blank(tmp_path, value):
    with open_writer(tmp_path / "bad-blank.xlsx") as writer:
        writer.start_sheet("Data")
        with pytest.raises(TypeError):
            writer.write_row(["fine", value])
        writer.end_sheet()


def test_out_of_order_call_raises(tmp_path):
    from excelreader import ExcelReaderError

    path = tmp_path / "bad.xlsx"
    with open_writer(path) as writer:
        with pytest.raises(ExcelReaderError):
            writer.write_row(["too early"])
        writer.start_sheet("Data")
        writer.write_row([])
        writer.end_sheet()


def test_closing_twice_is_safe(tmp_path):
    path = tmp_path / "twice.xlsx"
    writer = open_writer(path)
    writer.start_sheet("Data")
    writer.write_row([1])
    writer.end_sheet()
    writer.close()
    writer.close()


def test_write_row_infers_every_supported_type(tmp_path):
    path = tmp_path / "inferred.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(
            [
                "text",
                42,
                1.5,
                True,
                datetime.date(2021, 6, 7),
                datetime.time(8, 9, 10),
                datetime.datetime(2021, 6, 7, 8, 9, 10),
                None,
            ]
        )
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[(cell.type, cell.value) for cell in row] for row in workbook.sheets[0].rows()]

    types = [cell_type for cell_type, _ in rows[0]]
    values = [value for _, value in rows[0]]

    assert values[0] == "text"
    assert values[1] == "42"
    assert types[3] == CellType.BOOL, f"bool dispatched as {types[3]}"
    assert values[6] != values[4], "datetime and date produced the same cell"


def test_write_row_rejects_an_unsupported_type(tmp_path):
    path = tmp_path / "bad-type.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        with pytest.raises(TypeError) as excinfo:
            writer.write_row(["fine", object()])
        assert "1" in str(excinfo.value), "the message names the position"


def test_an_aware_datetime_is_written_in_utc(tmp_path):
    path = tmp_path / "aware.xlsx"
    offset = datetime.timezone(datetime.timedelta(hours=-3))
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(["at"])
        writer.write_row([datetime.datetime(2021, 6, 7, 5, 0, tzinfo=offset)])
        writer.write_row([datetime.datetime(2021, 6, 7, 8, 0)])
        writer.end_sheet()

    with open_workbook(path) as workbook:
        stamps = workbook.sheets[0].parse([ColumnSpec(ColumnType.TIMESTAMP, name="at")]).columns[0]

    assert stamps[0] == stamps[1]


def test_a_writer_over_a_bytesio_hands_it_the_workbook_on_close():
    import io

    from excelreader import open_workbook

    buffer = io.BytesIO()
    with open_writer(buffer, "xlsx") as writer:
        writer.start_sheet("Data")
        writer.write_row(["name", "qty"])
        writer.write_row(["widget", 7])
        writer.end_sheet()
        assert buffer.getvalue() == b""

    assert not buffer.closed
    with open_workbook(buffer) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert rows[0] == ["name", "qty"]
    assert rows[1][0] == "widget"


def test_a_writer_over_a_bytesio_finishes_a_sheet_left_open():
    import io

    from excelreader import open_workbook

    buffer = io.BytesIO()
    with open_writer(buffer, "xlsb") as writer:
        writer.start_sheet("Data")
        writer.write_row(["a"])

    with open_workbook(buffer) as workbook:
        assert [row[0].value for row in workbook.sheets[0].rows()] == ["a"]


def test_a_writer_over_a_stream_finishes_a_partial_write():
    class Trickle:
        def __init__(self):
            self.data = bytearray()

        def write(self, chunk):
            self.data += bytes(chunk[:5])
            return min(5, len(chunk))

    destination = Trickle()
    with open_writer(destination, "xlsx") as writer:
        writer.start_sheet("Data")
        writer.write_row(["widget", 7])

    with open_workbook(bytes(destination.data)) as workbook:
        assert [cell.value for cell in next(workbook.sheets[0].rows())] == ["widget", "7"]


def test_a_writer_over_a_stream_that_stalls_raises_and_still_closes():
    class Stalled:
        def write(self, chunk):
            return 0

    writer = open_writer(Stalled(), "xlsx")
    writer.start_sheet("Data")
    with pytest.raises(OSError, match="no progress"):
        writer.close()
    writer.close()


def test_a_writer_over_a_stream_needs_a_format_and_a_write_method():
    import io

    with pytest.raises(ValueError, match="format"):
        open_writer(io.BytesIO())
    with pytest.raises(ValueError, match="unknown format"):
        open_writer(io.BytesIO(), "parquet")
    with pytest.raises(TypeError, match="destination"):
        open_writer(42, "xlsx")


def test_write_workbook_to_a_bytesio_matches_write_workbook_to_a_file(tmp_path, xlsx_path):
    import io

    from excelreader import open_workbook, write_workbook

    with open_workbook(xlsx_path) as workbook:
        schema = workbook.sheets[0].infer_schema()
        table = workbook.sheets[0].parse(schema)

    types = [spec.type for spec in schema]

    path = tmp_path / "from-file.xlsx"
    write_workbook(path, table, types, format="xlsx")
    buffer = io.BytesIO()
    write_workbook(buffer, table, types, format="xlsx")
    payload = buffer.getvalue()

    with open_workbook(path) as workbook:
        from_file = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]
    with open_workbook(payload, format="xlsx") as workbook:
        from_memory = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert from_file == from_memory


@pytest.mark.parametrize("value", [2**63, -(2**63) - 1, 2**64 + 5])
def test_an_integer_beyond_64_bits_is_refused_not_wrapped(tmp_path, value):
    with open_writer(tmp_path / "big.xlsx") as writer:
        writer.start_sheet("Data")
        with pytest.raises(OverflowError):
            writer.write_row(["fine", value])
        writer.write_row([2**63 - 1, -(2**63)])
        writer.end_sheet()
