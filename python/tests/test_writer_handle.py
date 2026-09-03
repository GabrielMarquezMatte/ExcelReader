from __future__ import annotations

import datetime

import pytest

from excelreader import CellType, ColumnType, open_workbook, open_writer


def test_round_trips_typed_cells(tmp_path):
    path = tmp_path / "out.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.start_row()
        writer.write_str("name")
        writer.write_str("qty")
        writer.end_row()
        writer.start_row()
        writer.write_str("widget")
        writer.write_i64(7)
        writer.end_row()
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.rows()]

    assert rows[0] == ["name", "qty"]
    assert rows[1][0] == "widget"
    assert rows[1][1] == "7"


@pytest.mark.xfail(
    strict=True,
    reason=(
        "native reader bug: XlsxReader.Enumerator.cs ParseCellSpan skips self-closing <c/> elements "
        "instead of recording them as blank cells, so a trailing blank cell written by write_null is "
        "missing from the decoded row (len 7, not 8). Pre-existing in the native core, out of scope "
        "for Task 6 (Python bindings only) - remove this marker once the native reader is fixed."
    ),
)
def test_writes_every_scalar_type(tmp_path):
    path = tmp_path / "types.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.start_row()
        writer.write_str("s")
        writer.write_i64(1)
        writer.write_f64(1.5)
        writer.write_bool(True)
        writer.write_date(datetime.date(2020, 3, 4))
        writer.write_time(datetime.time(13, 45, 6))
        writer.write_timestamp(datetime.datetime(2020, 3, 4, 13, 45, 6))
        writer.write_null(ColumnType.I64)
        writer.end_row()
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.rows()]

    assert rows[0][0] == "s"
    assert rows[0][1] == "1"
    assert len(rows[0]) == 8


def test_write_str_none_is_a_blank_cell(tmp_path):
    path = tmp_path / "blank.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.start_row()
        writer.write_str(None)
        writer.write_str("after")
        writer.end_row()
        writer.end_sheet()

    with open_workbook(path) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.rows()]

    assert rows[0][-1] == "after"


def test_out_of_order_call_raises(tmp_path):
    from excelreader import ExcelReaderError

    path = tmp_path / "bad.xlsx"
    with open_writer(path) as writer:
        # A cell write before start_row is rejected by the native side.
        with pytest.raises(ExcelReaderError):
            writer.write_str("too early")
        # The writer stays usable after a rejected call (see SheetWriter's docstring). Finish a
        # minimal sheet so the implicit close on exit succeeds: an XLSX workbook with zero sheets is
        # itself rejected on close by the native side, same as rust/excelreader/tests/writer_handle.rs
        # ends its equivalent test with a start_sheet call before the handle drops.
        writer.start_sheet("Data")
        writer.start_row()
        writer.end_row()
        writer.end_sheet()


def test_closing_twice_is_safe(tmp_path):
    path = tmp_path / "twice.xlsx"
    writer = open_writer(path)
    writer.start_sheet("Data")
    writer.start_row()
    writer.write_i64(1)
    writer.end_row()
    writer.end_sheet()
    writer.close()
    writer.close()


def test_write_row_infers_every_supported_type(tmp_path):
    path = tmp_path / "inferred.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        # True and the datetime are the two subclass traps: bool subclasses int, and
        # datetime.datetime subclasses datetime.date.
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
        rows = [[(cell.type, cell.value) for cell in row] for row in workbook.rows()]

    types = [cell_type for cell_type, _ in rows[0]]
    values = [value for _, value in rows[0]]

    assert values[0] == "text"
    assert values[1] == "42"
    # bool must NOT have been written as the integer 1.
    assert types[3] == CellType.BOOL, f"bool dispatched as {types[3]}"
    # datetime must NOT have been written as a plain date.
    assert values[6] != values[4], "datetime and date produced the same cell"


def test_write_row_rejects_an_unsupported_type(tmp_path):
    path = tmp_path / "bad-type.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        with pytest.raises(TypeError) as excinfo:
            writer.write_row(["fine", object()])
        assert "1" in str(excinfo.value), "the message names the position"


def test_write_row_matches_explicit_calls(tmp_path):
    sugar = tmp_path / "sugar.xlsx"
    explicit = tmp_path / "explicit.xlsx"

    with open_writer(sugar) as writer:
        writer.start_sheet("Data")
        writer.write_row(["a", 1, 2.5, False])
        writer.end_sheet()

    with open_writer(explicit) as writer:
        writer.start_sheet("Data")
        writer.start_row()
        writer.write_str("a")
        writer.write_i64(1)
        writer.write_f64(2.5)
        writer.write_bool(False)
        writer.end_row()
        writer.end_sheet()

    def read(path):
        with open_workbook(path) as workbook:
            return [[(cell.type, cell.value) for cell in row] for row in workbook.rows()]

    assert read(sugar) == read(explicit)


def test_in_memory_writer_round_trips():
    from excelreader import open_bytes, open_writer_to_memory

    with open_writer_to_memory("xlsx") as writer:
        writer.start_sheet("Data")
        writer.write_row(["name", "qty"])
        writer.write_row(["widget", 7])
        writer.end_sheet()
        payload = writer.bytes()

    assert isinstance(payload, bytes)
    assert len(payload) > 0

    with open_bytes(payload, format="xlsx") as workbook:
        rows = [[cell.value for cell in row] for row in workbook.rows()]

    assert rows[0] == ["name", "qty"]
    assert rows[1][0] == "widget"


def test_bytes_on_a_file_writer_raises(tmp_path):
    from excelreader import ExcelReaderError

    path = tmp_path / "file.xlsx"
    with open_writer(path) as writer:
        writer.start_sheet("Data")
        writer.write_row(["a"])
        writer.end_sheet()
        with pytest.raises(ExcelReaderError):
            writer.bytes()


def test_write_workbook_to_bytes_matches_write_workbook(tmp_path, xlsx_path):
    from excelreader import open_bytes, write_workbook, write_workbook_to_bytes

    with open_workbook(xlsx_path) as workbook:
        schema = workbook.infer_schema()
        table = workbook.parse_typed(schema)

    types = [spec.type for spec in schema]

    path = tmp_path / "from-file.xlsx"
    write_workbook(path, table, types, format="xlsx")
    payload = write_workbook_to_bytes(table, types, format="xlsx")

    # Compare by reading both back: an XLSX ZIP is not guaranteed to be byte-identical.
    with open_workbook(path) as workbook:
        from_file = [[cell.value for cell in row] for row in workbook.rows()]
    with open_bytes(payload, format="xlsx") as workbook:
        from_memory = [[cell.value for cell in row] for row in workbook.rows()]

    assert from_file == from_memory
