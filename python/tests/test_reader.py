import gc
import shutil

import pytest
from excelreader import (
    Cell,
    CellType,
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    OpenOptions,
    decode_cell,
    open_bytes,
    open_workbook,
)
from excelreader import reader as reader_module


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


def test_sheet_names_matches_sheet_count_and_current_name(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        names = workbook.sheet_names
        assert len(names) == workbook.sheet_count
        assert names[0] == workbook.sheet_name
        assert list(workbook.sheets()) == list(enumerate(names))


def test_sheet_names_does_not_disturb_row_enumeration(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        first = next(rows)
        second = next(rows)

        # sheet_names must read every sheet name without touching the current sheet or the row
        # cursor — unlike move_to_sheet, which resets enumeration back to the sheet's first row.
        assert workbook.sheet_names

        third = next(rows)

    assert first != second
    assert second != third


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


def test_read_all_matches_row_by_row_iteration(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        all_rows = list(workbook.rows())

    with open_workbook(xlsx_path) as workbook:
        bulk_rows = workbook.read_all()

    assert bulk_rows == all_rows


def test_read_all_returns_empty_list_at_end_of_sheet(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        workbook.read_all()  # drain the sheet
        assert workbook.read_all() == []


def _columnar_to_rows(sheet) -> list[list[Cell]]:
    rows = []
    for row_index in range(len(sheet.row_offsets) - 1):
        start, end = sheet.row_offsets[row_index], sheet.row_offsets[row_index + 1]
        rows.append([decode_cell(sheet, cell_index) for cell_index in range(start, end)])
    return rows


@pytest.mark.parametrize("fixture_name", ["xlsx_path", "xlsb_path"])
def test_read_all_columnar_matches_read_all(fixture_name, request):
    path = request.getfixturevalue(fixture_name)

    with open_workbook(path) as workbook:
        expected = workbook.read_all()

    with open_workbook(path) as workbook:
        sheet = workbook.read_all_columnar()

    assert _columnar_to_rows(sheet) == expected


def test_read_all_columnar_without_numpy(xlsx_path, monkeypatch):
    monkeypatch.setattr(reader_module, "_numpy", None)

    with open_workbook(xlsx_path) as workbook:
        sheet = workbook.read_all_columnar()

    from array import array

    assert isinstance(sheet.columns, array)
    with open_workbook(xlsx_path) as workbook:
        assert _columnar_to_rows(sheet) == workbook.read_all()


def test_read_all_columnar_grows_the_buffer(xlsx_path, monkeypatch):
    # Force the initial buffer far too small so the first native call returns XL_BUFFER_TOO_SMALL and
    # read_all_columnar must retry with the required capacity, per xl_read_all_blob's contract that no
    # rows are lost across that retry.
    monkeypatch.setattr(reader_module, "_INITIAL_ALL_ROWS_BUFFER", 4)

    with open_workbook(xlsx_path) as workbook:
        grown = workbook.read_all_columnar()

    with open_workbook(xlsx_path) as workbook:
        expected = workbook.read_all()

    assert _columnar_to_rows(grown) == expected


def test_read_all_columnar_on_empty_sheet(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        workbook.read_all_columnar()  # drain the sheet
        sheet = workbook.read_all_columnar()

    assert list(sheet.row_offsets) == [0]
    assert len(sheet.columns) == 0
    assert sheet.values == b""


# --- parse_typed / to_arrow ---------------------------------------------------------------------


@pytest.fixture
def typed_csv(tmp_path):
    path = tmp_path / "typed.csv"
    path.write_text(
        "name,qty,price,flag,day,clock,stamp\n"
        "widget,3,1.5,TRUE,2020-01-02,01:00:00,2020-01-02 01:00:00\n"
        "gadget,7,2.5,FALSE,1970-01-01,00:00:01,1970-01-01 00:00:01\n",
        encoding="utf-8",
    )
    return path


_TYPED_SCHEMA = [
    ColumnSpec(ColumnType.STRING, name="name"),
    ColumnSpec(ColumnType.I64, name="qty"),
    ColumnSpec(ColumnType.F64, name="price"),
    ColumnSpec(ColumnType.BOOL, name="flag"),
    ColumnSpec(ColumnType.DATE, name="day"),
    ColumnSpec(ColumnType.TIME, name="clock"),
    ColumnSpec(ColumnType.TIMESTAMP, name="stamp"),
]


def test_parse_typed_returns_every_column_type(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.parse_typed(_TYPED_SCHEMA)

    assert table.row_count == 2
    assert table.names == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    name, qty, price, flag, day, clock, stamp = table.columns
    assert list(name) == ["widget", "gadget"]
    assert list(qty) == [3, 7]
    assert list(price) == [1.5, 2.5]
    assert list(flag) == [1, 0]
    assert list(day) == [18263, 0]  # days since 1970-01-01
    assert list(clock) == [3_600_000_000, 1_000_000]  # microseconds since midnight
    assert list(stamp) == [1_577_926_800_000_000, 1_000_000]  # microseconds since the epoch


def test_parse_typed_resolves_columns_by_index_when_name_is_none(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.parse_typed([ColumnSpec(ColumnType.I64, index=1)])

    assert list(table.columns[0]) == [3, 7]
    assert table.names == ["1"]


def test_parse_typed_string_column_supports_len_and_indexing(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.parse_typed([ColumnSpec(ColumnType.STRING, name="name")])

    column = table.columns[0]
    assert len(column) == 2
    assert column[0] == "widget"
    assert column[-1] == "gadget"


def test_parse_typed_reports_no_validity_bitmap_when_nothing_is_null(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.parse_typed([ColumnSpec(ColumnType.I64, name="qty")])

    assert table.validity == [None]


def test_parse_typed_builds_a_validity_bitmap_for_a_nullable_column(tmp_path):
    path = tmp_path / "nullable.csv"
    path.write_text("qty\n3\n\n9\n", encoding="utf-8")

    with open_workbook(path) as workbook:
        table = workbook.parse_typed([ColumnSpec(ColumnType.I64, name="qty", nullable=True)])

    assert table.row_count == 3
    assert list(table.columns[0]) == [3, 0, 9]
    # Arrow-style bit-packed, least-significant bit first: valid, null, valid.
    assert table.validity[0][0] & 0b111 == 0b101


def test_parse_typed_raises_when_a_non_nullable_column_fails_to_convert(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("qty\n3\nnot-a-number\n", encoding="utf-8")

    with open_workbook(path) as workbook, pytest.raises(ExcelReaderError):
        workbook.parse_typed([ColumnSpec(ColumnType.I64, name="qty")])


def test_parse_typed_raises_for_an_unknown_column_name(typed_csv):
    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.parse_typed([ColumnSpec(ColumnType.I64, name="nope")])


def test_parse_typed_rejects_an_empty_schema(typed_csv):
    with open_workbook(typed_csv) as workbook, pytest.raises(ValueError):
        workbook.parse_typed([])


def test_parse_typed_reads_the_whole_sheet_regardless_of_the_row_cursor(typed_csv):
    # xl_parse_typed restarts at the sheet's first row — an advanced rows() cursor must not shorten it.
    with open_workbook(typed_csv) as workbook:
        next(workbook.rows())
        table = workbook.parse_typed([ColumnSpec(ColumnType.I64, name="qty")])

    assert table.row_count == 2


def test_to_arrow_returns_a_struct_array_matching_parse_typed(typed_csv):
    pa = pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook:
        array = workbook.to_arrow(_TYPED_SCHEMA)

    assert isinstance(array, pa.StructArray)
    assert len(array) == 2
    batch = pa.RecordBatch.from_struct_array(array)
    assert batch.schema.names == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert batch.column("name").to_pylist() == ["widget", "gadget"]
    assert batch.column("qty").to_pylist() == [3, 7]
    assert batch.column("price").to_pylist() == [1.5, 2.5]
    assert batch.column("flag").to_pylist() == [True, False]
    assert [d.isoformat() for d in batch.column("day").to_pylist()] == ["2020-01-02", "1970-01-01"]


def test_to_arrow_survives_the_workbook_being_closed(typed_csv):
    # pyarrow owns the exported buffers via ArrowArray.release, not the workbook handle — the data
    # must stay readable after close(), which is the whole point of handing ownership over.
    pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook:
        array = workbook.to_arrow([ColumnSpec(ColumnType.I64, name="qty")])

    gc.collect()
    assert array.field(0).to_pylist() == [3, 7]


def test_to_arrow_raises_for_an_unknown_column_name(typed_csv):
    pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.to_arrow([ColumnSpec(ColumnType.I64, name="nope")])


def test_to_record_batch_returns_a_column_named_batch_matching_parse_typed(typed_csv):
    pa = pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook:
        batch = workbook.to_record_batch(_TYPED_SCHEMA)

    assert isinstance(batch, pa.RecordBatch)
    assert batch.schema.names == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert batch.column("qty").to_pylist() == [3, 7]


def test_to_pandas_returns_a_dataframe_matching_parse_typed(typed_csv):
    pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    with open_workbook(typed_csv) as workbook:
        df = workbook.to_pandas(_TYPED_SCHEMA)

    assert isinstance(df, pd.DataFrame)
    assert list(df.columns) == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert df["name"].tolist() == ["widget", "gadget"]
    assert df["qty"].tolist() == [3, 7]


def test_to_polars_returns_a_dataframe_matching_parse_typed(typed_csv):
    pytest.importorskip("pyarrow")
    pl = pytest.importorskip("polars")

    with open_workbook(typed_csv) as workbook:
        df = workbook.to_polars(_TYPED_SCHEMA)

    assert isinstance(df, pl.DataFrame)
    assert df.columns == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert df["name"].to_list() == ["widget", "gadget"]
    assert df["qty"].to_list() == [3, 7]


def test_to_polars_raises_for_an_unknown_column_name(typed_csv):
    pytest.importorskip("pyarrow")
    pytest.importorskip("polars")

    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.to_polars([ColumnSpec(ColumnType.I64, name="nope")])


def test_open_options_reaches_the_csv_reader(tmp_path):
    # A semicolon file parses as ONE column under the default comma dialect and as three under the
    # override. Asserting the column split, rather than just that the call succeeded, is what proves
    # the option travelled all the way into the reader instead of being silently dropped.
    path = tmp_path / "semicolons.csv"
    path.write_text("name;qty;price\nwidget;3;9.99\n", encoding="utf-8")

    with open_workbook(path, format="csv") as workbook:
        assert [cell.value for cell in next(workbook.rows())] == ["name;qty;price"]

    with open_workbook(path, format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
        assert [cell.value for cell in next(workbook.rows())] == ["name", "qty", "price"]


def test_open_options_apply_to_open_bytes_too(tmp_path):
    data = b"name;qty\nwidget;3\n"

    with open_bytes(data, format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
        assert [cell.value for cell in next(workbook.rows())] == ["name", "qty"]


def test_open_options_default_to_the_library_defaults(xlsx_path):
    # An all-None OpenOptions must behave exactly like passing none at all: every field decodes to
    # the ABI's "use the default" sentinel rather than to a zero that means something else.
    with open_workbook(xlsx_path) as plain, open_workbook(xlsx_path, options=OpenOptions()) as explicit:
        assert [c.value for c in next(plain.rows())] == [c.value for c in next(explicit.rows())]


def test_open_options_rejects_an_out_of_range_value(csv_path):
    # Validation lives on the native side; the wrapper's job is to surface its reason, not to
    # re-implement the bound.
    with pytest.raises(ExcelReaderError):
        open_workbook(csv_path, format="csv", options=OpenOptions(csv_delimiter=999))


def test_open_options_limit_actually_aborts_an_oversized_read(tmp_path):
    # The max_* fields are resource limits, not tuning knobs, so this asserts one BITES: the same file
    # reads fine at the default and raises under a lower cap. Only the pair proves the limit did the
    # rejecting rather than the file simply being broken.
    #
    # The cell has to exceed the reader's 64 KiB starting buffer, because these caps bound buffer
    # GROWTH — a value that fits the initial allocation never consults them.
    path = tmp_path / "wide-cell.csv"
    path.write_text("value\n" + ("x" * 200_000) + "\n", encoding="utf-8")

    with open_workbook(path, format="csv") as workbook:
        assert len(list(workbook.rows())[1][0].value) == 200_000

    with pytest.raises(ExcelReaderError), open_workbook(path, format="csv", options=OpenOptions(csv_max_cell_bytes=100_000)) as workbook:
            list(workbook.rows())
