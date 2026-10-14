import ctypes
import gc
import shutil
import sys
import threading
import time
from pathlib import Path

import pytest
from excelreader import (
    Cell,
    CellType,
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    OpenOptions,
    PasswordIncorrectError,
    PasswordRequiredError,
    open_workbook,
)
from excelreader import _native
from excelreader import arrow_stream as arrow_stream_module

ENCRYPTED = Path(__file__).resolve().parents[2] / "tests" / "ExcelReader.Tests" / "data" / "encrypted"


def test_reads_the_xlsx_header_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        header = next(workbook.sheets[0].rows())

    assert len(header) == 18
    assert header[0] == Cell(column=0, type=CellType.STRING, value="Coluna1")
    assert header[17].value == "Coluna18"


def test_reads_every_xlsx_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = list(workbook.sheets[0].rows())

    assert len(rows) == 101


def test_reports_xlsx_cell_types(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.sheets[0].rows()
        next(rows)
        first_data_row = next(rows)

    assert first_data_row[0].type is CellType.STRING
    assert first_data_row[1].type is CellType.DATE
    assert first_data_row[2].type is CellType.NUMBER


def test_reads_csv_when_the_extension_is_csv(csv_path):
    with open_workbook(csv_path) as workbook:
        rows = list(workbook.sheets[0].rows())

    assert [cell.value for cell in rows[0]] == ["name", "qty"]
    assert [cell.value for cell in rows[2]] == ["gadget", "9"]


def test_reads_xlsb(xlsb_path):
    with open_workbook(xlsb_path) as workbook:
        first = next(workbook.sheets[0].rows())

    assert len(first) > 0


def test_reads_xls(xls_path):
    with open_workbook(xls_path) as workbook:
        rows = workbook.sheets[0].rows()
        sampled = [next(rows) for _ in range(10)]

    assert all(len(row) > 0 for row in sampled)


def test_reads_from_bytes(xlsx_path):
    with open_workbook(xlsx_path.read_bytes()) as workbook:
        header = next(workbook.sheets[0].rows())

    assert header[0].value == "Coluna1"


def test_exposes_sheet_metadata(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        assert len(workbook.sheets) >= 1
        assert workbook.sheets[0].name
        assert workbook.is_date1904 is False


def test_listing_sheet_names_does_not_disturb_row_enumeration(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.sheets[0].rows()
        first = next(rows)
        second = next(rows)

        assert [sheet.name for sheet in workbook.sheets]

        third = next(rows)

    assert first != second
    assert second != third


def test_missing_file_raises(tmp_path):
    with pytest.raises(ExcelReaderError):
        open_workbook(tmp_path / "nope.xlsx")


def test_use_after_close_raises(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()

    with pytest.raises(ExcelReaderError):
        len(workbook.sheets)


def test_close_is_idempotent(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()
    workbook.close()


def test_dropping_a_workbook_without_close_still_releases_the_file(xlsx_path, tmp_path):
    copy_path = tmp_path / "dropped.xlsx"
    shutil.copyfile(xlsx_path, copy_path)

    workbook = open_workbook(copy_path)
    next(workbook.sheets[0].rows())
    del workbook
    gc.collect()

    copy_path.unlink()


def test_as_date_converts_the_serial_value(xlsx_path):
    from datetime import date

    with open_workbook(xlsx_path) as workbook:
        rows = workbook.sheets[0].rows()
        next(rows)
        first_data_row = next(rows)
        date1904 = workbook.is_date1904

    assert first_data_row[1].as_date(date1904) == date(2026, 1, 1)


def test_as_date_returns_none_for_non_date_cells(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.sheets[0].rows()
        next(rows)
        first_data_row = next(rows)

    assert first_data_row[0].as_date() is None


def test_rows_reads_the_whole_sheet_every_time(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        sheet = workbook.sheets[0]
        first = list(sheet.rows())
        assert first
        assert list(sheet.rows()) == first


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
        table = workbook.sheets[0].parse_typed(_TYPED_SCHEMA)

    assert table.row_count == 2
    assert table.names == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    name, qty, price, flag, day, clock, stamp = table.columns
    assert list(name) == ["widget", "gadget"]
    assert list(qty) == [3, 7]
    assert list(price) == [1.5, 2.5]
    assert list(flag) == [1, 0]
    assert list(day) == [18263, 0]  
    assert list(clock) == [3_600_000_000, 1_000_000]  
    assert list(stamp) == [1_577_926_800_000_000, 1_000_000]  


def test_parse_typed_resolves_columns_by_index_when_name_is_none(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, index=1)])

    assert list(table.columns[0]) == [3, 7]
    assert table.names == ["1"]


def test_parse_typed_string_column_supports_len_and_indexing(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.STRING, name="name")])

    column = table.columns[0]
    assert len(column) == 2
    assert column[0] == "widget"
    assert column[-1] == "gadget"


def test_parse_typed_reports_no_validity_bitmap_when_nothing_is_null(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, name="qty")])

    assert table.validity == [None]


def test_parse_typed_builds_a_validity_bitmap_for_a_nullable_column(tmp_path):
    path = tmp_path / "nullable.csv"
    path.write_text("qty\n3\n\n9\n", encoding="utf-8")

    with open_workbook(path) as workbook:
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, name="qty", nullable=True)])

    assert table.row_count == 3
    assert list(table.columns[0]) == [3, 0, 9]
    assert table.validity[0][0] & 0b111 == 0b101


def test_parse_typed_raises_when_a_non_nullable_column_fails_to_convert(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("qty\n3\nnot-a-number\n", encoding="utf-8")

    with open_workbook(path) as workbook, pytest.raises(ExcelReaderError):
        workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, name="qty")])


def test_parse_typed_raises_for_an_unknown_column_name(typed_csv):
    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, name="nope")])


def test_parse_typed_rejects_an_empty_schema(typed_csv):
    with open_workbook(typed_csv) as workbook, pytest.raises(ValueError):
        workbook.sheets[0].parse_typed([])


def test_parse_typed_reads_the_whole_sheet_regardless_of_the_row_cursor(typed_csv):
    with open_workbook(typed_csv) as workbook:
        rows = workbook.sheets[0].rows()
        next(rows)
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.I64, name="qty")])
        rest = list(rows)

    assert table.row_count == 2
    assert len(rest) == 2


def test_to_arrow_returns_a_record_batch_matching_parse_typed(typed_csv):
    pa = pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook:
        batch = workbook.sheets[0].to_arrow(_TYPED_SCHEMA)

    assert isinstance(batch, pa.RecordBatch)
    assert batch.num_rows == 2
    assert batch.schema.names == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert batch.column("name").to_pylist() == ["widget", "gadget"]
    assert batch.column("qty").to_pylist() == [3, 7]
    assert batch.column("price").to_pylist() == [1.5, 2.5]
    assert batch.column("flag").to_pylist() == [True, False]
    assert [d.isoformat() for d in batch.column("day").to_pylist()] == ["2020-01-02", "1970-01-01"]


def test_to_arrow_survives_the_workbook_being_closed(typed_csv):
    pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook:
        batch = workbook.sheets[0].to_arrow([ColumnSpec(ColumnType.I64, name="qty")])

    gc.collect()
    assert batch.column(0).to_pylist() == [3, 7]


def test_to_arrow_raises_for_an_unknown_column_name(typed_csv):
    pytest.importorskip("pyarrow")

    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.sheets[0].to_arrow([ColumnSpec(ColumnType.I64, name="nope")])


def test_to_pandas_returns_a_dataframe_matching_parse_typed(typed_csv):
    pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    with open_workbook(typed_csv) as workbook:
        df = workbook.sheets[0].to_pandas(_TYPED_SCHEMA)

    assert isinstance(df, pd.DataFrame)
    assert list(df.columns) == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert df["name"].tolist() == ["widget", "gadget"]
    assert df["qty"].tolist() == [3, 7]


def test_to_polars_returns_a_dataframe_matching_parse_typed(typed_csv):
    pl = pytest.importorskip("polars")

    with open_workbook(typed_csv) as workbook:
        df = workbook.sheets[0].to_polars(_TYPED_SCHEMA)

    assert isinstance(df, pl.DataFrame)
    assert df.columns == ["name", "qty", "price", "flag", "day", "clock", "stamp"]
    assert df["name"].to_list() == ["widget", "gadget"]
    assert df["qty"].to_list() == [3, 7]


def test_to_polars_raises_for_an_unknown_column_name(typed_csv):
    pytest.importorskip("polars")

    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.sheets[0].to_polars([ColumnSpec(ColumnType.I64, name="nope")])


def test_open_options_reaches_the_csv_reader(tmp_path):
    path = tmp_path / "semicolons.csv"
    path.write_text("name;qty;price\nwidget;3;9.99\n", encoding="utf-8")

    with open_workbook(path, format="csv") as workbook:
        assert [cell.value for cell in next(workbook.sheets[0].rows())] == ["name;qty;price"]

    with open_workbook(path, format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
        assert [cell.value for cell in next(workbook.sheets[0].rows())] == ["name", "qty", "price"]


def test_open_options_apply_to_bytes_too(tmp_path):
    data = b"name;qty\nwidget;3\n"

    with open_workbook(data, format="csv", options=OpenOptions(csv_delimiter=ord(";"))) as workbook:
        assert [cell.value for cell in next(workbook.sheets[0].rows())] == ["name", "qty"]


def test_open_options_default_to_the_library_defaults(xlsx_path):
    with open_workbook(xlsx_path) as plain, open_workbook(xlsx_path, options=OpenOptions()) as explicit:
        assert [c.value for c in next(plain.sheets[0].rows())] == [c.value for c in next(explicit.sheets[0].rows())]


def test_open_options_rejects_an_out_of_range_value(csv_path):
    with pytest.raises(ExcelReaderError):
        open_workbook(csv_path, format="csv", options=OpenOptions(csv_delimiter=999))


def test_open_options_limit_actually_aborts_an_oversized_read(tmp_path):
    path = tmp_path / "wide-cell.csv"
    path.write_text("value\n" + ("x" * 200_000) + "\n", encoding="utf-8")

    with open_workbook(path, format="csv") as workbook:
        assert len(list(workbook.sheets[0].rows())[1][0].value) == 200_000

    with pytest.raises(ExcelReaderError), open_workbook(path, format="csv", options=OpenOptions(max_cell_bytes=100_000)) as workbook:
            list(workbook.sheets[0].rows())


def test_opens_encrypted_workbook_with_password():
    with open_workbook(ENCRYPTED / "agile-aes256-sha512.xlsx", password="hunter2") as book:
        assert len(book.sheets) > 0


def test_reports_password_required_without_password():
    with pytest.raises(PasswordRequiredError):
        open_workbook(ENCRYPTED / "agile-aes256-sha512.xlsx")


def test_reports_password_incorrect_with_wrong_password():
    with pytest.raises(PasswordIncorrectError):
        open_workbook(ENCRYPTED / "agile-aes256-sha512.xlsx", password="wrong")


def test_encrypted_rows_match_plaintext():
    with open_workbook(ENCRYPTED / "agile-aes256-sha512.xlsx", password="hunter2") as enc:
        encrypted_rows = list(enc.sheets[0].rows())
    with open_workbook(ENCRYPTED / "agile-aes256-sha512.plain.xlsx") as plain:
        assert list(plain.sheets[0].rows()) == encrypted_rows


def test_password_is_not_in_repr():
    with open_workbook(ENCRYPTED / "agile-aes256-sha512.xlsx", password="hunter2") as book:
        assert "hunter2" not in repr(book)




@pytest.fixture
def batched_csv(tmp_path):
    path = tmp_path / "batched.csv"
    lines = ["name,qty"] + [f"row{i},{i}" for i in range(50)]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


_BATCH_SCHEMA = [
    ColumnSpec(ColumnType.STRING, name="name"),
    ColumnSpec(ColumnType.I64, name="qty"),
]


def _batch_rows(table):
    names, qtys = table.columns
    return [(names[i], qtys[i]) for i in range(table.row_count)]


@pytest.mark.parametrize("batch_size", [1, 7, 8, 9, 1000, 0])
def test_iter_parse_typed_matches_parse_typed(batched_csv, batch_size):
    with open_workbook(batched_csv) as workbook:
        expected = _batch_rows(workbook.sheets[0].parse_typed(_BATCH_SCHEMA))

    with open_workbook(batched_csv) as workbook:
        batches = list(workbook.sheets[0].iter_parse_typed(_BATCH_SCHEMA, batch_size=batch_size))

    assert [row for batch in batches for row in _batch_rows(batch)] == expected

    rows = len(expected)
    want = 1 if batch_size == 0 else (rows + batch_size - 1) // batch_size
    assert len(batches) == want
    if batch_size:
        assert all(batch.row_count <= batch_size for batch in batches)


def test_iter_parse_typed_rejects_a_negative_batch_size(batched_csv):
    with open_workbook(batched_csv) as workbook:
        with pytest.raises(ExcelReaderError):
            next(workbook.sheets[0].iter_parse_typed(_BATCH_SCHEMA, batch_size=-1))


def test_two_readers_on_one_sheet_each_read_every_row(batched_csv):
    with open_workbook(batched_csv) as workbook:
        sheet = workbook.sheets[0]
        first = sheet.iter_parse_typed(_BATCH_SCHEMA, batch_size=4)
        second = sheet.iter_parse_typed(_BATCH_SCHEMA, batch_size=4)

        assert next(first).row_count == 4
        assert next(second).row_count == 4
        assert sum(batch.row_count for batch in first) == 46
        assert sum(batch.row_count for batch in second) == 46


def test_a_whole_sheet_read_leaves_an_open_reader_alone(batched_csv):
    with open_workbook(batched_csv) as workbook:
        sheet = workbook.sheets[0]
        batches = sheet.iter_parse_typed(_BATCH_SCHEMA, batch_size=4)
        next(batches)
        assert sheet.parse_typed(_BATCH_SCHEMA).row_count == 50
        assert sum(batch.row_count for batch in batches) == 46


def _spy_on_native_close(monkeypatch, lib, name):
    closed = []
    real = getattr(lib, name)
    monkeypatch.setattr(lib, name, lambda handle: closed.append(handle.value) or real(handle))
    return closed


def test_abandoning_the_generator_closes_the_reader(batched_csv, monkeypatch):
    with open_workbook(batched_csv) as workbook:
        closed = _spy_on_native_close(monkeypatch, workbook._lib, "xl_typed_reader_close")
        batches = workbook.sheets[0].iter_parse_typed(_BATCH_SCHEMA, batch_size=4)
        next(batches)
        assert closed == []

        batches.close()

        assert len(closed) == 1
        assert closed[0]




def test_arrow_stream_read_by_pyarrow_matches_to_arrow(batched_csv):
    pa = pytest.importorskip("pyarrow")
    with open_workbook(batched_csv) as workbook:
        expected = workbook.sheets[0].to_arrow(_BATCH_SCHEMA)

    with open_workbook(batched_csv) as workbook:
        reader = pa.RecordBatchReader.from_stream(workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=8))
        assert reader.schema == expected.schema
        batches = list(reader)

    assert pa.Table.from_batches(batches).equals(pa.Table.from_batches([expected]))
    assert len(batches) == 7, "50 rows at batch size 8 is 7 batches"
    assert all(batch.num_rows <= 8 for batch in batches)


def test_arrow_stream_builds_a_polars_dataframe_in_batches(batched_csv):
    pl = pytest.importorskip("polars")
    with open_workbook(batched_csv) as workbook:
        df = pl.DataFrame(workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=8))

    assert df.columns == ["name", "qty"]
    assert df["qty"].to_list() == list(range(50))


def test_arrow_stream_is_single_use(batched_csv):
    with open_workbook(batched_csv) as workbook:
        stream = workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA)
        stream.__arrow_c_stream__()
        with pytest.raises(ExcelReaderError):
            stream.__arrow_c_stream__()


def _spy_on_stream_release(monkeypatch):
    released = []
    real = arrow_stream_module._release

    def spy(stream):
        released.append(bool(stream.release))
        real(stream)

    monkeypatch.setattr(arrow_stream_module, "_release", spy)
    return released


def test_an_unconsumed_arrow_stream_releases_the_read(batched_csv, monkeypatch):
    with open_workbook(batched_csv) as workbook:
        released = _spy_on_stream_release(monkeypatch)
        stream = workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=4)
        assert released == []

        del stream
        gc.collect()

        assert released == [True]


def test_a_dropped_capsule_releases_the_read(batched_csv, monkeypatch):
    with open_workbook(batched_csv) as workbook:
        released = _spy_on_stream_release(monkeypatch)
        stream = workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=4)
        capsule = stream.__arrow_c_stream__()
        del stream
        gc.collect()
        assert released == []

        del capsule
        gc.collect()

        assert released == [True]


def test_a_consumer_holding_the_gil_does_not_block_other_threads_while_a_batch_parses(tmp_path):
    path = tmp_path / "large.csv"
    rows = 1_000_000
    path.write_text("name,qty\n" + "".join(f"row{i},{i}\n" for i in range(rows)), encoding="utf-8")
    ticks = []
    stop = threading.Event()

    def tick():
        while not stop.is_set():
            ticks.append(time.perf_counter())

    with open_workbook(path) as workbook:
        capsule = workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=rows).__arrow_c_stream__()
        capsule_pointer = ctypes.PYFUNCTYPE(ctypes.c_void_p, ctypes.py_object, ctypes.c_char_p)(
            ("PyCapsule_GetPointer", ctypes.pythonapi)
        )
        address = capsule_pointer(capsule, b"arrow_array_stream")
        stream = _native.ArrowArrayStream.from_address(address)
        holding_gil = ctypes.PYFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
        get_next = holding_gil(ctypes.cast(stream.get_next, ctypes.c_void_p).value)
        batch = _native.ArrowArray()
        switch_interval = sys.getswitchinterval()
        sys.setswitchinterval(0.001)
        ticker = threading.Thread(target=tick)
        ticker.start()
        try:
            start = time.perf_counter()
            status = get_next(address, ctypes.addressof(batch))
            end = time.perf_counter()
        finally:
            stop.set()
            ticker.join()
            sys.setswitchinterval(switch_interval)
        release_batch = ctypes.CFUNCTYPE(None, ctypes.c_void_p)(batch.release)
        release_batch(ctypes.addressof(batch))
        del capsule

    assert status == 0
    assert batch.length == rows
    inside = [start] + [t for t in ticks if start < t < end] + [end]
    longest_gap = max(b - a for a, b in zip(inside, inside[1:]))
    assert end - start > 0.02, "the batch must parse for many switch intervals for the gap to mean something"
    assert longest_gap < (end - start) / 2


def test_arrow_stream_rejects_a_negative_batch_size(batched_csv):
    with open_workbook(batched_csv) as workbook:
        with pytest.raises(ExcelReaderError):
            workbook.sheets[0].to_arrow_stream(_BATCH_SCHEMA, batch_size=-1)


def test_a_stream_and_a_reader_on_one_sheet_each_read_every_row(batched_csv):
    pa = pytest.importorskip("pyarrow")
    with open_workbook(batched_csv) as workbook:
        sheet = workbook.sheets[0]
        reader = sheet.iter_parse_typed(_BATCH_SCHEMA, batch_size=4)
        next(reader)
        stream = pa.RecordBatchReader.from_stream(sheet.to_arrow_stream(_BATCH_SCHEMA, batch_size=4))

        assert sum(batch.row_count for batch in reader) == 46
        assert sum(batch.num_rows for batch in stream) == 50


def test_a_whole_sheet_read_leaves_an_open_stream_alone(batched_csv):
    pa = pytest.importorskip("pyarrow")
    with open_workbook(batched_csv) as workbook:
        sheet = workbook.sheets[0]
        reader = pa.RecordBatchReader.from_stream(sheet.to_arrow_stream(_BATCH_SCHEMA, batch_size=4))
        reader.read_next_batch()
        assert sheet.parse_typed(_BATCH_SCHEMA).row_count == 50

        assert sum(batch.num_rows for batch in reader) == 46


def test_to_polars_stays_chunked(batched_csv):
    polars = pytest.importorskip("polars")
    with open_workbook(batched_csv) as workbook:
        frame = workbook.sheets[0].to_polars(_BATCH_SCHEMA, batch_size=8)
    assert isinstance(frame, polars.DataFrame)
    assert frame.height == 50
    assert frame.n_chunks() > 1, "to_polars must not rechunk the streamed batches"


def test_to_pandas_still_reads_the_whole_sheet(batched_csv):
    pytest.importorskip("pyarrow")
    pytest.importorskip("pandas")
    with open_workbook(batched_csv) as workbook:
        frame = workbook.sheets[0].to_pandas(_BATCH_SCHEMA, batch_size=8)
    assert len(frame) == 50
    assert list(frame.columns) == ["name", "qty"]


def test_to_pandas_splits_blocks_per_column(batched_csv):
    pytest.importorskip("pyarrow")
    pytest.importorskip("pandas")
    schema = [*_BATCH_SCHEMA, ColumnSpec(ColumnType.I64, index=1)]
    with open_workbook(batched_csv) as workbook:
        frame = workbook.sheets[0].to_pandas(schema, batch_size=8)
    assert frame._mgr.nblocks == 3


_FIXTURE_CSV = Path(__file__).resolve().parents[2] / "tests" / "ExcelReader.Benchmarks" / "Data" / "65K_Records_Data.csv"

_FIXTURE_SCHEMA = [
    ColumnSpec(ColumnType.STRING, name="Region"),
    ColumnSpec(ColumnType.I64, name="Order ID"),
    ColumnSpec(ColumnType.DATE, name="Order Date"),
    ColumnSpec(ColumnType.F64, name="Total Profit"),
]


def test_parse_typed_in_parallel_matches_the_sequential_table():
    with open_workbook(_FIXTURE_CSV) as workbook:
        sequential = workbook.sheets[0].parse_typed(_FIXTURE_SCHEMA)
        parallel = workbook.sheets[0].parse_typed(_FIXTURE_SCHEMA, parallelism=0)

    assert parallel.row_count == sequential.row_count == 65_535
    for expected, actual in zip(sequential.columns, parallel.columns):
        assert list(actual) == list(expected)
    assert parallel.validity == sequential.validity


def test_to_polars_in_parallel_matches_the_streamed_frame():
    pytest.importorskip("polars")
    with open_workbook(_FIXTURE_CSV) as workbook:
        streamed = workbook.sheets[0].to_polars(_FIXTURE_SCHEMA)
        parallel = workbook.sheets[0].to_polars(_FIXTURE_SCHEMA, parallelism=0)

    assert parallel.equals(streamed)


def test_to_pandas_in_parallel_matches_the_streamed_frame():
    pytest.importorskip("pyarrow")
    pytest.importorskip("pandas")
    with open_workbook(_FIXTURE_CSV) as workbook:
        streamed = workbook.sheets[0].to_pandas(_FIXTURE_SCHEMA)
        parallel = workbook.sheets[0].to_pandas(_FIXTURE_SCHEMA, parallelism=0)

    assert parallel.equals(streamed)


def test_parse_typed_rejects_a_negative_parallelism(typed_csv):
    with open_workbook(typed_csv) as workbook, pytest.raises(ExcelReaderError):
        workbook.sheets[0].parse_typed(_TYPED_SCHEMA, parallelism=-1)


def test_infer_schema_parse_text_types_a_csv():
    with open_workbook(_FIXTURE_CSV) as workbook:
        schema = workbook.sheets[0].infer_schema(parse_text=True)
        table = workbook.sheets[0].parse_typed(schema)

    T = ColumnType
    assert [spec.type for spec in schema] == [T.STRING] * 5 + [T.DATE, T.I64, T.DATE, T.I64] + [T.F64] * 5
    assert table.row_count == 65_535


def test_infer_schema_without_parse_text_keeps_csv_columns_as_strings():
    with open_workbook(_FIXTURE_CSV) as workbook:
        schema = workbook.sheets[0].infer_schema()

    assert {spec.type for spec in schema} == {ColumnType.STRING}


@pytest.mark.parametrize(
    ("serial", "date1904", "expected"),
    [
        ("1", False, (1900, 1, 1)),
        ("59", False, (1900, 2, 28)),
        ("61", False, (1900, 3, 1)),
        ("45000.75", False, (2023, 3, 15)),
        ("2958465", False, (9999, 12, 31)),
        ("0", True, (1904, 1, 1)),
        ("43538", True, (2023, 3, 15)),
    ],
)
def test_as_date_follows_excels_calendar(serial, date1904, expected):
    from datetime import date

    assert Cell(0, CellType.DATE, serial).as_date(date1904) == date(*expected)


@pytest.mark.parametrize("serial", ["-1", "2958466", "1e400", "nan"])
def test_as_date_rejects_a_serial_excel_cannot_show(serial):
    with pytest.raises(ValueError, match="not an Excel date serial"):
        Cell(0, CellType.DATE, serial).as_date()


@pytest.mark.parametrize(
    ("text", "expected_micros"),
    [
        ("2024-01-15T13:30:05.0000000Z", 1_705_325_405_000_000),
        ("2024-01-15T13:30:05+03:00", 1_705_314_605_000_000),
        ("2024-01-15T13:30:05", 1_705_325_405_000_000),
    ],
)
def test_parse_typed_timestamp_text_does_not_depend_on_the_local_time_zone(tmp_path, text, expected_micros):
    path = tmp_path / "stamps.csv"
    path.write_text(f"at\n{text}\n", encoding="utf-8")

    with open_workbook(path) as workbook:
        table = workbook.sheets[0].parse_typed([ColumnSpec(ColumnType.TIMESTAMP, name="at")])

    assert list(table.columns[0]) == [expected_micros]
