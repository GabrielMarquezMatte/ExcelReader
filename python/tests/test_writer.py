import io
from array import array

import pytest
from excelreader import (
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    PasswordIncorrectError,
    StringColumn,
    TypedTable,
    WriteOptions,
    encrypt_package,
    open_workbook,
    write_workbook,
)

_SCHEMA = [
    ColumnSpec(ColumnType.STRING, name="name"),
    ColumnSpec(ColumnType.I64, name="qty"),
]
_TYPES = [spec.type for spec in _SCHEMA]


@pytest.fixture
def source_csv(tmp_path):
    path = tmp_path / "source.csv"
    path.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    return path


@pytest.mark.parametrize("extension", ["xlsx", "xlsb", "xls", "csv"])
def test_write_workbook_round_trips_every_format(source_csv, tmp_path, extension):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    out = tmp_path / f"out.{extension}"
    write_workbook(out, table, _TYPES)

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert result.row_count == 2
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_round_trips_a_nullable_column(tmp_path):
    source = tmp_path / "nullable.csv"
    source.write_text("qty\n3\n\n9\n", encoding="utf-8")
    schema = [ColumnSpec(ColumnType.I64, name="qty", nullable=True)]

    with open_workbook(source) as workbook:
        table = workbook.sheets[0].parse(schema)

    out = tmp_path / "out.xlsx"
    write_workbook(out, table, [spec.type for spec in schema])

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(schema)

    assert result.row_count == 3
    assert result.validity[0][0] & 0b111 == 0b101


def test_write_workbook_applies_the_sheet_name(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    out = tmp_path / "named.xlsx"
    write_workbook(out, table, _TYPES, options=WriteOptions(sheet_name="Vendas"))

    with open_workbook(out) as workbook:
        assert [sheet.name for sheet in workbook.sheets] == ["Vendas"]


def test_write_workbook_rejects_an_unknown_extension(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(ValueError, match="format"):
        write_workbook(tmp_path / "out.parquet", table, _TYPES)


@pytest.mark.parametrize("extension", ["xlsx", "xlsb"])
def test_write_workbook_in_parallel_round_trips(tmp_path, extension):
    source = tmp_path / "nullable.csv"
    source.write_text("qty\n0\n10\n\n30\n40\n50\n60\n", encoding="utf-8")
    schema = [ColumnSpec(ColumnType.I64, name="qty", nullable=True)]

    with open_workbook(source) as workbook:
        table = workbook.sheets[0].parse(schema)

    out = tmp_path / f"out.{extension}"
    options = WriteOptions(degree_of_parallelism=4, batch_rows=2)
    write_workbook(out, table, [spec.type for spec in schema], options=options)

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(schema)

    assert result.row_count == 7
    assert [result.columns[0][i] for i in range(7) if i != 2] == [0, 10, 30, 40, 50, 60]
    assert result.validity[0][0] & 0b0111_1111 == 0b0111_1011


def test_write_workbook_rejects_a_negative_degree_of_parallelism(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(ExcelReaderError, match="degree_of_parallelism"):
        write_workbook(tmp_path / "out.xlsx", table, _TYPES, options=WriteOptions(degree_of_parallelism=-1))


def test_write_workbook_reports_a_native_rejection(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(ExcelReaderError):
        write_workbook(tmp_path / "out.xlsx", table, _TYPES, options=WriteOptions(sheet_name="has/slash"))


def test_write_workbook_rejects_a_typed_table_without_types(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(TypeError, match="types"):
        write_workbook(tmp_path / "out.xlsx", table)


def test_write_workbook_rejects_types_for_arrow_data(tmp_path):
    pa = pytest.importorskip("pyarrow")

    batch = pa.RecordBatch.from_pydict({"name": ["widget"], "qty": [3]})
    with pytest.raises(TypeError, match="types"):
        write_workbook(tmp_path / "out.xlsx", batch, _TYPES)


def test_write_workbook_rejects_data_it_cannot_write(tmp_path):
    pytest.importorskip("pyarrow")

    with pytest.raises(TypeError, match="data"):
        write_workbook(tmp_path / "out.xlsx", [["widget", 3]])


@pytest.mark.parametrize("kind", ["record_batch", "table"])
def test_write_workbook_writes_arrow_data(tmp_path, kind):
    pa = pytest.importorskip("pyarrow")

    data = {"name": ["widget", "gadget"], "qty": [3, 7]}
    frame = pa.RecordBatch.from_pydict(data) if kind == "record_batch" else pa.table(data)
    out = tmp_path / "arrow.xlsx"
    write_workbook(out, frame)

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_leaves_a_pandas_index_out(tmp_path):
    pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    out = tmp_path / "pandas.xlsx"
    write_workbook(out, pd.DataFrame({"name": ["widget", "gadget"], "qty": [3, 7]}, index=[10, 20]))

    with open_workbook(out) as workbook:
        rows = [[cell.value for cell in row] for row in workbook.sheets[0].rows()]

    assert rows == [["name", "qty"], ["widget", "3"], ["gadget", "7"]]


def test_write_workbook_writes_any_arrow_c_stream(tmp_path):
    pa = pytest.importorskip("pyarrow")

    class Producer:
        def __arrow_c_stream__(self, requested_schema=None):
            return pa.table({"name": ["widget", "gadget"], "qty": [3, 7]}).__arrow_c_stream__(requested_schema)

    out = tmp_path / "stream.xlsx"
    write_workbook(out, Producer())

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_round_trips_a_multi_chunk_pandas_frame(tmp_path):
    pa = pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    chunked = pa.Table.from_batches(
        [
            pa.RecordBatch.from_pydict({"name": ["widget"], "qty": [3]}),
            pa.RecordBatch.from_pydict({"name": ["gadget"], "qty": [7]}),
        ]
    )
    frame = chunked.to_pandas(split_blocks=True)
    assert pa.Table.from_pandas(frame, preserve_index=False).column(0).num_chunks > 1

    out = tmp_path / "chunked.xlsx"
    write_workbook(out, frame)

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_writes_a_header_only_sheet_for_an_empty_frame(tmp_path):
    pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    out = tmp_path / "empty.xlsx"
    write_workbook(
        out,
        pd.DataFrame({"name": pd.Series([], dtype="string"), "qty": pd.Series([], dtype="int64")}),
    )

    with open_workbook(out) as workbook:
        rows = list(workbook.sheets[0].rows())

    assert [cell.value for cell in rows[0]] == ["name", "qty"]
    assert len(rows) == 1


def test_write_workbook_writes_a_polars_frame(tmp_path):
    pytest.importorskip("pyarrow")
    pl = pytest.importorskip("polars")

    out = tmp_path / "polars.xlsx"
    write_workbook(out, pl.DataFrame({"name": ["widget", "gadget"], "qty": [3, 7]}))

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_rejects_an_unsupported_arrow_type(tmp_path):
    pa = pytest.importorskip("pyarrow")

    batch = pa.RecordBatch.from_pydict({"nested": [[1, 2], [3]]})

    with pytest.raises(ValueError, match="type"):
        write_workbook(tmp_path / "bad.xlsx", batch)


def test_encrypt_package_round_trips_through_open_workbook(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    encrypted = tmp_path / "encrypted.xlsx"
    encrypt_package(plain, encrypted, "hunter2")

    with open_workbook(encrypted, password="hunter2") as workbook:
        result = workbook.sheets[0].parse(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_encrypt_package_rejects_the_wrong_password_on_open(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    encrypted = tmp_path / "encrypted.xlsx"
    encrypt_package(plain, encrypted, "hunter2")

    with pytest.raises(PasswordIncorrectError):
        open_workbook(encrypted, password="wrong")


def test_encrypt_package_rejects_an_empty_password(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    with pytest.raises(ExcelReaderError):
        encrypt_package(plain, tmp_path / "encrypted.xlsx", "")


def _parse_bytes(payload, format, password=None):
    with open_workbook(payload, format=format, password=password) as workbook:
        return workbook.sheets[0].parse(_SCHEMA)


def _to_bytes(table):
    buffer = io.BytesIO()
    write_workbook(buffer, table, _TYPES, format="xlsx")
    return buffer.getvalue()


@pytest.mark.parametrize("frame_module", [None, "pandas", "polars"])
def test_frames_round_trip_through_a_bytesio(frame_module):
    pa = pytest.importorskip("pyarrow")
    data = {"name": ["widget", "gadget"], "qty": [3, 7]}
    if frame_module is None:
        frame = pa.RecordBatch.from_pydict(data)
    else:
        frame = pytest.importorskip(frame_module).DataFrame(data)

    buffer = io.BytesIO()
    write_workbook(buffer, frame, format="xlsx", options=WriteOptions(sheet_name="Vendas"))
    result = _parse_bytes(buffer.getvalue(), "xlsx")

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


@pytest.mark.parametrize("package_kind", ["path", "bytes", "bytesio"])
@pytest.mark.parametrize("destination_kind", ["path", "bytesio"])
def test_encrypt_package_takes_paths_and_buffers_on_both_sides(
    tmp_path, source_csv, package_kind, destination_kind
):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)
    plain = _to_bytes(table)
    plain_path = tmp_path / "plain.xlsx"
    plain_path.write_bytes(plain)
    package = {"path": plain_path, "bytes": plain, "bytesio": io.BytesIO(plain)}[package_kind]
    destination = tmp_path / "secret.xlsx" if destination_kind == "path" else io.BytesIO()

    encrypt_package(package, destination, "hunter2")

    encrypted = destination.read_bytes() if destination_kind == "path" else destination.getvalue()
    result = _parse_bytes(encrypted, "xlsx", password="hunter2")
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]
    with pytest.raises(PasswordIncorrectError):
        open_workbook(encrypted, format="xlsx", password="wrong")


def test_encrypt_package_rejects_bytes_that_are_not_a_package():
    with pytest.raises(ExcelReaderError):
        encrypt_package(b"not a zip", io.BytesIO(), "hunter2")


def test_encrypt_package_rejects_an_empty_password_in_memory(source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(ExcelReaderError):
        encrypt_package(_to_bytes(table), io.BytesIO(), "")


def test_encrypt_package_rejects_inputs_it_cannot_use():
    with pytest.raises(TypeError, match="package"):
        encrypt_package(42, io.BytesIO(), "hunter2")
    with pytest.raises(TypeError, match="destination"):
        encrypt_package(b"x", 42, "hunter2")


def test_write_workbook_writes_into_an_open_binary_file(tmp_path, source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)
    path = tmp_path / "out.bin"

    with path.open("wb") as file:
        write_workbook(file, table, _TYPES, format="xlsb")

    result = _parse_bytes(path.read_bytes(), "xlsb")
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_to_a_stream_needs_a_format(source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(ValueError, match="format"):
        write_workbook(io.BytesIO(), table, _TYPES)
    with pytest.raises(ValueError, match="unknown format"):
        write_workbook(io.BytesIO(), table, _TYPES, format="parquet")


def test_write_workbook_rejects_a_destination_it_cannot_write(source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse(_SCHEMA)

    with pytest.raises(TypeError, match="destination"):
        write_workbook(42, table, _TYPES, format="xlsx")


def _one_column_table(column, rows=1, validity=None):
    return TypedTable(row_count=rows, names=["n"], columns=[column], validity=[validity])


@pytest.mark.parametrize(
    ("column", "column_type", "rows", "message"),
    [
        (array("q", [7]), ColumnType.I64, 10, "1 elements but 10"),
        (array("i", [1, 2]), ColumnType.I64, 1, "8-byte"),
        (array("d", [1.0]), ColumnType.I64, 1, "format"),
        (array("Q", [1]), ColumnType.I64, 1, "format"),
        (array("q", [1]), ColumnType.DATE, 1, "4-byte"),
        (StringColumn(array("i", [0]), b""), ColumnType.STRING, 1, "offsets"),
        (StringColumn(array("q", [0, 1]), b"a"), ColumnType.STRING, 1, "offsets"),
    ],
)
def test_write_workbook_rejects_a_column_buffer_that_does_not_match(column, column_type, rows, message):
    with pytest.raises((TypeError, ValueError), match=message):
        write_workbook(io.BytesIO(), _one_column_table(column, rows), [column_type], format="xlsx")


def test_write_workbook_rejects_a_short_validity_bitmap():
    table = _one_column_table(array("q", range(9)), rows=9, validity=b"\xff")

    with pytest.raises(ValueError, match="validity"):
        write_workbook(io.BytesIO(), table, [ColumnType.I64], format="xlsx")


def test_write_workbook_takes_numpy_columns():
    numpy = pytest.importorskip("numpy")
    table = TypedTable(
        row_count=2,
        names=["qty", "price", "day", "ok"],
        columns=[
            numpy.array([3, 7], dtype=numpy.int64),
            numpy.array([1.5, 2.5]),
            numpy.array([0, 1], dtype=numpy.int32),
            numpy.array([True, False]),
        ],
        validity=[None, None, None, None],
    )
    types = [ColumnType.I64, ColumnType.F64, ColumnType.DATE, ColumnType.BOOL]
    buffer = io.BytesIO()

    write_workbook(buffer, table, types, format="xlsx")

    schema = [ColumnSpec(t, name=n) for n, t in zip(table.names, types)]
    with open_workbook(buffer.getvalue(), format="xlsx") as workbook:
        result = workbook.sheets[0].parse(schema)
    assert list(result.columns[0]) == [3, 7]
    assert list(result.columns[1]) == [1.5, 2.5]
    assert list(result.columns[3]) == [1, 0]
