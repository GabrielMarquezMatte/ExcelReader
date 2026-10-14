import io

import pytest
from excelreader import (
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    PasswordIncorrectError,
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
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    out = tmp_path / f"out.{extension}"
    write_workbook(out, table, _TYPES)

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse_typed(_SCHEMA)

    assert result.row_count == 2
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_round_trips_a_nullable_column(tmp_path):
    source = tmp_path / "nullable.csv"
    source.write_text("qty\n3\n\n9\n", encoding="utf-8")
    schema = [ColumnSpec(ColumnType.I64, name="qty", nullable=True)]

    with open_workbook(source) as workbook:
        table = workbook.sheets[0].parse_typed(schema)

    out = tmp_path / "out.xlsx"
    write_workbook(out, table, [spec.type for spec in schema])

    with open_workbook(out) as workbook:
        result = workbook.sheets[0].parse_typed(schema)

    assert result.row_count == 3
    assert result.validity[0][0] & 0b111 == 0b101


def test_write_workbook_applies_the_sheet_name(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    out = tmp_path / "named.xlsx"
    write_workbook(out, table, _TYPES, options=WriteOptions(sheet_name="Vendas"))

    with open_workbook(out) as workbook:
        assert [sheet.name for sheet in workbook.sheets] == ["Vendas"]


def test_write_workbook_rejects_an_unknown_extension(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    with pytest.raises(ValueError, match="format"):
        write_workbook(tmp_path / "out.parquet", table, _TYPES)


def test_write_workbook_reports_a_native_rejection(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    with pytest.raises(ExcelReaderError):
        write_workbook(tmp_path / "out.xlsx", table, _TYPES, options=WriteOptions(sheet_name="has/slash"))


def test_write_workbook_rejects_a_typed_table_without_types(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

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
        result = workbook.sheets[0].parse_typed(_SCHEMA)

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
        result = workbook.sheets[0].parse_typed(_SCHEMA)

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
        result = workbook.sheets[0].parse_typed(_SCHEMA)

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
        result = workbook.sheets[0].parse_typed(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_rejects_an_unsupported_arrow_type(tmp_path):
    pa = pytest.importorskip("pyarrow")

    batch = pa.RecordBatch.from_pydict({"nested": [[1, 2], [3]]})

    with pytest.raises(ValueError, match="type"):
        write_workbook(tmp_path / "bad.xlsx", batch)


def test_encrypt_package_round_trips_through_open_workbook(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    encrypted = tmp_path / "encrypted.xlsx"
    encrypt_package(plain, encrypted, "hunter2")

    with open_workbook(encrypted, password="hunter2") as workbook:
        result = workbook.sheets[0].parse_typed(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_encrypt_package_rejects_the_wrong_password_on_open(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    encrypted = tmp_path / "encrypted.xlsx"
    encrypt_package(plain, encrypted, "hunter2")

    with pytest.raises(PasswordIncorrectError):
        open_workbook(encrypted, password="wrong")


def test_encrypt_package_rejects_an_empty_password(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    plain = tmp_path / "plain.xlsx"
    write_workbook(plain, table, _TYPES)

    with pytest.raises(ExcelReaderError):
        encrypt_package(plain, tmp_path / "encrypted.xlsx", "")


def _parse_bytes(payload, format, password=None):
    with open_workbook(payload, format=format, password=password) as workbook:
        return workbook.sheets[0].parse_typed(_SCHEMA)


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
        table = workbook.sheets[0].parse_typed(_SCHEMA)
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
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    with pytest.raises(ExcelReaderError):
        encrypt_package(_to_bytes(table), io.BytesIO(), "")


def test_encrypt_package_rejects_inputs_it_cannot_use():
    with pytest.raises(TypeError, match="package"):
        encrypt_package(42, io.BytesIO(), "hunter2")
    with pytest.raises(TypeError, match="destination"):
        encrypt_package(b"x", 42, "hunter2")


def test_write_workbook_writes_into_an_open_binary_file(tmp_path, source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)
    path = tmp_path / "out.bin"

    with path.open("wb") as file:
        write_workbook(file, table, _TYPES, format="xlsb")

    result = _parse_bytes(path.read_bytes(), "xlsb")
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_to_a_stream_needs_a_format(source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    with pytest.raises(ValueError, match="format"):
        write_workbook(io.BytesIO(), table, _TYPES)
    with pytest.raises(ValueError, match="unknown format"):
        write_workbook(io.BytesIO(), table, _TYPES, format="parquet")


def test_write_workbook_rejects_a_destination_it_cannot_write(source_csv):
    with open_workbook(source_csv) as workbook:
        table = workbook.sheets[0].parse_typed(_SCHEMA)

    with pytest.raises(TypeError, match="destination"):
        write_workbook(42, table, _TYPES, format="xlsx")
