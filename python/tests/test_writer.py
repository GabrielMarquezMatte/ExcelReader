import pytest

from excelreader import (
    ColumnSpec,
    ColumnType,
    ExcelReaderError,
    WriteOptions,
    open_workbook,
    write_arrow,
    write_pandas,
    write_polars,
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
        table = workbook.parse_typed(_SCHEMA)

    out = tmp_path / f"out.{extension}"
    write_workbook(out, table, _TYPES)

    with open_workbook(out) as workbook:
        result = workbook.parse_typed(_SCHEMA)

    assert result.row_count == 2
    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_workbook_round_trips_a_nullable_column(tmp_path):
    source = tmp_path / "nullable.csv"
    source.write_text("qty\n3\n\n9\n", encoding="utf-8")
    schema = [ColumnSpec(ColumnType.I64, name="qty", nullable=True)]

    with open_workbook(source) as workbook:
        table = workbook.parse_typed(schema)

    out = tmp_path / "out.xlsx"
    write_workbook(out, table, [spec.type for spec in schema])

    with open_workbook(out) as workbook:
        result = workbook.parse_typed(schema)

    assert result.row_count == 3
    # Bit 1 clear: valid, null, valid.
    assert result.validity[0][0] & 0b111 == 0b101


def test_write_workbook_applies_the_sheet_name(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.parse_typed(_SCHEMA)

    out = tmp_path / "named.xlsx"
    write_workbook(out, table, _TYPES, options=WriteOptions(sheet_name="Vendas"))

    with open_workbook(out) as workbook:
        assert workbook.sheet_names == ["Vendas"]


def test_write_workbook_rejects_an_unknown_extension(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.parse_typed(_SCHEMA)

    with pytest.raises(ValueError, match="format"):
        write_workbook(tmp_path / "out.parquet", table, _TYPES)


def test_write_workbook_reports_a_native_rejection(source_csv, tmp_path):
    with open_workbook(source_csv) as workbook:
        table = workbook.parse_typed(_SCHEMA)

    with pytest.raises(ExcelReaderError):
        write_workbook(tmp_path / "out.xlsx", table, _TYPES, options=WriteOptions(sheet_name="has/slash"))


def test_write_arrow_round_trips(tmp_path):
    pa = pytest.importorskip("pyarrow")

    batch = pa.RecordBatch.from_pydict({"name": ["widget", "gadget"], "qty": [3, 7]})
    out = tmp_path / "arrow.xlsx"
    write_arrow(out, batch)

    with open_workbook(out) as workbook:
        result = workbook.parse_typed(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_pandas_round_trips(tmp_path):
    pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")

    out = tmp_path / "pandas.xlsx"
    write_pandas(out, pd.DataFrame({"name": ["widget", "gadget"], "qty": [3, 7]}))

    with open_workbook(out) as workbook:
        result = workbook.parse_typed(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_polars_round_trips(tmp_path):
    pytest.importorskip("pyarrow")
    pl = pytest.importorskip("polars")

    out = tmp_path / "polars.xlsx"
    write_polars(out, pl.DataFrame({"name": ["widget", "gadget"], "qty": [3, 7]}))

    with open_workbook(out) as workbook:
        result = workbook.parse_typed(_SCHEMA)

    assert list(result.columns[0]) == ["widget", "gadget"]
    assert list(result.columns[1]) == [3, 7]


def test_write_arrow_rejects_an_unsupported_arrow_type(tmp_path):
    pa = pytest.importorskip("pyarrow")

    batch = pa.RecordBatch.from_pydict({"nested": [[1, 2], [3]]})

    with pytest.raises(ValueError, match="type"):
        write_arrow(tmp_path / "bad.xlsx", batch)
