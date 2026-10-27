import datetime

import pytest

pa = pytest.importorskip("pyarrow")

from excelreader.types import ColumnType, StringColumn
from excelreader.writer import _table_from_arrow

_EPOCH_DATE = datetime.date(1970, 1, 1)
_EPOCH = datetime.datetime(1970, 1, 1)


def _decode(column, column_type, validity, row_count):
    """Reads one converted column back into Python values, None where the validity bit is clear."""
    values = []
    for row in range(row_count):
        if validity is not None and not validity[row // 8] >> (row % 8) & 1:
            values.append(None)
            continue
        if column_type is ColumnType.STRING:
            values.append(column.data[column.offsets[row] : column.offsets[row + 1]].decode("utf-8"))
            continue
        raw = column[row]
        if column_type is ColumnType.DATE:
            values.append(_EPOCH_DATE + datetime.timedelta(days=raw))
        elif column_type is ColumnType.TIME:
            values.append((datetime.datetime.min + datetime.timedelta(microseconds=raw)).time())
        elif column_type is ColumnType.TIMESTAMP:
            values.append(_EPOCH + datetime.timedelta(microseconds=raw))
        elif column_type is ColumnType.BOOL:
            values.append(bool(raw))
        else:
            values.append(raw)
    return values


def _assert_round_trips(data, expected_types):
    table, types = _table_from_arrow(data)
    batch = data.combine_chunks() if isinstance(data, pa.Table) else data
    assert types == expected_types
    assert table.names == list(batch.schema.names)
    for index, column_type in enumerate(types):
        column, validity = table.columns[index], table.validity[index]
        expected = batch.column(index).to_pylist()
        assert _decode(column, column_type, validity, table.row_count) == expected
        assert (validity is None) == (None not in expected)
        if validity is not None:
            assert len(validity) == (table.row_count + 7) // 8
            spare = table.row_count % 8
            assert spare == 0 or validity[-1] >> spare == 0
        if column_type is ColumnType.STRING:
            assert isinstance(column, StringColumn)
            assert column.offsets[0] == 0 and column.offsets[-1] == len(column.data)
    return table


_TYPED = pa.record_batch(
    {
        "text": pa.array(["a", None, "ção ✓", "", "last"]),
        "large": pa.array(["x", "y", None, "zz", ""], pa.large_string()),
        "view": pa.array(["v1", None, "v3", "v4", "v5"], pa.string_view()),
        "i64": pa.array([1, None, -(2**63), 2**63 - 1, 0], pa.int64()),
        "i32": pa.array([7, -7, None, 2**31 - 1, 0], pa.int32()),
        "f64": pa.array([1.5, None, -0.0, 1e308, float("inf")], pa.float64()),
        "f32": pa.array([0.25, 2.5, None, -1.0, 3.0], pa.float32()),
        "flag": pa.array([True, False, None, True, False]),
        "day": pa.array([datetime.date(2020, 1, 2), None, datetime.date(1900, 3, 1), _EPOCH_DATE, datetime.date(9999, 12, 31)]),
        "clock": pa.array([datetime.time(0, 0), datetime.time(23, 59, 59, 999999), None, datetime.time(12, 30), datetime.time(1, 2, 3)]),
        "stamp": pa.array(
            [datetime.datetime(2024, 5, 6, 12, 30), None, datetime.datetime(1969, 12, 31, 23, 59, 59, 1), _EPOCH, datetime.datetime(2000, 1, 1)],
            pa.timestamp("us"),
        ),
    }
)
_TYPES = [
    ColumnType.STRING, ColumnType.STRING, ColumnType.STRING,
    ColumnType.I64, ColumnType.I64, ColumnType.F64, ColumnType.F64,
    ColumnType.BOOL, ColumnType.DATE, ColumnType.TIME, ColumnType.TIMESTAMP,
]


def test_every_supported_type_with_nulls_converts_value_for_value():
    _assert_round_trips(_TYPED, _TYPES)


@pytest.mark.parametrize(("offset", "length"), [(1, 3), (2, 3), (3, 2), (0, 0), (5, 0)])
def test_a_sliced_batch_converts_only_its_own_rows(offset, length):
    table = _assert_round_trips(_TYPED.slice(offset, length), _TYPES)
    assert table.row_count == length


def test_a_chunked_table_converts_as_one_batch():
    chunked = pa.Table.from_batches([_TYPED.slice(0, 2), _TYPED.slice(2, 3), _TYPED.slice(0, 0)])
    assert chunked.column(0).num_chunks > 1
    _assert_round_trips(chunked, _TYPES)


def test_an_empty_table_converts_to_zero_rows():
    table = _assert_round_trips(_TYPED.schema.empty_table(), _TYPES)
    assert table.row_count == 0


def test_columns_without_nulls_carry_no_validity_bitmap():
    batch = pa.record_batch({"n": pa.array(range(20), pa.int64()), "s": pa.array([str(i) for i in range(20)])})
    table = _assert_round_trips(batch, [ColumnType.I64, ColumnType.STRING])
    assert table.validity == [None, None]


def test_an_all_null_column_and_row_counts_past_a_byte_boundary_convert():
    batch = pa.record_batch(
        {
            "s": pa.array([None] * 13, pa.string()),
            "b": pa.array([None if i % 3 == 0 else i % 2 == 0 for i in range(13)]),
            "f": pa.array([None if i % 4 == 0 else i / 3 for i in range(13)]),
        }
    )
    _assert_round_trips(batch, [ColumnType.STRING, ColumnType.BOOL, ColumnType.F64])


def test_a_validity_mask_off_a_byte_boundary_converts(monkeypatch):
    import pyarrow.compute

    is_valid = pyarrow.compute.is_valid
    monkeypatch.setattr(pyarrow.compute, "is_valid", lambda column: pa.concat_arrays([pa.array([False] * 3), is_valid(column)]).slice(3))
    batch = pa.record_batch({"n": pa.array([None if i % 3 == 0 else i for i in range(11)], pa.int64())})
    _assert_round_trips(batch, [ColumnType.I64])


def test_an_unsupported_arrow_type_is_rejected():
    with pytest.raises(ValueError, match="int8"):
        _table_from_arrow(pa.record_batch({"small": pa.array([1], pa.int8())}))
