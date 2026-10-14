import gc
import io
import os
import threading
import weakref
from concurrent.futures import ThreadPoolExecutor

import pytest
from excelreader import ExcelReaderError, OpenOptions, open_workbook, open_writer_to_memory
from excelreader import sources as sources_module

SHEETS = 4
ROWS = 300


def _workbook_bytes() -> bytes:
    with open_writer_to_memory("xlsx") as writer:
        for sheet in range(SHEETS):
            writer.start_sheet(f"sheet{sheet}")
            for row in range(ROWS):
                writer.write_row([f"s{sheet}-r{row}"])
            writer.end_sheet()
        return writer.bytes()


def _first_column(sheet) -> list[str]:
    return [row[0].value for row in sheet.rows()]


class BytesSource:
    def __init__(self, data: bytes) -> None:
        self.data = data
        self.size = len(data)
        self.reads = 0
        self._lock = threading.Lock()

    def read_at(self, offset: int, buffer: memoryview) -> int:
        with self._lock:
            self.reads += 1
        chunk = self.data[offset : offset + len(buffer)]
        buffer[: len(chunk)] = chunk
        return len(chunk)


class FailingSource:
    size = 1024

    def read_at(self, offset: int, buffer: memoryview) -> int:
        raise OSError("disk on fire")


class OnlyRead:
    def __init__(self, data: bytes) -> None:
        self._inner = io.BytesIO(data)

    def read(self, size: int) -> bytes:
        return self._inner.read(min(size, 7))


def test_every_sheet_of_a_source_reads_in_parallel_like_memory():
    data = _workbook_bytes()
    with open_workbook(data) as reference:
        expected = [_first_column(sheet) for sheet in reference.sheets]

    with open_workbook(BytesSource(data)) as workbook, ThreadPoolExecutor(max_workers=SHEETS * 2) as pool:
        sheets = list(workbook.sheets)
        parallel = list(pool.map(_first_column, sheets + sheets))

    for position, rows in enumerate(parallel):
        assert rows == expected[position % SHEETS]


def test_the_source_is_dropped_after_the_workbook_closes():
    source = BytesSource(_workbook_bytes())
    finalized = weakref.finalize(source, lambda: None)
    before = set(sources_module._live)
    workbook = open_workbook(source, "xlsx")
    assert len(sources_module._live) == len(before) + 1
    del source
    _first_column(workbook.sheets[0])
    workbook.close()
    gc.collect()
    assert not finalized.alive
    assert set(sources_module._live) == before


def test_a_source_error_reaches_the_caller_with_its_message():
    with pytest.raises(ExcelReaderError, match="disk on fire"):
        open_workbook(FailingSource(), "xlsx")


def test_the_default_cache_fetches_a_small_workbook_once():
    data = _workbook_bytes()
    cached = BytesSource(data)
    with open_workbook(cached, "xlsx") as workbook:
        _first_column(workbook.sheets[0])
    assert cached.reads == 1

    uncached = BytesSource(data)
    with open_workbook(uncached, "xlsx", options=OpenOptions(source_block_size=-1)) as workbook:
        _first_column(workbook.sheets[0])
    assert uncached.reads > 1


def test_one_source_object_opened_twice_at_once_serves_both():
    source = BytesSource(_workbook_bytes())
    first = open_workbook(source, "xlsx")
    second = open_workbook(source, "xlsx")
    first.close()
    assert _first_column(second.sheets[2])[0] == "s2-r0"
    second.close()


def test_a_stream_reads_a_csv_from_a_pipe():
    read_fd, write_fd = os.pipe()

    def produce() -> None:
        with os.fdopen(write_fd, "wb") as out:
            out.write(b"name,qty\n")
            for i in range(1000):
                out.write(f"row{i},{i}\n".encode())

    producer = threading.Thread(target=produce)
    producer.start()
    with os.fdopen(read_fd, "rb") as pipe, open_workbook(pipe, "csv") as workbook:
        names = _first_column(workbook.sheets[0])
    producer.join()
    assert names[0] == "name"
    assert len(names) == 1001
    assert names[-1] == "row999"


def test_a_csv_stream_serves_a_schema_inference_and_then_every_row():
    data = b"name,qty\n" + b"".join(f"row{i},{i}\n".encode() for i in range(5000))
    with open_workbook(OnlyRead(data), "csv") as workbook:
        sheet = workbook.sheets[0]
        schema = sheet.infer_schema(parse_text=True)
        names = _first_column(sheet)
    assert len(schema) == 2
    assert len(names) == 5001
    assert names[-1] == "row4999"


def test_a_stream_reads_an_xlsx_from_an_object_with_only_read():
    with open_workbook(OnlyRead(_workbook_bytes())) as workbook:
        rows = _first_column(workbook.sheets[3])
    assert rows[0] == "s3-r0"
    assert len(rows) == ROWS


def test_open_workbook_rejects_an_object_it_cannot_read():
    with pytest.raises(TypeError, match="object"):
        open_workbook(object())


@pytest.mark.parametrize(
    "wrap",
    [
        pytest.param(lambda data, path: str(path), id="str"),
        pytest.param(lambda data, path: path, id="path"),
        pytest.param(lambda data, path: data, id="bytes"),
        pytest.param(lambda data, path: bytearray(data), id="bytearray"),
        pytest.param(lambda data, path: memoryview(data), id="memoryview"),
        pytest.param(lambda data, path: io.BytesIO(data), id="bytesio"),
        pytest.param(lambda data, path: BytesSource(data), id="source"),
        pytest.param(lambda data, path: OnlyRead(data), id="stream"),
        pytest.param(lambda data, path: path.open("rb"), id="file"),
    ],
)
def test_open_workbook_reads_the_same_rows_from_every_kind_of_input(tmp_path, wrap):
    data = _workbook_bytes()
    path = tmp_path / "book.xlsx"
    path.write_bytes(data)
    source = wrap(data, path)
    try:
        with open_workbook(source) as workbook:
            rows = [_first_column(sheet) for sheet in workbook.sheets]
    finally:
        if hasattr(source, "close"):
            source.close()
    assert rows == [[f"s{sheet}-r{row}" for row in range(ROWS)] for sheet in range(SHEETS)]


def test_a_bytesio_is_read_from_its_start_whatever_its_position():
    buffer = io.BytesIO(_workbook_bytes())
    buffer.seek(0, io.SEEK_END)
    with open_workbook(buffer) as workbook:
        assert _first_column(workbook.sheets[0])[0] == "s0-r0"


def test_a_failed_open_leaves_nothing_registered():
    before = set(sources_module._live)
    source = BytesSource(_workbook_bytes())
    with pytest.raises(ValueError):
        open_workbook(source, "nope")
    with pytest.raises(ValueError):
        open_workbook(source, "xlsx", password="\ud800")
    negative = BytesSource(b"")
    negative.size = -1
    with pytest.raises(ValueError):
        open_workbook(negative, "xlsx")
    with pytest.raises(ValueError):
        open_workbook(OnlyRead(b""), "nope")
    assert set(sources_module._live) == before


def test_a_stream_fails_on_a_text_stream_and_releases_it():
    before = set(sources_module._live)
    with pytest.raises(ExcelReaderError), open_workbook(io.StringIO("a,b\n1,2\n"), "csv") as workbook:
        _first_column(workbook.sheets[0])
    assert set(sources_module._live) == before


def test_a_stream_fails_when_read_returns_more_than_asked():
    class Greedy:
        def read(self, size: int) -> bytes:
            return b"x" * (size + 1)

    before = set(sources_module._live)
    with pytest.raises(ExcelReaderError, match="more bytes than requested"), open_workbook(Greedy(), "csv") as workbook:
        _first_column(workbook.sheets[0])
    assert set(sources_module._live) == before


def test_a_stream_fails_when_read_raises_keyboard_interrupt_and_releases_it():
    class Interrupted:
        def read(self, size: int) -> bytes:
            raise KeyboardInterrupt

    before = set(sources_module._live)
    with pytest.raises(ExcelReaderError), open_workbook(Interrupted(), "csv") as workbook:
        _first_column(workbook.sheets[0])
    assert set(sources_module._live) == before


def test_a_read_at_that_returns_none_reaches_the_caller_as_an_error():
    class Bad:
        size = 1024

        def read_at(self, offset: int, buffer: memoryview):
            return None

    with pytest.raises(ExcelReaderError):
        open_workbook(Bad(), "xlsx")
