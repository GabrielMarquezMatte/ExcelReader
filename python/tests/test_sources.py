import gc
import io
import os
import threading
import weakref
from concurrent.futures import ThreadPoolExecutor

import pytest
from excelreader import ExcelReaderError, OpenOptions, open_bytes, open_source, open_stream, open_writer_to_memory
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
    with open_bytes(data) as reference:
        expected = [_first_column(sheet) for sheet in reference.sheets]

    with open_source(BytesSource(data)) as workbook, ThreadPoolExecutor(max_workers=SHEETS * 2) as pool:
        sheets = list(workbook.sheets)
        parallel = list(pool.map(_first_column, sheets + sheets))

    for position, rows in enumerate(parallel):
        assert rows == expected[position % SHEETS]


def test_the_source_is_dropped_after_the_workbook_closes():
    source = BytesSource(_workbook_bytes())
    finalized = weakref.finalize(source, lambda: None)
    before = set(sources_module._live)
    workbook = open_source(source, "xlsx")
    assert len(sources_module._live) == len(before) + 1
    del source
    _first_column(workbook.sheets[0])
    workbook.close()
    gc.collect()
    assert not finalized.alive
    assert set(sources_module._live) == before


def test_a_source_error_reaches_the_caller_with_its_message():
    with pytest.raises(ExcelReaderError, match="disk on fire"):
        open_source(FailingSource(), "xlsx")


def test_the_default_cache_fetches_a_small_workbook_once():
    data = _workbook_bytes()
    cached = BytesSource(data)
    with open_source(cached, "xlsx") as workbook:
        _first_column(workbook.sheets[0])
    assert cached.reads == 1

    uncached = BytesSource(data)
    with open_source(uncached, "xlsx", options=OpenOptions(source_block_size=-1)) as workbook:
        _first_column(workbook.sheets[0])
    assert uncached.reads > 1


def test_one_source_object_opened_twice_at_once_serves_both():
    source = BytesSource(_workbook_bytes())
    first = open_source(source, "xlsx")
    second = open_source(source, "xlsx")
    first.close()
    assert _first_column(second.sheets[2])[0] == "s2-r0"
    second.close()


def test_open_stream_reads_a_csv_from_a_pipe():
    read_fd, write_fd = os.pipe()

    def produce() -> None:
        with os.fdopen(write_fd, "wb") as out:
            out.write(b"name,qty\n")
            for i in range(1000):
                out.write(f"row{i},{i}\n".encode())

    producer = threading.Thread(target=produce)
    producer.start()
    with os.fdopen(read_fd, "rb") as pipe, open_stream(pipe, "csv") as workbook:
        names = _first_column(workbook.sheets[0])
    producer.join()
    assert names[0] == "name"
    assert len(names) == 1001
    assert names[-1] == "row999"


def test_open_stream_reads_an_xlsx_from_an_object_with_only_read():
    with open_stream(OnlyRead(_workbook_bytes())) as workbook:
        rows = _first_column(workbook.sheets[3])
    assert rows[0] == "s3-r0"
    assert len(rows) == ROWS


def test_open_stream_rejects_an_object_that_cannot_be_read():
    with pytest.raises(TypeError):
        open_stream(object())


def test_a_failed_open_leaves_nothing_registered():
    before = set(sources_module._live)
    source = BytesSource(_workbook_bytes())
    with pytest.raises(ValueError):
        open_source(source, "nope")
    with pytest.raises(ValueError):
        open_source(source, "xlsx", password="\ud800")
    negative = BytesSource(b"")
    negative.size = -1
    with pytest.raises(ValueError):
        open_source(negative, "xlsx")
    with pytest.raises(ValueError):
        open_stream(io.BytesIO(), "nope")
    assert set(sources_module._live) == before


def test_open_stream_fails_on_a_text_stream_and_releases_it():
    before = set(sources_module._live)
    with pytest.raises(ExcelReaderError), open_stream(io.StringIO("a,b\n1,2\n"), "csv") as workbook:
        _first_column(workbook.sheets[0])
    assert set(sources_module._live) == before


def test_open_stream_fails_when_read_returns_more_than_asked():
    class Greedy:
        def read(self, size: int) -> bytes:
            return b"x" * (size + 1)

    before = set(sources_module._live)
    with pytest.raises(ExcelReaderError, match="more bytes than requested"), open_stream(Greedy(), "csv") as workbook:
        _first_column(workbook.sheets[0])
    assert set(sources_module._live) == before


def test_a_read_at_that_returns_none_reaches_the_caller_as_an_error():
    class Bad:
        size = 1024

        def read_at(self, offset: int, buffer: memoryview):
            return None

    with pytest.raises(ExcelReaderError):
        open_source(Bad(), "xlsx")
