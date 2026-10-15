"""Reading a workbook through callbacks: a positional `Source`, or a readable stream."""

from __future__ import annotations

import ctypes
import itertools
import operator
from typing import Protocol

from excelreader import _native
from excelreader.reader import Workbook, _check, _raw_options, _resolve_format
from excelreader.types import OpenOptions


class Source(Protocol):
    """Random-access bytes for `open_workbook()`. `read_at` is called from several threads at once."""

    size: int

    def read_at(self, offset: int, buffer: memoryview) -> int:
        """Fills `buffer` with the bytes starting at `offset` and returns how many it wrote.

        `buffer` is valid only during the call; do not retain it.
        """
        ...


_keys = itertools.count(1)
_live: dict[int, object] = {}


def _fail(error: BaseException) -> int:
    message = (str(error) or type(error).__name__).encode("utf-8", errors="replace")
    _native.load_library().xl_set_source_error(message, len(message))
    return -1


def _view(address: int, length: int) -> memoryview:
    return memoryview((ctypes.c_ubyte * length).from_address(address)).cast("B")


@_native.SourceReadAt
def _read_at(key: int, offset: int, buffer: int, length: int) -> int:
    try:
        return operator.index(_live[key].read_at(offset, _view(buffer, length)))
    except BaseException as error:  # noqa: BLE001
        return _fail(error)


@_native.StreamRead
def _read(key: int, buffer: int, length: int) -> int:
    try:
        stream = _live[key]
        readinto = getattr(stream, "readinto", None)
        if readinto is None:
            data = bytes(stream.read(length))
            if len(data) > length:
                raise OSError("the stream returned more bytes than requested")
            ctypes.memmove(buffer, data, len(data))
            return len(data)
        count = readinto(_view(buffer, length))
        if count is None:
            raise OSError("the stream has no data ready; open_workbook() needs a blocking stream")
        count = operator.index(count)
        if count > length:
            raise OSError("the stream returned more bytes than requested")
        return count
    except BaseException as error:  # noqa: BLE001
        return _fail(error)


@_native.SourceRelease
def _release(key: int) -> None:
    _live.pop(key, None)


def _open_source(
    source: Source, format: str | None, options: OpenOptions | None, password: str | None
) -> Workbook:
    size = operator.index(source.size)
    if size < 0:
        raise ValueError("source.size must not be negative")
    lib = _native.load_library()
    resolved = _resolve_format(format, None)
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    key = next(_keys)
    raw = _native.NativeSource(ctypes.sizeof(_native.NativeSource), key, size, _read_at, _release)
    handle = ctypes.c_void_p()
    _live[key] = source
    _check(lib.xl_open_source(ctypes.byref(raw), resolved, raw_options, ctypes.byref(handle)))
    return Workbook(handle)


def _open_stream(
    stream: object, format: str | None, options: OpenOptions | None, password: str | None
) -> Workbook:
    lib = _native.load_library()
    resolved = _resolve_format(format, None)
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    key = next(_keys)
    raw = _native.NativeStream(ctypes.sizeof(_native.NativeStream), key, _read, _release)
    handle = ctypes.c_void_p()
    _live[key] = stream
    _check(lib.xl_open_stream(ctypes.byref(raw), resolved, raw_options, ctypes.byref(handle)))
    return Workbook(handle)
