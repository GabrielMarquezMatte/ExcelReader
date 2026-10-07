"""Open a workbook from bytes the caller serves: a positional `Source`, or a readable stream."""

from __future__ import annotations

import ctypes
import itertools
from typing import Protocol

from excelreader import _native
from excelreader.reader import Workbook, _check, _raw_options, _resolve_format
from excelreader.types import OpenOptions


class Source(Protocol):
    """Random-access bytes for `open_source()`. `read_at` is called from several threads at once."""

    size: int

    def read_at(self, offset: int, buffer: memoryview) -> int:
        """Fills `buffer` with the bytes starting at `offset` and returns how many it wrote."""
        ...


_keys = itertools.count(1)
_live: dict[int, object] = {}


def _fail(error: Exception) -> int:
    message = (str(error) or type(error).__name__).encode("utf-8", errors="replace")
    _native.load_library().xl_set_source_error(message, len(message))
    return -1


def _view(address: int, length: int) -> memoryview:
    return memoryview((ctypes.c_ubyte * length).from_address(address)).cast("B")


@_native.SourceReadAt
def _read_at(key: int, offset: int, buffer: int, length: int) -> int:
    try:
        return _live[key].read_at(offset, _view(buffer, length))
    except Exception as error:  # noqa: BLE001
        return _fail(error)


@_native.StreamRead
def _read(key: int, buffer: int, length: int) -> int:
    try:
        stream = _live[key]
        readinto = getattr(stream, "readinto", None)
        if readinto is None:
            data = stream.read(length)
            ctypes.memmove(buffer, data, len(data))
            return len(data)
        count = readinto(_view(buffer, length))
        if count is None:
            raise OSError("the stream has no data ready; open_stream() needs a blocking stream")
        return count
    except Exception as error:  # noqa: BLE001
        return _fail(error)


@_native.SourceRelease
def _release(key: int) -> None:
    _live.pop(key, None)


def open_source(
    source: Source, format: str | None = None, *, options: OpenOptions | None = None, password: str | None = None
) -> Workbook:
    """Opens a workbook over bytes `source` serves; every sheet can be read on its own thread.

    The library keeps a reference to `source` until the workbook and everything read from it are
    closed. `read_at` runs with the GIL held, so a local file is better opened with `open_workbook()`.
    """
    size = source.size
    lib = _native.load_library()
    resolved = _resolve_format(format, None)
    key = next(_keys)
    _live[key] = source
    raw = _native.NativeSource(ctypes.sizeof(_native.NativeSource), key, size, _read_at, _release)
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    handle = ctypes.c_void_p()
    _check(lib.xl_open_source(ctypes.byref(raw), resolved, raw_options, ctypes.byref(handle)))
    return Workbook(handle)


def open_stream(
    stream: object, format: str | None = None, *, options: OpenOptions | None = None, password: str | None = None
) -> Workbook:
    """Opens a workbook over a stream that cannot seek, such as `sys.stdin.buffer` or a socket.

    A CSV (`format="csv"`) is read as it arrives and only once; an XLSX, XLSB or XLS is read whole
    before the workbook opens. `stream` needs `readinto` or `read`; the library never closes it.
    """
    if not (hasattr(stream, "readinto") or hasattr(stream, "read")):
        raise TypeError("open_stream() needs an object with readinto() or read()")
    lib = _native.load_library()
    resolved = _resolve_format(format, None)
    key = next(_keys)
    _live[key] = stream
    raw = _native.NativeStream(ctypes.sizeof(_native.NativeStream), key, _read, _release)
    raw_options, _raw, _password_buffer = _raw_options(options, password)
    handle = ctypes.c_void_p()
    _check(lib.xl_open_stream(ctypes.byref(raw), resolved, raw_options, ctypes.byref(handle)))
    return Workbook(handle)
