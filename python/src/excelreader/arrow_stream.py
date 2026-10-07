"""An Arrow C stream exposed through the Arrow PyCapsule interface, with no pyarrow dependency."""

from __future__ import annotations

import ctypes

from excelreader import _native
from excelreader.types import ExcelReaderError

_CAPSULE_NAME = b"arrow_array_stream"

_capsule_new = ctypes.pythonapi.PyCapsule_New
_capsule_new.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_void_p]
_capsule_new.restype = ctypes.py_object

_capsule_pointer = ctypes.pythonapi.PyCapsule_GetPointer
_capsule_pointer.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
_capsule_pointer.restype = ctypes.c_void_p

_ReadFn = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
_ErrorFn = ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p)
_ReleaseFn = ctypes.CFUNCTYPE(None, ctypes.c_void_p)


class _Target:
    """The native stream behind a proxy, with its callbacks re-typed so ctypes releases the GIL."""

    def __init__(self, stream: _native.ArrowArrayStream) -> None:
        self.stream = stream
        self.address = ctypes.addressof(stream)
        self.get_schema = _ReadFn(_address_of(stream.get_schema))
        self.get_next = _ReadFn(_address_of(stream.get_next))
        self.get_last_error = _ErrorFn(_address_of(stream.get_last_error))


def _address_of(function: object) -> int:
    return ctypes.cast(function, ctypes.c_void_p).value


# A consumer may move the proxy out of the capsule, so callbacks find the native stream through
# private_data, which travels with the move, never through the address they are called with.
_streams: dict[int, _Target] = {}
_proxies: dict[int, _native.ArrowArrayStream] = {}


def _target(stream: int) -> _Target:
    return _streams[_native.ArrowArrayStream.from_address(stream).private_data]


@_ReadFn
def _get_schema(stream: int, out: int) -> int:
    target = _target(stream)
    return target.get_schema(target.address, out)


@_ReadFn
def _get_next(stream: int, out: int) -> int:
    target = _target(stream)
    return target.get_next(target.address, out)


@_ErrorFn
def _get_last_error(stream: int) -> int | None:
    target = _target(stream)
    return target.get_last_error(target.address)


def _close(proxy: _native.ArrowArrayStream) -> None:
    _release(_streams.pop(proxy.private_data).stream)
    proxy.release = _native._ArrowStreamRelease()


@_ReleaseFn
def _release_proxy(stream: int) -> None:
    _close(_native.ArrowArrayStream.from_address(stream))


def _release(stream: _native.ArrowArrayStream) -> None:
    if stream.release:
        stream.release(ctypes.byref(stream))


def _proxy(stream: _native.ArrowArrayStream) -> _native.ArrowArrayStream:
    target = _Target(stream)
    _streams[id(target)] = target
    proxy = _native.ArrowArrayStream()
    proxy.get_schema = ctypes.cast(_get_schema, _native._ArrowStreamGetSchema)
    proxy.get_next = ctypes.cast(_get_next, _native._ArrowStreamGetNext)
    proxy.get_last_error = ctypes.cast(_get_last_error, _native._ArrowStreamGetLastError)
    proxy.release = ctypes.cast(_release_proxy, _native._ArrowStreamRelease)
    proxy.private_data = id(target)
    return proxy


def _destroy_capsule(capsule: int) -> None:
    proxy = _proxies.pop(_capsule_pointer(capsule, _CAPSULE_NAME))
    if proxy.release:
        _close(proxy)


_destroy_capsule_callback = ctypes.CFUNCTYPE(None, ctypes.c_void_p)(_destroy_capsule)


class ArrowStream:
    """A single-use stream of Arrow record batches implementing `__arrow_c_stream__`.

    Hand it to any consumer of the Arrow PyCapsule interface — `polars.DataFrame(stream)`,
    `polars.from_arrow(stream)`, `pyarrow.RecordBatchReader.from_stream(stream)`. Once a consumer
    takes it, the stream is spent; dropping it unconsumed releases the native read.
    """

    def __init__(self, stream: _native.ArrowArrayStream) -> None:
        self._stream: _native.ArrowArrayStream | None = stream

    def __arrow_c_stream__(self, requested_schema: object = None) -> object:
        if self._stream is None:
            raise ExcelReaderError("this Arrow stream was already consumed")
        stream, self._stream = self._stream, None
        proxy = _proxy(stream)
        address = ctypes.addressof(proxy)
        _proxies[address] = proxy
        return _capsule_new(address, _CAPSULE_NAME, _destroy_capsule_callback)

    def __del__(self) -> None:
        if getattr(self, "_stream", None) is not None:
            _release(self._stream)
