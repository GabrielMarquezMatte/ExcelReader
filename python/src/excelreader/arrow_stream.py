"""An Arrow C stream exposed through the Arrow PyCapsule interface, with no pyarrow dependency."""

from __future__ import annotations

import ctypes
from typing import TYPE_CHECKING

from excelreader import _native
from excelreader.types import ExcelReaderError

if TYPE_CHECKING:
    from excelreader.reader import Workbook

_CAPSULE_NAME = b"arrow_array_stream"

_capsule_new = ctypes.pythonapi.PyCapsule_New
_capsule_new.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_void_p]
_capsule_new.restype = ctypes.py_object

_capsule_pointer = ctypes.pythonapi.PyCapsule_GetPointer
_capsule_pointer.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
_capsule_pointer.restype = ctypes.c_void_p

_live: dict[int, tuple[_native.ArrowArrayStream, Workbook]] = {}


def _release(stream: _native.ArrowArrayStream) -> None:
    if stream.release:
        stream.release(ctypes.byref(stream))


def _destroy_capsule(capsule: int) -> None:
    address = _capsule_pointer(capsule, _CAPSULE_NAME)
    stream, _workbook = _live.pop(address)
    _release(stream)


_destroy_capsule_callback = ctypes.CFUNCTYPE(None, ctypes.c_void_p)(_destroy_capsule)


class ArrowStream:
    """A single-use stream of Arrow record batches implementing `__arrow_c_stream__`.

    Hand it to any consumer of the Arrow PyCapsule interface — `polars.DataFrame(stream)`,
    `polars.from_arrow(stream)`, `pyarrow.RecordBatchReader.from_stream(stream)`. Once a consumer
    takes it, the stream is spent; dropping it unconsumed releases the native read.
    """

    def __init__(self, stream: _native.ArrowArrayStream, workbook: Workbook) -> None:
        self._stream: _native.ArrowArrayStream | None = stream
        self._workbook = workbook

    def __arrow_c_stream__(self, requested_schema: object = None) -> object:
        if self._stream is None:
            raise ExcelReaderError("this Arrow stream was already consumed")
        stream, self._stream = self._stream, None
        address = ctypes.addressof(stream)
        _live[address] = (stream, self._workbook)
        return _capsule_new(address, _CAPSULE_NAME, _destroy_capsule_callback)

    def __del__(self) -> None:
        if getattr(self, "_stream", None) is not None:
            _release(self._stream)
