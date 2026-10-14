"""Forces a platform-specific wheel tag: the wheel bundles a native library per OS
(see _lib/), so it isn't pure Python even though every .py file in it is. The package
has no CPython C-extension (it's ctypes calling a NativeAOT shared library), so it's
ABI/interpreter-independent - only the platform component of the tag should be specific.

The platform component comes from the bundled library, not from the build machine: a
runner's own tag (macosx_26_0, manylinux_2_39) shuts out every older system the library
actually loads on."""

import platform
import re
import struct
from pathlib import Path

from hatchling.builders.hooks.plugin.interface import BuildHookInterface
from packaging.tags import platform_tags

LIB_DIR = Path(__file__).parent / "src" / "excelreader" / "_lib"
LC_BUILD_VERSION = 0x32


def glibc_floor(library: bytes) -> str:
    """The newest GLIBC_2.N symbol version the library needs, as manylinux's `2_N`."""
    minors = [int(minor) for minor in re.findall(rb"GLIBC_2\.(\d+)", library)]
    return f"2_{max(minors, default=17)}"


def macos_floor(library: bytes) -> str:
    """The minimum macOS a thin 64-bit Mach-O declares in LC_BUILD_VERSION, as `MAJOR_0`."""
    commands = struct.unpack_from("<I", library, 16)[0]
    offset = 32
    for _ in range(commands):
        command, size = struct.unpack_from("<II", library, offset)
        if command == LC_BUILD_VERSION:
            minos = struct.unpack_from("<I", library, offset + 12)[0]
            return f"{minos >> 16}_0"
        offset += size
    raise ValueError("the native library has no LC_BUILD_VERSION load command")


def platform_tag() -> str:
    machine = platform.machine().lower()
    system = platform.system()
    if system == "Linux":
        library = (LIB_DIR / "ExcelReader.Native.so").read_bytes()
        return f"manylinux_{glibc_floor(library)}_{machine}"
    if system == "Darwin":
        library = (LIB_DIR / "ExcelReader.Native.dylib").read_bytes()
        return f"macosx_{macos_floor(library)}_{machine}"
    return next(iter(platform_tags()))


class NativeLibraryBuildHook(BuildHookInterface):
    def initialize(self, version, build_data):
        build_data["pure_python"] = False
        build_data["infer_tag"] = True
        tag = next(iter(platform_tags())) if version == "editable" else platform_tag()
        build_data["tag"] = f"py3-none-{tag}"
