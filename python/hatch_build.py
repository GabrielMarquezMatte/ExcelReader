"""Forces a platform-specific wheel tag: the wheel bundles a native library per OS
(see _lib/), so it isn't pure Python even though every .py file in it is. The package
has no CPython C-extension (it's ctypes calling a NativeAOT shared library), so it's
ABI/interpreter-independent - only the platform component of the tag should be specific."""

import platform

from hatchling.builders.hooks.plugin.interface import BuildHookInterface
from packaging.tags import platform_tags


def _platform_tag() -> str:
    """Picks a platform tag PyPI actually accepts.

    packaging.tags.platform_tags() lists manylinux/musllinux-compatible tags first
    and the bare `linux_<machine>` tag last - but on a runner where it can't confirm
    manylinux compliance (no _manylinux metadata), the bare tag is all that's offered,
    and PyPI rejects raw `linux_*`/`freebsd_*`/etc. wheel uploads outright. Skip that
    bare tag and fall back to manylinux2014 (glibc 2.17, released 2012), a floor every
    current Linux distro satisfies - there's nothing to detect here since this wheel
    never links against glibc at build time (it dlopen()s a NativeAOT .so at runtime).
    """
    machine = platform.machine().lower()
    bare_linux = f"linux_{machine}"
    for tag in platform_tags():
        if tag != bare_linux:
            return tag
    return f"manylinux2014_{machine}"


class NativeLibraryBuildHook(BuildHookInterface):
    def initialize(self, version, build_data):
        build_data["pure_python"] = False
        build_data["infer_tag"] = True
        build_data["tag"] = f"py3-none-{_platform_tag()}"
