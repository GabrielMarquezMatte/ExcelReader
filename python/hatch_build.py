"""Forces a platform-specific wheel tag: the wheel bundles a native library per OS
(see _lib/), so it isn't pure Python even though every .py file in it is. The package
has no CPython C-extension (it's ctypes calling a NativeAOT shared library), so it's
ABI/interpreter-independent - only the platform component of the tag should be specific."""

from hatchling.builders.hooks.plugin.interface import BuildHookInterface
from packaging.tags import sys_tags


class NativeLibraryBuildHook(BuildHookInterface):
    def initialize(self, version, build_data):
        build_data["pure_python"] = False
        build_data["infer_tag"] = True
        build_data["tag"] = f"py3-none-{next(iter(sys_tags())).platform}"
