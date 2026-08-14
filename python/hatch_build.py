"""Forces a platform-specific wheel tag: the wheel bundles a native library per OS
(see _lib/), so it isn't pure Python even though every .py file in it is."""

from hatchling.builders.hooks.plugin.interface import BuildHookInterface


class NativeLibraryBuildHook(BuildHookInterface):
    def initialize(self, version, build_data):
        build_data["pure_python"] = False
        build_data["infer_tag"] = True
