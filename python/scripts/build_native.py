"""Publish ExcelReader.Native for this machine and copy the binary into the Python package.

Run from anywhere:  python python/scripts/build_native.py
"""

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
CSPROJ = REPO_ROOT / "src" / "ExcelReader.Native" / "ExcelReader.Native.csproj"
PACKAGE_LIB_DIR = REPO_ROOT / "python" / "src" / "excelreader" / "_lib"

LIB_NAMES = {
    "Windows": "ExcelReader.Native.dll",
    "Linux": "ExcelReader.Native.so",
    "Darwin": "ExcelReader.Native.dylib",
}


def default_rid() -> str:
    system = platform.system()
    machine = platform.machine().lower()
    arch = "arm64" if machine in {"arm64", "aarch64"} else "x64"
    if system == "Windows":
        return f"win-{arch}"
    if system == "Darwin":
        return f"osx-{arch}"
    return f"linux-{arch}"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rid", default=default_rid(), help="dotnet runtime identifier")
    parser.add_argument("--framework", default="net10.0")
    args = parser.parse_args()

    subprocess.run(
        ["dotnet", "publish", str(CSPROJ), "-c", "Release", "-f", args.framework, "-r", args.rid],
        check=True,
    )

    publish_dir = CSPROJ.parent / "bin" / "Release" / args.framework / args.rid / "publish"
    name = LIB_NAMES[platform.system()]
    source = publish_dir / name
    if not source.exists():
        print(f"error: expected native library at {source}", file=sys.stderr)
        return 1

    PACKAGE_LIB_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, PACKAGE_LIB_DIR / name)
    print(f"copied {source} -> {PACKAGE_LIB_DIR / name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
