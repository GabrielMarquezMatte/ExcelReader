"""Fails unless every wheel in the given directory installs on this machine.

Run after building a wheel:  python python/scripts/check_wheel_tag.py dist
"""

import sys
from pathlib import Path

from packaging.tags import sys_tags
from packaging.utils import parse_wheel_filename


def main() -> int:
    supported = set(sys_tags())
    wheels = sorted(Path(sys.argv[1]).glob("*.whl"))
    if not wheels:
        print(f"error: no wheel in {sys.argv[1]}", file=sys.stderr)
        return 1
    for wheel in wheels:
        tags = parse_wheel_filename(wheel.name)[3]
        if tags.isdisjoint(supported):
            print(f"error: {wheel.name} does not install on this machine", file=sys.stderr)
            return 1
        print(f"ok: {wheel.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
