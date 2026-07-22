#!/usr/bin/env python3
"""Moves every entry out of each PublicAPI.Unshipped.txt into its PublicAPI.Shipped.txt.

Run after a release ships: everything that was "unshipped" at release time is now part of
the public API contract. Operates on every TFM folder under src/*/PublicAPI/*/.
Exits with status 1 if nothing needed promoting, so the caller can skip the rest of the job.
"""

import glob
import sys

HEADER = "#nullable enable"


def read_entries(path):
    with open(path, encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip() and line.strip() != HEADER]


def write_entries(path, entries):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(HEADER + "\n")
        for entry in sorted(entries):
            f.write(entry + "\n")


def main():
    unshipped_files = sorted(glob.glob("src/*/PublicAPI/*/PublicAPI.Unshipped.txt"))
    if not unshipped_files:
        print("No PublicAPI/*/PublicAPI.Unshipped.txt files found.")
        return 1

    promoted_any = False
    for unshipped_path in unshipped_files:
        shipped_path = unshipped_path.replace("Unshipped.txt", "Shipped.txt")
        new_entries = read_entries(unshipped_path)
        if not new_entries:
            print(f"{unshipped_path}: nothing to promote.")
            continue

        shipped_entries = set(read_entries(shipped_path)) if glob.glob(shipped_path) else set()
        shipped_entries.update(new_entries)
        write_entries(shipped_path, shipped_entries)
        write_entries(unshipped_path, [])
        print(f"{unshipped_path}: promoted {len(new_entries)} entries to {shipped_path}.")
        promoted_any = True

    return 0 if promoted_any else 1


if __name__ == "__main__":
    sys.exit(main())
