#!/usr/bin/env python3
"""Adds every symbol RS0016 flags as missing to the right PublicAPI.Unshipped.txt.

`dotnet format analyzers --diagnostics RS0016` does NOT do this: that code fix edits an
AdditionalFile (the .txt), not a .cs document, and dotnet-format's CLI fixer only applies
fixes that land in source documents. So this replicates it: build with the warning-as-error
gate relaxed, parse RS0016's message (locale-agnostic: it always quotes the symbol right
after the diagnostic id), and append to the matching TFM's Unshipped.txt.

Run this locally whenever `dotnet build` fails on RS0016 for a member you meant to add.
"""

import glob
import os
import re
import subprocess
import sys

HEADER = "#nullable enable"
RS0016_RE = re.compile(r"RS0016[^']*'([^']+)'.*TargetFramework=([\w.]+)\]")


def discover_tracked_projects():
    """Maps each PublicAPI-tracked project dir to its set of tracked TFMs."""
    projects = {}
    for unshipped_path in glob.glob("src/*/PublicAPI/*/PublicAPI.Unshipped.txt"):
        parts = unshipped_path.replace("\\", "/").split("/")
        project_dir, tfm = parts[0] + "/" + parts[1], parts[3]
        projects.setdefault(project_dir, set()).add(tfm)
    return projects


def read_entries(path):
    if not os.path.exists(path):
        return set()
    with open(path, encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip() and line.strip() != HEADER}


def write_entries(path, entries):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(HEADER + "\n")
        for entry in sorted(entries):
            f.write(entry + "\n")


def main():
    projects = discover_tracked_projects()
    if not projects:
        print("No src/*/PublicAPI/*/PublicAPI.Unshipped.txt files found — nothing to track.")
        return 1

    added_any = False
    for project_dir, tfms in projects.items():
        csproj = glob.glob(f"{project_dir}/*.csproj")[0]
        result = subprocess.run(
            [
                "dotnet", "build", csproj, "--configuration", "Release", "--nologo",
                "-p:TreatWarningsAsErrors=false", "-p:CodeAnalysisTreatWarningsAsErrors=false",
            ],
            capture_output=True, text=True, check=False,
        )
        missing = {tfm: set() for tfm in tfms}
        for line in result.stdout.splitlines():
            m = RS0016_RE.search(line)
            if m and m.group(2) in missing:
                missing[m.group(2)].add(m.group(1))

        for tfm, symbols in missing.items():
            if not symbols:
                continue
            path = f"{project_dir}/PublicAPI/{tfm}/PublicAPI.Unshipped.txt"
            current = read_entries(path)
            write_entries(path, current | symbols)
            print(f"{path}: added {len(symbols)} entries.")
            added_any = True

    if not added_any:
        print("Nothing missing — build already passes RS0016 clean.")
    return 0 if added_any else 1


if __name__ == "__main__":
    sys.exit(main())
