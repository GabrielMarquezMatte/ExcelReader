#!/usr/bin/env bash
# Fails when a benchmark is not matched by any group filter in benchmark-groups.json.
#
# The workflows build their job matrix from that file, so a benchmark missing from it is never run
# and never published - silently, which is how ChunkedParseBenchmark and DataReaderBenchmark went a
# whole release cycle without a single CI measurement.
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
groups_file="$here/benchmark-groups.json"
project="$here/ExcelReader.Benchmarks.csproj"

# Both lists land in variables rather than a process substitution: under `set -e` a pipeline that
# dies inside <(...) leaves the loop reading an empty list, which would pass this check (or, worse,
# fail it for every benchmark at once) instead of reporting the real failure.
#
# Both are stripped of carriage returns: on Windows the shell and `dotnet` hand back CRLF, and a
# trailing \r inside a pattern silently fails every regex match but the last one.
filters="$(python3 -c '
import json, sys
with open(sys.argv[1], encoding="utf-8") as f:
    for group in json.load(f):
        for pattern in group["filter"].split():
            print(pattern)
' "$groups_file" | tr -d '\r')"

ids="$(dotnet run --project "$project" -c Release --no-build -- --list flat |
    tr -d '\r' | grep '^ExcelReader\.Benchmarks\.')"

regexes=()
while IFS= read -r filter; do
    escaped=${filter//./\\.}
    regexes+=("^${escaped//\*/.*}$")
done <<<"$filters"

missing=()
while IFS= read -r id; do
    matched=0
    for regex in "${regexes[@]}"; do
        if [[ $id =~ $regex ]]; then
            matched=1
            break
        fi
    done
    ((matched)) || missing+=("$id")
done <<<"$ids"

if ((${#missing[@]})); then
    echo "Benchmarks matched by no group in benchmark-groups.json:" >&2
    printf '  %s\n' "${missing[@]}" >&2
    echo "Add them to a group, or add a new group." >&2
    exit 1
fi

echo "Every benchmark is covered by a group."
