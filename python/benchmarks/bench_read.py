"""Benchmark ExcelReader's Python bindings against polars.

Usage:  python python/benchmarks/bench_read.py [path] [--n N]

`path` defaults to tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb, resolved relative to
the repository root. Reports min and median over N iterations (default 10) for every available
read path. A variant that reads zero rows or zero cells is a harness bug, not a result — this
script asserts against that instead of silently publishing a zero.

NOTE: read_all()/rows() build one Cell tuple per cell; polars.read_excel returns a typed columnar
DataFrame with type inference. That is not matched work (see STYLEGUIDE.md "Tests and
Benchmarks") — this script labels the comparison, it does not pretend it is apples-to-apples.
The parse_typed/to_arrow variants are the closest thing to matched work against polars: they
do real native-side type conversion into columns, against a fixed schema rather than inferring one.
"""

from __future__ import annotations

import argparse
import statistics
import sys
import timeit
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PATH = REPO_ROOT / "tests" / "ExcelReader.Benchmarks" / "Data" / "65K_Records_Data.xlsb"

sys.path.insert(0, str(REPO_ROOT / "python" / "src"))

import excelreader


def _report(name: str, times: list[float], n: int) -> None:
    print(f"{name}: min={min(times):.4f}s median={statistics.median(times):.4f}s (n={n})")


def bench_read_all(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        all_rows = workbook.read_all()
    row_count = len(all_rows)
    cell_count = sum(len(row) for row in all_rows)
    return row_count, cell_count


def bench_rows(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        row_count = 0
        cell_count = 0
        for row in workbook.rows():
            row_count += 1
            cell_count += len(row)
    return row_count, cell_count


def bench_read_all_columnar(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        sheet = workbook.read_all_columnar()
    return len(sheet.row_offsets) - 1, len(sheet.columns)


# The 14 columns of the default fixture, in file order. parse_typed()/to_arrow() take a schema;
# this is the schema that matches 65K_Records_Data.xlsb. Passing --path a different file makes both
# typed variants meaningless, so they are skipped there rather than silently benchmarking a failure.
_T = excelreader.ColumnType
_FIXTURE_SCHEMA = [
    excelreader.ColumnSpec(type_, name=name)
    for name, type_ in [
        ("Region", _T.STRING),
        ("Country", _T.STRING),
        ("Item Type", _T.STRING),
        ("Sales Channel", _T.STRING),
        ("Order Priority", _T.STRING),
        ("Order Date", _T.DATE),
        ("Order ID", _T.I64),
        ("Ship Date", _T.DATE),
        ("Units Sold", _T.I64),
        ("Unit Price", _T.F64),
        ("Unit Cost", _T.F64),
        ("Total Revenue", _T.F64),
        ("Total Cost", _T.F64),
        ("Total Profit", _T.F64),
    ]
]


def bench_parse_typed(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        table = workbook.parse_typed(_FIXTURE_SCHEMA)
    return table.row_count, table.row_count * len(table.columns)


def bench_to_arrow(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        array = workbook.to_arrow(_FIXTURE_SCHEMA)
    return len(array), len(array) * array.type.num_fields


def bench_polars(path: Path) -> tuple[int, int]:
    import polars as pl

    df = pl.read_excel(path)
    return df.shape[0], df.shape[0] * df.shape[1]


def _time_and_assert(label: str, func, path: Path, n: int) -> None:
    row_count = 0
    cell_count = 0

    def run() -> None:
        nonlocal row_count, cell_count
        row_count, cell_count = func(path)

    times = timeit.repeat(run, repeat=n, number=1)
    if row_count == 0 or cell_count == 0:
        raise AssertionError(f"{label} read zero work (rows={row_count}, cells={cell_count}) — harness is broken")
    print(f"  rows={row_count} cells={cell_count}")
    _report(label, times, n)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, default=DEFAULT_PATH)
    parser.add_argument("--n", type=int, default=10)
    args = parser.parse_args()

    if not args.path.exists():
        print(f"error: fixture not found: {args.path}", file=sys.stderr)
        return 1

    print(f"file: {args.path}")
    print()

    print("excelreader.read_all():")
    _time_and_assert("  read_all", bench_read_all, args.path, args.n)
    print()

    print("excelreader.rows() (per-row iteration):")
    _time_and_assert("  rows", bench_rows, args.path, args.n)
    print()

    print("excelreader.read_all_columnar() [no per-cell Cell/str objects — the fast path]:")
    _time_and_assert("  read_all_columnar", bench_read_all_columnar, args.path, args.n)
    print()

    if args.path.resolve() == DEFAULT_PATH.resolve():
        print("parse_typed() [schema-driven typed columns, native-side conversion]:")
        _time_and_assert("  parse_typed", bench_parse_typed, args.path, args.n)
        print()

        try:
            import pyarrow  # noqa: F401
        except ImportError:
            print("pyarrow not installed — skipping the to_arrow() variant")
            print()
        else:
            print("to_arrow() [same parse, handed to pyarrow zero-copy]:")
            _time_and_assert("  to_arrow", bench_to_arrow, args.path, args.n)
            print()
    else:
        print("typed/Arrow variants skipped — their schema only matches the default fixture")
        print()

    try:
        import polars  # noqa: F401
    except ImportError:
        print("polars not installed — skipping polars.read_excel comparison")
        return 0

    print("polars.read_excel() [NOTE: typed columnar DataFrame with type inference — not matched work]:")
    _time_and_assert("  polars", bench_polars, args.path, args.n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
