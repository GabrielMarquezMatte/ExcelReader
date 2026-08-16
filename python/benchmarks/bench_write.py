"""Benchmark ExcelReader's Python write bindings against pandas/polars.

Usage:  python python/benchmarks/bench_write.py [path] [--n N]

`path` is the READ source used to build the data this script writes back out — it defaults to the
same 65K_Records_Data.xlsb fixture bench_read.py uses, resolved relative to the repository root.
The fixture is parsed ONCE via parse_typed()/to_pandas()/to_polars() before any timing starts;
every benchmark below writes that same in-memory table repeatedly, so what's measured is the write
path alone, not read+write. Reports min and median over N iterations (default 10) for every
available write path, into a fresh temp file per run. A write that produces an empty or missing
file is a harness bug, not a result — this script asserts against that instead of publishing a zero.

NOTE: pandas.DataFrame.to_excel() and polars.DataFrame.write_excel() do their own dtype-to-cell-
format translation and neither one can write XLSB or XLS at all, so only the XLSX pandas/polars
comparisons below are apples-to-apples (see STYLEGUIDE.md "Tests and Benchmarks"); everything else
is this library writing a format the competitor cannot produce, and is labeled as such rather than
presented as a win.
"""

from __future__ import annotations

import argparse
import statistics
import sys
import tempfile
import timeit
from importlib.util import find_spec
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PATH = REPO_ROOT / "tests" / "ExcelReader.Benchmarks" / "Data" / "65K_Records_Data.xlsb"

sys.path.insert(0, str(REPO_ROOT / "python" / "src"))

import excelreader

# The 14 columns of the default fixture, in file order — same schema bench_read.py uses, so both
# scripts describe the same data. Passing --path a different file makes this schema meaningless, so
# the whole script is skipped there rather than silently benchmarking a failure.
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
_FIXTURE_TYPES = [spec.type for spec in _FIXTURE_SCHEMA]


def _report(name: str, times: list[float], n: int) -> None:
    print(f"{name}: min={min(times):.4f}s median={statistics.median(times):.4f}s (n={n})")


def _time_and_assert(label: str, func, out_path: Path, n: int) -> None:
    def run() -> None:
        func(out_path)

    times = timeit.repeat(run, repeat=n, number=1)
    size = out_path.stat().st_size if out_path.exists() else 0
    if size == 0:
        raise AssertionError(f"{label} produced an empty or missing file at {out_path} — harness is broken")
    print(f"  bytes={size}")
    _report(label, times, n)


def benchmark_pandas(args: argparse.Namespace, tmp: Path):
    pd = find_spec("pandas")
    if pd is None:
        print("pandas not installed — skipping write_pandas/to_excel comparison")
        print()
        return
    with excelreader.open_workbook(args.path, format="xlsb") as workbook:
        pandas_df = workbook.to_pandas(_FIXTURE_SCHEMA)

    print("write_pandas() [pandas.DataFrame -> Arrow -> pylist -> native columns -> xlsx]:")
    _time_and_assert(
        "  write_pandas",
        lambda p: excelreader.write_pandas(p, pandas_df),
        tmp / "out_pandas.xlsx",
        args.n,
    )
    print()

    openpyxl = find_spec("openpyxl")
    if openpyxl is None:
        print("openpyxl not installed — skipping write_pandas/write_workbook comparison")
        print()
        return
    print("pandas.DataFrame.to_excel() [same DataFrame, matched work — both write xlsx]:")
    _time_and_assert(
        "  to_excel",
        lambda p: pandas_df.to_excel(p, index=False),
        tmp / "out_pandas_native.xlsx",
        args.n,
    )
    print()

def benchmark_polars(args: argparse.Namespace, tmp: Path):
    pl = find_spec("polars")
    if pl is None:
        print("polars not installed — skipping write_polars/write_excel comparison")
        print()
        return
    with excelreader.open_workbook(args.path, format="xlsb") as workbook:
        schema = workbook.infer_schema(sample_size=10)
        polars_df = workbook.to_polars(schema)

    print("write_polars() [polars.DataFrame -> Arrow -> pylist -> native columns -> xlsx]:")
    _time_and_assert(
        "  write_polars",
        lambda p: excelreader.write_polars(p, polars_df),
        tmp / "out_polars.xlsx",
        args.n,
    )
    print()
    xlsxwriter = find_spec("xlsxwriter")
    if xlsxwriter is None:
        print("xlsxwriter not installed — skipping polars.DataFrame.write_excel() comparison")
        print()
        return
    print("polars.DataFrame.write_excel() [same DataFrame, matched work — both write xlsx]:")
    _time_and_assert(
        "  write_excel",
        lambda p: polars_df.write_excel(p),
        tmp / "out_polars_native.xlsx",
        args.n,
    )

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, default=DEFAULT_PATH)
    parser.add_argument("--n", type=int, default=10)
    args = parser.parse_args()

    if not args.path.exists():
        print(f"error: fixture not found: {args.path}", file=sys.stderr)
        return 1

    if args.path.resolve() != DEFAULT_PATH.resolve():
        print("this benchmark's schema only matches the default fixture; pass no path to run it")
        return 1

    print(f"source file: {args.path}")

    with excelreader.open_workbook(args.path, format="xlsb") as workbook:
        table = workbook.parse_typed(_FIXTURE_SCHEMA)
    print(f"parsed {table.row_count} rows x {len(table.columns)} columns once, before any timing")
    print()

    with tempfile.TemporaryDirectory(prefix="excelreader-bench-write-") as tmp_dir:
        tmp = Path(tmp_dir)

        print("write_workbook() [xlsx, native typed columns -> file]:")
        _time_and_assert(
            "  write_workbook(xlsx)",
            lambda p: excelreader.write_workbook(p, table, _FIXTURE_TYPES),
            tmp / "out.xlsx",
            args.n,
        )
        print()

        print("write_workbook() [xlsb]:")
        _time_and_assert(
            "  write_workbook(xlsb)",
            lambda p: excelreader.write_workbook(p, table, _FIXTURE_TYPES),
            tmp / "out.xlsb",
            args.n,
        )
        print()

        print("write_workbook() [xls]:")
        _time_and_assert(
            "  write_workbook(xls)",
            lambda p: excelreader.write_workbook(p, table, _FIXTURE_TYPES),
            tmp / "out.xls",
            args.n,
        )
        print()

        print("write_workbook() [csv, no cell formatting to compute]:")
        _time_and_assert(
            "  write_workbook(csv)",
            lambda p: excelreader.write_workbook(p, table, _FIXTURE_TYPES),
            tmp / "out.csv",
            args.n,
        )
        print()

        try:
            import pyarrow  # noqa: F401
        except ImportError:
            print("pyarrow not installed — skipping write_pandas/write_polars/pandas/polars comparisons")
            return 0

        benchmark_pandas(args, tmp)
        benchmark_polars(args, tmp)
    return 0

if __name__ == "__main__":
    sys.exit(main())
