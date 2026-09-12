"""Benchmark ExcelReader's Python bindings against polars.

Usage:  python python/benchmarks/bench_read.py [path] [--n N]

`path` defaults to tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb, resolved relative to
the repository root. Reports min and median over N iterations (default 10) for every available
read path. A variant that reads zero rows or zero cells is a harness bug, not a result — this
script asserts against that instead of silently publishing a zero.

NOTE: read_all()/rows() build one Cell tuple per cell; polars.read_excel returns a typed columnar
DataFrame with type inference. That is not matched work (see STYLEGUIDE.md "Tests and
Benchmarks") — this script labels the comparison, it does not pretend it is apples-to-apples.
The parse_typed/parse_arrow variants are the closest thing to matched work against polars: they
do real native-side type conversion into columns, against a fixed schema rather than inferring one.
"""

from __future__ import annotations

import argparse
import ctypes
import statistics
import subprocess
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


def bench_iter_parse_typed(path: Path) -> tuple[int, int]:
    """Batched counterpart to bench_parse_typed: one batch resident instead of the whole sheet."""
    rows = 0
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        for table in workbook.iter_parse_typed(_FIXTURE_SCHEMA, batch_size=10000):
            rows += table.row_count
    return rows, rows * len(_FIXTURE_SCHEMA)


def bench_record_batch_reader(path: Path) -> tuple[int, int]:
    rows = 0
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        for batch in workbook.to_record_batch_reader(_FIXTURE_SCHEMA, batch_size=10000):
            rows += batch.num_rows
    return rows, rows * len(_FIXTURE_SCHEMA)


# --- peak memory ------------------------------------------------------------------------------
#
# Every other leg here is wall-clock, which cannot see the one property batching exists for: bytes
# live at once. These two legs measure it.
#
# NOT pyarrow.total_allocated_bytes(): the batch buffers are allocated by the native reader and
# handed to pyarrow through the Arrow C Data Interface, so they never pass through Arrow's pool and
# it reports 0 for both paths. The process's private bytes do see them — confirmed by a control that
# retains every batch, which reads ABOVE the whole-sheet peak rather than at one batch.
#
# Both legs include the same native workbook parse, so the floor is common to them; it is the
# DIFFERENCE between the two peaks that is the exported table, bounded to one batch or not.


class _ProcessMemoryCountersEx(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_uint32),
        ("PageFaultCount", ctypes.c_uint32),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
        ("PrivateUsage", ctypes.c_size_t),
    ]


def _private_bytes() -> int:
    kernel32 = ctypes.WinDLL("kernel32")
    counters = _ProcessMemoryCountersEx()
    counters.cb = ctypes.sizeof(counters)
    kernel32.GetCurrentProcess.restype = ctypes.c_void_p
    if not kernel32.K32GetProcessMemoryInfo(
        ctypes.c_void_p(kernel32.GetCurrentProcess()), ctypes.byref(counters), counters.cb
    ):
        raise OSError("K32GetProcessMemoryInfo failed")
    return counters.PrivateUsage


def peak_leg_whole_sheet(path: Path) -> tuple[int, int]:
    """Whole-sheet Arrow read, sampled once while the full table is live. Returns (rows, peak)."""
    base = _private_bytes()
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        batch = workbook.to_record_batch(_FIXTURE_SCHEMA)
        peak = _private_bytes() - base
        rows = batch.num_rows
    return rows, peak


def peak_leg_streamed(path: Path) -> tuple[int, int]:
    """The same read streamed, sampled after every batch. Returns (rows, the maximum sample)."""
    base = _private_bytes()
    rows = 0
    peak = 0
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        for batch in workbook.to_record_batch_reader(_FIXTURE_SCHEMA, batch_size=10000):
            rows += batch.num_rows
            peak = max(peak, _private_bytes() - base)
            # Released before the next batch is pulled. Holding it across the pull would measure two
            # batches live and hide exactly the defect this leg exists to catch.
            del batch
    return rows, peak


_PEAK_LEGS = {"whole_sheet": peak_leg_whole_sheet, "streamed": peak_leg_streamed}


def _measure_peaks(path: Path) -> dict[str, int]:
    """Runs each peak leg in a FRESH subprocess and collects its number.

    In-process would read zero: the legs above allocate far more than these reads do (read_all alone
    builds ~900k Cell tuples), a heap is not handed back to the OS on free, and so by the time we got
    here the growth we are trying to see would already be absorbed by slack.
    """
    peaks: dict[str, int] = {}
    for name in _PEAK_LEGS:
        proc = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), str(path), "--peak-leg", name],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            raise AssertionError(f"peak leg {name} failed:\n{proc.stdout}{proc.stderr}")
        rows, peak = (int(field) for field in proc.stdout.split())
        if rows == 0:
            raise AssertionError(f"peak leg {name} read zero rows — harness is broken")
        print(f"  {name}: rows={rows} peak={peak} bytes ({peak / 1024 / 1024:.1f} MiB)")
        peaks[name] = peak
    return peaks


def bench_to_polars(path: Path) -> tuple[int, int]:
    with excelreader.open_workbook(path, format="xlsb") as workbook:
        schema = workbook.infer_schema(sample_size=10)
        table = workbook.to_polars(schema)
        return table.shape[0], table.shape[0] * table.shape[1]

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
    # Internal re-entry point for _measure_peaks(): it shells out to this same script so each peak
    # leg gets a fresh process (see _measure_peaks' docstring for why in-process reads zero).
    parser.add_argument("--peak-leg", choices=sorted(_PEAK_LEGS), help=argparse.SUPPRESS)
    args = parser.parse_args()

    if not args.path.exists():
        print(f"error: fixture not found: {args.path}", file=sys.stderr)
        return 1

    if args.peak_leg is not None:
        rows, peak = _PEAK_LEGS[args.peak_leg](args.path)
        print(rows, peak)
        return 0

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

        print("excelreader.iter_parse_typed() (batched):")
        _time_and_assert("  iter_parse_typed", bench_iter_parse_typed, args.path, args.n)
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

            print("excelreader.to_record_batch_reader() (batched):")
            _time_and_assert("  record_batch_reader", bench_record_batch_reader, args.path, args.n)
            print()

            if sys.platform == "win32":
                print("peak memory [to_record_batch whole sheet vs to_record_batch_reader(batch_size=10000)]:")
                peaks = _measure_peaks(args.path)
                saved = peaks["whole_sheet"] - peaks["streamed"]
                print(f"  streaming holds {saved} bytes ({saved / 1024 / 1024:.1f} MiB) less at the peak")
            else:
                print("peak memory leg skipped — its private-bytes probe is Windows-only")
            print()
    else:
        print("typed/Arrow variants skipped — their schema only matches the default fixture")
        print()

    try:
        import polars  # noqa: F401
    except ImportError:
        print("polars not installed — skipping polars.read_excel comparison")
        return 0

    print("excelreader.to_polars() [typed columnar DataFrame with schema inference]:")
    _time_and_assert("  to_polars", bench_to_polars, args.path, args.n)
    print()

    print("polars.read_excel() [NOTE: typed columnar DataFrame with type inference — not matched work]:")
    _time_and_assert("  polars", bench_polars, args.path, args.n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
