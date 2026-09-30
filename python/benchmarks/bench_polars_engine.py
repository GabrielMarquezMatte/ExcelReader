"""excelreader as a polars Excel engine: excelreader vs polars.read_excel(engine="calamine").

Usage:  python python/benchmarks/bench_polars_engine.py [--rows N] [--n N] [--data-dir DIR]

Needs polars, fastexcel, pyarrow, psutil, xlsxwriter, msoffcrypto-tool and excelreader-native. It
imports whichever excelreader is installed; put python/src on PYTHONPATH to measure the repo instead.

Generates one deterministic sheet (7 columns: two string, i64, f64, bool, date, datetime) and caches
it under --data-dir as:
  - plain.xlsx      written by xlsxwriter (polars.write_excel), so no reader grades its own output
  - plain.xlsb      written by excelreader, the only Python XLSB writer
  - encrypted.xlsx / encrypted.xlsb   the above, agile-encrypted (ECMA-376, Excel's defaults)

The XLSB writer must be one that encodes integers in RK's float form, as Excel does (excelreader after
5.1.0): calamine ignores a date style on the int form, so older-written XLSB dates read back as ints.
The data files are cached, so generate them once with the repo on PYTHONPATH, then time any version.

Every case reads into a polars.DataFrame with the same fixed schema, so both sides do the same type
conversion and neither infers. Before timing, every case's frame is checked equal to the generated one.
Encrypted files have no calamine path, so the baseline there is the usual workaround: msoffcrypto-tool
decrypting into memory, then polars.read_excel(engine="calamine") on the bytes — decryption is timed.

Each case runs N times in its own fresh subprocess. Peak RSS is the process peak minus its RSS after
imports, so it counts the read (native heap included), not the interpreter or the libraries.
"""

from __future__ import annotations

import argparse
import io
import json
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from importlib.metadata import version
from pathlib import Path

import polars as pl

PASSWORD = "benchmark"

POLARS_SCHEMA = {
    "category": pl.String,
    "name": pl.String,
    "qty": pl.Int64,
    "price": pl.Float64,
    "flag": pl.Boolean,
    "day": pl.Date,
    "stamp": pl.Datetime("us"),
}


def _excelreader_schema() -> list:
    from excelreader import ColumnSpec, ColumnType

    types = {
        pl.String: ColumnType.STRING,
        pl.Int64: ColumnType.I64,
        pl.Float64: ColumnType.F64,
        pl.Boolean: ColumnType.BOOL,
        pl.Date: ColumnType.DATE,
    }
    return [
        ColumnSpec(ColumnType.TIMESTAMP if isinstance(dtype, pl.Datetime) else types[dtype], name=name)
        for name, dtype in POLARS_SCHEMA.items()
    ]


def _frame(rows: int) -> pl.DataFrame:
    i = pl.int_range(rows, dtype=pl.Int64)
    return pl.select(
        category=pl.format("category-{}", i % 50),
        name=pl.format("customer-{}", (i * 7919) % 200_000),
        qty=(i * 31) % 10_000,
        price=((i * 17) % 100_000) / 100,
        flag=i % 3 == 0,
        day=pl.date(2020, 1, 1) + pl.duration(days=i % 3650),
        stamp=pl.datetime(2020, 1, 1, time_unit="us") + pl.duration(seconds=(i * 37) % (86_400 * 3650)),
    )


def _generate(data_dir: Path, rows: int) -> dict[str, Path]:
    from excelreader import encrypt_package, write_polars

    data_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "plain.xlsx": data_dir / "plain.xlsx",
        "plain.xlsb": data_dir / "plain.xlsb",
        "encrypted.xlsx": data_dir / "encrypted.xlsx",
        "encrypted.xlsb": data_dir / "encrypted.xlsb",
    }
    missing = [name for name, path in paths.items() if not path.exists()]
    if not missing:
        return paths

    df = _frame(rows)
    print(f"generating {', '.join(missing)} ({rows} rows) into {data_dir}, cached for later runs...", flush=True)
    if not paths["plain.xlsx"].exists():
        df.write_excel(paths["plain.xlsx"], autofit=False)
    if not paths["plain.xlsb"].exists():
        write_polars(paths["plain.xlsb"], df)
    for suffix in ("xlsx", "xlsb"):
        if not paths[f"encrypted.{suffix}"].exists():
            encrypt_package(paths[f"plain.{suffix}"], paths[f"encrypted.{suffix}"], PASSWORD)
    return paths


def read_excelreader(path: Path) -> pl.DataFrame:
    from excelreader import open_workbook

    password = PASSWORD if path.stem.startswith("encrypted") else None
    with open_workbook(path, password=password) as workbook:
        return workbook.to_polars(_excelreader_schema())


def read_calamine(path: Path | io.BytesIO) -> pl.DataFrame:
    return pl.read_excel(path, engine="calamine", schema_overrides=POLARS_SCHEMA)


def read_msoffcrypto_calamine(path: Path) -> pl.DataFrame:
    import msoffcrypto

    decrypted = io.BytesIO()
    with path.open("rb") as handle:
        office = msoffcrypto.OfficeFile(handle)
        office.load_key(password=PASSWORD)
        office.decrypt(decrypted)
    decrypted.seek(0)
    return read_calamine(decrypted)


CASES = {
    "excelreader": read_excelreader,
    "calamine": read_calamine,
    "msoffcrypto+calamine": read_msoffcrypto_calamine,
}

PLAN = {
    "plain.xlsx": ["excelreader", "calamine"],
    "plain.xlsb": ["excelreader", "calamine"],
    "encrypted.xlsx": ["excelreader", "msoffcrypto+calamine"],
    "encrypted.xlsb": ["excelreader", "msoffcrypto+calamine"],
}


def _normalize(df: pl.DataFrame) -> pl.DataFrame:
    return df.cast(POLARS_SCHEMA).with_columns(pl.col("stamp").dt.round("1s")).rechunk()


def _verify(paths: dict[str, Path], rows: int) -> None:
    from polars.testing import assert_frame_equal

    expected = _normalize(_frame(rows))
    for file, cases in PLAN.items():
        for case in cases:
            try:
                assert_frame_equal(_normalize(CASES[case](paths[file])), expected)
            except AssertionError as error:
                raise AssertionError(f"{case} on {file} does not match the generated frame:\n{error}") from None
    print("verified: every case returns the generated frame", flush=True)


def _rss() -> int:
    import psutil

    return psutil.Process().memory_info().rss


def _peak_rss() -> int:
    if sys.platform == "win32":
        import psutil

        return psutil.Process().memory_info().peak_wset
    import resource

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak if sys.platform == "darwin" else peak * 1024


def _worker(case: str, path: Path, n: int) -> None:
    import excelreader  # noqa: F401
    import fastexcel  # noqa: F401
    import msoffcrypto  # noqa: F401
    import pyarrow  # noqa: F401

    read = CASES[case]
    baseline = _rss()
    times = []
    rows = 0
    for _ in range(n):
        start = time.perf_counter()
        df = read(path)
        times.append(time.perf_counter() - start)
        rows = df.height
        del df
    print(json.dumps({"times": times, "rows": rows, "peak": _peak_rss() - baseline}))


def _run(case: str, path: Path, n: int) -> dict:
    proc = subprocess.run(
        [sys.executable, str(Path(__file__).resolve()), "--worker", case, str(path), "--n", str(n)],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise AssertionError(f"{case} on {path.name} failed:\n{proc.stdout}{proc.stderr}")
    result = json.loads(proc.stdout.strip().splitlines()[-1])
    if result["rows"] == 0:
        raise AssertionError(f"{case} on {path.name} read zero rows — harness is broken")
    return result


def _environment() -> str:
    import excelreader

    packages = ["polars", "fastexcel", "excelreader-native", "msoffcrypto-tool", "pyarrow"]
    lines = [
        f"- OS: {platform.platform()}",
        f"- CPU: {platform.processor() or platform.machine()}",
        f"- Python: {platform.python_version()}",
    ]
    lines += [f"- {package}: {version(package)}" for package in packages]
    lines.append(f"- excelreader imported from: {Path(excelreader.__file__).parent}")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=1_000_000)
    parser.add_argument("--n", type=int, default=5)
    parser.add_argument("--data-dir", type=Path, default=Path(tempfile.gettempdir()) / "excelreader-polars-bench")
    parser.add_argument("--skip-verify", action="store_true")
    parser.add_argument("--worker", nargs=2, metavar=("CASE", "PATH"), help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.worker is not None:
        _worker(args.worker[0], Path(args.worker[1]), args.n)
        return 0

    if args.rows > 1_048_575:
        parser.error("--rows cannot exceed Excel's 1,048,575 data rows")

    paths = _generate(args.data_dir / str(args.rows), args.rows)
    if not args.skip_verify:
        _verify(paths, args.rows)

    print()
    print(f"{args.rows:,} rows x {len(POLARS_SCHEMA)} columns, n={args.n} runs per case")
    print()
    print("| file | size | reader | median | min | rows/s | peak RSS |")
    print("|---|---:|---|---:|---:|---:|---:|")
    for file, cases in PLAN.items():
        path = paths[file]
        size = path.stat().st_size / 1024 / 1024
        for case in cases:
            result = _run(case, path, args.n)
            median = statistics.median(result["times"])
            print(
                f"| {file} | {size:.1f} MiB | {case} | {median:.3f} s | {min(result['times']):.3f} s "
                f"| {result['rows'] / median / 1e6:.2f} M | {result['peak'] / 1024 / 1024:.0f} MiB |",
                flush=True,
            )
    print()
    print(_environment())
    return 0


if __name__ == "__main__":
    sys.exit(main())
