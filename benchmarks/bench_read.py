"""Read benchmark: pyrust_xlsxwriter.open_workbook vs openpyxl (read-only).

Writes a 200,000 x 10 mixed-type workbook with px.Workbook, then times

* ``px.open_workbook(path).worksheet(0).values()``
* ``openpyxl.load_workbook(path, read_only=True, data_only=True)`` +
  ``iter_rows(values_only=True)`` (``data_only`` so formula cells give their
  cached value, as ``values()`` does)

and checks both readers return the same values. Build the extension in
release mode first (``maturin develop --release``) for meaningful numbers.

Usage: python benchmarks/bench_read.py [rows] [--keep PATH]
"""

from __future__ import annotations

import argparse
import datetime as dt
import tempfile
import time
from collections.abc import Sequence
from pathlib import Path

import openpyxl

import pyrust_xlsxwriter as px

COLS = 10


def row_values(i: int) -> list[int | float | str | bool | dt.date | None]:
    return [
        i,  # int
        i * 1.5,  # float
        f"name {i % 1000}",  # repeated (shared) string
        f"unique text {i}",  # unique string
        i % 3 == 0,  # bool
        dt.datetime(2020, 1, 1) + dt.timedelta(minutes=i),  # datetime
        dt.date(2000, 1, 1) + dt.timedelta(days=i % 9000),  # date
        -i / 7,  # float
        None if i % 5 == 0 else i % 97,  # sparse ints
        f"=A{i + 2}*2",  # formula (cached result 0)
    ]


def write(path: Path, rows: int) -> float:
    start = time.perf_counter()
    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_row(0, 0, [f"col{c}" for c in range(COLS)])
    ws.write_rows(1, 0, (row_values(i) for i in range(rows)))
    wb.save(path)
    return time.perf_counter() - start


def time_px(path: Path) -> tuple[float, Sequence[Sequence[object]]]:
    start = time.perf_counter()
    values = px.open_workbook(path).worksheet(0).values()
    return time.perf_counter() - start, values


def time_openpyxl(path: Path) -> tuple[float, Sequence[Sequence[object]]]:
    start = time.perf_counter()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    return time.perf_counter() - start, rows


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("rows", nargs="?", type=int, default=200_000)
    parser.add_argument("--keep", type=Path, help="write the workbook here and keep it")
    args = parser.parse_args()

    with tempfile.TemporaryDirectory() as tmp:
        path = args.keep or Path(tmp) / "bench_read.xlsx"
        elapsed = write(path, args.rows)
        size = path.stat().st_size / 1e6
        print(f"wrote {args.rows:,} x {COLS} cells ({size:.1f} MB) in {elapsed:.2f}s")

        t_px, ours = time_px(path)
        t_op, theirs = time_openpyxl(path)

        same = len(ours) == len(theirs) and all(list(a) == list(b) for a, b in zip(ours, theirs, strict=False))
        print(f"pyrust_xlsxwriter open_workbook().worksheet(0).values(): {t_px:8.2f}s")
        print(f"openpyxl load_workbook(read_only=True).iter_rows():      {t_op:8.2f}s")
        print(f"speed-up: {t_op / t_px:.1f}x   identical values: {same}")


if __name__ == "__main__":
    main()
