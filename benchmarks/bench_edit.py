"""Edit benchmark: load an existing workbook, change cells, save.

Creates a ROWS x 10 mixed-type workbook with the writer, then for each
engine (in a fresh subprocess, so peak memory is per engine):

1. load the file,
2. overwrite EDITS cells spread across the sheet and append 100 rows,
3. save to a new file.

    python benchmarks/bench_edit.py --rows 100000 --edits 1000
"""

from __future__ import annotations

import argparse
import json
import resource
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def make_source(path: Path, rows: int) -> None:
    import datetime as dt

    import pyrust_xlsxwriter as px

    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    bold = px.Format(bold=True)
    ws.write_row(0, 0, [f"col{c}" for c in range(10)], bold)
    base = dt.date(2024, 1, 1)
    ws.write_rows(
        1,
        0,
        (
            [r, r * 1.5, f"name {r % 1000}", r % 2 == 0, base, r % 7, f"x{r}", r / 3, "const", r * 2]
            for r in range(rows)
        ),
    )
    wb.save(path)


def run_engine(engine: str, src: Path, dst: Path, edits: int, rows: int) -> dict[str, float]:
    start = time.perf_counter()
    step = max(rows // edits, 1)
    if engine == "pyrust":
        import pyrust_xlsxwriter as px

        book = px.load_workbook(src)
        ws = book["Data"]
        t_load = time.perf_counter() - start
        for i in range(edits):
            ws.write(1 + i * step, 2, f"edited {i}")
        ws.write_rows(rows + 1, 0, [[i, "appended"] for i in range(100)])
        t_edit = time.perf_counter() - start - t_load
        book.save(dst)
    else:
        import openpyxl

        book = openpyxl.load_workbook(src)
        ws = book["Data"]
        t_load = time.perf_counter() - start
        for i in range(edits):
            ws.cell(row=2 + i * step, column=3, value=f"edited {i}")
        for i in range(100):
            ws.append([i, "appended"])
        t_edit = time.perf_counter() - start - t_load
        book.save(dst)
    total = time.perf_counter() - start
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak_mb = peak / (1024 * 1024) if sys.platform == "darwin" else peak / 1024
    return {"load": t_load, "edit": t_edit, "save": total - t_load - t_edit, "total": total, "peak_mb": peak_mb}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=100_000)
    parser.add_argument("--edits", type=int, default=1_000)
    parser.add_argument("--engine", help=argparse.SUPPRESS)
    parser.add_argument("--src", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--dst", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.engine:  # child process
        print(json.dumps(run_engine(args.engine, args.src, args.dst, args.edits, args.rows)))
        return

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source.xlsx"
        make_source(src, args.rows)
        size_mb = src.stat().st_size / 1e6
        print(
            f"source: {args.rows:,} rows x 10 cols ({size_mb:.1f} MB); editing {args.edits:,} cells + appending 100 rows"
        )
        print(f"{'engine':10s} {'load':>8s} {'edit':>8s} {'save':>8s} {'total':>8s} {'peak MB':>9s}")
        results = {}
        for engine in ("pyrust", "openpyxl"):
            dst = Path(tmp) / f"{engine}.xlsx"
            out = subprocess.run(
                [
                    sys.executable,
                    __file__,
                    "--engine",
                    engine,
                    "--src",
                    str(src),
                    "--dst",
                    str(dst),
                    "--edits",
                    str(args.edits),
                    "--rows",
                    str(args.rows),
                ],
                capture_output=True,
                text=True,
                check=True,
            )
            r = results[engine] = json.loads(out.stdout)
            print(
                f"{engine:10s} {r['load']:8.2f} {r['edit']:8.2f} {r['save']:8.2f} {r['total']:8.2f} {r['peak_mb']:9.0f}"
            )
        speedup = results["openpyxl"]["total"] / results["pyrust"]["total"]
        print(f"pyrust is {speedup:.1f}x faster end to end")


if __name__ == "__main__":
    main()
