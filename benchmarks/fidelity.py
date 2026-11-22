"""Load-and-save fidelity across editing engines.

For each input file, an engine opens and saves it without changes. We then
report, per engine:

* lost parts   - parts that disappeared, counted per content type so that
                 renamed parts don't count (document properties and calcChain
                 are ignored: rewriting them is normal);
* invalid      - outputs that fail tests/ooxml_check.py's structural checks;
* errors       - files the engine could not open or save;
* time         - total load + save time.

Engines: ``pyrust`` (this project's editor), ``openpyxl``, and ``umya`` (a
umya-spreadsheet round-trip binary passed with --umya-bin).

    python benchmarks/fidelity.py tests/corpus/*.xls* --engines pyrust openpyxl
"""

from __future__ import annotations

import argparse
import collections
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))

from xml.etree import ElementTree as ET

from ooxml_check import check_package

NS_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
IGNORED_TYPES = ("core-properties", "extended-properties", "custom-properties", "calcChain")


def parts(path: Path) -> collections.Counter[str]:
    """Count the package's parts by content type."""
    with zipfile.ZipFile(path) as zf:
        ct = ET.fromstring(zf.read("[Content_Types].xml"))
        defaults = {d.get("Extension", "").lower(): d.get("ContentType", "") for d in ct.iter(f"{NS_CT}Default")}
        overrides = {o.get("PartName", "").lstrip("/"): o.get("ContentType", "") for o in ct.iter(f"{NS_CT}Override")}
        counts: collections.Counter[str] = collections.Counter()
        for name in zf.namelist():
            if name.endswith("/") or name == "[Content_Types].xml":
                continue
            kind = overrides.get(name) or defaults.get(name.rsplit(".", 1)[-1].lower(), "unknown")
            if not any(t in kind for t in IGNORED_TYPES):
                counts[kind.split("/")[-1].replace("vnd.openxmlformats-officedocument.", "")] += 1
        return counts


def run_pyrust(src: Path, dst: Path) -> None:
    import pyrust_xlsxwriter as px

    px.load_workbook(src).save(dst)


def run_openpyxl(src: Path, dst: Path) -> None:
    import openpyxl

    keep_vba = src.suffix.lower() == ".xlsm"
    openpyxl.load_workbook(src, keep_vba=keep_vba).save(dst)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--engines", nargs="+", default=["pyrust", "openpyxl"])
    parser.add_argument("--umya-bin", type=Path, help="path to the umya round-trip binary")
    parser.add_argument("--verbose", action="store_true", help="list lost parts and problems per file")
    args = parser.parse_args()

    files = sorted(f for f in args.files if f.suffix.lower() in {".xlsx", ".xlsm"})
    print(f"{len(files)} files")
    for engine in args.engines:
        lost_total = invalid = errors = 0
        files_losing = 0
        elapsed = 0.0
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp)
            if engine == "umya":
                if not args.umya_bin:
                    print("umya: --umya-bin not given, skipped")
                    continue
                start = time.perf_counter()
                proc = subprocess.run(
                    [str(args.umya_bin), str(out_dir), *map(str, files)], capture_output=True, text=True, check=False
                )
                elapsed = time.perf_counter() - start
                failed = {line.split("\t")[1] for line in proc.stdout.splitlines() if not line.startswith("OK")}
                errors = len(failed)
            for src in files:
                dst = out_dir / src.name
                if engine != "umya":
                    start = time.perf_counter()
                    try:
                        {"pyrust": run_pyrust, "openpyxl": run_openpyxl}[engine](src, dst)
                    except Exception as exc:
                        errors += 1
                        if args.verbose:
                            print(f"  {engine} ERROR {src.name}: {type(exc).__name__}: {exc}")
                        continue
                    finally:
                        elapsed += time.perf_counter() - start
                if not dst.exists():
                    continue
                lost = parts(src) - parts(dst)
                problems = check_package(dst)
                lost_total += sum(lost.values())
                files_losing += bool(lost)
                invalid += bool(problems)
                if args.verbose and (lost or problems):
                    print(f"  {engine} {src.name}: lost={dict(lost)} problems={problems[:3]}")
        print(
            f"{engine:9s} errors={errors:3d}  files losing parts={files_losing:3d}  "
            f"parts lost={lost_total:4d}  invalid outputs={invalid:3d}  time={elapsed:.2f}s"
        )


if __name__ == "__main__":
    main()
