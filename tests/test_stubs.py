"""The hand-written type stubs must match the compiled module and the way it is used.

* mypy stubtest imports the extension and compares it with the stubs: names,
  parameters, ``__new__`` vs ``__init__``, ``@final`` and positional-only markers.
  pyrefly has no equivalent, which is the only reason mypy is a dev dependency.
* pyrefly type-checks the package, tests and benchmarks against the stubs.
  Tests that pass wrong types on purpose carry ``# pyrefly: ignore[...]``, and an
  unused ignore is an error (see ``[tool.pyrefly]`` in pyproject.toml), so the
  stubs must accept every real call and reject every deliberate bad one.
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
ALLOWLIST = Path(__file__).parent / "stubtest_allowlist.txt"


def test_stubs_match_runtime():
    result = subprocess.run(
        [sys.executable, "-m", "mypy.stubtest", "pyrust_xlsxwriter", "--allowlist", str(ALLOWLIST)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_project_type_checks():
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "pyrefly",
            "check",
            "--python-interpreter-path",
            sys.executable,
            "--output-format",
            "min-text",
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
