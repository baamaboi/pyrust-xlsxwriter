"""Shared pytest fixtures.

* ``out``: a factory for output paths in a temporary directory.
* ``reopen``: load a saved file with openpyxl (an independent reader).
* ``part``: read one XML part from a saved package.
* ``corpus``: the directory of Excel-made reference files (tests/corpus).

Every file a test saves should also pass ``assert_valid`` from ooxml_check.
"""

from __future__ import annotations

import sys
import zipfile
from collections.abc import Callable
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).parent))

from ooxml_check import assert_valid

CORPUS = Path(__file__).parent / "corpus"


@pytest.fixture
def out(tmp_path: Path) -> Callable[[str], Path]:
    def make(name: str = "out.xlsx") -> Path:
        return tmp_path / name

    return make


@pytest.fixture
def reopen() -> Callable[..., openpyxl.Workbook]:
    def load(path: Path, **kwargs: object) -> openpyxl.Workbook:
        assert_valid(path)
        return openpyxl.load_workbook(path, **kwargs)

    return load


@pytest.fixture
def part() -> Callable[[Path, str], str]:
    def read(path: Path, name: str) -> str:
        with zipfile.ZipFile(path) as zf:
            return zf.read(name).decode("utf-8")

    return read


@pytest.fixture
def corpus() -> Path:
    return CORPUS
