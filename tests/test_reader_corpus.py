"""Reader vs. Excel-made files: everything reads, and agrees with openpyxl."""

from __future__ import annotations

import warnings
from pathlib import Path

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

import pyrust_xlsxwriter as px

CORPUS = Path(__file__).parent / "corpus"
FILES = sorted(p for p in CORPUS.iterdir() if p.suffix in {".xlsx", ".xlsm"})
# calamine decodes the `_xHHHH_` escapes Excel uses for control characters;
# openpyxl leaves them as text. Checked separately below.
ESCAPED = {"shared_strings01.xlsx"}


def trim(rows) -> list[list[object]]:
    """Rows as lists without trailing empty rows/columns (openpyxl also
    counts formatted empty cells in its dimensions)."""
    rows = [list(r) for r in rows]
    width = max((i + 1 for r in rows for i, v in enumerate(r) if v is not None), default=0)
    rows = [(r + [None] * width)[:width] for r in rows]
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    return rows


def load(path: Path, **kwargs: object) -> openpyxl.Workbook:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(path, **kwargs)


def test_corpus_present():
    assert len(FILES) >= 40


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_reads_everything(path: Path):
    with px.open_workbook(path) as book:
        assert book.format == "xlsx"
        assert len(book) == len(book.sheet_names) == len(book.sheets)
        assert isinstance(book.properties, dict)
        assert isinstance(book.defined_names, list)
        for table in book.tables:
            assert len(table.values()) >= 0
        for ext, data in book.pictures:
            assert isinstance(ext, str) and isinstance(data, bytes) and data
        modules = book.vba_modules
        assert modules is None or all(isinstance(v, str) for v in modules.values())
        for ws in book:
            values = ws.values()
            assert len(values) == ws.max_row
            assert all(len(r) == ws.max_column for r in values)
            formulas = ws.formulas()
            assert len(formulas) >= ws.max_row
            for r, row in enumerate(values):
                for c, _ in enumerate(row):
                    ws.style(r, c)
            assert all(len(m) == 4 for m in ws.merged_ranges)
            assert all(isinstance(h, px.HyperlinkInfo) for h in ws.hyperlinks)
            assert list(ws.iter_rows()) == [tuple(r) for r in values]


@pytest.mark.parametrize("path", [p for p in FILES if p.name not in ESCAPED], ids=lambda p: p.name)
def test_values_match_openpyxl(path: Path):
    expected = load(path, data_only=True)
    book = px.open_workbook(path)
    assert book.sheet_names == expected.sheetnames
    for ws in expected.worksheets:
        got = trim(book[ws.title].values())
        want = trim(ws.iter_rows(values_only=True))
        assert got == want, ws.title
        # Same types too (int vs float, datetime vs time, ...).
        assert [[type(v) for v in r] for r in got] == [[type(v) for v in r] for r in want]


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_formulas_merges_links_visibility_match_openpyxl(path: Path):
    expected = load(path)
    book = px.open_workbook(path)
    for ws in expected.worksheets:
        ours = book[ws.title]
        want = []
        for row in ws.iter_rows():
            want.append(
                [
                    (c.value.text if isinstance(c.value, ArrayFormula) else c.value) if c.data_type == "f" else None
                    for c in row
                ]
            )
        assert trim(ours.formulas()) == trim(want), ws.title

        merged = sorted((m.min_row - 1, m.min_col - 1, m.max_row - 1, m.max_col - 1) for m in ws.merged_cells.ranges)
        assert sorted(ours.merged_ranges) == merged

        links = {
            (c.hyperlink.ref, c.hyperlink.target, c.hyperlink.location)
            for row in ws.iter_rows()
            for c in row
            if c.hyperlink is not None
        }
        assert {(h.reference, h.target, h.location) for h in ours.hyperlinks} == links

        state = {"visible": "visible", "hidden": "hidden", "veryHidden": "very_hidden"}[ws.sheet_state]
        assert ours.visibility == state
        assert ours.kind == "worksheet"
    for name in expected.sheetnames:
        if name not in [ws.title for ws in expected.worksheets]:
            assert book[name].kind == "chartsheet"
            assert book[name].values() == []


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_defined_names_match_openpyxl(path: Path):
    expected = load(path)
    want = {(name, dn.attr_text) for name, dn in expected.defined_names.items()}
    for ws in expected.worksheets:
        want |= {(name, dn.attr_text) for name, dn in ws.defined_names.items()}
    got = {(name, value) for name, value in px.open_workbook(path).defined_names}
    # openpyxl moves print areas/titles and filters out of the defined names.
    assert {n for n in got if not n[0].startswith("_xlnm.")} == want


def test_control_character_escapes():
    ws = px.open_workbook(CORPUS / "shared_strings01.xlsx")[0]
    values = [r[0] for r in ws.values()]
    assert "\x00" in values and "\x1f" in values
    assert all(isinstance(v, str) for v in values)


def test_chartsheet_and_hyperlink_details():
    book = px.open_workbook(CORPUS / "chart_bar14.xlsx")
    assert [s.kind for s in book.sheets] == ["worksheet", "worksheet", "worksheet", "chartsheet"]
    (link,) = book["Sheet2"].hyperlinks
    assert link.cell == (5, 0) and link.range == (5, 0, 5, 0) and link.reference == "A6"
    assert link.target == "http://www.perl.com/" and link.location is None


def test_properties_match_openpyxl():
    path = CORPUS / "properties04.xlsx"
    props = px.open_workbook(path).properties
    core = load(path).properties
    assert props["author"] == core.creator
    assert props["created"] == core.created
    assert props["modified"] == core.modified
    assert props["last_modified_by"] == core.lastModifiedBy
    assert props["custom"]["Department"] == "Finance"
    assert "title" not in props  # only keys that are set


def test_tables():
    book = px.open_workbook(CORPUS / "table30.xlsx")
    (table,) = book.tables
    want = load(CORPUS / "table30.xlsx").active.tables["Table1"]
    assert table.name == want.name == "Table1"
    assert table.sheet_name == "Sheet1"
    assert table.columns == [c.name for c in want.tableColumns]
    assert px.cell_range(*table.range) == want.ref
    assert table.has_header and not table.has_totals
    assert table.values() == [[None] * 4] * 10
    assert table.records()[0] == dict.fromkeys(table.columns)
    assert book.table("table1").name == "Table1"
    with pytest.raises(KeyError):
        book.table("missing")
    assert px.open_workbook(CORPUS / "format01.xlsx").tables == []


def test_vba_and_pictures():
    modules = px.open_workbook(CORPUS / "macro01.xlsm").vba_modules
    assert modules is not None and sorted(modules) == ["Module1", "Sheet1", "ThisWorkbook"]
    assert "Sub Macro1()" in modules["Module1"]
    assert px.open_workbook(CORPUS / "format01.xlsx").vba_modules is None
    ((ext, data),) = px.open_workbook(CORPUS / "image19.xlsx").pictures
    assert ext in {"jpeg", "jpg"} and data[:3] == b"\xff\xd8\xff"


def test_styles_match_openpyxl():
    path = CORPUS / "format01.xlsx"
    expected = load(path)
    book = px.open_workbook(path)
    checked = 0
    for ws in expected.worksheets:
        ours = book[ws.title]
        for row in ws.iter_rows():
            for c in row:
                fmt = ours.style(c.row - 1, c.column - 1)
                if c.style_id == 0:
                    assert fmt is None
                    continue
                checked += 1
                assert fmt is not None
                assert fmt.bold == bool(c.font.b)
                assert fmt.italic == bool(c.font.i)
                assert fmt.font_name == c.font.name
                assert fmt.font_size == c.font.sz
                assert (fmt.num_format or "General") == c.number_format
    assert checked > 0
