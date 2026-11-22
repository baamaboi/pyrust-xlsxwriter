"""convert_to_xlsx(): any readable format to xlsx."""

import datetime as dt
import io
from pathlib import Path

import openpyxl
from ooxml_check import assert_valid
from test_reader import make_ods

import pyrust_xlsxwriter as px
from pyrust_xlsxwriter._convert import _odf_formula

CORPUS = Path(__file__).parent / "corpus"


def test_odf_formulas():
    assert _odf_formula("=[.B1]*2") == "=B1*2"
    assert _odf_formula("=SUM([.A1:.A3];[Data.B1])") == "=SUM(A1:A3,Data!B1)"
    assert _odf_formula("=[$'My Sheet'.$A$1]") == "='My Sheet'!$A$1"
    assert _odf_formula('=IF([.A1]>0;"a;b";"c")') == '=IF(A1>0,"a;b","c")'


def test_ods_to_xlsx():
    data = px.convert_to_xlsx(make_ods())
    assert_valid(data)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert wb.sheetnames == ["Data", "Second"]
    ws = wb["Data"]
    assert [c.value for c in ws[1]] == ["name", 3, 2.5, True]
    assert ws["A2"].value == dt.datetime(2024, 1, 2)
    assert ws["D2"].value == "=B1*2"
    assert wb["Second"]["A1"].value == "x"


def test_xlsx_round_trip_keeps_values_styles_merges_links(tmp_path):
    for name in [
        "format01.xlsx",
        "merge_range02.xlsx",
        "hyperlink35.xlsx",
        "rich_string12.xlsx",
        "defined_name01.xlsx",
    ]:
        src = CORPUS / name
        out = tmp_path / f"converted_{name}"
        assert px.convert_to_xlsx(src, out) is None
        assert_valid(out)
        before, after = openpyxl.load_workbook(src), openpyxl.load_workbook(out)
        for ws in before.worksheets:
            other = after[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    if type(cell).__name__ == "MergedCell":
                        continue
                    assert other[cell.coordinate].value == cell.value, (name, cell.coordinate)
                    if cell.value is not None:
                        assert other[cell.coordinate].font.b == cell.font.b, (name, cell.coordinate)
                        assert other[cell.coordinate].number_format == cell.number_format, (name, cell.coordinate)
            assert {str(r) for r in other.merged_cells.ranges} == {str(r) for r in ws.merged_cells.ranges}
            for coord, link in ((c.coordinate, c.hyperlink) for r in ws.iter_rows() for c in r if c.hyperlink):
                assert other[coord].hyperlink is not None and other[coord].hyperlink.target == link.target


def test_text_that_looks_like_a_formula_stays_text():
    wb = px.Workbook()
    wb.add_worksheet().write_string("A1", "=not a formula")
    data = px.convert_to_xlsx(wb.save_to_buffer())
    cell = openpyxl.load_workbook(io.BytesIO(data)).active["A1"]
    assert cell.value == "=not a formula" and cell.data_type == "s"
