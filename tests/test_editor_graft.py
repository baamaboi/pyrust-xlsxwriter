"""Adding objects to existing workbooks (grafting writer output into the package)."""

import io
import re
import zipfile
from pathlib import Path

import openpyxl
import pytest
from ooxml_check import assert_valid

import pyrust_xlsxwriter as px

CORPUS = Path(__file__).parent / "corpus"


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {n: zf.read(n) for n in zf.namelist()}


def png() -> bytes:
    with zipfile.ZipFile(CORPUS / "image19.xlsx") as zf:
        name = next(n for n in zf.namelist() if n.startswith("xl/media/"))
        return zf.read(name)


def save(book) -> tuple[bytes, openpyxl.Workbook]:
    data = book.save_to_buffer()
    assert_valid(data)
    return data, openpyxl.load_workbook(io.BytesIO(data))


def data_book(tmp_path) -> Path:
    path = tmp_path / "data.xlsx"
    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_row("A1", ["Item", "Qty", "Price"])
    ws.write_rows("A2", [["a", 1, 10], ["b", 2, 20], ["c", 3, 30]])
    ws.add_conditional_format(
        "B2:B4",
        px.ConditionalFormatCell(rule=px.ConditionalFormatCellRule.greater_than(1), format=px.Format(italic=True)),
    )
    wb.save(path)
    return path


def test_conditional_formats_and_validation(tmp_path):
    book = px.load_workbook(data_book(tmp_path))
    ws = book["Data"]
    ws.add_conditional_format(
        "C2:C4",
        px.ConditionalFormatCell(
            rule=px.ConditionalFormatCellRule.less_than(25), format=px.Format(bold=True, font_color="red")
        ),
    )
    ws.add_conditional_format("A2:A4", px.ConditionalFormatDataBar())
    ws.add_conditional_format("C2:C4", px.ConditionalFormatDataBar())
    ws.add_data_validation("B2:B4", px.DataValidation(whole_number=px.DataValidationRule.between(0, 10)))
    data, wb = save(book)
    sheet = parts(data)["xl/worksheets/sheet1.xml"].decode()
    priorities = [int(p) for p in re.findall(r'priority="(\d+)"', sheet)]
    assert len(priorities) == len(set(priorities)) == 4
    ids = re.findall(r"<x14:id>(\{[^}]+\})</x14:id>", sheet)
    assert len(ids) == len(set(ids)) == 2  # data bar links are unique
    ranges = {str(r.sqref) for r in wb["Data"].conditional_formatting}
    assert {"B2:B4", "C2:C4", "A2:A4"} <= ranges
    assert str(wb["Data"].data_validations.dataValidation[0].sqref) == "B2:B4"
    styles = parts(data)["xl/styles.xml"].decode()
    assert styles.count("<dxf>") == 2  # the original italic + the new bold red


def test_validation_merges_into_existing(tmp_path):
    book = px.load_workbook(CORPUS / "data_validation03.xlsx")
    ws = book[0]
    before = len(
        openpyxl.load_workbook(CORPUS / "data_validation03.xlsx").worksheets[0].data_validations.dataValidation
    )
    ws.add_data_validation("H1:H5", px.DataValidation(list_strings=["yes", "no"]))
    _, wb = save(book)
    assert len(wb.worksheets[0].data_validations.dataValidation) == before + 1


def test_table_with_unique_name_and_headers(tmp_path):
    book = px.load_workbook(CORPUS / "table30.xlsx")
    ws = book[0]
    ws.add_table("K2:L6", px.Table(columns=["Region", "Sales"]))
    data, wb = save(book)
    tables = [n for n in parts(data) if n.startswith("xl/tables/")]
    assert len(tables) == 2
    names = {t.displayName for t in wb.worksheets[0].tables.values()}
    assert len(names) == 2
    assert wb.worksheets[0]["K2"].value == "Region" and wb.worksheets[0]["L2"].value == "Sales"
    # Default names (TableN) are renamed on conflict; explicit duplicates are errors.
    b2 = px.load_workbook(data)
    b2[0].add_table("N2:O4", px.Table(name="Sales"))
    b2[0].add_table("Q2:R4", px.Table(name="Sales"))
    with pytest.raises(px.EditError, match="already exists"):
        b2.save_to_buffer()


def test_images_and_charts(tmp_path):
    # Into a sheet that already has a drawing.
    book = px.load_workbook(CORPUS / "image19.xlsx")
    book[0].insert_image(20, 2, px.Image.from_bytes(png()))
    data, _ = save(book)
    drawing = parts(data)["xl/drawings/drawing1.xml"].decode()
    assert drawing.count("<xdr:pic>") == 2
    ids = re.findall(r'<xdr:cNvPr id="(\d+)"', drawing)
    assert len(ids) == len(set(ids))
    # A chart into a sheet without a drawing, reading the sheet's data.
    book = px.load_workbook(data_book(tmp_path))
    chart = px.Chart(px.ChartType.Column)
    chart.add_series(values="=Data!$B$2:$B$4", categories="=Data!$A$2:$A$4", name="Qty")
    book["Data"].insert_chart("E2", chart)
    data, wb = save(book)
    names = parts(data)
    assert "xl/charts/chart1.xml" in names and "xl/drawings/drawing1.xml" in names
    assert "Data!$B$2:$B$4" in names["xl/charts/chart1.xml"].decode()
    assert len(wb["Data"]._charts) == 1


def test_sparklines_shapes_and_notes(tmp_path):
    book = px.load_workbook(data_book(tmp_path))
    ws = book["Data"]
    ws.add_sparkline("D2", px.Sparkline("Data!B2:C2"))
    ws.insert_shape("F10", px.Shape.textbox(text="Hello"))
    ws.insert_note("A1", "a note")
    data, wb = save(book)
    sheet = parts(data)["xl/worksheets/sheet1.xml"].decode()
    assert "x14:sparklineGroups" in sheet and "<xm:sqref>D2</xm:sqref>" in sheet
    assert wb["Data"]["A1"].comment.text.strip().endswith("a note")
    # A second note on the same sheet isn't supported yet.
    book2 = px.load_workbook(data)
    book2["Data"].insert_note("B1", "another")
    with pytest.raises(px.EditError, match="already has notes"):
        book2.save_to_buffer()


def test_grafts_move_with_structural_edits(tmp_path):
    book = px.load_workbook(data_book(tmp_path))
    ws = book["Data"]
    ws.add_conditional_format("C2:C4", px.ConditionalFormatDataBar())
    ws.add_table("E1:F3", px.Table(columns=["X", "Y"]))
    ws.insert_rows(0, 2)
    _data, wb = save(book)
    sheet = wb["Data"]
    assert "C4:C6" in {str(r.sqref) for r in sheet.conditional_formatting}
    table = next(iter(sheet.tables.values()))
    assert table.ref == "E3:F5"
    assert sheet["E3"].value == "X"


def test_graft_into_new_sheet(tmp_path):
    book = px.load_workbook(data_book(tmp_path))
    new = book.add_worksheet("Charts")
    chart = px.Chart(px.ChartType.Line)
    chart.add_series(values="=Data!$C$2:$C$4")
    new.insert_chart("B2", chart)
    _data, wb = save(book)
    assert len(wb["Charts"]._charts) == 1
