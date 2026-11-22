"""Editor structural edits: inserting/deleting rows and columns, renames."""

import io
import zipfile

import openpyxl
import pytest
from ooxml_check import assert_valid

import pyrust_xlsxwriter as px


def build(tmp_path, chart=False):
    path = tmp_path / "source.xlsx"
    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_row("A1", ["Item", "Qty", "Price"])
    ws.write_rows("A2", [["a", 1, 10], ["b", 2, 20], ["c", 3, 30], ["d", 4, 40]])
    ws.write_formula("D2", "=B2*C2")
    ws.write_formula("B7", "=SUM(B2:B5)")
    ws.merge_range("A9:C9", "merged")
    ws.add_conditional_format(
        "B2:B5",
        px.ConditionalFormatCell(rule=px.ConditionalFormatCellRule.greater_than(2), format=px.Format(bold=True)),
    )
    ws.add_data_validation("C2:C5", px.DataValidation(whole_number=px.DataValidationRule.between(0, 100)))
    other = wb.add_worksheet("Summary")
    other.write_formula("A1", "=Data!B7")
    other.write_formula("A2", "=SUM(Data!C2:C5)")
    wb.define_name("Prices", "=Data!$C$2:$C$5")
    if chart:
        c = px.Chart(px.ChartType.Column)
        c.add_series(values="=Data!$B$2:$B$5", categories="=Data!$A$2:$A$5")
        ws.insert_chart(1, 5, c)
    wb.save(path)
    return path


def reopen(data):
    assert_valid(data)
    return openpyxl.load_workbook(io.BytesIO(data))


def xml(data, name):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return zf.read(name).decode()


def test_insert_rows_moves_everything(tmp_path):
    book = px.load_workbook(build(tmp_path, chart=True))
    ws = book["Data"]
    ws.insert_rows(2, 2)  # before row 3 ("b")
    assert ws.cell("A3") is None and ws.cell("A5") == "b"
    assert ws.formula("B9") == "=SUM(B2:B7)"
    assert ws.merged_ranges == [(10, 0, 10, 2)]
    ws.write("A3", "new")
    data = book.save_to_buffer()
    wb = reopen(data)
    d, s = wb["Data"], wb["Summary"]
    assert [d.cell(r, 1).value for r in range(1, 8)] == ["Item", "a", "new", None, "b", "c", "d"]
    assert d["B9"].value == "=SUM(B2:B7)" and d["D2"].value == "=B2*C2"
    assert {str(r) for r in d.merged_cells.ranges} == {"A11:C11"}
    assert s["A1"].value == "=Data!B9" and s["A2"].value == "=SUM(Data!C2:C7)"
    assert wb.defined_names["Prices"].attr_text == "Data!$C$2:$C$7"
    assert str(d.conditional_formatting._cf_rules.keys().__iter__().__next__().sqref) == "B2:B7"
    assert str(d.data_validations.dataValidation[0].sqref) == "C2:C7"
    chart = xml(data, "xl/charts/chart1.xml")
    assert "Data!$B$2:$B$7" in chart and "Data!$A$2:$A$7" in chart


def test_delete_rows_makes_ref_errors(tmp_path):
    book = px.load_workbook(build(tmp_path))
    ws = book["Data"]
    ws.delete_rows(1, 1)  # remove "a"
    ws.delete_rows(5)  # the SUM row (now row 6)
    data = book.save_to_buffer()
    wb = reopen(data)
    d, s = wb["Data"], wb["Summary"]
    assert [d.cell(r, 1).value for r in range(1, 5)] == ["Item", "b", "c", "d"]
    assert d["D2"].value is None  # the formula lived in the deleted row
    assert s["A1"].value == "=Data!#REF!"  # Excel keeps the sheet name
    assert s["A2"].value == "=SUM(Data!C2:C4)"


def test_insert_and_delete_columns(tmp_path):
    book = px.load_workbook(build(tmp_path))
    ws = book["Data"]
    ws.set_column_width("C", 20)
    ws.insert_cols("B", 1)
    assert ws.cell("C2") == 1 and ws.cell("B2") is None
    ws.write("B1", "Inserted")
    data = book.save_to_buffer()
    wb = reopen(data)
    d = wb["Data"]
    assert [c.value for c in d[1]][:5] == ["Item", "Inserted", "Qty", "Price", None]
    assert d["E2"].value == "=C2*D2" and wb["Summary"]["A2"].value == "=SUM(Data!D2:D5)"
    assert d.column_dimensions["D"].width == 20
    ws.delete_cols("B")
    data = book.save_to_buffer()
    d = reopen(data)["Data"]
    assert [c.value for c in d[1]][:3] == ["Item", "Qty", "Price"]
    assert d["D2"].value == "=B2*C2"


def test_rename_updates_formulas_everywhere(tmp_path):
    book = px.load_workbook(build(tmp_path, chart=True))
    book["Data"].name = "Sales Data"
    assert book["Summary"].formula("A1") == "='Sales Data'!B7"
    data = book.save_to_buffer()
    wb = reopen(data)
    assert wb["Summary"]["A2"].value == "=SUM('Sales Data'!C2:C5)"
    assert wb.defined_names["Prices"].attr_text == "'Sales Data'!$C$2:$C$5"
    assert "'Sales Data'!$B$2:$B$5" in xml(data, "xl/charts/chart1.xml")


def test_structural_errors(tmp_path):
    ws = px.load_workbook(build(tmp_path))["Data"]
    with pytest.raises(ValueError):
        ws.insert_rows(0, 0)
    with pytest.raises(px.RangeError):
        ws.delete_rows(1_048_576)
    ws.write("A1048576", "bottom")
    with pytest.raises(px.EditError, match="off the end"):
        ws.insert_rows(0)
