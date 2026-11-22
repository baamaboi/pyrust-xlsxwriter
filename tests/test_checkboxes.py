"""Checkboxes (Excel 365 boolean cells shown as check boxes)."""

import re

import pytest

import pyrust_xlsxwriter as px


def test_insert_checkbox(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_checkbox("B2", True)
    ws.insert_checkbox(2, 1, False)
    ws.insert_checkbox(3, "B", checked=True, format=px.Format(bold=True, font_color="red"))
    ws.insert_checkbox("B5", False, None)
    path = out()
    wb.save(path)
    sheet = reopen(path).active
    assert [sheet[f"B{r}"].value for r in range(2, 6)] == [True, False, True, False]
    assert sheet["B4"].font.b is True
    xml = part(path, "xl/worksheets/sheet1.xml")
    assert re.search(r'<c r="B2" s="(\d+)" t="b"><v>1</v></c>', xml)
    styles = part(path, "xl/styles.xml")
    assert "xfpb:xfComplement" in styles  # the checkbox cell style
    assert "XFControls" in part(path, "xl/featurePropertyBag/featurePropertyBag.xml")


def test_checkbox_same_style_as_format_checkbox(out, part):
    """insert_checkbox is equivalent to a boolean with Format(checkbox=True)."""
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_checkbox("A1", True)
    ws.write_boolean("A2", True, px.Format(checkbox=True))
    path = out()
    wb.save(path)
    xml = part(path, "xl/worksheets/sheet1.xml")
    s1 = re.search(r'<c r="A1" s="(\d+)"', xml)
    s2 = re.search(r'<c r="A2" s="(\d+)"', xml)
    assert s1 and s2 and s1.group(1) == s2.group(1)


def test_insert_checkbox_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError, match="missing required argument: 'checked'"):
        ws.insert_checkbox("A1")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError):
        ws.insert_checkbox("A1", "yes")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        ws.insert_checkbox("A1", True, "bold")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.RangeError):
        ws.insert_checkbox(0, 16_384, True)
