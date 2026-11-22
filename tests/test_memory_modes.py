"""Constant-memory and low-memory worksheets (`add_worksheet(memory=...)`)."""

import gc
import io

import openpyxl
import pytest
from ooxml_check import assert_valid

import pyrust_xlsxwriter as px


def save(wb, out, name="out.xlsx"):
    path = out(name)
    wb.save(path)
    return path


def values(sheet):
    return [[c.value for c in row] for row in sheet.iter_rows()]


@pytest.mark.parametrize("memory", ["constant", "low"])
def test_basic(out, reopen, part, memory):
    wb = px.Workbook()
    ws = wb.add_worksheet("Data", memory=memory, zoom=120, tab_color="#FF0000")
    bold = px.Format(bold=True)
    ws.write_row("A1", ["name", "value"], bold)
    ws.write_row(1, 0, ["a", 1.5])
    ws.write(2, 0, "b")
    ws.write(2, 1, 2)
    path = save(wb, out)
    sh = reopen(path)["Data"]
    assert values(sh) == [["name", "value"], ["a", 1.5], ["b", 2]]
    assert sh["A1"].font.b and not sh["A2"].font.b
    assert sh.sheet_view.zoomScale == 120
    xml = part(path, "xl/worksheets/sheet1.xml")
    if memory == "constant":
        assert 't="inlineStr"' in xml  # strings stored inline
    else:
        assert 't="s"' in xml and "inlineStr" not in xml  # shared string table


def test_mixed_with_standard_sheets(out, reopen):
    wb = px.Workbook()
    std = wb.add_worksheet("Std")
    const = wb.add_worksheet("Const", memory="constant")
    low = wb.add_worksheet("Low", memory="low")
    fmt = px.Format(num_format="0.00", italic=True)
    for i, ws in enumerate((std, const, low)):
        ws.write("A1", f"sheet {i}")
        ws.write("B2", i + 0.5, fmt)
        ws.write("C3", "shared")
    std.write("A5", "out of order")
    std.write("A4", "is fine here")
    book = reopen(save(wb, out))
    for i, name in enumerate(("Std", "Const", "Low")):
        sh = book[name]
        assert sh["A1"].value == f"sheet {i}" and sh["C3"].value == "shared"
        assert sh["B2"].value == i + 0.5
        assert sh["B2"].number_format == "0.00" and sh["B2"].font.i
    assert book["Std"]["A4"].value == "is fine here"


@pytest.mark.parametrize("memory", ["constant", "low"])
def test_100k_rows(out, reopen, memory):
    rows = 100_000
    wb = px.Workbook()
    ws = wb.add_worksheet(memory=memory)
    ws.write_row(0, 0, ["id", "label", "half"])
    for r in range(1, rows + 1):
        ws.write_row(r, 0, [r, f"item {r % 50}", r / 2])
    path = save(wb, out)
    book = reopen(path, read_only=True)
    sh = book["Sheet1"]
    count = 0
    last = None
    for row in sh.iter_rows(values_only=True):
        count += 1
        last = row
    assert count == rows + 1
    assert last == (rows, f"item {rows % 50}", rows / 2)
    book.close()


def test_save_twice_and_write_after_save(out, reopen):
    wb = px.Workbook()
    std = wb.add_worksheet("Std")
    const = wb.add_worksheet("Const", memory="constant")
    low = wb.add_worksheet("Low", memory="low")
    bold = px.Format(bold=True)
    for ws in (std, const, low):
        ws.write("A1", "head", bold)
        ws.write_row("A2", [1, "two"])
    first = save(wb, out, "first.xlsx")
    second = save(wb, out, "second.xlsx")
    for path in (first, second):
        book = reopen(path)
        for name in ("Std", "Const", "Low"):
            assert values(book[name]) == [["head", None], [1, "two"]]
            assert book[name]["A1"].font.b
    # Keep writing (in row order) after a save, change a property, save again.
    for ws in (std, const, low):
        ws.write_row("A3", ["after", 3], bold)
        ws.zoom = 150
    third = save(wb, out, "third.xlsx")
    book = reopen(third)
    for name in ("Std", "Const", "Low"):
        assert values(book[name]) == [["head", None], [1, "two"], ["after", 3]]
        assert book[name]["A3"].font.b and book[name].sheet_view.zoomScale == 150
    # Buffer and file-object saves work the same way.
    data = wb.save_to_buffer()
    assert_valid(data)
    buf = io.BytesIO()
    wb.save_to_file(buf)
    loaded = openpyxl.load_workbook(io.BytesIO(buf.getvalue()))
    assert values(loaded["Low"]) == [["head", None], [1, "two"], ["after", 3]]


def test_earlier_rows_are_rejected(out, reopen):
    """In memory modes data must be written in row order: writing above the
    current row raises instead of being silently dropped."""
    wb = px.Workbook()
    ws = wb.add_worksheet(memory="constant")
    ws.write("A5", "row 5")
    ws.write("B5", "same row is fine")
    with pytest.raises(px.ParameterError, match="row order"):
        ws.write("A2", "too late")
    with pytest.raises(px.ParameterError, match="row order"):
        ws.write_row("A3", [1, 2])
    first = save(wb, out)
    # A save completes the current row: continue on a later one.
    with pytest.raises(px.ParameterError, match="saved"):
        ws.write("C5", "same row, after save")
    ws.write("A6", "next row")
    sh = reopen(first)["Sheet1"]
    assert sh["A5"].value == "row 5" and sh["A2"].value is None
    sh = reopen(save(wb, out, "second.xlsx"))["Sheet1"]
    assert sh["A6"].value == "next row"


def test_standard_sheets_have_no_row_order():
    ws = px.Workbook().add_worksheet()
    ws.write("A5", 1)
    ws.write("A1", 2)  # any order is fine in the default mode


def test_very_hidden_survives_property_changes(out, part):
    wb = px.Workbook()
    wb.add_worksheet("Visible")
    ws = wb.add_worksheet("Secret")
    ws.set_very_hidden()
    ws.zoom = 120
    ws.tab_color = "red"
    path = save(wb, out)
    assert 'state="veryHidden"' in part(path, "xl/workbook.xml")


def test_untouched_header_footer_adds_no_xml(out, part):
    wb = px.Workbook()
    wb.add_worksheet().write("A1", 1)
    path = save(wb, out)
    sheet = part(path, "xl/worksheets/sheet1.xml")
    assert "headerFooter" not in sheet


def test_memory_sheet_can_not_move_to_another_workbook():
    wb = px.Workbook()
    ws = wb.add_worksheet("Const", memory="constant")
    other = px.Workbook()
    with pytest.raises(px.ParameterError, match="belongs to the workbook"):
        other.push_worksheet(ws)
    with pytest.raises(px.ParameterError, match="already in the workbook"):
        wb.push_worksheet(ws)
    # Still refused once the creating workbook is gone.
    del wb
    gc.collect()
    with pytest.raises(px.ParameterError, match="belongs to the workbook"):
        other.push_worksheet(ws)
    assert len(other) == 0


def test_dropped_workbooks_do_not_block_new_sheets(out, reopen):
    for _ in range(20):
        wb = px.Workbook()
        wb.add_worksheet(memory="low").write("A1", 1)
        del wb
    gc.collect()
    target = px.Workbook()
    for i in range(20):
        sheet = px.Worksheet(f"S{i}")
        sheet.write("A1", i)
        target.push_worksheet(sheet)
    book = reopen(save(target, out))
    assert [book[f"S{i}"]["A1"].value for i in range(20)] == list(range(20))


def test_memory_argument_errors():
    wb = px.Workbook()
    with pytest.raises(px.ParameterError, match="memory must be"):
        wb.add_worksheet(memory="tiny")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="unexpected keyword"):
        wb.add_worksheet("X", memory="constant", colour="red")  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(px.SheetNameError):
        wb.add_worksheet("Bad[name]", memory="low")
    assert len(wb) == 0
    wb.add_worksheet("X", memory="constant")
    with pytest.raises(px.SheetNameError):
        wb.add_worksheet("x", memory="low")


def test_set_tempdir(out, reopen, tmp_path):
    wb = px.Workbook()
    tempdir = tmp_path / "temp"
    tempdir.mkdir()
    wb.set_tempdir(tempdir)
    ws = wb.add_worksheet(memory="constant")
    ws.write("A1", "via tempdir")
    assert reopen(save(wb, out))["Sheet1"]["A1"].value == "via tempdir"
    with pytest.raises(px.FileError):
        wb.set_tempdir(tmp_path / "does-not-exist")
    with pytest.raises(OSError):
        wb.set_tempdir(str(tmp_path / "does-not-exist"))


def test_memory_sheets_use_default_format(out, reopen, part):
    wb = px.Workbook()
    wb.set_default_format(px.Format(font_name="Arial", font_size=10), 18, 72)
    const = wb.add_worksheet("Const", memory="constant")
    low = wb.add_worksheet("Low", memory="low", default_row_height=30)
    for ws in (const, low):
        ws.write("A1", "x")
        ws.write("A2", "y", px.Format(bold=True, font_name="Arial", font_size=10))
    path = save(wb, out)
    book = reopen(path)
    for name in ("Const", "Low"):
        assert book[name]["A1"].font.name == "Arial" and not book[name]["A1"].font.b
        assert book[name]["A2"].font.b
    assert '<sheetFormatPr defaultRowHeight="13.5"/>' in part(path, "xl/worksheets/sheet1.xml")
    assert 'defaultRowHeight="30"' in part(path, "xl/worksheets/sheet2.xml")


def test_records_in_constant_memory(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet(memory="constant")
    rows = ({"id": i, "name": f"n{i}"} for i in range(1000))
    assert ws.write_records("B2", rows) == (1, 1, 1001, 2)  # header + 1000 rows
    sh = reopen(save(wb, out))["Sheet1"]
    assert sh["B2"].value == "id" and sh["C1002"].value == "n999"
