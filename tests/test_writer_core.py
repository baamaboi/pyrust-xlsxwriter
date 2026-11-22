"""Writer core: value dispatch, cell writing, rows/columns, workbook lifecycle."""

import datetime as dt
import decimal
import io
import zipfile
from pathlib import Path

import openpyxl
import pytest

import pyrust_xlsxwriter as px

VBA = Path(__file__).parent / "corpus" / "vbaProject.bin"


def test_write_dispatch(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet("S")
    values = [
        True,
        None,
        dt.date(2024, 1, 2),
        dt.datetime(2024, 1, 2, 3, 4, 5),
        dt.time(6, 30),
        decimal.Decimal("1.5"),
        2**60,
        1.25,
        "text",
        "=1+1",
        dt.timedelta(hours=36),
    ]
    ws.write_column("A1", values)
    path = out()
    wb.save(path)
    sh = reopen(path)["S"]
    got = [sh.cell(r + 1, 1) for r in range(len(values))]
    assert got[0].value is True and got[0].data_type == "b"
    assert got[1].value is None
    assert got[2].value == dt.datetime(2024, 1, 2) and got[2].number_format == "yyyy-mm-dd"
    assert got[3].value == dt.datetime(2024, 1, 2, 3, 4, 5)
    assert got[4].number_format == "hh:mm:ss"
    assert got[5].value == 1.5
    assert got[6].value == float(2**60)
    assert got[7].value == 1.25
    assert got[8].value == "text"
    assert got[9].value == "=1+1"
    assert got[10].number_format == "[h]:mm:ss"


def test_tz_aware_datetime_rejected():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.ParameterError, match="time zones"):
        ws.write(0, 0, dt.datetime(2024, 1, 1, tzinfo=dt.UTC))


def test_unsupported_type():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError, match="unsupported type"):
        ws.write(0, 0, object())  # pyrefly: ignore[no-matching-overload]


def test_a1_and_keyword_calls(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    bold = px.Format(bold=True)
    ws.write("B2", "a", bold)
    ws.write(2, 1, value="b", format=bold)
    ws.write_number("B4", 4)
    ws.write_string(4, "B", "c")  # column letters are accepted too
    ws.write_formula("B6", "=B4*2", value=8)
    ws.write_url("B7", "https://example.com", text="Example", tip="tip")
    ws.write_rich_string("B8", ["plain ", (bold, "bold")])
    ws.write_boolean("B9", False)
    ws.write_blank("B10", bold)
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    assert sh["B2"].value == "a" and sh["B2"].font.b
    assert sh["B3"].value == "b" and sh["B3"].font.b
    assert sh["B4"].value == 4
    assert sh["B5"].value == "c"
    assert sh["B6"].value == "=B4*2"
    assert sh["B7"].hyperlink.target == "https://example.com" and sh["B7"].value == "Example"
    assert sh["B9"].value is False


def test_argument_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError, match="unexpected keyword"):
        ws.write(0, 0, 1, colour=None)  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(TypeError, match="missing required"):
        ws.write(0, 0)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.RangeError):
        ws.write(1_048_576, 0, 1)
    with pytest.raises(px.RangeError):
        ws.merge_range("C3:A1", "x")


def test_rows_columns_and_matrix(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_row("A1", [1, 2, 3])
    ws.write_column(1, 0, ["a", "b"])
    ws.write_rows("C2", [[1, 2], [3, 4]])
    ws.write_columns("F1", [[1, 2], [3, 4]])
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    assert [c.value for c in sh[1][:3]] == [1, 2, 3]
    assert sh["A2"].value == "a" and sh["A3"].value == "b"
    assert sh["C2"].value == 1 and sh["D3"].value == 4
    assert sh["F2"].value == 2 and sh["G1"].value == 3


def test_merge_and_formats(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    title = px.Format(bold=True, align=px.FormatAlign.Center, background_color="#DDEBF7")
    ws.merge_range("A1:D1", "Title", title)
    ws.merge_range(2, 0, 2, 3, 42.5)
    ws.write("A5", 1)
    ws.set_cell_format("A5", px.Format(num_format="0.00"))
    ws.set_range_format("A6:B7", px.Format(italic=True))
    ws.set_range_format_with_border("D6:E8", px.Format(), px.Format(border=px.FormatBorder.Medium))
    ws.write("A9", 9)
    ws.clear_cell("A9")
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    merged = {str(r) for r in sh.merged_cells.ranges}
    assert merged == {"A1:D1", "A3:D3"}
    assert sh["A1"].value == "Title" and sh["A1"].font.b
    assert sh["A3"].value == 42.5
    assert sh["A5"].number_format == "0.00"
    assert sh["A6"].font.i
    assert sh["D6"].border.top.style == "medium" and sh["E8"].border.bottom.style == "medium"
    assert sh["A9"].value is None


def test_format_unset_and_all_properties(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    f = px.Format(bold=True, italic=True)
    f.bold = False
    f.font_color = None
    ws.write("A1", "x", f)
    full = px.Format(
        strikethrough=True,
        font_script=px.FormatScript.Superscript,
        pattern=px.FormatPattern.DarkGrid,
        foreground_color="red",
        background_color="blue",
        border_top=px.FormatBorder.Thick,
        border_top_color=px.Color.Green,
        border_diagonal=px.FormatBorder.Thin,
        border_diagonal_type=px.FormatDiagonalBorder.BorderUp,
        quote_prefix=True,
        rotation=45,
        indent=2,
        text_wrap=True,
        shrink=False,
        locked=False,
        hidden=True,
    )
    ws.write("A2", "y", full)
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    assert sh["A1"].font.b is False and sh["A1"].font.i is True
    a2 = sh["A2"]
    assert a2.font.strike and a2.font.vertAlign == "superscript"
    assert a2.fill.patternType == "darkGrid"
    assert a2.border.top.style == "thick"
    assert a2.border.diagonalUp
    assert a2.alignment.textRotation == 45 and a2.alignment.indent == 2 and a2.alignment.wrap_text
    assert a2.protection.locked is False and a2.protection.hidden is True
    assert a2.quotePrefix


def test_columns_and_rows(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_column_width("B:D", 20)
    ws.set_column_width(5, 12, px.Format(bold=True))
    ws.set_column_width_pixels("H", 100)
    ws.set_column_hidden("J")
    ws.set_row_height(1, 30)
    ws.set_row_height_pixels(2, 40)
    ws.set_row_hidden(3)
    ws.group_rows(5, 7, collapsed=True)
    ws.group_columns("L:M")
    ws.freeze_panes("B2")
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    dim = sh.column_dimensions["B"]
    assert (dim.min, dim.max) == (2, 4) and dim.width == pytest.approx(20.71, abs=0.1)
    assert sh.column_dimensions["J"].hidden
    assert sh.row_dimensions[2].height == 30
    assert sh.row_dimensions[4].hidden
    assert sh.row_dimensions[6].outlineLevel == 1
    assert sh.freeze_panes == "B2"


def test_worksheet_properties(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet("Props", tab_color="#FF0000", zoom=150, paper_size=px.PaperSize.A4)
    assert ws.tab_color == px.Color("#FF0000")
    assert ws.paper_size == px.PaperSize.A4
    ws.paper_size = 1
    assert ws.paper_size == px.PaperSize.Letter
    ws.hidden = False
    ws.right_to_left = True
    with pytest.raises(ValueError):
        ws.zoom = 5
    wb.add_worksheet()
    path = out()
    wb.save(path)
    book = reopen(path)
    assert book.sheetnames == ["Props", "Sheet2"]
    assert book["Props"].sheet_properties.tabColor.rgb == "FFFF0000"


def test_workbook_lookup_and_names():
    wb = px.Workbook()
    a = wb.add_worksheet()
    b = wb.add_worksheet("Data")
    assert wb.sheet_names == ["Sheet1", "Data"]
    assert wb["Data"] is b or wb["Data"].name == "Data"
    assert wb[0].name == a.name and wb[-1].name == "Data"
    assert "Data" in wb and len(wb) == 2
    with pytest.raises(px.SheetNameError):
        wb.add_worksheet("data")
    with pytest.raises(KeyError):
        wb["missing"]
    standalone = px.Worksheet()
    wb.push_worksheet(standalone)
    assert standalone.name == "Sheet3"
    with pytest.raises(px.ParameterError):
        wb.push_worksheet(standalone)


def test_save_variants(out, tmp_path):
    wb = px.Workbook()
    wb.add_worksheet().write("A1", 1)
    wb.save(out("a.xlsx"))  # pathlib.Path
    wb.save(str(out("b.xlsx")))  # str
    data = wb.save_to_buffer()
    assert data[:2] == b"PK"
    buffer = io.BytesIO()
    wb.save_to_file(buffer)
    assert openpyxl.load_workbook(io.BytesIO(buffer.getvalue())).active["A1"].value == 1
    # Saving twice keeps worksheets usable and writable.
    wb[0].write("A2", 2)
    wb.save(out("c.xlsx"))
    assert openpyxl.load_workbook(out("c.xlsx")).active["A2"].value == 2


def test_context_manager(tmp_path):
    path = tmp_path / "ctx.xlsx"
    with px.Workbook(path) as wb:
        wb.add_worksheet().write("A1", "ctx")
    assert openpyxl.load_workbook(path).active["A1"].value == "ctx"
    path2 = tmp_path / "not_saved.xlsx"
    with pytest.raises(RuntimeError), px.Workbook(path2) as wb:
        wb.add_worksheet()
        raise RuntimeError("boom")
    assert not path2.exists()


@pytest.mark.skipif(not VBA.exists(), reason="corpus vbaProject.bin missing")
def test_vba_project_is_saved(out):
    wb = px.Workbook()
    wb.add_worksheet().set_vba_name("Sheet1")
    wb.add_vba_project(VBA)
    wb.set_vba_name("ThisWorkbook")
    path = out("macro.xlsm")
    wb.save(path)
    wb.save(path)  # second save must keep it too
    with zipfile.ZipFile(path) as zf:
        assert "xl/vbaProject.bin" in zf.namelist()
