"""Worksheet page setup, views, printing, protection and ignored errors."""

import copy
import re

import pytest
from ooxml_check import assert_valid

import pyrust_xlsxwriter as px

SHEET1 = "xl/worksheets/sheet1.xml"


def excel_hash(password: str) -> str:
    """Excel's legacy 16-bit password hash (independent implementation)."""
    h = 0
    for ch in reversed(password):
        h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF)
        h ^= ord(ch)
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF)
    h ^= len(password)
    h ^= 0xCE4B
    return f"{h:04X}"


def save(wb, out, name="out.xlsx"):
    path = out(name)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Views
# ---------------------------------------------------------------------------


def test_selection_and_top_left_cell(out, reopen):
    wb = px.Workbook()
    a = wb.add_worksheet("A")
    a.set_selection("B3:D6")
    a.set_top_left_cell("C20")
    b = wb.add_worksheet("B")
    b.set_selection(5, 3, 2, 1)  # reversed: D6 is the active cell
    c = wb.add_worksheet("C")
    c.set_selection(1, "C")
    d = wb.add_worksheet("D")
    d.set_selection("E5")
    d.set_top_left_cell(9, 1)
    wb_ = reopen(save(wb, out))
    sel = wb_["A"].sheet_view.selection[0]
    assert sel.sqref == "B3:D6" and sel.activeCell == "B3"
    assert wb_["A"].sheet_view.topLeftCell == "C20"
    sel = wb_["B"].sheet_view.selection[0]
    assert sel.sqref == "B3:D6" and sel.activeCell == "D6"
    assert wb_["C"].sheet_view.selection[0].sqref == "C2"
    assert wb_["D"].sheet_view.selection[0].sqref == "E5"
    assert wb_["D"].sheet_view.topLeftCell == "B10"


def test_freeze_panes_top_cell(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.freeze_panes("A2")
    ws.set_freeze_panes_top_cell("A20")
    pane = reopen(save(wb, out))["Sheet1"].sheet_view.pane
    assert pane.state == "frozen" and pane.topLeftCell == "A20" and pane.ySplit == 1


@pytest.mark.parametrize(
    ("method", "view"),
    [("set_view_page_layout", "pageLayout"), ("set_view_page_break_preview", "pageBreakPreview")],
)
def test_views(out, reopen, method, view):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    getattr(ws, method)()
    assert reopen(save(wb, out))["Sheet1"].sheet_view.view == view
    ws.set_view_normal()
    assert reopen(save(wb, out, "normal.xlsx"))["Sheet1"].sheet_view.view in (None, "normal")


def test_hidden_zeroes_first_tab_and_zoom_to_fit(out, reopen, part):
    wb = px.Workbook()
    ws1 = wb.add_worksheet()
    ws2 = wb.add_worksheet()
    ws1.set_hidden_zeroes()
    ws2.set_first_tab()
    ws2.active = True
    ws1.set_zoom_to_fit()  # only used by chartsheets; must not break a worksheet
    path = save(wb, out)
    book = reopen(path)
    assert book["Sheet1"].sheet_view.showZeros is False
    assert book.views[0].firstSheet == 1
    ws1.set_hidden_zeroes(False)
    path = save(wb, out, "b.xlsx")
    assert reopen(path)["Sheet1"].sheet_view.showZeros in (None, True)


def test_very_hidden(out, reopen):
    wb = px.Workbook()
    wb.add_worksheet("Visible")
    ws = wb.add_worksheet("Secret")
    ws.set_very_hidden()
    assert ws.hidden is True
    assert reopen(save(wb, out))["Secret"].sheet_state == "veryHidden"
    ws.set_very_hidden(False)
    assert ws.hidden is False
    assert reopen(save(wb, out, "b.xlsx"))["Secret"].sheet_state == "visible"


def test_ignore_errors(out, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_string("A1", "123")
    ws.ignore_error("A1", px.IgnoreError.NumberStoredAsText)
    ws.ignore_error(1, 0, error_type=px.IgnoreError.NumberStoredAsText)
    ws.ignore_error_range("C1:D5", px.IgnoreError.FormulaDiffers)
    ws.ignore_error_range(9, 0, 9, 2, px.IgnoreError.TwoDigitTextYear)
    path = save(wb, out)
    assert_valid(path)
    xml = part(path, SHEET1)
    assert re.search(r'<ignoredError sqref="A1 A2" numberStoredAsText="1"/>', xml)
    assert 'sqref="C1:D5" formula="1"' in xml
    assert 'sqref="A10:C10" twoDigitTextYear="1"' in xml
    with pytest.raises(TypeError):
        ws.ignore_error("A1", "NumberStoredAsText")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="missing required"):
        ws.ignore_error("A1")  # pyrefly: ignore[no-matching-overload]


def test_ignore_error_enum():
    members = [
        "NumberStoredAsText",
        "FormulaError",
        "FormulaDiffers",
        "FormulaRefersToEmptyCells",
        "FormulaOmitsCells",
        "DataValidationError",
        "TwoDigitTextYear",
        "UnlockedCellsWithFormula",
        "InconsistentColumnFormula",
    ]
    values = [getattr(px.IgnoreError, m) for m in members]
    assert len(set(values)) == len(members)
    assert [int(v) for v in values] == list(range(len(members)))


# ---------------------------------------------------------------------------
# Printing
# ---------------------------------------------------------------------------


def test_margins(out, reopen):
    wb = px.Workbook()
    a = wb.add_worksheet("A")
    a.set_margins(left=1.0, right=1.25, top=1.5, bottom=1.75, header=0.5, footer=0.6)
    b = wb.add_worksheet("B")
    b.set_margins(top=2.0)
    book = reopen(save(wb, out))
    m = book["A"].page_margins
    assert (m.left, m.right, m.top, m.bottom, m.header, m.footer) == (1.0, 1.25, 1.5, 1.75, 0.5, 0.6)
    m = book["B"].page_margins
    assert (m.left, m.right, m.top, m.bottom, m.header, m.footer) == (0.7, 0.7, 2.0, 0.75, 0.3, 0.3)
    with pytest.raises(px.ParameterError, match="margin left"):
        a.set_margins(left=-1)
    with pytest.raises(ValueError, match="margin footer"):
        a.set_margins(footer=float("nan"))


def test_print_area_and_titles(out, reopen):
    wb = px.Workbook()
    a = wb.add_worksheet("A")
    a.set_print_area("A1:H20")
    a.set_repeat_rows(0, 1)
    a.set_repeat_columns("A:B")
    b = wb.add_worksheet("B")
    b.set_print_area(1, 1, 9, "D")
    b.set_repeat_rows(2)
    b.set_repeat_columns(2)
    c = wb.add_worksheet("C")
    c.set_repeat_columns(1, 3)
    d = wb.add_worksheet("D")
    d.set_repeat_columns("E")
    book = reopen(save(wb, out))
    assert book["A"].print_area == "'A'!$A$1:$H$20"
    assert book["A"].print_title_rows == "$1:$2"
    assert book["A"].print_title_cols == "$A:$B"
    assert book["B"].print_area == "'B'!$B$2:$D$10"
    assert book["B"].print_title_rows == "$3:$3"
    assert book["B"].print_title_cols == "$C:$C"
    assert book["C"].print_title_cols == "$B:$D"
    assert book["D"].print_title_cols == "$E:$E"


def test_print_area_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.RangeError):
        ws.set_print_area("D4:A1")
    with pytest.raises(px.ParameterError):
        ws.set_print_area("nonsense")
    with pytest.raises(TypeError, match="at most 0"):
        ws.set_print_area("A1:B2", "extra")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.RangeError):
        ws.set_repeat_rows(5, 2)
    with pytest.raises(px.RangeError):
        ws.set_repeat_rows(-1)
    with pytest.raises(px.RangeError):
        ws.set_repeat_rows(0, 1_048_576)
    with pytest.raises(px.RangeError):
        ws.set_repeat_columns(3, 1)
    with pytest.raises(px.ParameterError):
        ws.set_repeat_columns("1:2")


def test_page_breaks(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_page_breaks([40, 20, 20, 0])  # sorted, deduplicated, row 0 dropped
    ws.set_vertical_page_breaks([5])
    ws.set_page_breaks([])  # ignored: keeps the previous breaks
    ws.set_page_breaks([0])  # only row 0: ignored (rust_xlsxwriter would panic)
    sh = reopen(save(wb, out))["Sheet1"]
    assert [b.id for b in sh.row_breaks.brk] == [20, 40]
    assert [b.id for b in sh.col_breaks.brk] == [5]
    with pytest.raises(px.RangeError):
        ws.set_page_breaks([1_048_576])
    with pytest.raises(px.RangeError):
        ws.set_vertical_page_breaks([-1])
    with pytest.raises(px.RangeError):
        ws.set_vertical_page_breaks([16_384])
    with pytest.raises(px.ParameterError, match="1023"):
        ws.set_page_breaks(list(range(1, 1100)))


def test_print_options(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_print_scale(75)
    ws.set_print_first_page_number(1)
    ws.set_page_order(False)
    ws.set_print_black_and_white()
    ws.set_print_draft()
    ws.set_header_footer_scale_with_doc(False)
    ws.set_header_footer_align_with_page(False)
    ws.header = "&CTitle"
    sh = reopen(save(wb, out))["Sheet1"]
    ps = sh.page_setup
    assert ps.scale == 75
    assert ps.useFirstPageNumber
    assert ps.pageOrder == "overThenDown"
    assert ps.blackAndWhite and ps.draft
    assert sh.HeaderFooter.scaleWithDoc is False
    assert sh.HeaderFooter.alignWithMargins is False


def test_print_options_reset(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_print_black_and_white()
    ws.set_print_black_and_white(False)
    ws.set_print_draft()
    ws.set_print_draft(False)
    ws.set_page_order(True)
    ps = reopen(save(wb, out))["Sheet1"].page_setup
    assert not ps.blackAndWhite and not ps.draft
    assert ps.pageOrder in (None, "downThenOver")


def test_print_scale_errors():
    ws = px.Workbook().add_worksheet()
    # rust_xlsxwriter 0.103 writes other first page numbers as invalid XML.
    with pytest.raises(px.ParameterError, match="not supported"):
        ws.set_print_first_page_number(3)
    with pytest.raises(px.ParameterError, match=r"10\.\.=400"):
        ws.set_print_scale(5)
    with pytest.raises(px.ParameterError):
        ws.set_print_scale(401)


# ---------------------------------------------------------------------------
# Protection
# ---------------------------------------------------------------------------

PROTECTION_FIELDS = [
    "select_locked_cells",
    "select_unlocked_cells",
    "format_cells",
    "format_columns",
    "format_rows",
    "insert_columns",
    "insert_rows",
    "insert_links",
    "delete_columns",
    "delete_rows",
    "sort",
    "use_autofilter",
    "use_pivot_tables",
    "edit_scenarios",
    "edit_objects",
    "contents",
]


def test_protection_options_class():
    opts = px.ProtectionOptions()
    defaults = opts.to_dict()
    assert list(defaults) == PROTECTION_FIELDS
    assert {k for k, v in defaults.items() if v} == {"select_locked_cells", "select_unlocked_cells", "contents"}
    assert repr(opts) == "ProtectionOptions()"
    full = px.ProtectionOptions(**dict.fromkeys(PROTECTION_FIELDS, True))
    assert all(getattr(full, f) for f in PROTECTION_FIELDS)
    for field in PROTECTION_FIELDS:
        o = px.ProtectionOptions()
        setattr(o, field, not getattr(o, field))
        assert getattr(o, field) is not defaults[field]
        assert o != opts
    o = px.ProtectionOptions().set(sort=True, insert_rows=True)
    assert o.sort and o.insert_rows
    assert repr(o) == "ProtectionOptions(insert_rows=True, sort=True)"
    assert o == px.ProtectionOptions(sort=True, insert_rows=True)
    assert o.copy() == o and copy.copy(o) == o and copy.deepcopy(o) == o
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.ProtectionOptions(objects=True)  # pyrefly: ignore[unexpected-keyword]


def test_protect_plain(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.protect()
    p = reopen(save(wb, out))["Sheet1"].protection
    assert p.sheet and p.password is None
    assert p.formatCells and p.insertRows and p.sort  # protected = not allowed


def test_protect_with_password_and_options(out, reopen):
    wb = px.Workbook()
    a = wb.add_worksheet("A")
    a.protect(
        password="secret", options=px.ProtectionOptions(format_cells=True, insert_rows=True, select_locked_cells=False)
    )
    b = wb.add_worksheet("B")
    b.protect_with_password("secret")
    c = wb.add_worksheet("C")
    c.protect_with_options(px.ProtectionOptions(sort=True, use_autofilter=True, edit_objects=True))
    book = reopen(save(wb, out))
    pa = book["A"].protection
    assert pa.sheet and pa.password == excel_hash("secret") == book["B"].protection.password
    assert not pa.formatCells and not pa.insertRows and pa.selectLockedCells
    assert pa.formatColumns and pa.deleteRows
    pc = book["C"].protection
    assert pc.sheet and pc.password is None
    assert not pc.sort and not pc.autoFilter and not pc.objects


def test_unprotect_ranges(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.protect()
    ws.unprotect_range("A1:B5")
    ws.unprotect_range(9, 0, 9, 3, title="Inputs", password="pw")
    ws.unprotect_range("D1", None, None)
    ws.unprotect_range_with_options("F1:F3", "Named", "pw")
    path = save(wb, out)
    assert reopen(path)["Sheet1"].protection.sheet
    xml = part(path, SHEET1)
    pw = excel_hash("pw")
    assert re.findall(r"<protectedRange [^>]*/>", xml) == [
        '<protectedRange sqref="A1:B5" name="Range1"/>',
        f'<protectedRange password="{pw}" sqref="A10:D10" name="Inputs"/>',
        '<protectedRange sqref="D1" name="Range3"/>',
        f'<protectedRange password="{pw}" sqref="F1:F3" name="Named"/>',
    ]
    with pytest.raises(px.RangeError):
        ws.unprotect_range("C3:A1")
    with pytest.raises(TypeError, match="missing required"):
        ws.unprotect_range_with_options("A1:B2", "x")  # pyrefly: ignore[no-matching-overload]


# ---------------------------------------------------------------------------
# Argument handling
# ---------------------------------------------------------------------------


def test_cell_argument_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.RangeError):
        ws.set_top_left_cell(1_048_576, 0)
    with pytest.raises(px.ParameterError):
        ws.set_top_left_cell("A0")
    with pytest.raises(TypeError, match="at most 0"):
        ws.set_top_left_cell("A1", 3)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="missing"):
        ws.set_selection()  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError, match="missing"):
        ws.set_selection(1, 2, 3)  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError):
        ws.set_selection("ZZZZ1")
    with pytest.raises(TypeError, match="at most 0"):
        ws.set_freeze_panes_top_cell(1, 1, 1)  # pyrefly: ignore[no-matching-overload]


def test_settings_survive_resave(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_print_area("A1:C3")
    ws.set_repeat_rows(0)
    ws.protect(password="secret")
    for name in ("a.xlsx", "b.xlsx"):
        sh = reopen(save(wb, out, name))["Sheet1"]
        assert sh.print_area == "'Sheet1'!$A$1:$C$3"
        assert sh.print_title_rows == "$1:$1"
        assert sh.protection.sheet and sh.protection.password == excel_hash("secret")
