"""The in-place editor: fidelity, cell edits, styles, structure, and errors."""

import datetime as dt
import io
import zipfile
from pathlib import Path

import openpyxl
import pytest
from ooxml_check import assert_valid

import pyrust_xlsxwriter as px

CORPUS = Path(__file__).parent / "corpus"
FILES = sorted(p for p in CORPUS.glob("*.xls*"))


def parts(data: bytes | Path) -> dict[str, bytes]:
    src = io.BytesIO(data) if isinstance(data, bytes) else data
    with zipfile.ZipFile(src) as zf:
        return {n: zf.read(n) for n in zf.namelist()}


def make_book(tmp_path: Path, build) -> Path:
    """Create a source workbook with the writer."""
    path = tmp_path / "source.xlsx"
    wb = px.Workbook()
    build(wb)
    wb.save(path)
    return path


def patch_part(path: Path, name: str, transform) -> None:
    """Rewrite one XML part of a package in place (to craft test inputs)."""
    items = parts(path)
    items[name] = transform(items[name].decode()).encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, data in items.items():
            zf.writestr(n, data)


# ---------------------------------------------------------------------------
# Fidelity
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_load_save_without_edits_is_byte_identical(path):
    data = px.load_workbook(path).save_to_buffer()
    assert parts(data) == parts(path)


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_editing_a_cell_keeps_every_other_part(path):
    book = px.load_workbook(path)
    target = next((ws for ws in book.worksheets if _is_worksheet(ws)), None)
    if target is None:
        pytest.skip("no worksheet")
    target.write("ZZ1000", "edited")
    data = book.save_to_buffer()
    assert_valid(data)
    before, after = parts(path), parts(data)
    assert set(before) - {"xl/calcChain.xml"} <= set(after)
    changed = {n for n in before if n in after and before[n] != after[n]}
    allowed = {
        "xl/sharedStrings.xml",
        "xl/workbook.xml",
        "[Content_Types].xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
    }
    assert changed - allowed == {_sheet_part(path, target.name)}
    reread = openpyxl.load_workbook(io.BytesIO(data))
    assert reread[target.name]["ZZ1000"].value == "edited"


def _is_worksheet(ws) -> bool:
    try:
        ws.dimensions  # noqa: B018 - raises for chartsheets
        return True
    except px.EditError:
        return False


def _sheet_part(path: Path, name: str) -> str:
    wb = openpyxl.load_workbook(path)
    index = wb.sheetnames.index(name)
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
        workbook = zf.read("xl/workbook.xml").decode()
    import re

    rids = re.findall(r'<sheet [^>]*r:id="(rId\d+)"', workbook)
    target = re.search(rf'Id="{rids[index]}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rids[index]}"', rels)
    assert target
    t = target.group(1) or target.group(2)
    return t.lstrip("/") if t.startswith("/") else "xl/" + t


def test_chart_and_vba_survive_edits(tmp_path):
    src = CORPUS / "chart_bar14.xlsx"
    book = px.load_workbook(src)
    book[0].write("A1", 999)
    out = tmp_path / "out.xlsx"
    book.save(out)
    before, after = parts(src), parts(out)
    for name in before:
        if "/charts/" in name or "/drawings/" in name:
            assert after[name] == before[name]
    macro = next(CORPUS.glob("macro*.xlsm"))
    book = px.load_workbook(macro)
    book[0].write("B2", "x")
    data = book.save_to_buffer()
    assert parts(data)["xl/vbaProject.bin"] == parts(macro)["xl/vbaProject.bin"]
    assert_valid(data)


# ---------------------------------------------------------------------------
# Values
# ---------------------------------------------------------------------------


def test_write_and_read_values(tmp_path):
    src = make_book(tmp_path, lambda wb: wb.add_worksheet("Data").write_row("A1", ["name", 1, 2.5, True]))
    book = px.load_workbook(src)
    ws = book["Data"]
    assert ws.values() == [["name", 1, 2.5, True]]
    ws.write("A2", "new")
    ws.write(1, 1, 42)
    ws.write("C2", dt.date(2024, 5, 6))
    ws.write("D2", dt.datetime(2024, 5, 6, 7, 8, 9))
    ws.write("E2", dt.time(12, 30))
    ws.write(
        "F2",
        "=B1+B2",
    )
    ws.write_formula("G2", "=B2*2", value=84)
    ws.write("H2", False)
    ws.write("A1", None)
    ws.write_rows("A4", [[1, 2], [3, 4]])
    ws.write_column("J1", ["a", "b"])
    assert ws.cell("B2") == 42
    assert ws.cell("C2") == dt.datetime(2024, 5, 6)
    assert ws.cell("E2") == dt.time(12, 30)
    assert ws.cell("A1") is None
    assert ws.formula("G2") == "=B2*2"
    assert ws.dimensions == (0, 0, 4, 9)
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    sh = openpyxl.load_workbook(out)["Data"]
    assert sh["A1"].value is None
    assert sh["A2"].value == "new" and sh["B2"].value == 42
    assert sh["C2"].value == dt.datetime(2024, 5, 6) and sh["C2"].number_format == "yyyy-mm-dd"
    assert sh["D2"].value == dt.datetime(2024, 5, 6, 7, 8, 9)
    assert sh["F2"].value == "=B1+B2" and sh["G2"].value == "=B2*2"
    assert sh["H2"].value is False
    assert sh["B5"].value == 4 and sh["J2"].value == "b"
    with zipfile.ZipFile(out) as zf:
        assert 'fullCalcOnLoad="1"' in zf.read("xl/workbook.xml").decode()
    # The editor re-reads its own output.
    again = px.load_workbook(out)["Data"]
    assert again.cell("A2") == "new" and again.cell("H2") is False


def test_writing_keeps_the_cell_style(tmp_path):
    def build(wb):
        ws = wb.add_worksheet()
        ws.write("A1", 1, px.Format(bold=True, num_format="0.00", background_color="#FFFF00"))
        ws.write("A2", dt.date(2020, 1, 1), px.Format(num_format="dd/mm/yyyy"))

    book = px.load_workbook(make_book(tmp_path, build))
    ws = book[0]
    ws.write("A1", 2)
    ws.write("A2", dt.date(2030, 12, 31))
    style = ws.style("A1")
    assert style is not None and style.bold and style.num_format == "0.00"
    out = tmp_path / "out.xlsx"
    book.save(out)
    sh = openpyxl.load_workbook(out).active
    assert sh["A1"].value == 2 and sh["A1"].font.b and sh["A1"].number_format == "0.00"
    assert sh["A1"].fill.fgColor.rgb == "FFFFFF00"
    assert sh["A2"].number_format == "dd/mm/yyyy" and sh["A2"].value == dt.datetime(2030, 12, 31)


def test_explicit_and_merged_formats(tmp_path):
    def build(wb):
        ws = wb.add_worksheet()
        ws.write("A1", "x", px.Format(italic=True, font_color="#0000FF"))
        ws.write("A2", "y", px.Format(italic=True))

    book = px.load_workbook(make_book(tmp_path, build))
    ws = book[0]
    ws.set_cell_format("A1", px.Format(bold=True), merge=True)  # keeps italic + colour
    ws.write("A2", "z", px.Format(bold=True))  # replaces the style
    ws.set_range_format("B1:B3", px.Format(border=px.FormatBorder.Thin, background_color="#DDEBF7"))
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    sh = openpyxl.load_workbook(out).active
    assert sh["A1"].font.b and sh["A1"].font.i and sh["A1"].font.color.rgb == "FF0000FF"
    assert sh["A2"].font.b and not sh["A2"].font.i
    assert sh["B2"].border.left.style == "thin" and sh["B3"].fill.fgColor.rgb == "FFDDEBF7"
    # Repeating the same edit doesn't add new styles.
    with zipfile.ZipFile(out) as zf:
        styles = zf.read("xl/styles.xml").decode()
    assert styles.count("<xf ") < 12


def test_style_merge_preserves_unknown_font_bits(tmp_path):
    src = make_book(tmp_path, lambda wb: wb.add_worksheet().write("A1", "x", px.Format(font_name="Arial")))
    patch_part(src, "xl/styles.xml", lambda x: x.replace('<name val="Arial"/>', '<outline/><name val="Arial"/>', 1))
    book = px.load_workbook(src)
    book[0].set_cell_format("A1", px.Format(bold=True), merge=True)
    with zipfile.ZipFile(io.BytesIO(book.save_to_buffer())) as zf:
        styles = zf.read("xl/styles.xml").decode()
    assert "<b/><outline/>" in styles or ("<b/>" in styles and "<outline/>" in styles)


# ---------------------------------------------------------------------------
# Formulas
# ---------------------------------------------------------------------------


def test_shared_formula_master_edit_expands_dependents(tmp_path):
    def build(wb):
        ws = wb.add_worksheet()
        ws.write_column("A1", [1, 2, 3])
        ws.write_formula("B1", "=A1*2", value=2)
        ws.write_formula("B2", "=A2*2", value=4)
        ws.write_formula("B3", "=A3*2", value=6)

    src = make_book(tmp_path, build)

    def to_shared(xml: str) -> str:
        xml = xml.replace("<f>A1*2</f>", '<f t="shared" ref="B1:B3" si="0">A1*2</f>')
        xml = xml.replace("<f>A2*2</f>", '<f t="shared" si="0"/>').replace("<f>A3*2</f>", '<f t="shared" si="0"/>')
        return xml

    patch_part(src, "xl/worksheets/sheet1.xml", to_shared)
    book = px.load_workbook(src)
    book[0].write("B1", 100)
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    sh = openpyxl.load_workbook(out).active
    assert sh["B1"].value == 100
    assert sh["B2"].value == "=A2*2" and sh["B3"].value == "=A3*2"


def test_array_formula_rules(tmp_path):
    book = px.load_workbook(CORPUS / "array_formula01.xlsx")
    ws = book[0]
    ws.write_array_formula("F1:F3", "{=SUM(B1:C1*B2:C2)}")
    data = book.save_to_buffer()
    assert_valid(data)
    names = parts(data)
    assert "xl/calcChain.xml" not in names
    assert 't="array" ref="F1:F3"' in names["xl/worksheets/sheet1.xml"].decode()


# ---------------------------------------------------------------------------
# Sheet layout
# ---------------------------------------------------------------------------


def test_merges_columns_rows_panes(tmp_path):
    src = make_book(tmp_path, lambda wb: wb.add_worksheet().merge_range("A1:B1", "m"))
    book = px.load_workbook(src)
    ws = book[0]
    assert ws.merged_ranges == [(0, 0, 0, 1)]
    with pytest.raises(px.OverlapError):
        ws.merge_range("B1:C2", "x")
    ws.unmerge_range("A1:B1")
    ws.merge_range("A3:D3", "Title", px.Format(bold=True, align=px.FormatAlign.Center))
    ws.set_column_width("B:C", 25)
    ws.set_column_hidden("E")
    ws.set_row_height(4, 30)
    ws.set_row_hidden(5)
    ws.freeze_panes("A2")
    ws.tab_color = "#FF0000"
    ws.zoom = 150
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    sh = openpyxl.load_workbook(out).active
    assert {str(r) for r in sh.merged_cells.ranges} == {"A3:D3"}
    assert sh["A3"].value == "Title" and sh["A3"].font.b
    assert sh.column_dimensions["B"].width == 25 and sh.column_dimensions["E"].hidden
    assert sh.row_dimensions[5].height == 30 and sh.row_dimensions[6].hidden
    assert sh.freeze_panes == "A2"
    assert sh.sheet_properties.tabColor.rgb == "FFFF0000"
    assert sh.sheet_view.zoomScale == 150


def test_hyperlinks(tmp_path):
    book = px.load_workbook(make_book(tmp_path, lambda wb: wb.add_worksheet("S")))
    ws = book["S"]
    ws.write_url("A1", "https://example.com", text="Example", tip="Go")
    ws.write_url("A2", "internal:S!B5")
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    sh = openpyxl.load_workbook(out)["S"]
    assert sh["A1"].value == "Example" and sh["A1"].hyperlink.target == "https://example.com"
    assert sh["A1"].font.u == "single"
    assert sh["A2"].hyperlink.location == "S!B5"


def test_table_header_edit_updates_table(tmp_path):
    def build(wb):
        ws = wb.add_worksheet()
        ws.write_rows("A1", [["Name", "Qty"], ["a", 1], ["b", 2]])

    src = make_book(tmp_path, build)
    # Add a table part by hand (the writer's table API lives in another package).
    table = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" '
        'displayName="Table1" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2">'
        '<tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/></tableColumns>'
        '<tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/></table>'
    )
    items = parts(src)
    items["xl/tables/table1.xml"] = table.encode()
    items["xl/worksheets/_rels/sheet1.xml.rels"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
        b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
        b'Target="../tables/table1.xml"/></Relationships>'
    )
    items["xl/worksheets/sheet1.xml"] = items["xl/worksheets/sheet1.xml"].replace(
        b"</worksheet>", b'<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
    )
    assert b"xmlns:r=" in items["xl/worksheets/sheet1.xml"]
    items["[Content_Types].xml"] = items["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
    )
    with zipfile.ZipFile(src, "w") as zf:
        for n, d in items.items():
            zf.writestr(n, d)
    book = px.load_workbook(src)
    book[0].write("B1", "Quantity")
    book[0].write("A1", 2024)
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    t = parts(out)["xl/tables/table1.xml"].decode()
    assert 'name="Quantity"' in t and 'name="2024"' in t
    assert openpyxl.load_workbook(out).active["A1"].value == "2024"


# ---------------------------------------------------------------------------
# Workbook structure
# ---------------------------------------------------------------------------


def test_sheets_add_rename_move_remove(tmp_path):
    def build(wb):
        wb.add_worksheet("One").write("A1", 1)
        wb.add_worksheet("Two").write("A1", 2)

    src = make_book(tmp_path, build)
    book = px.load_workbook(src)
    book.define_name("Total", "=One!$A$1+Two!$A$1")
    book.define_name("Local", "=Two!$A$1", sheet="Two")
    new = book.add_worksheet("New", index=0)
    new.write_row("A1", ["fresh", 3])
    book["One"].name = "First"
    assert ("Total", "First!$A$1+Two!$A$1", None) in book.defined_names
    book.move_worksheet("Two", 0)
    assert book.sheet_names == ["Two", "New", "First"]
    assert ("Local", "Two!$A$1", "Two") in book.defined_names
    book["New"].state = "hidden"
    book.active = 2
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert_valid(out)
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Two", "New", "First"]
    assert wb["New"]["A1"].value == "fresh" and wb["New"].sheet_state == "hidden"
    assert wb["First"]["A1"].value == 1
    assert wb.active.title == "First"
    book.remove_worksheet("Two")
    assert ("Local", "Two!$A$1", "Two") not in book.defined_names
    data = book.save_to_buffer()
    assert_valid(data)
    assert openpyxl.load_workbook(io.BytesIO(data)).sheetnames == ["New", "First"]


def test_sheet_errors(tmp_path):
    book = px.load_workbook(make_book(tmp_path, lambda wb: wb.add_worksheet("Only")))
    with pytest.raises(px.SheetNameError):
        book["Only"].name = "bad/name"
    with pytest.raises(px.SheetNameError):
        book.add_worksheet("only")
    with pytest.raises(px.EditError):
        book.remove_worksheet("Only")
    with pytest.raises(px.EditError):
        book["Only"].hidden = True
    with pytest.raises(KeyError):
        book["missing"]
    removed = book.add_worksheet("Tmp")
    book.remove_worksheet(removed)
    with pytest.raises(px.EditError, match="removed"):
        removed.write("A1", 1)


def test_properties(tmp_path):
    book = px.load_workbook(make_book(tmp_path, lambda wb: wb.add_worksheet()))
    book.set_properties(title="Report", author="Me", created=dt.datetime(2024, 1, 2, 3, 4, 5))
    assert book.properties["title"] == "Report"
    data = book.save_to_buffer()
    props = openpyxl.load_workbook(io.BytesIO(data)).properties
    assert props.title == "Report" and props.creator == "Me"
    with pytest.raises(TypeError):
        book.set_properties(colour="red")  # pyrefly: ignore[unexpected-keyword]


# ---------------------------------------------------------------------------
# Lifecycle
# ---------------------------------------------------------------------------


def test_sources_saves_and_context_manager(tmp_path):
    src = make_book(tmp_path, lambda wb: wb.add_worksheet().write("A1", "v"))
    raw = src.read_bytes()
    assert px.load_workbook(raw)[0].cell("A1") == "v"
    assert px.load_workbook(io.BytesIO(raw))[0].cell(0, 0) == "v"
    book = px.load_workbook(str(src))
    book[0].write("A2", "w")
    first, second = book.save_to_buffer(), book.save_to_buffer()
    assert parts(first) == parts(second)  # saving is repeatable
    with px.load_workbook(src) as book:
        book[0].write("A3", "ctx")
    assert openpyxl.load_workbook(src).active["A3"].value == "ctx"
    with pytest.raises(TypeError):
        px.load_workbook(raw).save()


def test_errors_reading():
    with pytest.raises(px.FileError):
        px.load_workbook("/does/not/exist.xlsx")
    with pytest.raises(px.ReadError):
        px.load_workbook(b"not a zip")


def test_many_edits_are_fast_and_valid(tmp_path):
    def build(wb):
        ws = wb.add_worksheet()
        ws.write_rows("A1", [[r, r * 2, f"row {r}"] for r in range(20_000)])

    book = px.load_workbook(make_book(tmp_path, build))
    ws = book[0]
    ws.write_rows("D1", [[r % 7] for r in range(20_000)])
    import time

    start = time.perf_counter()
    data = book.save_to_buffer()
    assert time.perf_counter() - start < 5
    assert_valid(data)
    sh = openpyxl.load_workbook(io.BytesIO(data), read_only=True).active
    rows = list(sh.iter_rows(min_row=19_999, max_row=20_000, values_only=True))
    assert rows[-1][:4] == (19_999, 39_998, "row 19999", 19_999 % 7)


def test_writing_inside_a_merged_range_is_rejected(tmp_path):
    book = px.load_workbook(make_book(tmp_path, lambda wb: wb.add_worksheet().merge_range("B2:D4", "title")))
    ws = book[0]
    ws.write("B2", "new title")  # the top-left cell is fine
    with pytest.raises(px.ParameterError, match="merged range B2:D4"):
        ws.write("C3", "hidden")
    ws.unmerge_range("B2:D4")
    ws.write("C3", "now visible")
    sh = openpyxl.load_workbook(io.BytesIO(book.save_to_buffer())).active
    assert sh["B2"].value == "new title" and sh["C3"].value == "now visible"


def test_iter_rows_reads_rows_as_it_goes(tmp_path: Path):
    path = make_book(tmp_path, lambda wb: wb.add_worksheet("S").write_column(0, 0, [1, 2, 3]))
    sheet = px.load_workbook(path)["S"]
    rows = sheet.iter_rows(max_col=0)
    assert next(rows) == (1,)
    sheet.write(1, 0, "changed")  # later rows see edits made while iterating
    assert list(rows) == [("changed",), (3,)]
    assert list(sheet.iter_rows(min_row=5)) == []
