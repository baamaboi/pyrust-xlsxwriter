"""Reader: values, formulas, styles, sources and errors, round-tripped
through files written by px.Workbook and checked against openpyxl."""

from __future__ import annotations

import datetime as dt
import decimal
import io
import threading
import warnings
import zipfile
from pathlib import Path

import openpyxl
import pytest

import pyrust_xlsxwriter as px

F = px.Format


def build(path: Path | None = None) -> bytes:
    """A workbook with every value type, styles, merges, links and a hidden sheet."""
    wb = px.Workbook()
    ws = wb.add_worksheet("Values")
    ws.write_row("A1", ["text", 42, 1.25, True, None, -7])
    ws.write("A2", dt.date(2024, 1, 2))
    ws.write("B2", dt.datetime(2024, 1, 2, 3, 4, 5, 250_000))
    ws.write("C2", dt.time(6, 30))
    ws.write("D2", dt.timedelta(hours=36, minutes=15))
    ws.write("E2", decimal.Decimal("2.5"))
    ws.write("F2", False)
    ws.write("A3", "=B1*2")
    ws.write_formula("B3", "=SUM(B1:C1)", value=43.25)
    ws.write_formula("C3", "=1/0", value="#DIV/0!")
    ws.write("D3", 2**60)
    ws.write("E3", "naïve ☃ text")
    ws.write("F3", 1e-7)
    ws.write("H5", "far corner")
    ws.merge_range("A7:C8", "merged", F(align=px.FormatAlign.Center))
    ws.write_url("A10", "https://example.com/x", text="site", tip="Go")
    ws.write_url("B10", "internal:'Other Sheet'!B2", text="jump")

    st = wb.add_worksheet("Styles")
    st.write("A1", "bold", F(bold=True))
    st.write("B1", 1.5, F(num_format="0.00", background_color="#FFC7CE", font_color="#9C0006"))
    st.write("C1", "box", F(border=px.FormatBorder.Thin, border_color="#0000FF"))
    st.write(
        "D1",
        "aligned",
        F(align=px.FormatAlign.Right, valign=px.FormatAlign.Top, text_wrap=True, rotation=-45, indent=2),
    )
    st.write(
        "E1",
        "fonty",
        F(italic=True, underline=px.FormatUnderline.Double, strikethrough=True, font_name="Arial", font_size=14),
    )
    st.write(
        "F1", "pattern", F(pattern=px.FormatPattern.DarkGrid, foreground_color="#00FF00", background_color="#FFFF00")
    )
    st.write(
        "G1",
        "prot",
        F(
            locked=False,
            hidden=True,
            border_diagonal=px.FormatBorder.Dashed,
            border_diagonal_type=px.FormatDiagonalBorder.BorderUpDown,
            border_diagonal_color="#FF0000",
        ),
    )
    st.write("H1", "plain")
    st.write("I1", "shrink", F(shrink=True, align=px.FormatAlign.CenterAcross, border_bottom=px.FormatBorder.Double))

    wb.add_worksheet("Other Sheet", hidden=True).write("B2", "target")
    wb.add_worksheet("Empty")
    data = wb.save_to_buffer()
    if path is not None:
        path.write_bytes(data)
    return data


@pytest.fixture(scope="module")
def data() -> bytes:
    return build()


@pytest.fixture
def book(data: bytes) -> px.ReadWorkbook:
    return px.open_workbook(data)


@pytest.fixture
def xlsx(tmp_path: Path, data: bytes) -> Path:
    path = tmp_path / "read.xlsx"
    path.write_bytes(data)
    return path


def set_1904(data: bytes) -> bytes:
    """Switch a package to the 1904 date system (the serials stay the same)."""
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text = content.decode()
                assert "<workbookPr" in text
                content = text.replace("<workbookPr", '<workbookPr date1904="1"', 1).encode()
            dst.writestr(item, content)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Workbook-level information
# ---------------------------------------------------------------------------


def test_workbook_metadata(book: px.ReadWorkbook):
    assert book.format == "xlsx"
    assert book.sheet_names == ["Values", "Styles", "Other Sheet", "Empty"]
    assert [s.visibility for s in book.sheets] == ["visible", "visible", "hidden", "visible"]
    assert all(s.kind == "worksheet" for s in book.sheets)
    assert [s.index for s in book.sheets] == [0, 1, 2, 3]
    assert book.is_1904 is False
    assert len(book) == 4
    assert "Styles" in book and "styles" not in book and 3 not in book
    assert [ws.name for ws in book] == book.sheet_names
    assert book.tables == [] and book.pictures == [] and book.vba_modules is None
    assert "Values" in repr(book) and "SheetInfo" in repr(book.sheets[0])


def test_sheet_lookup(book: px.ReadWorkbook):
    ws = book.worksheet("Styles")
    assert ws is book["Styles"] is book[1] is book[-3]  # cached
    assert book["STYLES"] is ws  # case-insensitive fallback, as in Excel
    assert ws.name == "Styles" and ws.index == 1 and ws.kind == "worksheet"
    assert book["Other Sheet"].visibility == "hidden"
    with pytest.raises(KeyError, match="no sheet named"):
        book["Nope"]
    with pytest.raises(IndexError):
        book[4]
    with pytest.raises(IndexError):
        book.worksheet(-5)
    with pytest.raises(TypeError):
        book[1.5]  # pyrefly: ignore[bad-index]
    with pytest.raises(TypeError):
        book[True]


# ---------------------------------------------------------------------------
# Values
# ---------------------------------------------------------------------------


def test_value_types(book: px.ReadWorkbook):
    ws = book["Values"]
    assert ws.dimensions == (0, 0, 9, 7)
    assert (ws.max_row, ws.max_column) == (10, 8)
    rows = ws.values()
    assert len(rows) == 10 and all(len(r) == 8 for r in rows)
    assert rows[0][:6] == ["text", 42, 1.25, True, None, -7]
    assert type(rows[0][1]) is int and type(rows[0][2]) is float and rows[0][3] is True
    assert rows[1][0] == dt.datetime(2024, 1, 2)
    assert rows[1][1] == dt.datetime(2024, 1, 2, 3, 4, 5, 250_000)
    assert rows[1][2] == dt.time(6, 30)
    assert rows[1][3] == dt.timedelta(hours=36, minutes=15)
    assert rows[1][4] == 2.5 and rows[1][5] is False
    assert rows[2][0] == 0  # formula without a cached result
    assert rows[2][1] == 43.25
    assert rows[2][2] == "#DIV/0!"
    assert rows[2][3] == float(2**60) and type(rows[2][3]) is float  # beyond 2**53 stays a float
    assert rows[2][4] == "naïve ☃ text"
    assert rows[2][5] == 1e-7
    assert rows[4][7] == "far corner"
    assert rows[6][0] == "merged" and rows[6][1] is None
    assert rows[3] == [None] * 8


def test_values_match_openpyxl(xlsx: Path):
    want = openpyxl.load_workbook(xlsx, data_only=True)
    book = px.open_workbook(xlsx)
    for name in ["Values", "Styles", "Other Sheet"]:
        expected = [list(r) for r in want[name].iter_rows(values_only=True)]
        got = book[name].values()
        assert got == expected, name
        # rust_xlsxwriter stores 42 as "42.0", which openpyxl reads as a float;
        # calamine cannot see the difference and integral numbers come back as
        # int (as openpyxl does for Excel's own "42"). Other types must match.
        kind = lambda v: float if type(v) is int else type(v)  # noqa: E731
        assert [[kind(v) for v in r] for r in got] == [[kind(v) for v in r] for r in expected]
    assert book["Empty"].values() == []


def test_absolute_coordinates(tmp_path: Path):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write("C4", "x")
    ws.write("E6", 5)
    path = tmp_path / "offset.xlsx"
    wb.save(path)
    sheet = px.open_workbook(path)[0]
    assert sheet.dimensions == (3, 2, 5, 4)
    values = sheet.values()
    assert len(values) == 6 and len(values[0]) == 5
    assert values[3][2] == "x" and values[5][4] == 5
    assert values[0] == [None] * 5
    assert sheet.cell(3, 2) == sheet.cell("C4") == sheet.cell(3, "C") == "x"
    assert sheet.cell(0, 0) is None and sheet.cell(1000, 1000) is None


def test_cell_arguments(book: px.ReadWorkbook):
    ws = book["Values"]
    assert ws.cell("A1") == "text"
    assert ws.cell(0, 1) == 42
    with pytest.raises(px.ParameterError):
        ws.cell("not a cell")
    with pytest.raises(TypeError):
        ws.cell(0, 1, 2)  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError):
        ws.cell()  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.RangeError):
        ws.cell(-1, 0)


def test_dates_1904(data: bytes):
    patched = set_1904(data)
    book = px.open_workbook(patched)
    assert book.is_1904 is True
    want = openpyxl.load_workbook(io.BytesIO(patched), data_only=True)["Values"]
    got = book["Values"].values()
    # Same serials, 1462 days later; times and durations do not move.
    assert got[1][0] == dt.datetime(2028, 1, 3) == want["A2"].value
    assert got[1][1] == dt.datetime(2028, 1, 3, 3, 4, 5, 250_000) == want["B2"].value
    assert got[1][2] == dt.time(6, 30) == want["C2"].value
    assert got[1][3] == dt.timedelta(hours=36, minutes=15) == want["D2"].value


def test_excel_1900_leap_bug(tmp_path: Path):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    date = F(num_format="yyyy-mm-dd")
    for row, serial in enumerate([1, 59, 60, 61, 0.75, -1.0, 3e6]):
        ws.write_number(row, 0, serial, date)
    path = tmp_path / "serials.xlsx"
    wb.save(path)
    got = [r[0] for r in px.open_workbook(path)[0].values()]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # openpyxl warns about the out-of-range serial
        want = [r[0] for r in openpyxl.load_workbook(path).active.iter_rows(values_only=True)]
    assert (
        got[:5]
        == want[:5]
        == [
            dt.datetime(1900, 1, 1),
            dt.datetime(1900, 2, 28),
            dt.datetime(1900, 2, 28),
            dt.datetime(1900, 3, 1),
            dt.time(18, 0),
        ]
    )
    assert got[5] == dt.datetime(1899, 12, 29)
    assert got[6] == 3e6  # beyond year 9999: left as a number


# ---------------------------------------------------------------------------
# iter_rows / records
# ---------------------------------------------------------------------------


def test_iter_rows(book: px.ReadWorkbook):
    ws = book["Values"]
    assert list(ws.iter_rows()) == [tuple(r) for r in ws.values()]
    assert list(ws.iter_rows(min_row=1, max_row=2, min_col=0, max_col=2)) == [
        (dt.datetime(2024, 1, 2), dt.datetime(2024, 1, 2, 3, 4, 5, 250_000), dt.time(6, 30)),
        (0, 43.25, "#DIV/0!"),
    ]
    # Explicit bounds beyond the data are padded with None.
    assert list(ws.iter_rows(min_row=4, max_row=5, min_col=7, max_col=9)) == [
        ("far corner", None, None),
        (None, None, None),
    ]
    assert list(ws.iter_rows(min_row=5, max_row=4)) == []
    cells = next(ws.iter_rows(max_row=0, values_only=False))
    assert [c.value for c in cells][:3] == ["text", 42, 1.25]
    assert (cells[1].row, cells[1].col, cells[1].coordinate) == (0, 1, "B1")
    assert "B1" in repr(cells[1])
    it = ws.iter_rows(max_row=0)
    assert iter(it) is it
    next(it)
    with pytest.raises(StopIteration):
        next(it)
    assert list(book["Empty"].iter_rows()) == []
    with pytest.raises(px.RangeError):
        ws.iter_rows(min_row=-1)
    with pytest.raises(px.RangeError):
        ws.iter_rows(max_col=20_000)


def test_records(tmp_path: Path):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_row("B2", ["name", "qty", None, "qty"])
    ws.write_row("B3", ["apple", 3, "x", 4])
    ws.write_row("B5", ["pear", 5])
    path = tmp_path / "records.xlsx"
    wb.save(path)
    sheet = px.open_workbook(path)[0]
    assert sheet.records(header_row=1) == [
        {"name": "apple", "qty": 3, "D": "x", "qty.1": 4},
        {"name": "pear", "qty": 5, "D": None, "qty.1": None},
    ]
    assert sheet.records(header_row=4) == []
    # An empty header row: every column is keyed by its letter.
    assert sheet.records()[0] == {"B": "name", "C": "qty", "D": None, "E": "qty"}
    with pytest.raises(px.RangeError):
        sheet.records(header_row=-1)


# ---------------------------------------------------------------------------
# Formulas, merged ranges, hyperlinks
# ---------------------------------------------------------------------------


def test_formulas(book: px.ReadWorkbook, xlsx: Path):
    ws = book["Values"]
    assert ws.formula("A3") == "=B1*2"
    assert ws.formula(2, 1) == "=SUM(B1:C1)"
    assert ws.formula("A1") is None and ws.formula(500, 500) is None
    formulas = ws.formulas()
    assert len(formulas) == 10 and all(len(r) == 8 for r in formulas)
    assert formulas[2][:3] == ["=B1*2", "=SUM(B1:C1)", "=1/0"]
    want = openpyxl.load_workbook(xlsx)["Values"]
    expected = [[c.value if c.data_type == "f" else None for c in row] for row in want.iter_rows()]
    assert formulas == expected
    assert book["Empty"].formulas() == []


def test_merged_and_hyperlinks(book: px.ReadWorkbook, xlsx: Path):
    ws = book["Values"]
    assert ws.merged_ranges == [(6, 0, 7, 2)]
    links = {h.reference: h for h in ws.hyperlinks}
    assert links["A10"].target == "https://example.com/x"
    assert links["A10"].tooltip == "Go" and links["A10"].cell == (9, 0)
    assert links["B10"].target is None and links["B10"].location == "'Other Sheet'!B2"
    want = openpyxl.load_workbook(xlsx)["Values"]
    assert {(h.reference, h.target, h.location) for h in ws.hyperlinks} == {
        (c.hyperlink.ref, c.hyperlink.target, c.hyperlink.location)
        for row in want.iter_rows()
        for c in row
        if c.hyperlink
    }
    assert book["Styles"].merged_ranges == [] and book["Styles"].hyperlinks == []


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------


def test_styles(book: px.ReadWorkbook):
    ws = book["Styles"]
    bold = ws.style("A1")
    assert isinstance(bold, px.Format) and bold.bold and not bold.italic
    assert bold.font_name == "Calibri" and bold.font_size == 11

    money = ws.style(0, 1)
    assert money is not None
    assert money.num_format == "0.00"
    assert money.pattern == px.FormatPattern.Solid
    assert money.background_color == px.Color("#FFC7CE")
    assert money.font_color == px.Color("#9C0006")

    box = ws.style("C1")
    assert box is not None
    assert box.border == px.FormatBorder.Thin
    assert box.border_color == px.Color("#0000FF")

    aligned = ws.style("D1")
    assert aligned is not None
    assert aligned.align == px.FormatAlign.Right
    assert aligned.valign == px.FormatAlign.Top
    assert aligned.text_wrap and aligned.rotation == -45 and aligned.indent == 2

    fonty = ws.style("E1")
    assert fonty is not None
    assert fonty.italic and fonty.strikethrough
    assert fonty.underline == px.FormatUnderline.Double
    assert fonty.font_name == "Arial" and fonty.font_size == 14

    pattern = ws.style("F1")
    assert pattern is not None
    assert pattern.pattern == px.FormatPattern.DarkGrid
    assert pattern.foreground_color == px.Color("#00FF00")
    assert pattern.background_color == px.Color("#FFFF00")

    prot = ws.style("G1")
    assert prot is not None
    assert prot.locked is False and prot.hidden is True
    assert prot.border_diagonal == px.FormatBorder.Dashed
    assert prot.border_diagonal_type == px.FormatDiagonalBorder.BorderUpDown
    assert prot.border_diagonal_color == px.Color("#FF0000")

    assert ws.style("H1") is None  # default format
    assert ws.style("Z99") is None  # outside the used range

    shrink = ws.style("I1")
    assert shrink is not None
    assert shrink.shrink and shrink.align == px.FormatAlign.CenterAcross
    assert shrink.border_bottom == px.FormatBorder.Double and shrink.border_top is None

    date_style = book["Values"].style("A2")
    assert date_style is not None and date_style.num_format == "yyyy-mm-dd"
    assert book["Values"].style("A1") is None


def test_styles_round_trip_through_writer(book: px.ReadWorkbook, tmp_path: Path):
    """A style read back can be written again and reads back the same."""
    src = book["Styles"]
    wb = px.Workbook()
    ws = wb.add_worksheet()
    originals = []
    for col in range(9):
        fmt = src.style(0, col)
        originals.append(fmt)
        ws.write(0, col, "x", fmt)
    path = tmp_path / "copy.xlsx"
    wb.save(path)
    again = px.open_workbook(path)[0]
    for col, fmt in enumerate(originals):
        assert again.style(0, col) == fmt, col


def test_styles_match_openpyxl(xlsx: Path, book: px.ReadWorkbook):
    want = openpyxl.load_workbook(xlsx)["Styles"]
    ws = book["Styles"]
    for cell in next(want.iter_rows()):
        fmt = ws.style(0, cell.column - 1)
        if cell.style_id == 0:
            assert fmt is None
            continue
        assert fmt is not None
        assert fmt.bold == cell.font.b and fmt.italic == cell.font.i
        assert fmt.font_name == cell.font.name and fmt.font_size == cell.font.sz
        assert (fmt.num_format or "General") == cell.number_format
        assert fmt.text_wrap == bool(cell.alignment.wrap_text)
        assert fmt.locked == cell.protection.locked


# ---------------------------------------------------------------------------
# Sources and errors
# ---------------------------------------------------------------------------


def test_sources(xlsx: Path, data: bytes):
    expected = px.open_workbook(xlsx)["Values"].values()
    assert px.open_workbook(str(xlsx))["Values"].values() == expected
    assert px.open_workbook(data)["Values"].values() == expected
    assert px.open_workbook(bytearray(data))["Values"].values() == expected
    assert px.open_workbook(memoryview(data))["Values"].values() == expected
    assert px.open_workbook(io.BytesIO(data))["Values"].values() == expected
    with open(xlsx, "rb") as fh:
        assert px.open_workbook(fh)["Values"].values() == expected
    assert px.ReadWorkbook(xlsx)["Values"].values() == expected


def test_source_type_errors(xlsx: Path):
    with pytest.raises(TypeError, match="path"):
        px.open_workbook(42)  # pyrefly: ignore[bad-argument-type]
    with open(xlsx, encoding="latin-1") as fh, pytest.raises(TypeError, match="binary mode"):
        px.open_workbook(fh)  # pyrefly: ignore[bad-argument-type]


def test_missing_file(tmp_path: Path):
    with pytest.raises(px.FileError) as info:
        px.open_workbook(tmp_path / "missing.xlsx")
    assert isinstance(info.value, OSError)
    with pytest.raises(px.FileError, match="directory"):
        px.open_workbook(tmp_path)


@pytest.mark.parametrize(
    "content",
    [b"", b"hello, world", b"PK\x03\x04 not really a zip", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1garbage"],
    ids=["empty", "text", "bad-zip", "bad-ole"],
)
def test_not_a_spreadsheet(content: bytes):
    with pytest.raises(px.ReadError):
        px.open_workbook(content)


def test_zip_without_workbook():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("readme.txt", "hi")
    with pytest.raises(px.ReadError, match="without a workbook"):
        px.open_workbook(buf.getvalue())


def test_damaged_xlsx(data: bytes):
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                content = content[: len(content) // 2]  # truncated XML
            dst.writestr(item, content)
    book = px.open_workbook(out.getvalue())
    with pytest.raises(px.ReadError, match="Values"):
        book["Values"]
    assert book["Styles"].cell("A1") == "bold"  # other sheets are unaffected


def test_close(xlsx: Path):
    with px.open_workbook(xlsx) as book:
        loaded = book["Values"]
        assert not book.closed
    assert book.closed and "closed" in repr(book)
    assert loaded.values()[0][0] == "text"  # loaded sheets stay readable
    with pytest.raises(px.ReadError, match="closed"):
        book["Styles"]
    with pytest.raises(px.ReadError, match="closed"):
        loaded.formulas()
    book.close()  # idempotent
    xlsx.unlink()  # the file handle was released


def test_close_frees_sheets_nothing_holds(xlsx: Path):
    book = px.open_workbook(xlsx)
    held = book["Values"]
    book["Styles"].values()  # parsed, then no longer referenced
    book.close()
    assert book["Values"].values() == held.values()  # still held: still readable
    with pytest.raises(px.ReadError, match="closed"):
        book["Styles"]  # freed by close()


def test_repeated_text_shares_one_str():
    wb = px.Workbook()
    ws = wb.add_worksheet("S")
    ws.write_row(0, 0, ["name", "currency"])
    for row in range(1, 4):
        ws.write_row(row, 0, [f"item{row}", "USD"])
    sheet = px.open_workbook(wb.save_to_buffer())["S"]
    values = sheet.values()
    assert values[1][1] == "USD" and values[1][1] is values[2][1] is values[3][1]
    records = sheet.records()
    assert [r["name"] for r in records] == ["item1", "item2", "item3"]
    assert records[0]["currency"] is records[2]["currency"]


def test_threads(data: bytes):
    book = px.open_workbook(data)
    results: dict[int, list] = {}

    def read(i: int) -> None:
        results[i] = book[i % 2].values()

    threads = [threading.Thread(target=read, args=(i,)) for i in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert all(results[i] == results[i % 2] for i in range(8))
    assert results[0][0][0] == "text" and results[1][0][0] == "bold"


# ---------------------------------------------------------------------------
# OpenDocument
# ---------------------------------------------------------------------------

ODS_CONTENT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
<table:table-cell office:value-type="string"><text:p>name</text:p></table:table-cell>
<table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
<table:table-cell office:value-type="float" office:value="2.5"><text:p>2.5</text:p></table:table-cell>
<table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>
</table:table-row>
<table:table-row>
<table:table-cell office:value-type="date" office:date-value="2024-01-02"><text:p>x</text:p></table:table-cell>
<table:table-cell office:value-type="date" office:date-value="2024-01-02T03:04:05"><text:p>x</text:p></table:table-cell>
<table:table-cell office:value-type="time" office:time-value="PT36H30M00S"><text:p>x</text:p></table:table-cell>
<table:table-cell table:formula="of:=[.B1]*2" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell>
</table:table-row>
</table:table>
<table:table table:name="Second">
<table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row>
</table:table>
</office:spreadsheet></office:body></office:document-content>
"""


def make_ods() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", zipfile.ZIP_STORED)
        # calamine expects no whitespace between the table elements.
        content = "".join(line.strip() for line in ODS_CONTENT.splitlines())
        zf.writestr("content.xml", content, zipfile.ZIP_DEFLATED)
        zf.writestr(
            "META-INF/manifest.xml",
            '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
            'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>',
        )
    return buf.getvalue()


def test_ods():
    book = px.open_workbook(make_ods())
    assert book.format == "ods"
    assert book.sheet_names == ["Data", "Second"]
    ws = book["Data"]
    assert ws.values() == [
        ["name", 3, 2.5, True],
        [dt.datetime(2024, 1, 2), dt.datetime(2024, 1, 2, 3, 4, 5), dt.timedelta(hours=36, minutes=30), 6],
    ]
    assert ws.formula("D2") == "=[.B1]*2"
    assert ws.style("A1") is None  # styles are read from xlsx only
    assert ws.merged_ranges == [] and ws.hyperlinks == []
    assert book.tables == [] and book.is_1904 is False
    assert book["Second"].cell("A1") == "x"
