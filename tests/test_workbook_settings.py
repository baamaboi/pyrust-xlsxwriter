"""Workbook-level settings: document properties, defined names, default
format and themes, calculation mode, tab ratio and zip options."""

import copy
import datetime as dt
import re
import zipfile
from typing import Any

import pytest

import pyrust_xlsxwriter as px

TEXT_FIELDS = [
    "title",
    "subject",
    "manager",
    "company",
    "category",
    "author",
    "keywords",
    "comment",
    "status",
    "hyperlink_base",
]


def save(wb, out, name="out.xlsx"):
    path = out(name)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# DocProperties
# ---------------------------------------------------------------------------


def test_doc_properties_class():
    p = px.DocProperties()
    assert all(getattr(p, f) == "" for f in TEXT_FIELDS)
    assert p.creation_datetime is None and p.custom_properties == {}
    assert repr(p) == "DocProperties()"
    upper: dict[str, Any] = {f: f.upper() for f in TEXT_FIELDS}
    full = px.DocProperties(**upper)
    assert all(getattr(full, f) == f.upper() for f in TEXT_FIELDS)
    for field in TEXT_FIELDS:
        q = px.DocProperties()
        setattr(q, field, "x")
        assert getattr(q, field) == "x" and q != p
    q = px.DocProperties().set(title="T", author="A")
    assert (q.title, q.author) == ("T", "A")
    assert repr(q) == "DocProperties(title='T', author='A')"
    assert q == px.DocProperties(author="A", title="T")
    assert q != "not properties"
    assert q.copy() == q and copy.copy(q) == q and copy.deepcopy(q) == q
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.DocProperties(owner="x")  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(TypeError):
        px.DocProperties(title=5)  # pyrefly: ignore[bad-argument-type]


def test_doc_properties_datetimes():
    naive = dt.datetime(2024, 1, 2, 3, 4, 5)
    assert px.DocProperties(creation_datetime=naive).creation_datetime == naive
    aware = dt.datetime(2024, 1, 2, 3, 4, 5, tzinfo=dt.timezone(dt.timedelta(hours=2)))
    assert px.DocProperties(creation_datetime=aware).creation_datetime == dt.datetime(2024, 1, 2, 1, 4, 5)
    p = px.DocProperties(creation_datetime=dt.date(2024, 5, 6))
    assert p.creation_datetime == dt.datetime(2024, 5, 6)
    assert "creation_datetime=datetime.datetime(2024, 5, 6, 0, 0)" in repr(p)
    p.creation_datetime = None
    assert p.creation_datetime is None
    with pytest.raises(TypeError, match="datetime or date"):
        px.DocProperties(creation_datetime="2024-01-01")  # pyrefly: ignore[bad-argument-type]


def test_custom_properties():
    p = px.DocProperties()
    assert p.set_custom_property("Checked by", "Ann") is p
    when = dt.datetime(2020, 1, 2, 3, 4, 5)
    p.set_custom_property("Revision", 3).set_custom_property("Ratio", 0.5)
    p.set_custom_property("Final", True).set_custom_property("Approved", when)
    p.set_custom_property("Day", dt.date(2021, 2, 3))
    p.set_custom_property("Big", 2**40)  # outside 32 bits: stored as a number
    p.set_custom_property("Revision", 4)  # replaces, keeps the position
    assert p.custom_properties == {
        "Checked by": "Ann",
        "Revision": 4,
        "Ratio": 0.5,
        "Final": True,
        "Approved": when,
        "Day": dt.datetime(2021, 2, 3),
        "Big": float(2**40),
    }
    assert list(p.custom_properties) == ["Checked by", "Revision", "Ratio", "Final", "Approved", "Day", "Big"]
    assert type(p.custom_properties["Final"]) is bool
    p.remove_custom_property("Big")
    assert "Big" not in p.custom_properties
    with pytest.raises(KeyError):
        p.remove_custom_property("Big")
    p.custom_properties = [("a", 1), ("b", "x")]
    assert p.custom_properties == {"a": 1, "b": "x"}
    q = px.DocProperties(custom_properties={"n": 1})
    assert q.custom_properties == {"n": 1}
    assert repr(q) == "DocProperties(custom_properties={'n': 1})"
    with pytest.raises(TypeError, match="custom property values"):
        p.set_custom_property("bad", [1])  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.ParameterError, match="empty"):
        p.set_custom_property("", 1)
    with pytest.raises(TypeError):
        p.custom_properties = 5  # pyrefly: ignore[bad-argument-type]


def test_set_properties(out, reopen, part):
    wb = px.Workbook()
    values: dict[str, Any] = {f: f"{f} value" for f in TEXT_FIELDS}
    props = px.DocProperties(**values)
    props.creation_datetime = dt.datetime(2023, 4, 5, 6, 7, 8)
    props.set_custom_property("Checked by", "Ann").set_custom_property("Revision", 3)
    props.set_custom_property("Ratio", 0.5).set_custom_property("Final", True)
    props.set_custom_property("Approved", dt.datetime(2020, 1, 2, 3, 4, 5))
    wb.set_properties(props)
    wb.add_worksheet()
    path = save(wb, out)
    book = reopen(path)
    core = book.properties
    assert (core.title, core.subject, core.creator, core.keywords) == (
        "title value",
        "subject value",
        "author value",
        "keywords value",
    )
    assert (core.description, core.category, core.contentStatus) == ("comment value", "category value", "status value")
    assert core.created == dt.datetime(2023, 4, 5, 6, 7, 8)
    app = part(path, "docProps/app.xml")
    assert "<Manager>manager value</Manager>" in app
    assert "<Company>company value</Company>" in app
    assert "<HyperlinkBase>hyperlink_base value</HyperlinkBase>" in app
    custom = {p.name: p.value for p in book.custom_doc_props.props}
    assert custom == {
        "Checked by": "Ann",
        "Revision": 3,
        "Ratio": 0.5,
        "Final": True,
        "Approved": dt.datetime(2020, 1, 2, 3, 4, 5),
    }
    xml = part(path, "docProps/custom.xml")
    assert "<vt:i4>3</vt:i4>" in xml and "<vt:r8>0.5</vt:r8>" in xml
    assert "<vt:filetime>2020-01-02T03:04:05Z</vt:filetime>" in xml


def test_set_properties_from_dict_and_errors(out, reopen):
    wb = px.Workbook()
    wb.set_properties({"title": "From dict", "custom_properties": {"k": "v"}})
    wb.add_worksheet()
    book = reopen(save(wb, out))
    assert book.properties.title == "From dict"
    assert {p.name: p.value for p in book.custom_doc_props.props} == {"k": "v"}
    with pytest.raises(TypeError, match="DocProperties or a dict"):
        wb.set_properties("title")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="unexpected keyword"):
        wb.set_properties({"nope": 1})


def test_default_creation_datetime_is_now(out, reopen):
    before = dt.datetime.now(dt.UTC).replace(tzinfo=None, microsecond=0)
    wb = px.Workbook()
    wb.set_properties(px.DocProperties(title="t"))
    wb.add_worksheet()
    created = reopen(save(wb, out)).properties.created
    assert before - dt.timedelta(seconds=5) <= created <= before + dt.timedelta(minutes=5)


# ---------------------------------------------------------------------------
# Defined names
# ---------------------------------------------------------------------------


def test_define_name(out, reopen):
    wb = px.Workbook()
    wb.add_worksheet()
    wb.add_worksheet("My Data")
    wb.define_name("Rate", "=0.05")
    wb.define_name("Sales", "=Sheet1!$A$1:$A$10")
    wb.define_name("Sheet1!Local", "=Sheet1!$B$1")
    wb.define_name("'My Data'!Other", "='My Data'!$C$1")
    book = reopen(save(wb, out))
    assert book.defined_names["Rate"].attr_text == "0.05"
    assert book.defined_names["Sales"].attr_text == "Sheet1!$A$1:$A$10"
    assert book["Sheet1"].defined_names["Local"].attr_text == "Sheet1!$B$1"
    assert book["My Data"].defined_names["Other"].attr_text == "'My Data'!$C$1"


@pytest.mark.parametrize("name", ["!bad", "1abc", "A1", "has space", ""])
def test_define_name_invalid(name):
    wb = px.Workbook()
    with pytest.raises(px.ParameterError):
        wb.define_name(name, "=1")


# ---------------------------------------------------------------------------
# Default format and themes
# ---------------------------------------------------------------------------


def first_font(xml):
    match = re.search(r"<fonts[^>]*><font>(.*?)</font>", xml)
    assert match
    return match.group(1)


def test_set_default_format(out, reopen, part):
    wb = px.Workbook()
    wb.set_default_format(px.Format(font_name="Arial", font_size=10), 18, 72)
    std = wb.add_worksheet("Std")
    tall = wb.add_worksheet("Tall", default_row_height=30)
    pushed = px.Worksheet("Pushed")
    pushed.write("A1", "p")
    wb.push_worksheet(pushed)
    std.write("A1", "x")
    tall.write("A1", "y")
    path = save(wb, out)
    book = reopen(path)
    font = first_font(part(path, "xl/styles.xml"))
    assert '<sz val="10"/>' in font and '<name val="Arial"/>' in font
    assert book["Std"]["A1"].font.name == "Arial" and book["Pushed"]["A1"].font.name == "Arial"
    sheets = [part(path, f"xl/worksheets/sheet{i}.xml") for i in (1, 2, 3)]
    assert '<sheetFormatPr defaultRowHeight="13.5"/>' in sheets[0]
    assert 'defaultRowHeight="30"' in sheets[1]
    assert '<sheetFormatPr defaultRowHeight="13.5"/>' in sheets[2]


def test_set_default_format_errors():
    wb = px.Workbook()
    with pytest.raises(px.ParameterError, match="column width 50"):
        wb.set_default_format(px.Format(), 20, 50)
    with pytest.raises(px.ParameterError, match="row height"):
        wb.set_default_format(px.Format(), 0, 64)
    wb.add_worksheet()
    with pytest.raises(px.ParameterError, match="before any worksheet"):
        wb.set_default_format(px.Format(), 20, 64)
    with pytest.raises(px.ParameterError, match="before any worksheet"):
        wb.use_excel_2023_theme()


def test_default_format_and_pushed_formatted_sheet():
    wb = px.Workbook()
    wb.set_default_format(px.Format(font_name="Arial"), 20, 64)
    ws = px.Worksheet("Formatted")
    ws.write("A1", "x", px.Format(bold=True))
    with pytest.raises(px.ParameterError, match="default format"):
        wb.push_worksheet(ws)
    assert len(wb) == 0


def test_excel_2023_theme(out, reopen, part):
    wb = px.Workbook()
    wb.use_excel_2023_theme()
    ws = wb.add_worksheet()
    ws.write("A1", "Hello")
    path = save(wb, out)
    book = reopen(path)
    assert "Aptos" in part(path, "xl/theme/theme1.xml")
    font = first_font(part(path, "xl/styles.xml"))
    assert '<name val="Aptos Narrow"/>' in font and '<scheme val="minor"/>' in font
    assert book["Sheet1"]["A1"].font.name == "Aptos Narrow"
    assert 'defaultThemeVersion="202300"' in part(path, "xl/workbook.xml")


def test_custom_theme(out, reopen, part, corpus, tmp_path):
    source = corpus / "theme02.xlsx"
    with zipfile.ZipFile(source) as zf:
        theme = zf.read("xl/theme/theme1.xml").decode("utf-8")
    wb = px.Workbook()
    wb.use_custom_theme(source)
    wb.add_worksheet().write("A1", 1)
    path = save(wb, out)
    reopen(path)
    assert part(path, "xl/theme/theme1.xml").strip() == theme.strip()
    # A theme1.xml file extracted from a package works too.
    xml_file = tmp_path / "theme1.xml"
    xml_file.write_text(theme, encoding="utf-8")
    wb = px.Workbook()
    wb.use_custom_theme(str(xml_file))
    wb.add_worksheet()
    path = save(wb, out, "b.xlsx")
    assert part(path, "xl/theme/theme1.xml").strip() == theme.strip()


def test_custom_theme_errors(tmp_path):
    wb = px.Workbook()
    with pytest.raises(px.FileError):
        wb.use_custom_theme(tmp_path / "missing.thmx")
    bad = tmp_path / "bad.xml"
    bad.write_text("not a theme")
    with pytest.raises(px.ParameterError, match="theme"):
        wb.use_custom_theme(bad)


# ---------------------------------------------------------------------------
# Other workbook settings
# ---------------------------------------------------------------------------


def test_read_only_recommended(out, reopen, part):
    wb = px.Workbook()
    wb.read_only_recommended()
    wb.add_worksheet()
    path = save(wb, out)
    reopen(path)
    assert '<fileSharing readOnlyRecommended="1"/>' in part(path, "xl/workbook.xml")


@pytest.mark.parametrize(
    ("mode", "expected"),
    [(px.CalculationMode.Manual, "manual"), (px.CalculationMode.Partial, "autoNoTable")],
)
def test_calculation_mode(out, reopen, mode, expected):
    wb = px.Workbook()
    wb.set_calculation_mode(mode)
    wb.add_worksheet().write("A1", "=1+1")
    assert reopen(save(wb, out)).calculation.calcMode == expected


def test_calculation_mode_automatic(out, reopen):
    wb = px.Workbook()
    wb.set_calculation_mode(px.CalculationMode.Manual)
    wb.set_calculation_mode(px.CalculationMode.Automatic)
    wb.add_worksheet()
    calc = reopen(save(wb, out)).calculation
    assert calc.calcMode in (None, "auto") and calc.fullCalcOnLoad
    assert [int(m) for m in (px.CalculationMode.Automatic, px.CalculationMode.Manual, px.CalculationMode.Partial)] == [
        0,
        1,
        2,
    ]
    with pytest.raises(TypeError):
        wb.set_calculation_mode("manual")  # pyrefly: ignore[bad-argument-type]


def test_tab_ratio(out, reopen):
    wb = px.Workbook()
    wb.set_tab_ratio(35)
    wb.add_worksheet()
    assert reopen(save(wb, out)).views[0].tabRatio == 350
    with pytest.raises(px.ParameterError, match=r"0\.\.=100"):
        wb.set_tab_ratio(101)
    with pytest.raises(OverflowError):
        wb.set_tab_ratio(-1)


def test_zip_compression_level(out, reopen):
    sizes = {}
    for level in (1, 9):
        wb = px.Workbook()
        wb.set_zip_compression_level(level)
        ws = wb.add_worksheet()
        for row in range(2000):
            ws.write_row(row, 0, [row, f"text {row % 37}", row * 1.5])
        path = save(wb, out, f"level{level}.xlsx")
        assert reopen(path)["Sheet1"]["A2000"].value == 1999
        sizes[level] = path.stat().st_size
    assert sizes[9] < sizes[1]
    wb = px.Workbook()
    for level in (0, 10):
        with pytest.raises(px.ParameterError, match=r"1\.\.=9"):
            wb.set_zip_compression_level(level)


def test_zip_large_file(out, reopen):
    wb = px.Workbook()
    wb.use_zip_large_file()
    wb.add_worksheet().write("A1", "big")
    path = save(wb, out)
    assert reopen(path)["Sheet1"]["A1"].value == "big"
    with zipfile.ZipFile(path) as zf:
        assert all(info.extract_version >= 45 for info in zf.infolist())  # ZIP64
    wb.use_zip_large_file(False)
    path = save(wb, out, "small.xlsx")
    with zipfile.ZipFile(path) as zf:
        assert all(info.extract_version < 45 for info in zf.infolist())
