"""Sparklines: Sparkline, SparklineType, SparklineEmptyCells, add_sparkline(_group)."""

import re

import pytest

import pyrust_xlsxwriter as px

pytestmark = pytest.mark.filterwarnings("ignore:.*extension is not supported:UserWarning")


def groups(part, path, sheet=1):
    """The <x14:sparklineGroup ...> elements (as strings) of a worksheet."""
    xml = part(path, f"xl/worksheets/sheet{sheet}.xml")
    return re.findall(r"<x14:sparklineGroup\b.*?</x14:sparklineGroup>", xml)


def attrs(group):
    head = re.match(r"<x14:sparklineGroup([^>]*)>", group)
    assert head
    return dict(re.findall(r'(\w+)="([^"]*)"', head.group(1)))


def book():
    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_row("A1", [1, -2, 3, 4, 5])
    ws.write_row("A2", [5, 4, -3, 2, 1])
    ws.write_row("A3", [2, 2, 2, 2, 2])
    return wb, ws


def test_add_sparkline_forms(out, reopen, part):
    wb, ws = book()
    ws.add_sparkline("F1", px.Sparkline("A1:E1"))  # sheet name is filled in
    ws.add_sparkline(1, 5, px.Sparkline(range="Data!A2:E2", sparkline_type=px.SparklineType.Column))
    ws.add_sparkline(2, "F", sparkline=px.Sparkline(("Data", 2, 0, 2, 4), px.SparklineType.WinLoss))
    path = out()
    wb.save(path)
    reopen(path)
    gs = groups(part, path)
    assert len(gs) == 3
    found = {sqref: g for g in gs for sqref in re.findall(r"<xm:sqref>(\w+)</xm:sqref>", g)}
    assert "<xm:f>Data!A1:E1</xm:f>" in found["F1"] and "type" not in attrs(found["F1"])
    assert attrs(found["F2"])["type"] == "column"
    assert attrs(found["F3"])["type"] == "stacked"
    assert "<xm:f>Data!A3:E3</xm:f>" in found["F3"]


def test_all_options(out, reopen, part):
    wb, ws = book()
    ws2 = wb.add_worksheet("Dates")
    ws2.write_row("A1", [45000, 45001, 45002, 45003, 45004])
    sl = px.Sparkline(
        "A1:E1",
        style=12,
        high_point=True,
        low_point=True,
        first_point=True,
        last_point=True,
        negative_points=True,
        markers=True,
        axis=True,
        hidden_data=True,
        empty_cells=px.SparklineEmptyCells.Connected,
        right_to_left=True,
        sparkline_color="#112233",
        markers_color="#223344",
        line_weight=2.25,
        custom_min=-5,
        custom_max=10,
        date_range="Dates!A1:E1",
    )
    ws.add_sparkline("F1", sl)
    path = out()
    wb.save(path)
    reopen(path)
    (g,) = groups(part, path)
    a = attrs(g)
    for flag in ["high", "low", "first", "last", "negative", "markers", "displayXAxis", "displayHidden", "rightToLeft"]:
        assert a[flag] == "1", flag
    assert a["displayEmptyCellsAs"] == "span"
    assert a["lineWeight"] == "2.25"
    assert (a["minAxisType"], a["manualMin"]) == ("custom", "-5")
    assert (a["maxAxisType"], a["manualMax"]) == ("custom", "10")
    assert 'x14:colorSeries rgb="FF112233"' in g
    assert 'x14:colorMarkers rgb="FF223344"' in g
    assert a["dateAxis"] == "1" and "<xm:f>Dates!A1:E1</xm:f>" in g


def test_point_colours_turn_points_on(out, reopen, part):
    wb, ws = book()
    sl = px.Sparkline(
        "A1:E1",
        high_point_color="red",
        low_point_color=(0, 0, 255),
        first_point_color=0x00FF00,
        last_point_color="#010203",
        negative_points_color=px.Color.Orange,
    )
    assert sl.high_point and sl.low_point and sl.first_point and sl.last_point and sl.negative_points
    ws.add_sparkline("F1", sl)
    path = out()
    wb.save(path)
    reopen(path)
    (g,) = groups(part, path)
    a = attrs(g)
    assert all(a[f] == "1" for f in ["high", "low", "first", "last", "negative"])
    assert 'x14:colorHigh rgb="FFFF0000"' in g
    assert 'x14:colorLow rgb="FF0000FF"' in g
    assert 'x14:colorFirst rgb="FF00FF00"' in g
    assert 'x14:colorLast rgb="FF010203"' in g
    assert 'x14:colorNegative rgb="FFFF6600"' in g


def test_group_and_column_order(out, reopen, part):
    wb, ws = book()
    ws.add_sparkline_group("F1:F3", px.Sparkline("A1:E3", group_max=True, group_min=True))
    ws.add_sparkline_group(4, 0, 4, 2, px.Sparkline("Data!A1:C3", column_order=True))
    path = out()
    wb.save(path)
    reopen(path)
    gs = groups(part, path)
    assert len(gs) == 2
    first = next(g for g in gs if "<xm:sqref>F1</xm:sqref>" in g)
    assert first.count("<x14:sparkline>") == 3
    a = attrs(first)
    assert a["minAxisType"] == "group" and a["maxAxisType"] == "group"
    second = next(g for g in gs if "<xm:sqref>A5</xm:sqref>" in g)
    assert "<xm:f>Data!A1:A3</xm:f>" in second  # columns, not rows


def test_style_changes_colours(out, part):
    wb, ws = book()
    ws.add_sparkline("F1", px.Sparkline("A1:E1", style=1))
    ws.add_sparkline("F2", px.Sparkline("A2:E2", style=2))
    path = out()
    wb.save(path)
    gs = groups(part, path)
    assert re.search(r'colorSeries theme="4"', gs[0] + gs[1])
    assert re.search(r'colorSeries theme="5"', gs[0] + gs[1])


def test_properties_round_trip():
    sl = px.Sparkline()
    assert sl.range is None and sl.sparkline_type == px.SparklineType.Line and sl.style == 1
    assert not (sl.high_point or sl.low_point or sl.first_point or sl.last_point or sl.markers or sl.negative_points)
    assert not (sl.axis or sl.hidden_data or sl.right_to_left or sl.column_order or sl.group_min or sl.group_max)
    assert sl.empty_cells == px.SparklineEmptyCells.Gaps
    assert sl.line_weight is None and sl.custom_min is None and sl.date_range is None
    assert px.Sparkline("").range is None  # the old default value still works
    sl.range = "Sheet1!A1:E1"
    sl.sparkline_type = px.SparklineType.WinLoss
    sl.high_point = True
    sl.style = 5
    assert (sl.range, sl.sparkline_type, sl.high_point, sl.style) == ("Sheet1!A1:E1", px.SparklineType.WinLoss, True, 5)
    sl.custom_max = 3
    assert sl.custom_max == 3 and not sl.group_max
    sl.group_max = True
    assert sl.custom_max is None and sl.group_max
    sl.custom_min = 1.5
    sl.group_min = True
    assert sl.custom_min is None
    sl.sparkline_color = "#ABCDEF"
    assert sl.sparkline_color == px.Color("#ABCDEF")
    clone = sl.copy()
    assert clone == sl
    clone.markers = True
    assert clone != sl
    assert sl.set(axis=True) is sl and sl.axis
    assert px.Sparkline(("My Data", 0, 0, 0, 4)).range == "'My Data'!$A$1:$E$1"


def test_errors():
    _wb, ws = book()
    with pytest.raises(px.ParameterError, match="no data range"):
        ws.add_sparkline("F1", px.Sparkline())
    with pytest.raises(px.ParameterError, match="1D"):
        ws.add_sparkline("F1", px.Sparkline("A1:E3"))
    with pytest.raises(px.ParameterError, match="2D"):
        ws.add_sparkline_group("F1:F3", px.Sparkline("A1:E1"))
    with pytest.raises(px.ParameterError, match="1..=36"):
        px.Sparkline("A1:E1", style=37)
    with pytest.raises(px.ParameterError, match="line_weight"):
        px.Sparkline("A1:E1", line_weight=0)
    with pytest.raises(px.ParameterError):
        px.Sparkline("A1:E1", custom_min=float("nan"))
    with pytest.raises(TypeError):
        px.Sparkline(42)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.Sparkline("A1:E1", colour="red")  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(px.RangeError):
        ws.add_sparkline(1_048_576, 0, px.Sparkline("A1:E1"))
    with pytest.raises(TypeError):
        ws.add_sparkline("F1", "A1:E1")  # pyrefly: ignore[bad-argument-type]
