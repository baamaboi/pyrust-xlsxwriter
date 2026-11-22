"""Chart axes: names, scaling, units, gridlines, ticks, labels, crossing, date axes."""

from __future__ import annotations

import datetime as dt
import re

import pytest
from charts_util import chart_xml, column_chart, plot_area, save_chart

import pyrust_xlsxwriter as px


def axes(xml: str) -> dict[str, str]:
    """The axis elements of the plot area by tag, e.g. {"catAx": "...", "valAx": "..."}."""
    found: dict[str, str] = {}
    for tag, body in re.findall(r"<c:(catAx|valAx|dateAx)>(.*?)</c:\1>", plot_area(xml)):
        found.setdefault(tag, body)
    return found


def test_axis_defaults_keep_chart_type_defaults(out):
    chart = column_chart()
    assert chart.x_axis.major_gridlines is None and chart.y_axis.hidden is None
    ax = axes(chart_xml(save_chart(out, chart)))
    assert "<c:majorGridlines/>" in ax["valAx"] and "majorGridlines" not in ax["catAx"]
    assert '<c:auto val="1"/>' in ax["catAx"]
    assert repr(chart.y_axis) == "ChartAxis()"


def test_axis_names_and_name_formatting(out, reopen):
    chart = column_chart(px.ChartType.Line)
    chart.x_axis.name = "Month"
    chart.x_axis.name_font.italic = True
    chart.x_axis.name_format.solid_fill.color = "yellow"
    chart.x_axis.label_layout = px.ChartLayout(offset=(0.5, 0.9))
    chart.y_axis.name = ("Sheet1", 0, 1)
    path = save_chart(out, chart)
    reopen(path)
    ax = axes(chart_xml(path))
    cat = ax["catAx"]
    assert "<a:t>Month</a:t>" in cat and 'i="1"' in cat
    assert '<a:srgbClr val="FFFF00"/>' in cat
    assert '<c:x val="0.5"/><c:y val="0.9"/>' in cat
    assert "<c:f>Sheet1!$B$1</c:f>" in ax["valAx"]
    assert chart.y_axis_name == "=Sheet1!$B$1"


def test_value_axis_scaling(out, reopen):
    chart = column_chart()
    y = chart.y_axis
    y.min, y.max = 0, 100
    y.major_unit, y.minor_unit = 20, 5
    y.log_base = 10
    y.reverse = True
    y.num_format = "$#,##0"
    y.display_unit_type = px.ChartAxisDisplayUnitType.Thousands
    y.display_units_visible = True
    path = save_chart(out, chart)
    reopen(path)
    val = axes(chart_xml(path))["valAx"]
    assert (
        '<c:scaling><c:logBase val="10"/><c:orientation val="maxMin"/><c:max val="100"/><c:min val="0"/></c:scaling>'
        in val
    )
    assert '<c:majorUnit val="20"/><c:minorUnit val="5"/>' in val
    assert '<c:numFmt formatCode="$#,##0" sourceLinked="0"/>' in val
    assert '<c:builtInUnit val="thousands"/><c:dispUnitsLbl>' in val


def test_display_units_hidden(out):
    chart = column_chart()
    chart.y_axis.display_unit_type = px.ChartAxisDisplayUnitType.Millions
    chart.y_axis.display_units_visible = False
    val = axes(chart_xml(save_chart(out, chart)))["valAx"]
    assert '<c:builtInUnit val="millions"/>' in val and "dispUnitsLbl" not in val


def test_gridlines(out):
    chart = column_chart()
    chart.y_axis.major_gridlines_line = px.ChartLine(color="#D9D9D9", dash_type=px.ChartLineDashType.RoundDot)
    chart.y_axis.minor_gridlines_line.color = "#EEEEEE"  # setting the line shows the gridlines
    chart.x_axis.major_gridlines = True
    ax = axes(chart_xml(save_chart(out, chart)))
    assert (
        '<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill><a:prstDash val="sysDot"/>'
        in ax["valAx"]
    )
    assert '<c:minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="EEEEEE"/>' in ax["valAx"]
    assert "<c:majorGridlines/>" in ax["catAx"]
    chart.y_axis.major_gridlines = False
    chart.y_axis.major_gridlines_line = None
    assert "majorGridlines" not in axes(chart_xml(save_chart(out, chart, "off.xlsx")))["valAx"]


def test_ticks_labels_and_intervals(out):
    chart = column_chart()
    x = chart.x_axis
    x.major_tick_type = px.ChartAxisTickType.Cross
    x.minor_tick_type = px.ChartAxisTickType.Inside
    x.label_position = px.ChartAxisLabelPosition.Low
    x.label_alignment = px.ChartAxisLabelAlignment.Left
    x.label_interval = 2
    x.tick_interval = 3
    x.position_between_ticks = False
    x.font = px.ChartFont(size=8, rotation=-45)
    x.format.line.color = "red"
    ax = axes(chart_xml(save_chart(out, chart)))
    cat = ax["catAx"]
    assert '<c:majorTickMark val="cross"/><c:minorTickMark val="in"/><c:tickLblPos val="low"/>' in cat
    assert '<c:lblAlgn val="l"/>' in cat and '<c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/>' in cat
    assert '<a:bodyPr rot="-2700000" vert="horz"/>' in cat and 'sz="800"' in cat
    assert '<a:ln><a:solidFill><a:srgbClr val="FF0000"/>' in cat
    assert '<c:crossBetween val="midCat"/>' in ax["valAx"]


@pytest.mark.parametrize(
    "tick,code",
    [
        (px.ChartAxisTickType.None_, "none"),
        (px.ChartAxisTickType.Inside, "in"),
        (px.ChartAxisTickType.Outside, "out"),
        (px.ChartAxisTickType.Cross, "cross"),
    ],
)
def test_tick_types(out, tick, code):
    chart = column_chart()
    chart.y_axis.major_tick_type = tick
    assert f'<c:majorTickMark val="{code}"/>' in axes(chart_xml(save_chart(out, chart)))["valAx"]


@pytest.mark.parametrize(
    "position,code",
    [
        (px.ChartAxisLabelPosition.NextTo, "nextTo"),
        (px.ChartAxisLabelPosition.High, "high"),
        (px.ChartAxisLabelPosition.Low, "low"),
        (px.ChartAxisLabelPosition.None_, "none"),
    ],
)
def test_label_positions(out, position, code):
    chart = column_chart()
    chart.y_axis.label_position = position
    assert f'<c:tickLblPos val="{code}"/>' in axes(chart_xml(save_chart(out, chart)))["valAx"]


def test_crossing(out):
    chart = column_chart()
    chart.x_axis.crossing = px.ChartAxisCrossing.CategoryNumber(2)
    chart.y_axis.crossing = px.ChartAxisCrossing.Max
    ax = axes(chart_xml(save_chart(out, chart)))
    # rust_xlsxwriter writes an axis' crossing on the perpendicular axis
    assert '<c:crossesAt val="2"/>' in ax["valAx"]
    assert '<c:crosses val="max"/>' in ax["catAx"]
    chart.x_axis.crossing = px.ChartAxisCrossing.AxisValue(3.5)
    chart.y_axis.crossing = px.ChartAxisCrossing.Min
    ax = axes(chart_xml(save_chart(out, chart, "c2.xlsx")))
    assert '<c:crossesAt val="3.5"/>' in ax["valAx"] and '<c:crosses val="min"/>' in ax["catAx"]


def test_crossing_value_class():
    c = px.ChartAxisCrossing
    assert c.Automatic.kind == "Automatic" and c.Automatic.value is None
    assert c.CategoryNumber(3).kind == "CategoryNumber" and c.CategoryNumber(3).value == 3
    assert c.AxisValue(2.5).value == 2.5
    assert c.AxisValue(2.5) == c.AxisValue(2.5) and c.AxisValue(2.5) != c.AxisValue(3)
    assert c.Max == c.Max and c.Max != c.Min and c.Min != "Min"
    assert len({c.Max, c.Max, c.AxisValue(1), c.AxisValue(1.0)}) == 2
    assert repr(c.AxisValue(2.5)) == "ChartAxisCrossing.AxisValue(2.5)"
    assert repr(c.Max) == "ChartAxisCrossing.Max"
    with pytest.raises(px.ParameterError):
        c.AxisValue(float("nan"))


def test_date_axis(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    for i in range(5):
        ws.write(i, 0, dt.date(2024, 1, 1) + dt.timedelta(days=30 * i))
        ws.write(i, 1, i * 10)
    chart = px.Chart(px.ChartType.Line)
    chart.add_series(values="=Sheet1!$B$1:$B$5", categories="=Sheet1!$A$1:$A$5")
    x = chart.x_axis
    x.date_axis = True
    x.min_date = dt.date(2024, 1, 1)
    x.max_date = dt.datetime(2024, 6, 1, 12, 0)
    x.major_unit = 1
    x.major_unit_date_type = px.ChartAxisDateUnitType.Months
    x.minor_unit = 7
    x.minor_unit_date_type = px.ChartAxisDateUnitType.Days
    x.num_format = "mmm yy"
    assert x.min_date == dt.datetime(2024, 1, 1) and x.max_date == dt.datetime(2024, 6, 1, 12)
    ws.insert_chart("D2", chart)
    path = out()
    wb.save(path)
    reopen(path)
    date_ax = axes(chart_xml(path))["dateAx"]
    assert '<c:max val="45444.5"/><c:min val="45292"/>' in date_ax
    assert (
        '<c:majorUnit val="1"/><c:majorTimeUnit val="months"/><c:minorUnit val="7"/><c:minorTimeUnit val="days"/>'
        in date_ax
    )
    assert '<c:numFmt formatCode="mmm yy" sourceLinked="0"/>' in date_ax


def test_min_max_and_dates_replace_each_other():
    axis = px.ChartAxis(max=10)
    axis.max_date = dt.date(2024, 1, 1)
    assert axis.max is None
    axis.max = 5
    assert axis.max_date is None and axis.max == 5
    axis.min_date = dt.date(2020, 1, 1)
    axis.min = None
    assert axis.min_date is None
    with pytest.raises(px.ParameterError, match="time zones"):
        axis.max_date = dt.datetime(2024, 1, 1, tzinfo=dt.UTC)
    with pytest.raises(TypeError, match="ChartAxis.max_date"):
        px.ChartAxis(max_date="2024-01-01")  # pyrefly: ignore[bad-argument-type]


def test_text_and_automatic_axis(out):
    chart = column_chart()
    chart.x_axis.text_axis = True
    assert '<c:auto val="1"/>' not in axes(chart_xml(save_chart(out, chart)))["catAx"]
    chart.x_axis.text_axis = None
    chart.x_axis.automatic_axis = False
    assert '<c:auto val="1"/>' in axes(chart_xml(save_chart(out, chart, "a.xlsx")))["catAx"]


def test_hidden_axis_and_secondary_axes(out, reopen):
    chart = column_chart(px.ChartType.Line, series=2)
    chart.series[1].secondary_axis = True
    chart.x_axis.hidden = True
    chart.y2_axis.name = "Secondary"
    chart.y2_axis.max = 1000
    chart.x2_axis.hidden = False
    path = save_chart(out, chart)
    reopen(path)
    plot = plot_area(chart_xml(path))
    assert plot.count("<c:lineChart>") == 2
    assert plot.count("<c:valAx>") == 2 and plot.count("<c:catAx>") == 2
    assert '<c:delete val="1"/>' in plot
    assert "<a:t>Secondary</a:t>" in plot and '<c:max val="1000"/>' in plot


@pytest.mark.parametrize(
    "prop,value",
    [
        ("log_base", 1),
        ("major_unit", -1.0),
        ("minor_unit", -0.5),
        ("label_interval", 0),
        ("tick_interval", 0),
        ("max", float("inf")),
    ],
)
def test_axis_validation(prop, value):
    axis = px.ChartAxis()
    with pytest.raises(px.ParameterError):
        setattr(axis, prop, value)
    assert getattr(axis, prop) is None


def test_axis_keywords_and_repr():
    axis = px.ChartAxis(
        name="Y", min=0, reverse=True, major_gridlines_line={"color": "red"}, crossing=px.ChartAxisCrossing.Max
    )
    text = repr(axis)
    assert text.startswith("ChartAxis(name='Y', crossing=ChartAxisCrossing.Max, reverse=True")
    assert "major_gridlines_line=ChartLine(color=Color.Red)" in text and text.endswith("min=0.0)")
