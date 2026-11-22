"""Series options: markers, data labels, points, trendlines, error bars, bar/line options."""

from __future__ import annotations

import pytest
from charts_util import chart_xml, column_chart, normalized, plot_area, save_chart

import pyrust_xlsxwriter as px


def first_series(xml: str) -> str:
    plot = plot_area(xml)
    return plot[plot.index("<c:ser>") : plot.index("</c:ser>")]


def test_marker(out, reopen):
    chart = column_chart(px.ChartType.Line)
    series = chart.series[0]
    series.marker = px.ChartMarker(type=px.ChartMarkerType.Diamond, size=9, format={"solid_fill": {"color": "blue"}})
    assert series.marker.type == px.ChartMarkerType.Diamond and series.marker.size == 9
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    assert '<c:marker><c:symbol val="diamond"/><c:size val="9"/><c:spPr><a:solidFill><a:srgbClr val="0000FF"/>' in ser


def test_line_series_have_no_markers_unless_set(out):
    chart = column_chart(px.ChartType.Line)
    assert '<c:symbol val="none"/>' in first_series(chart_xml(save_chart(out, chart)))
    chart.series[0].marker.size = 5  # touching the live marker uses it
    assert '<c:symbol val="none"/>' not in first_series(chart_xml(save_chart(out, chart, "m.xlsx")))
    chart.series[0].marker = {"none": True}
    assert '<c:symbol val="none"/>' in first_series(chart_xml(save_chart(out, chart, "n.xlsx")))


@pytest.mark.parametrize(
    "name,symbol",
    [
        ("Square", "square"),
        ("Diamond", "diamond"),
        ("Triangle", "triangle"),
        ("X", "x"),
        ("Star", "star"),
        ("ShortDash", "dot"),
        ("LongDash", "dash"),
        ("Circle", "circle"),
        ("PlusSign", "plus"),
    ],
)
def test_marker_types(out, name, symbol):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].marker.type = getattr(px.ChartMarkerType, name)
    assert f'<c:symbol val="{symbol}"/>' in chart_xml(save_chart(out, chart))


def test_marker_validation():
    with pytest.raises(px.ParameterError):
        px.ChartMarker(size=1)
    marker = px.ChartMarker()
    with pytest.raises(px.ParameterError):
        marker.size = 73
    assert marker.size is None


def test_data_labels(out, reopen):
    chart = column_chart(px.ChartType.Line)
    label = chart.series[0].data_label
    label.show_value = True
    label.show_category_name = True
    label.show_series_name = True
    label.show_legend_key = True
    label.position = px.ChartDataLabelPosition.Above
    label.num_format = "0.0"
    label.separator = ";"
    label.font.bold = True
    label.format.border.color = "red"
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    dlbls = ser[ser.index("<c:dLbls>") : ser.index("</c:dLbls>")]
    assert '<c:numFmt formatCode="0.0" sourceLinked="0"/>' in dlbls
    assert '<a:srgbClr val="FF0000"/>' in dlbls and '<a:defRPr b="1" i="0"/>' in dlbls
    assert (
        '<c:dLblPos val="t"/><c:showLegendKey val="1"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:separator>; </c:separator>'
        in dlbls
    )


def test_data_label_aliases_and_pie_options(out):
    label = px.ChartDataLabel(show_y_value=True, show_x_value=True)
    assert label.show_value and label.show_category_name
    chart = column_chart(px.ChartType.Pie)
    chart.series[0].data_label = {
        "show_percentage": True,
        "show_leader_lines": True,
        "position": px.ChartDataLabelPosition.InsideEnd,
    }
    ser = first_series(chart_xml(save_chart(out, chart)))
    assert '<c:dLblPos val="inEnd"/>' in ser and '<c:showPercent val="1"/><c:showLeaderLines val="1"/>' in ser
    # the chart type's default position (BestFit for pies) is not written
    chart.series[0].data_label.position = px.ChartDataLabelPosition.BestFit
    assert "dLblPos" not in first_series(chart_xml(save_chart(out, chart, "best.xlsx")))


@pytest.mark.parametrize(
    "name,code",
    [
        ("Center", "ctr"),
        ("Left", "l"),
        ("Above", "t"),
        ("Below", "b"),
    ],
)
def test_data_label_positions(out, name, code):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].data_label = px.ChartDataLabel(show_value=True, position=getattr(px.ChartDataLabelPosition, name))
    assert f'<c:dLblPos val="{code}"/>' in chart_xml(save_chart(out, chart))


def test_data_label_validation():
    with pytest.raises(px.ParameterError, match="separator"):
        px.ChartDataLabel(separator="|")
    with pytest.raises(ValueError):
        px.ChartDataLabel(separator="ab")
    assert px.ChartDataLabel(separator="\n").separator == "\n"


def test_custom_data_labels(out, reopen):
    chart = column_chart(px.ChartType.Pie)
    series = chart.series[0]
    series.data_label.show_percentage = True
    base = px.ChartDataLabel(value="Alpha", font={"italic": True})
    copy = base.to_custom()
    assert copy is not base and copy.value == "Alpha" and copy.font is not base.font and copy.font.italic
    copy.font = None
    assert base.font.italic
    series.custom_data_labels = [copy, None, {"value": ("Sheet1", 0, 1)}, px.ChartDataLabel(hidden=True)]
    labels = series.custom_data_labels
    third = labels[2]
    assert labels[0] is copy and labels[1] is None and third is not None and third.value == ("Sheet1", 0, 1)
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    assert '<c:dLbl><c:idx val="0"/><c:layout/><c:tx><c:rich>' in ser and "<a:t>Alpha</a:t>" in ser
    assert '<c:idx val="1"/>' not in ser.split("<c:dLbls>")[1].split("<c:dLbl>")[1]  # None: default label
    assert '<c:dLbl><c:idx val="2"/><c:layout/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>' in ser
    assert '<c:dLbl><c:idx val="3"/><c:delete val="1"/></c:dLbl>' in ser
    with pytest.raises(TypeError, match="expected a ChartDataLabel or a dict"):
        series.custom_data_labels = ["text"]  # pyrefly: ignore[bad-argument-type]


def test_corpus_custom_data_labels(out, corpus):
    """Same chart as Excel's chart_data_labels28.xlsx."""
    chart = px.Chart(px.ChartType.Column)
    s = chart.add_series(values="=Sheet1!$A$1:$A$5")
    s.data_label.show_value = True
    s.custom_data_labels = [
        px.ChartDataLabel(value="123"),
        px.ChartDataLabel(value="abc"),
        None,
        px.ChartDataLabel(value="=Sheet1!$D$1"),
        px.ChartDataLabel(value="=Sheet1!$D$2"),
    ]
    chart.add_series(values="=Sheet1!$B$1:$B$5")
    chart.add_series(values="=Sheet1!$C$1:$C$5")
    path = save_chart(out, chart)
    assert normalized(chart_xml(path)) == normalized(chart_xml(corpus / "chart_data_labels28.xlsx"))


def test_points_and_point_colors(out, reopen):
    chart = column_chart(px.ChartType.Pie)
    series = chart.series[0]
    series.point_colors = ["red", None, "#00FF00"]
    assert series.point_colors == [px.Color.Red, None, px.Color("#00FF00")]
    first = series.points[0]
    assert first is not None and first.format.solid_fill is not None
    assert first.format.solid_fill.color == px.Color.Red and series.points[1] is None
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    assert '<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/>' in ser
    assert '<c:dPt><c:idx val="1"/>' not in ser
    assert '<c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/>' in ser
    series.points = [
        px.ChartPoint(format=px.ChartPatternFill(pattern=px.ChartPatternFillType.Wave)),
        {"format": {"no_fill": True}},
    ]
    assert series.point_colors == [None, None]
    ser = first_series(chart_xml(save_chart(out, chart, "p.xlsx")))
    assert '<a:pattFill prst="wave">' in ser and '<c:dPt><c:idx val="1"/><c:spPr><a:noFill/>' in ser
    series.points = None
    assert series.points == []


@pytest.mark.parametrize(
    "trend,xml_type,extra",
    [
        (px.ChartTrendlineType.Linear, "linear", ""),
        (px.ChartTrendlineType.Exponential, "exp", ""),
        (px.ChartTrendlineType.Logarithmic, "log", ""),
        (px.ChartTrendlineType.Power, "power", ""),
        (px.ChartTrendlineType.Polynomial(3), "poly", '<c:order val="3"/>'),
        (px.ChartTrendlineType.MovingAverage(2), "movingAvg", '<c:period val="2"/>'),
    ],
)
def test_trendline_types(out, trend, xml_type, extra):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].trendline.type = trend
    ser = first_series(chart_xml(save_chart(out, chart)))
    assert f'<c:trendlineType val="{xml_type}"/>{extra}' in ser


def test_trendline_options(out, reopen):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].trendline = px.ChartTrendline(
        type=px.ChartTrendlineType.Polynomial(3),
        name="Fit",
        forward_period=0.5,
        backward_period=0.5,
        display_equation=True,
        display_r_squared=True,
        intercept=1.5,
        delete_from_legend=True,
        label_font={"size": 8},
        label_format={"no_fill": True},
        format={"line": {"color": "green"}},
    )
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    ser = first_series(xml)
    assert "<c:trendline><c:name>Fit</c:name>" in ser
    assert '<a:srgbClr val="008000"/>' in ser
    assert (
        '<c:forward val="0.5"/><c:backward val="0.5"/><c:intercept val="1.5"/><c:dispRSqr val="1"/><c:dispEq val="1"/>'
        in ser
    )
    assert "<c:trendlineLbl>" in ser and 'sz="800"' in ser
    assert '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>' in xml


def test_trendline_type_value_class():
    t = px.ChartTrendlineType
    assert t.None_.kind == "None_" and t.Linear.kind == "Linear" and t.Linear.value is None
    assert t.Polynomial(4).value == 4 and t.MovingAverage(3).kind == "MovingAverage"
    assert t.Polynomial(2) == t.Polynomial(2) and t.Polynomial(2) != t.Polynomial(3)
    assert repr(t.Polynomial(2)) == "ChartTrendlineType.Polynomial(2)" and repr(t.Power) == "ChartTrendlineType.Power"
    assert len({t.Linear, t.Linear, t.Power}) == 2
    for bad in (1, 6):
        with pytest.raises(px.ParameterError, match="Polynomial order"):
            t.Polynomial(bad)
    with pytest.raises(px.ParameterError, match="MovingAverage period"):
        t.MovingAverage(4)
    assert px.ChartTrendline().type == t.None_


def test_error_bars(out, reopen):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].y_error_bars = px.ChartErrorBars(
        type=px.ChartErrorBarsType.Percentage(5),
        direction=px.ChartErrorBarsDirection.Plus,
        end_cap=False,
        format={"line": {"color": "gray"}},
    )
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    assert (
        '<c:errBars><c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="percentage"/><c:noEndCap val="1"/><c:val val="5"/>'
        in ser
    )
    assert '<a:srgbClr val="808080"/>' in ser


def test_error_bar_types_and_x_error_bars(out, reopen):
    chart = px.Chart(px.ChartType.Scatter)
    s = chart.add_series(values="=Sheet1!$B$1:$B$5", categories="=Sheet1!$A$1:$A$5")
    s.x_error_bars = px.ChartErrorBars(type=px.ChartErrorBarsType.Custom("=Sheet1!$C$1:$C$5", ("Sheet1", 0, 3, 4, 3)))
    s.y_error_bars = {"type": px.ChartErrorBarsType.FixedValue(2), "direction": px.ChartErrorBarsDirection.Minus}
    path = save_chart(out, chart)
    reopen(path)
    ser = first_series(chart_xml(path))
    assert '<c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="cust"/>' in ser
    assert (
        "<c:plus><c:numRef><c:f>Sheet1!$C$1:$C$5</c:f>" in ser
        and "<c:minus><c:numRef><c:f>Sheet1!$D$1:$D$5</c:f>" in ser
    )
    assert '<c:errDir val="y"/><c:errBarType val="minus"/><c:errValType val="fixedVal"/><c:val val="2"/>' in ser
    s.y_error_bars = px.ChartErrorBars(type=px.ChartErrorBarsType.StandardDeviation(1.5))
    s.x_error_bars = False
    ser = first_series(chart_xml(save_chart(out, chart, "sd.xlsx")))
    assert '<c:errValType val="stdDev"/><c:val val="1.5"/>' in ser and 'errDir val="x"' not in ser


def test_error_bars_true_means_standard_error(out):
    chart = column_chart(px.ChartType.Line)
    assert "<c:errBars>" not in chart_xml(save_chart(out, chart))
    chart.series[0].y_error_bars = True
    assert '<c:errValType val="stdErr"/>' in chart_xml(save_chart(out, chart, "e.xlsx"))


def test_error_bars_type_value_class():
    t = px.ChartErrorBarsType
    assert t.StandardError.kind == "StandardError" and t.StandardError.value is None
    assert t.FixedValue(2).value == 2 and t.Percentage(5).kind == "Percentage"
    custom = t.Custom("=Sheet1!$A$1:$A$2", ("Sheet1", 0, 1, 1, 1))
    assert custom.plus == "=Sheet1!$A$1:$A$2" and custom.minus == ("Sheet1", 0, 1, 1, 1) and custom.value is None
    assert t.FixedValue(2) == t.FixedValue(2.0) and t.FixedValue(2) != t.Percentage(2)
    assert repr(t.StandardDeviation(1.0)) == "ChartErrorBarsType.StandardDeviation(1.0)"
    assert "Custom('=Sheet1!$A$1:$A$2', ('Sheet1', 0, 1, 1, 1))" in repr(custom)
    assert len({t.StandardError, t.StandardError, t.FixedValue(1)}) == 2
    with pytest.raises(px.ParameterError):
        t.FixedValue(0)
    with pytest.raises(px.ParameterError):
        t.Percentage(-1)
    with pytest.raises(px.ParameterError):
        t.StandardDeviation(-0.5)
    assert px.ChartErrorBars().type == t.StandardError and px.ChartErrorBars().end_cap


def test_bar_options(out, reopen):
    chart = column_chart(series=2)
    s = chart.series[0]
    s.gap = 50
    s.overlap = 20
    s.invert_if_negative_color = "red"
    chart.series[1].delete_from_legend = True
    chart.series[1].invert_if_negative = True
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    plot = plot_area(xml)
    assert '<c:gapWidth val="50"/><c:overlap val="20"/>' in plot
    assert plot.count('<c:invertIfNegative val="1"/>') == 2
    assert "<c14:invertSolidFillFmt>" in plot and '<a:srgbClr val="FF0000"/>' in plot
    assert '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>' in xml


@pytest.mark.parametrize("prop,value", [("gap", 501), ("overlap", 101), ("overlap", -101)])
def test_bar_option_validation(prop, value):
    series = px.ChartSeries()
    with pytest.raises(px.ParameterError):
        setattr(series, prop, value)


def test_smooth_line(out):
    chart = column_chart(px.ChartType.Line)
    chart.series[0].smooth = True
    assert '<c:smooth val="1"/>' in first_series(chart_xml(save_chart(out, chart)))


def test_series_name_forms(out):
    chart = column_chart(series=3)
    chart.series[0].name = "Literal"
    chart.series[1].name = "=Sheet1!$B$1"
    chart.series[2].name = ("Sheet1", 0, 2)
    plot = plot_area(chart_xml(save_chart(out, chart)))
    assert "<c:tx><c:v>Literal</c:v></c:tx>" in plot
    assert "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>" in plot
    assert "<c:tx><c:strRef><c:f>Sheet1!$C$1</c:f>" in plot
