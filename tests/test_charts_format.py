"""Chart formatting: lines, fills, ChartFormat, fonts, layouts, title, legend, areas, data table."""

from __future__ import annotations

import pytest
from charts_util import chart_xml, column_chart, plot_area, save_chart

import pyrust_xlsxwriter as px


def series_sppr(xml: str, index: int = 0) -> str:
    """The <c:spPr> of the index-th series."""
    ser = plot_area(xml).split("<c:ser>")[index + 1]
    return ser[ser.index("<c:spPr>") : ser.index("</c:spPr>")]


def test_chart_line(out, reopen):
    chart = column_chart(series=2)
    line = px.ChartLine(color="red", width=2.25, dash_type=px.ChartLineDashType.Dash, transparency=50)
    assert (line.color, line.width, line.transparency) == (px.Color.Red, 2.25, 50)
    assert repr(line) == "ChartLine(color=Color.Red, width=2.25, transparency=50, dash_type=ChartLineDashType.Dash)"
    chart.series[0].format.line = line
    chart.series[1].format.border = px.ChartLine(hidden=True)
    assert chart.series[1].format.line is chart.series[1].format.border
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    sp0 = series_sppr(xml, 0)
    assert '<a:ln w="28575">' in sp0
    assert '<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>' in sp0
    assert '<a:prstDash val="dash"/>' in sp0
    assert "<a:ln><a:noFill/></a:ln>" in series_sppr(xml, 1)


@pytest.mark.parametrize("dash", [n for n in dir(px.ChartLineDashType) if not n.startswith("_")])
def test_all_dash_types(out, dash):
    chart = column_chart()
    chart.series[0].format.line.dash_type = getattr(px.ChartLineDashType, dash)
    xml = chart_xml(save_chart(out, chart))
    if dash == "Solid":
        assert "prstDash" not in xml
    else:
        assert "<a:prstDash val=" in xml


def test_solid_fill_and_theme_colors(out):
    chart = column_chart(series=2)
    chart.series[0].format.solid_fill = px.ChartSolidFill(color=px.Color.theme(4, 1), transparency=25)
    chart.series[1].format = px.ChartSolidFill(color=(0, 128, 255))  # a bare fill becomes a ChartFormat
    assert isinstance(chart.series[1].format, px.ChartFormat)
    assert chart.series[1].format.solid_fill.color == px.Color("#0080FF")
    xml = chart_xml(save_chart(out, chart))
    assert (
        '<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/><a:alpha val="75000"/>'
        in series_sppr(xml, 0)
    )
    assert '<a:solidFill><a:srgbClr val="0080FF"/></a:solidFill>' in series_sppr(xml, 1)


def test_pattern_fill(out):
    chart = column_chart()
    chart.series[0].format = px.ChartPatternFill(
        pattern=px.ChartPatternFillType.Weave, foreground_color="#00FF00", background_color="yellow"
    )
    xml = chart_xml(save_chart(out, chart))
    assert (
        '<a:pattFill prst="weave"><a:fgClr><a:srgbClr val="00FF00"/></a:fgClr><a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr></a:pattFill>'
        in xml
    )


def test_all_pattern_types_build(out):
    names = [n for n in dir(px.ChartPatternFillType) if not n.startswith("_")]
    assert len(names) == 48
    chart = column_chart()
    for name in names:
        chart.series[0].format.pattern_fill.pattern = getattr(px.ChartPatternFillType, name)
        chart.validate()
    assert "<a:pattFill" in chart_xml(save_chart(out, chart))


def test_gradient_fill(out, reopen):
    chart = column_chart(series=2)
    stops = [
        px.ChartGradientStop("#FF0000", 0),
        px.ChartGradientStop("#00FF00", 50),
        px.ChartGradientStop("#0000FF", 100),
    ]
    linear = px.ChartGradientFill(gradient_stops=stops, angle=45)
    assert linear.gradient_stops == stops and linear.angle == 45 and linear.type is None
    chart.series[0].format.gradient_fill = linear
    chart.series[1].format = px.ChartGradientFill(type=px.ChartGradientFillType.Radial, gradient_stops=stops[:2])
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    sp0 = series_sppr(xml, 0)
    assert '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' in sp0
    assert '<a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>' in sp0
    assert '<a:lin ang="2700000" scaled="0"/>' in sp0
    assert '<a:path path="circle">' in series_sppr(xml, 1)


@pytest.mark.parametrize(
    "kind,fragment",
    [
        (px.ChartGradientFillType.Linear, "<a:lin "),
        (px.ChartGradientFillType.Radial, 'path="circle"'),
        (px.ChartGradientFillType.Rectangular, 'path="rect"'),
        (px.ChartGradientFillType.Path, 'path="shape"'),
    ],
)
def test_gradient_types(out, kind, fragment):
    chart = column_chart()
    chart.series[0].format.gradient_fill.type = kind
    chart.series[0].format.gradient_fill.gradient_stops = [
        px.ChartGradientStop("red", 0),
        px.ChartGradientStop("blue", 100),
    ]
    assert fragment in chart_xml(save_chart(out, chart))


def test_gradient_stop_value_class():
    stop = px.ChartGradientStop("red", 40)
    assert stop.color == px.Color.Red and stop.position == 40
    assert stop == px.ChartGradientStop(px.Color.Red, 40) and hash(stop) == hash(px.ChartGradientStop("red", 40))
    assert repr(stop) == "ChartGradientStop(Color.Red, 40)"
    with pytest.raises(AttributeError):
        stop.position = 5  # pyrefly: ignore[read-only]
    with pytest.raises(px.ParameterError):
        px.ChartGradientStop("red", 101)


def test_no_fill_no_line_and_aliases(out):
    chart = column_chart()
    fmt = px.ChartFormat(no_fill=True, no_border=True)
    assert fmt.no_line and fmt.no_border and fmt.no_fill
    chart.series[0].format = fmt
    assert "<a:noFill/><a:ln><a:noFill/></a:ln>" in series_sppr(chart_xml(save_chart(out, chart)))


def test_format_coercion_and_reset():
    series = px.ChartSeries("=Sheet1!$A$1:$A$5")
    line = px.ChartLine(color="red")
    series.format = line
    assert series.format.line is line  # wrapped, the line stays live
    series.format = {"line": {"width": 3}, "solid_fill": {"color": "blue"}}
    assert series.format.line.width == 3 and series.format.solid_fill.color == px.Color.Blue
    series.format = None
    assert series.format.line.width is None
    with pytest.raises(TypeError, match="expected a ChartFormat or a dict"):
        series.format = "red"  # pyrefly: ignore[bad-argument-type]
    assert "format=" not in repr(series)  # an untouched format is not used


@pytest.mark.parametrize(
    "cls,prop,value",
    [
        (px.ChartLine, "transparency", 101),
        (px.ChartLine, "width", 2000.0),
        (px.ChartLine, "width", -1.0),
        (px.ChartSolidFill, "transparency", 150),
        (px.ChartGradientFill, "angle", 360),
        (px.ChartGradientFill, "gradient_stops", [px.ChartGradientStop("red", 0)]),
        (px.ChartFont, "size", 0.0),
        (px.ChartFont, "rotation", 100),
        (px.ChartLayout, "offset", (1.5, 0.0)),
        (px.ChartLayout, "dimensions", (0.5, -0.1)),
    ],
)
def test_format_validation(cls, prop, value):
    obj = cls()
    with pytest.raises(px.ParameterError):
        setattr(obj, prop, value)
    with pytest.raises(px.ParameterError):
        cls(**{prop: value})


def test_color_inputs():
    line = px.ChartLine()
    for value, expected in [
        ("red", px.Color.Red),
        ("#123456", px.Color("#123456")),
        (0x00FF00, px.Color(0x00FF00)),
        ((1, 2, 3), px.Color.rgb(1, 2, 3)),
    ]:
        line.color = value
        assert line.color == expected
    with pytest.raises(px.ParameterError):
        line.color = "notacolor"


def test_font(out, reopen):
    chart = column_chart()
    chart.title = "Title"
    chart.title.font = px.ChartFont(
        bold=False,
        italic=True,
        size=14,
        color="navy",
        name="Arial",
        rotation=-45,
        underline=True,
        strikethrough=True,
        pitch_family=34,
        character_set=0,
    )
    font = chart.title.font
    assert font.bold is False and font.italic and font.size == 14 and font.name == "Arial"
    chart.legend.font.bold = True
    chart.x_axis.font = {"color": "#FF0000", "rotation": 270}
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    assert '<a:bodyPr rot="-2700000" vert="horz"/>' in xml
    assert '<a:defRPr sz="1400" b="0" i="1" u="sng" baseline="0">' in xml
    assert '<a:srgbClr val="000080"/>' in xml and '<a:latin typeface="Arial" pitchFamily="34"' in xml
    assert '<a:defRPr b="1" i="0"/>' in xml  # legend font
    assert 'vert="wordArtVert"' in xml  # x axis rotation 270: stacked text


def test_title(out, reopen):
    chart = column_chart()
    chart.title = px.ChartTitle("Sales", overlay=True, format={"solid_fill": {"color": "#FFFF00"}})
    chart.title.layout = px.ChartLayout(offset=(0.1, 0.2))
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    title = xml[xml.index("<c:title>") : xml.index("</c:title>")]
    assert "<a:t>Sales</a:t>" in title
    assert '<c:x val="0.1"/><c:y val="0.2"/>' in title
    assert '<c:overlay val="1"/>' in title
    assert '<a:srgbClr val="FFFF00"/>' in title


def test_title_from_cell_and_hidden(out):
    chart = column_chart()
    chart.title.name = "=Sheet1!$A$1"
    xml = chart_xml(save_chart(out, chart))
    assert "<c:strRef><c:f>Sheet1!$A$1</c:f>" in xml
    chart = column_chart()
    chart.series[0].name = "Only series"
    chart.title.hidden = True
    xml = chart_xml(save_chart(out, chart, "hidden.xlsx"))
    assert '<c:autoTitleDeleted val="1"/>' in xml


@pytest.mark.parametrize(
    "position,code",
    [
        (px.ChartLegendPosition.Right, "r"),
        (px.ChartLegendPosition.Left, "l"),
        (px.ChartLegendPosition.Top, "t"),
        (px.ChartLegendPosition.Bottom, "b"),
        (px.ChartLegendPosition.TopRight, "tr"),
    ],
)
def test_legend_positions(out, position, code):
    chart = column_chart()
    chart.legend.position = position
    assert f'<c:legendPos val="{code}"/>' in chart_xml(save_chart(out, chart))


def test_legend_options(out, reopen):
    chart = column_chart(series=3)
    legend = chart.legend
    legend.delete_entries = [0, 2]
    legend.overlay = True
    legend.layout = px.ChartLayout(offset=(0.8, 0.1), dimensions=(0.2, 0.5))
    legend.format.border.color = "#000000"
    legend.font = px.ChartFont(size=9)
    assert legend.delete_entries == [0, 2]
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    leg = xml[xml.index("<c:legend>") : xml.index("</c:legend>")]
    assert '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>' in leg
    assert '<c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry>' in leg
    assert '<c:overlay val="1"/>' in leg and '<c:w val="0.2"/><c:h val="0.5"/>' in leg
    assert '<a:srgbClr val="000000"/>' in leg and 'sz="900"' in leg


def test_legend_hidden(out):
    chart = column_chart()
    chart.legend.hidden = True
    assert "<c:legend>" not in chart_xml(save_chart(out, chart))


def test_chart_area_and_plot_area(out, reopen):
    chart = column_chart()
    chart.chart_area.format = px.ChartFormat(border=px.ChartLine(color="black"), solid_fill={"color": "#F0F0F0"})
    chart.plot_area.format.solid_fill.color = "#EEEEEE"
    chart.plot_area.layout.offset = (0.1, 0.1)
    chart.plot_area.layout.dimensions = (0.6, 0.6)
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    after_chart = xml[xml.index("</c:chart>") :]
    assert (
        '<c:spPr><a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/>'
        in after_chart
    )
    plot = plot_area(xml)
    assert '<c:layoutTarget val="inner"/>' in plot and '<c:w val="0.6"/>' in plot
    assert '<c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr>' in plot


def test_data_table(out, reopen):
    chart = column_chart()
    assert "<c:dTable>" not in chart_xml(save_chart(out, chart))
    chart.data_table = True
    xml = chart_xml(save_chart(out, chart, "t1.xlsx"))
    assert '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/></c:dTable>' in xml
    chart.data_table = px.ChartDataTable(show_vertical_borders=False, show_legend_keys=True, font={"italic": True})
    table = chart.data_table
    assert table.show_horizontal_borders and not table.show_vertical_borders and table.show_legend_keys
    path = save_chart(out, chart, "t2.xlsx")
    reopen(path)
    xml = chart_xml(path)
    assert "<c:showVertBorder" not in xml and '<c:showKeys val="1"/>' in xml and '<a:defRPr i="1"/>' in xml
    chart.data_table = False
    assert "<c:dTable>" not in chart_xml(save_chart(out, chart, "t3.xlsx"))
    chart.data_table.show_legend_keys = True  # changing a property also shows the table
    assert "<c:dTable>" in chart_xml(save_chart(out, chart, "t4.xlsx"))
