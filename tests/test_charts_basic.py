"""Charts: types, insertion, chart-level properties, live objects, compatibility, errors."""

from __future__ import annotations

import pytest
from charts_util import chart_xml, column_chart, plot_area, save_chart, workbook_with_data

import pyrust_xlsxwriter as px

CT = px.ChartType

# chart type -> fragments expected in the chart XML
TYPE_XML = {
    CT.Area: ["<c:areaChart>", '<c:grouping val="standard"/>'],
    CT.AreaStacked: ["<c:areaChart>", '<c:grouping val="stacked"/>'],
    CT.AreaPercentStacked: ["<c:areaChart>", '<c:grouping val="percentStacked"/>'],
    CT.Bar: ['<c:barDir val="bar"/>', '<c:grouping val="clustered"/>'],
    CT.BarStacked: ['<c:barDir val="bar"/>', '<c:grouping val="stacked"/>', '<c:overlap val="100"/>'],
    CT.BarPercentStacked: ['<c:barDir val="bar"/>', '<c:grouping val="percentStacked"/>'],
    CT.Column: ['<c:barDir val="col"/>', '<c:grouping val="clustered"/>'],
    CT.ColumnStacked: ['<c:barDir val="col"/>', '<c:grouping val="stacked"/>'],
    CT.ColumnPercentStacked: ['<c:barDir val="col"/>', '<c:grouping val="percentStacked"/>'],
    CT.Doughnut: ["<c:doughnutChart>", '<c:holeSize val="50"/>'],
    CT.Line: ["<c:lineChart>", '<c:grouping val="standard"/>'],
    CT.LineStacked: ["<c:lineChart>", '<c:grouping val="stacked"/>'],
    CT.LinePercentStacked: ["<c:lineChart>", '<c:grouping val="percentStacked"/>'],
    CT.Pie: ["<c:pieChart>"],
    CT.Radar: ["<c:radarChart>", '<c:radarStyle val="marker"/>', '<c:symbol val="none"/>'],
    CT.RadarWithMarkers: ["<c:radarChart>", '<c:radarStyle val="marker"/>'],
    CT.RadarFilled: ["<c:radarChart>", '<c:radarStyle val="filled"/>'],
    CT.Scatter: ["<c:scatterChart>", '<c:scatterStyle val="lineMarker"/>'],
    CT.ScatterStraight: ["<c:scatterChart>", '<c:scatterStyle val="lineMarker"/>', '<c:symbol val="none"/>'],
    CT.ScatterStraightWithMarkers: ["<c:scatterChart>", '<c:scatterStyle val="lineMarker"/>'],
    CT.ScatterSmooth: ["<c:scatterChart>", '<c:scatterStyle val="smoothMarker"/>', '<c:symbol val="none"/>'],
    CT.ScatterSmoothWithMarkers: ["<c:scatterChart>", '<c:scatterStyle val="smoothMarker"/>'],
    CT.Stock: ["<c:stockChart>", "<c:dateAx>"],
}


def test_all_23_chart_types_are_exposed():
    names = [n for n in dir(CT) if not n.startswith("_")]
    assert len(TYPE_XML) == 23
    assert set(TYPE_XML) == {getattr(CT, n) for n in names}


@pytest.mark.parametrize("chart_type", list(TYPE_XML), ids=lambda t: repr(t).split(".")[-1])
def test_chart_types(out, reopen, chart_type):
    chart = px.Chart(chart_type)
    if chart_type == CT.Stock:
        for col in "BCD":
            chart.add_series(values=f"=Sheet1!${col}$1:${col}$5", categories="=Sheet1!$A$1:$A$5")
    else:
        chart.add_series(values="=Sheet1!$B$1:$B$5", categories="=Sheet1!$A$1:$A$5")
    assert chart.chart_type == chart_type
    path = save_chart(out, chart)
    wb = reopen(path)
    assert len(wb["Sheet1"]._charts) == 1
    xml = chart_xml(path)
    for fragment in TYPE_XML[chart_type]:
        assert fragment in xml, fragment


def test_insert_chart_forms(out, reopen, part):
    wb, ws = workbook_with_data()
    chart = column_chart()
    ws.insert_chart(1, 4, chart)  # row, col
    ws.insert_chart("E20", chart, x_offset=10, y_offset=5)
    ws.insert_chart(40, "E", chart, 3, 4)  # column letter, positional offsets
    ws.insert_chart_with_offset(60, 4, chart, 7, 8)  # old API
    ws.insert_chart_with_offset("E80", chart=chart, x_offset=1, y_offset=2)
    path = out()
    wb.save(path)
    assert len(reopen(path)["Sheet1"]._charts) == 5
    drawing = part(path, "xl/drawings/drawing1.xml")
    assert drawing.count("<xdr:graphicFrame") == 5
    assert "<xdr:row>19</xdr:row>" in drawing and "<xdr:colOff>95250</xdr:colOff>" in drawing
    assert "<xdr:row>40</xdr:row>" in drawing
    for n in range(1, 6):
        assert "<c:barChart>" in part(path, f"xl/charts/chart{n}.xml")


def test_insert_chart_bad_arguments():
    _, ws = workbook_with_data()
    chart = column_chart()
    with pytest.raises(TypeError, match="must be a Chart"):
        ws.insert_chart(0, 0, "not a chart")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError, match="missing required argument: 'chart'"):
        ws.insert_chart("A1")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.ParameterError, match="invalid cell reference"):
        ws.insert_chart("A0", chart)
    with pytest.raises(px.RangeError):
        ws.insert_chart(2_000_000, 0, chart)


def test_backward_compatibility(out, reopen):
    """The original minimal API keeps working."""
    chart = px.Chart(chart_type=px.ChartType.Column)
    chart.add_series(values="=Sheet1!$B$1:$B$5", categories="=Sheet1!$A$1:$A$5", name="Sales")
    chart.add_series("=Sheet1!$C$1:$C$5", "", "")
    chart.title = "Fruit Sales"
    chart.x_axis_name = "Fruit"
    chart.y_axis_name = "Amount"
    chart.width = 600
    chart.height = 400
    chart.style = 10
    chart.legend_position = px.ChartLegendPosition.Bottom
    assert str(chart.title) == "Fruit Sales"
    assert chart.title.name == "Fruit Sales"
    assert chart.x_axis_name == "Fruit" and chart.x_axis.name == "Fruit"
    assert chart.y_axis_name == "Amount"
    assert (chart.width, chart.height, chart.style) == (600, 400, 10)
    assert chart.legend_position == px.ChartLegendPosition.Bottom
    assert chart.legend.position == px.ChartLegendPosition.Bottom
    assert chart.series[1].values == "=Sheet1!$C$1:$C$5" and chart.series[1].name is None
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    for text in ["Fruit Sales", "Fruit", "Amount", '<c:legendPos val="b"/>', '<c:style val="10"/>', "Sheet1!$C$1:$C$5"]:
        assert text in xml


def test_defaults():
    chart = px.Chart()
    assert chart.chart_type == CT.Column
    assert (chart.width, chart.height, chart.style) == (480, 288, 2)
    assert (chart.scale_width, chart.scale_height) == (1.0, 1.0)
    assert chart.object_movement == px.ObjectMovement.MoveAndSizeWithCells
    assert chart.series == [] and chart.combined_chart is None
    assert chart.title.name is None and str(chart.title) == ""
    assert chart.x_axis_name == "" and chart.legend_position == px.ChartLegendPosition.Right
    assert chart.rotation is None and chart.hole_size is None
    assert not (chart.up_down_bars or chart.high_low_lines or chart.drop_lines)


def test_constructor_keywords_and_repr():
    chart = px.Chart(
        CT.Pie,
        width=500,
        height=300,
        rotation=90,
        title="Pie",
        name="My pie",
        alt_text="alt",
        legend={"position": px.ChartLegendPosition.Left},
    )
    assert chart.rotation == 90 and chart.name == "My pie" and chart.alt_text == "alt"
    assert chart.title.name == "Pie"
    assert chart.legend.position == px.ChartLegendPosition.Left
    text = repr(chart)
    assert text.startswith('Chart(ChartType.Pie, title="Pie", series=0')
    assert "rotation=90" in text and "legend=ChartLegend(position=ChartLegendPosition.Left)" in text
    with pytest.raises(TypeError, match="unexpected keyword argument 'colour'"):
        px.Chart(colour="red")  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(TypeError, match="Chart.width"):
        px.Chart(width="wide")  # pyrefly: ignore[bad-argument-type]


def test_nested_objects_are_live(out):
    chart = column_chart()
    assert chart.x_axis is chart.x_axis
    assert chart.title is chart.title and chart.legend is chart.legend
    s = chart.series[0]
    assert s is chart.series[0]
    assert s.format.line is s.format.line
    axis = chart.y_axis
    axis.min = 0
    axis.max = 50
    chart.title = "Live"
    title = chart.title
    title.font.color = "#FF0000"
    s.marker.size = 7  # bar charts ignore markers, but the object is live
    assert chart.y_axis.max == 50 and chart.series[0].marker.size == 7
    xml = chart_xml(save_chart(out, chart))
    assert '<c:max val="50"/>' in xml and '<c:min val="0"/>' in xml
    assert "Live" in xml and 'val="FF0000"' in xml


def test_assign_objects_dicts_and_none():
    chart = px.Chart()
    axis = px.ChartAxis(name="X", reverse=True)
    chart.x_axis = axis
    assert chart.x_axis is axis
    chart.y_axis = {"name": "Y", "min": 1}
    assert chart.y_axis.name == "Y" and chart.y_axis.min == 1
    chart.y_axis = None  # reset to a fresh default axis
    assert chart.y_axis.name is None and chart.y_axis is not axis
    title = px.ChartTitle("T", overlay=True)
    chart.title = title
    assert chart.title is title
    chart.title = ("Sheet1", 0, 0)  # sets the name of the existing title
    assert chart.title is title and title.name == ("Sheet1", 0, 0)
    chart.title = None
    assert chart.title.name is None
    with pytest.raises(TypeError, match="expected a ChartAxis or a dict"):
        chart.x_axis = 5  # pyrefly: ignore[bad-argument-type]


def test_chart_level_properties(out, reopen, part):
    chart = column_chart()
    chart.width = 720
    chart.height = 360
    chart.scale_width = 1.5
    chart.scale_height = 0.5
    chart.style = 42
    chart.name = "Revenue chart"
    chart.alt_text = "Revenue by month"
    chart.object_movement = px.ObjectMovement.DontMoveOrSizeWithCells
    chart.show_empty_cells_as = px.ChartEmptyCells.Connected
    chart.show_na_as_empty_cell = True
    chart.show_hidden_data = True
    path = save_chart(out, chart)
    reopen(path)
    xml = chart_xml(path)
    assert '<c:style val="42"/>' in xml
    assert '<c:dispBlanksAs val="span"/>' in xml
    assert "plotVisOnly" not in xml  # rust_xlsxwriter drops "plot visible cells only"
    assert 'c16r3:dispNaAsBlank val="1"' in xml
    drawing = part(path, "xl/drawings/drawing1.xml")
    assert 'name="Revenue chart"' in drawing and 'descr="Revenue by month"' in drawing
    assert 'editAs="absolute"' in drawing
    # 720 * 1.5 = 1080 px wide from column E (16 columns of 64 px + 56 px),
    # 360 * 0.5 = 180 px high from row 2 (9 rows of 20 px); 9525 EMU per pixel.
    to = "<xdr:to><xdr:col>20</xdr:col><xdr:colOff>533400</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
    assert to in drawing


def test_decorative(out, part):
    chart = column_chart()
    chart.decorative = True
    path = save_chart(out, chart)
    assert "decorative" in part(path, "xl/drawings/drawing1.xml")


@pytest.mark.parametrize(
    "chart_type,prop,value,fragment",
    [
        (CT.Pie, "rotation", 90, '<c:firstSliceAng val="90"/>'),
        (CT.Doughnut, "rotation", 360, '<c:firstSliceAng val="360"/>'),
        (CT.Doughnut, "hole_size", 75, '<c:holeSize val="75"/>'),
    ],
)
def test_pie_and_doughnut_options(out, chart_type, prop, value, fragment):
    chart = column_chart(chart_type)
    setattr(chart, prop, value)
    assert fragment in chart_xml(save_chart(out, chart))


@pytest.mark.parametrize(
    "prop,value",
    [
        ("width", 0),
        ("height", 0),
        ("scale_width", 0.0),
        ("scale_height", -1.0),
        ("style", 0),
        ("style", 49),
        ("rotation", 361),
        ("hole_size", 91),
        ("alt_text", "x" * 256),
    ],
)
def test_chart_property_validation(prop, value):
    chart = px.Chart()
    before = getattr(chart, prop)
    with pytest.raises(px.ParameterError):
        setattr(chart, prop, value)
    assert getattr(chart, prop) == before  # an invalid value is not stored


def test_chart_validation_errors():
    _, ws = workbook_with_data()
    with pytest.raises(px.ParameterError, match="at least one series"):
        ws.insert_chart(0, 0, px.Chart())
    chart = px.Chart()
    chart.add_series(categories="=Sheet1!$A$1:$A$5")
    with pytest.raises(px.ParameterError, match="'values' range"):
        ws.insert_chart(0, 0, chart)
    with pytest.raises(px.ParameterError, match="'values' range"):
        chart.validate()
    scatter = px.Chart(CT.Scatter)
    scatter.add_series(values="=Sheet1!$B$1:$B$5")
    with pytest.raises(px.ParameterError, match="'categories' range"):
        ws.insert_chart(0, 0, scatter)
    ok = column_chart()
    ok.validate()


def test_range_validation():
    chart = px.Chart()
    with pytest.raises(px.ParameterError, match="sheet name"):
        chart.add_series(values="A1:A5")
    with pytest.raises(px.RangeError, match="reversed"):
        chart.add_series(values=("Sheet1", 4, 0, 0, 0))
    with pytest.raises(px.RangeError):
        chart.add_series(values=("Sheet1", 0, 0, 2_000_000, 0))
    with pytest.raises(px.ParameterError, match="non-empty sheet name"):
        chart.add_series(values=("", 0, 0, 4, 0))
    with pytest.raises(TypeError, match="expected a range string"):
        chart.add_series(values=("Sheet1", 0))  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="expected a range string"):
        chart.add_series(values=5)  # pyrefly: ignore[bad-argument-type]
    s = chart.add_series(values=("My Sheet", 0, 1, 4, 1))
    with pytest.raises(px.ParameterError, match="sheet name"):
        s.categories = "B1:B5"
    assert s.values == ("My Sheet", 0, 1, 4, 1)


def test_tuple_ranges_and_quoted_sheet_names(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet("My Data")
    ws.write_rows("A1", [[i, i * i] for i in range(1, 6)])
    chart = px.Chart(CT.Line)
    chart.add_series(values=("My Data", 0, 1, 4, 1), categories=("My Data", 0, 0, 4, 0), name=("My Data", 0, 0))
    ws.insert_chart("D2", chart)
    path = out()
    wb.save(path)
    reopen(path)
    xml = chart_xml(path)
    assert "<c:f>'My Data'!$B$1:$B$5</c:f>" in xml
    assert "<c:f>'My Data'!$A$1:$A$5</c:f>" in xml
    assert "<c:f>'My Data'!$A$1</c:f>" in xml


def test_series_list_and_push_series(out):
    chart = px.Chart()
    s1 = chart.add_series(values="=Sheet1!$A$1:$A$5", name="first")
    s2 = px.ChartSeries("=Sheet1!$B$1:$B$5", name="second")
    chart.push_series(s2)
    assert chart.series == [s1, s2]
    s2.name = "renamed after push"  # pushed series stay live
    xml = chart_xml(save_chart(out, chart))
    assert "renamed after push" in xml and "first" in xml
    chart.series = [s2]
    assert chart.series == [s2]
    xml = chart_xml(save_chart(out, chart, "again.xlsx"))
    assert "first" not in xml


def test_chart_can_be_changed_and_reinserted(out, reopen):
    wb, ws = workbook_with_data()
    chart = column_chart()
    chart.title = "First"
    ws.insert_chart("E2", chart)
    chart.title = "Second"
    chart.chart_type = CT.Line
    ws.insert_chart("E20", chart)
    path = out()
    wb.save(path)
    reopen(path)
    one, two = chart_xml(path, 1), chart_xml(path, 2)
    assert "First" in one and "<c:barChart>" in one
    assert "Second" in two and "<c:lineChart>" in two


def test_series_constructor_and_add_series_keywords():
    chart = px.Chart()
    s = chart.add_series(
        values="=Sheet1!$A$1:$A$5",
        secondary_axis=True,
        gap=200,
        overlap=-10,
        smooth=True,
        format={"solid_fill": {"color": "red"}},
        marker={"type": px.ChartMarkerType.Diamond},
    )
    assert s.secondary_axis and s.gap == 200 and s.overlap == -10 and s.smooth
    assert s.format.solid_fill.color == px.Color.Red
    assert s.marker.type == px.ChartMarkerType.Diamond
    assert "ChartSeries(values='=Sheet1!$A$1:$A$5'" in repr(s)
    copy = s.copy()
    copy.format.solid_fill.color = "blue"
    assert s.format.solid_fill.color == px.Color.Red
    with pytest.raises(TypeError, match="unexpected keyword argument 'value'"):
        px.ChartSeries(value="=Sheet1!$A$1:$A$5")  # pyrefly: ignore[unexpected-keyword]


def test_plot_area_only_has_expected_series(out):
    chart = column_chart(series=3)
    xml = plot_area(chart_xml(save_chart(out, chart)))
    assert xml.count("<c:ser>") == 3
