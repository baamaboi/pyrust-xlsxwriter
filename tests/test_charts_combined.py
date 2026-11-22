"""Combined charts, secondary axes, chartsheets, stock-chart lines/bars and Excel corpus comparisons."""

from __future__ import annotations

import datetime as dt

import pytest
from charts_util import chart_xml, column_chart, normalized, plot_area, save_chart, workbook_with_data

import pyrust_xlsxwriter as px

STOCK_ROWS = [
    (39083, 27.2, 23.49, 25.45),
    (39084, 25.03, 19.55, 23.05),
    (39085, 19.05, 15.12, 17.32),
    (39086, 20.34, 17.84, 20.45),
    (39087, 18.5, 16.34, 17.34),
]


def stock_workbook() -> tuple[px.Workbook, px.Worksheet]:
    wb = px.Workbook()
    ws = wb.add_worksheet()
    date_format = px.Format(num_format="dd/mm/yyyy")
    for row, (serial, high, low, close) in enumerate(STOCK_ROWS):
        ws.write(row, 0, dt.date(1899, 12, 30) + dt.timedelta(days=serial), date_format)
        ws.write_row(row, 1, [high, low, close])
    return wb, ws


def assert_matches_corpus(path, corpus, name: str) -> None:
    assert normalized(chart_xml(path)) == normalized(chart_xml(corpus / f"{name}.xlsx"))


# ---------------------------------------------------------------------------
# Combined charts and secondary axes
# ---------------------------------------------------------------------------


def test_combined_column_and_line_on_secondary_axis(out, reopen):
    column = column_chart(series=1)
    line = px.Chart(px.ChartType.Line)
    line.add_series(values="=Sheet1!$B$1:$B$5", secondary_axis=True, name="Line")
    line.y2_axis.name = "ignored"  # axes of a combined chart are set on the primary chart
    column.y2_axis.name = "Secondary"
    column.combine(line)
    assert column.combined_chart is line
    path = save_chart(out, column)
    reopen(path)
    plot = plot_area(chart_xml(path))
    assert "<c:barChart>" in plot and "<c:lineChart>" in plot
    assert plot.count("<c:valAx>") == 2
    assert '<c:axPos val="r"/>' in plot and '<c:crosses val="max"/>' in plot
    assert "<a:t>Secondary</a:t>" in plot and "ignored" not in plot


def test_combined_chart_on_same_axis(out):
    column = column_chart()
    line = px.Chart(px.ChartType.Line)
    line.add_series(values="=Sheet1!$C$1:$C$5")
    column.combined_chart = line
    plot = plot_area(chart_xml(save_chart(out, column)))
    assert plot.count("<c:valAx>") == 1 and "<c:lineChart>" in plot
    # the combined chart is live too
    line.series[0].format.line.color = "red"
    assert '<a:srgbClr val="FF0000"/>' in chart_xml(save_chart(out, column, "live.xlsx"))


def test_combine_errors_and_removal(out):
    chart = column_chart()
    with pytest.raises(px.ParameterError, match="itself"):
        chart.combine(chart)
    with pytest.raises(px.ParameterError, match="itself"):
        chart.combined_chart = chart
    other = column_chart(px.ChartType.Line)
    third = column_chart(px.ChartType.Area)
    other.combine(third)  # nested combinations are ignored (Excel supports one level)
    chart.combine(other)
    plot = plot_area(chart_xml(save_chart(out, chart)))
    assert "<c:lineChart>" in plot and "<c:areaChart>" not in plot
    other.combine(chart)  # a cycle does not recurse
    chart_xml(save_chart(out, chart, "cycle.xlsx"))
    chart.combine(None)
    assert chart.combined_chart is None
    assert "<c:lineChart>" not in chart_xml(save_chart(out, chart, "none.xlsx"))


def test_secondary_axis_bar_chart(out, reopen):
    chart = column_chart(px.ChartType.Bar, series=2)
    chart.series[1].secondary_axis = True
    chart.x2_axis.name = "Top axis"
    path = save_chart(out, chart)
    reopen(path)
    plot = plot_area(chart_xml(path))
    assert plot.count("<c:barChart>") == 2 and "<a:t>Top axis</a:t>" in plot


# ---------------------------------------------------------------------------
# Up/down bars, high-low lines, drop lines
# ---------------------------------------------------------------------------


def test_up_down_bars_high_low_and_drop_lines(out, reopen):
    chart = column_chart(px.ChartType.Line, series=2)
    chart.up_down_bars = True
    chart.high_low_lines = True
    chart.drop_lines = True
    path = save_chart(out, chart)
    reopen(path)
    line = plot_area(chart_xml(path))
    assert "<c:dropLines/>" in line and "<c:hiLowLines/>" in line
    assert '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' in line


def test_line_and_bar_formats_enable_them(out):
    chart = column_chart(px.ChartType.Line, series=2)
    assert not chart.up_down_bars
    chart.up_bar_format.solid_fill.color = "#00FF00"
    chart.down_bar_format = px.ChartSolidFill(color="#FF0000")
    chart.high_low_lines_format = px.ChartLine(color="#0000FF")
    chart.drop_lines_format.line.dash_type = px.ChartLineDashType.Dash
    assert chart.up_down_bars and chart.high_low_lines and chart.drop_lines
    line = plot_area(chart_xml(save_chart(out, chart)))
    assert '<c:upBars><c:spPr><a:solidFill><a:srgbClr val="00FF00"/>' in line
    assert '<c:downBars><c:spPr><a:solidFill><a:srgbClr val="FF0000"/>' in line
    assert '<c:hiLowLines><c:spPr><a:ln><a:solidFill><a:srgbClr val="0000FF"/>' in line
    assert '<c:dropLines><c:spPr><a:ln><a:prstDash val="dash"/>' in line
    chart.up_bar_format = None
    chart.down_bar_format = None
    assert not chart.up_down_bars


# ---------------------------------------------------------------------------
# Chartsheets
# ---------------------------------------------------------------------------


def test_chartsheet(out, reopen, part):
    wb, _ = workbook_with_data()
    cs = wb.add_chartsheet()
    assert cs.name == "Chart1" and wb.sheet_names == ["Sheet1", "Chart1"]
    assert wb["Chart1"] is cs
    chart = column_chart()
    chart.title = "On a chartsheet"
    cs.insert_chart(chart)
    named = wb.add_chartsheet("Summary", tab_color="#FF0000")
    named.insert_chart(5, 5, chart)  # the position is ignored on a chartsheet
    path = out()
    wb.save(path)
    loaded = reopen(path)
    assert loaded.sheetnames == ["Sheet1", "Chart1", "Summary"]
    assert [c.title for c in loaded.chartsheets] == ["Chart1", "Summary"]
    workbook_xml = part(path, "xl/workbook.xml")
    assert 'name="Chart1"' in workbook_xml
    assert "chartsheet+xml" in part(path, "[Content_Types].xml")
    sheet = part(path, "xl/chartsheets/sheet2.xml")
    assert '<tabColor rgb="FFFF0000"/>' in sheet
    assert "<xdr:absoluteAnchor>" in part(path, "xl/drawings/drawing2.xml")
    assert "On a chartsheet" in chart_xml(path)


def test_chartsheet_names():
    wb = px.Workbook()
    wb.add_worksheet("Chart1")
    assert wb.add_chartsheet().name == "Chart2"
    with pytest.raises(px.SheetNameError, match="already used"):
        wb.add_chartsheet("chart2")
    with pytest.raises(px.SheetNameError):
        wb.add_chartsheet("bad[name]")


def test_corpus_chartsheet10(out, corpus):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_rows("A1", [[1, 2, 3], [2, 4, 6], [3, 6, 9], [4, 8, 12], [5, 10, 15]])
    cs = wb.add_chartsheet()
    chart = px.Chart(px.ChartType.Bar)
    for col in "ABC":
        chart.add_series(values=f"=Sheet1!${col}$1:${col}$5")
    cs.insert_chart(chart)
    path = out()
    wb.save(path)
    assert_matches_corpus(path, corpus, "chartsheet10")


# ---------------------------------------------------------------------------
# Corpus comparisons (Excel-made files)
# ---------------------------------------------------------------------------


def test_corpus_combined_doughnut_and_pie(out, reopen, corpus):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write("H2", "Donut")
    ws.write("I2", "Pie")
    ws.write_column("H3", [25, 50, 25, 100])
    ws.write_column("I3", [75, 1, 124])
    doughnut = px.Chart(px.ChartType.Doughnut, rotation=270, show_empty_cells_as=px.ChartEmptyCells.Gaps)
    s = doughnut.add_series(values="=Sheet1!$H$3:$H$6", name="=Sheet1!$H$2")
    s.points = [
        px.ChartPoint(format=px.ChartSolidFill(color="#FF0000")),
        px.ChartPoint(format=px.ChartSolidFill(color="#FFC000")),
        px.ChartPoint(format=px.ChartSolidFill(color="#00B050")),
        px.ChartPoint(format={"no_fill": True}),
    ]
    doughnut.legend.hidden = True
    doughnut.chart_area.format = px.ChartFormat(no_border=True, no_fill=True)
    pie = px.Chart(px.ChartType.Pie, rotation=270)
    pie_series = pie.add_series(values="=Sheet1!$I$3:$I$6", name="=Sheet1!$I$2")
    pie_series.points = [
        {"format": {"no_fill": True}},
        {"format": px.ChartSolidFill(color="#FF0000")},
        {"format": {"no_fill": True}},
    ]
    doughnut.combine(pie)
    ws.insert_chart(0, 0, doughnut)
    path = out()
    wb.save(path)
    reopen(path)
    assert_matches_corpus(path, corpus, "chart_combined11")


def test_corpus_stock_with_secondary_axis(out, reopen, corpus):
    wb, ws = stock_workbook()
    chart = px.Chart(px.ChartType.Stock, high_low_lines=True)
    hidden_line = px.ChartLine(width=2.25, hidden=True)
    for col in "BCD":
        chart.add_series(
            values=f"=Sheet1!${col}$1:${col}$5",
            categories="=Sheet1!$A$1:$A$5",
            format=px.ChartFormat(line=hidden_line),
            marker={"none": True},
        )
    chart.series[2].secondary_axis = True
    chart.series[2].marker = px.ChartMarker(type=px.ChartMarkerType.ShortDash, size=3)
    ws.insert_chart("E9", chart)
    path = out()
    wb.save(path)
    reopen(path)
    assert_matches_corpus(path, corpus, "chart_stock02")


def test_corpus_stock_with_error_bars(out, reopen, corpus):
    wb, ws = stock_workbook()
    chart = px.Chart(px.ChartType.Stock)
    chart.high_low_lines = True
    for col in "BCD":
        chart.add_series(
            values=f"=Sheet1!${col}$1:${col}$5",
            categories="=Sheet1!$A$1:$A$5",
            format=px.ChartLine(width=2.25, hidden=True),
            marker={"none": True},
            y_error_bars=True,
        )
    chart.series[2].marker = px.ChartMarker(type=px.ChartMarkerType.ShortDash, size=3)
    ws.insert_chart("E9", chart)
    path = out()
    wb.save(path)
    reopen(path)
    assert_matches_corpus(path, corpus, "chart_errorbars07")


def test_corpus_scatter_axis_names_from_cells(out, reopen, corpus):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_rows("A1", [["X", "Y"], [1, 10], [3, 30]])
    chart = px.Chart(px.ChartType.Scatter)
    chart.add_series(values="=Sheet1!$B$2:$B$3", categories="=Sheet1!$A$2:$A$3")
    chart.x_axis.name = "=Sheet1!$A$1"
    chart.x_axis.name_font.italic = True
    chart.y_axis.name = ("Sheet1", 0, 1)
    ws.insert_chart("E9", chart)
    path = out()
    wb.save(path)
    reopen(path)
    assert_matches_corpus(path, corpus, "chart_scatter15")
