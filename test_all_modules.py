import pyrust_xlsxwriter
from pyrust_xlsxwriter import (
    Workbook,
    Chart,
    ChartType,
    ChartLegendPosition,
    Sparkline,
    SparklineType,
    ConditionalFormat2ColorScale,
    ConditionalFormat3ColorScale,
    ConditionalFormatDataBar,
    Color,
    PaperSize,
    VbaProject,
)


def test_chart_sparkline_conditional_format_macros():
    wb = Workbook()
    ws = wb.add_worksheet(name="Data")

    # Test default paper_size and tab_color
    assert ws.paper_size == PaperSize.PrinterDefault
    assert ws.tab_color == Color.Default

    # Test setting valid values
    ws.paper_size = PaperSize.A4
    assert ws.paper_size == PaperSize.A4

    ws.tab_color = Color.Red
    assert ws.tab_color == Color.Red

    import logging

    class LogHandler(logging.Handler):
        def __init__(self):
            super().__init__()
            self.records = []

        def emit(self, record):
            self.records.append(record)

    logger = logging.getLogger()
    handler = LogHandler()
    logger.addHandler(handler)

    ws.paper_size = None
    assert any("None is not allowed for paper_size" in r.getMessage() for r in handler.records)
    assert ws.paper_size == PaperSize.PrinterDefault

    handler.records.clear()
    ws.tab_color = None
    assert any("None is not allowed for tab_color" in r.getMessage() for r in handler.records)
    assert ws.tab_color == Color.Default

    logger.removeHandler(handler)

    # Write data for Chart & Sparkline
    ws.write_row(0, 0, [10, 20, 30, 40, 50])
    ws.write_row(1, 0, [15, 25, 35, 45, 55])

    # 1. Test Chart
    chart = Chart(chart_type=ChartType.Column)
    chart.title = "Sales Performance"
    chart.x_axis_name = "Quarter"
    chart.y_axis_name = "Revenue"
    chart.width = 600
    chart.height = 350
    chart.style = 2
    chart.legend_position = ChartLegendPosition.Bottom

    assert chart.title == "Sales Performance"
    assert chart.x_axis_name == "Quarter"
    assert chart.y_axis_name == "Revenue"
    assert chart.width == 600
    assert chart.height == 350
    assert chart.style == 2
    assert chart.legend_position == ChartLegendPosition.Bottom

    chart.add_series(values="=Data!$A$1:$E$1", name="Series 1")
    chart.add_series(values="=Data!$A$2:$E$2", name="Series 2")

    ws.insert_chart(3, 0, chart)
    ws.insert_chart_with_offset(3, 8, chart, 10, 20)

    # 2. Test Sparkline
    sparkline = Sparkline(range="=Data!$A$1:$E$1", sparkline_type=SparklineType.Line)
    sparkline.high_point = True
    sparkline.low_point = True
    sparkline.markers = True
    sparkline.style = 5

    assert sparkline.range == "=Data!$A$1:$E$1"
    assert sparkline.sparkline_type == SparklineType.Line
    assert sparkline.high_point is True
    assert sparkline.low_point is True
    assert sparkline.markers is True
    assert sparkline.style == 5

    ws.add_sparkline(0, 5, sparkline)

    # 3. Test Conditional Format 2 Color Scale
    cf2 = ConditionalFormat2ColorScale()
    cf2.min_color = Color.White
    cf2.max_color = Color.Green
    assert cf2.min_color == Color.White
    assert cf2.max_color == Color.Green
    ws.add_conditional_format_2color_scale(0, 0, 1, 4, cf2)

    # 4. Test Conditional Format 3 Color Scale
    cf3 = ConditionalFormat3ColorScale()
    cf3.min_color = Color.Red
    cf3.mid_color = Color.Yellow
    cf3.max_color = Color.Green
    assert cf3.min_color == Color.Red
    assert cf3.mid_color == Color.Yellow
    assert cf3.max_color == Color.Green
    ws.add_conditional_format_3color_scale(10, 0, 15, 4, cf3)

    # 5. Test Conditional Format Data Bar
    cf_db = ConditionalFormatDataBar()
    cf_db.fill_color = Color.Blue
    assert cf_db.fill_color == Color.Blue
    ws.add_conditional_format_data_bar(20, 0, 25, 4, cf_db)

    # 6. Test Macros / VBA Project
    vba_proj = VbaProject(path="vbaProject.bin")
    assert vba_proj.path == "vbaProject.bin"
    assert vba_proj.signature_path is None

    ws.set_vba_name("MySheetModule")
    wb.set_vba_name("MyWorkbookModule")

    wb.save("test_all_modules_output.xlsx")
    print(
        "ALL modules (Chart, Sparkline, Conditional Format, Macros, Workbook) tested & verified successfully!"
    )


if __name__ == "__main__":
    test_chart_sparkline_conditional_format_macros()
