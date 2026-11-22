"""Helpers shared by the chart tests (tests/test_charts_*.py)."""

from __future__ import annotations

import re
import zipfile
from collections.abc import Callable
from pathlib import Path

import pyrust_xlsxwriter as px

# Sheet1!A1:D5: column A = 1..5, B = 2x, C = 3x, D = 4x.
DATA = [[r * c for c in range(1, 5)] for r in range(1, 6)]


def workbook_with_data(sheet: str = "Sheet1") -> tuple[px.Workbook, px.Worksheet]:
    wb = px.Workbook()
    ws = wb.add_worksheet(sheet)
    ws.write_rows("A1", DATA)
    return wb, ws


def column_chart(chart_type: px.ChartType = px.ChartType.Column, series: int = 1) -> px.Chart:
    """A chart with `series` series over Sheet1 columns A, B, C, ..."""
    chart = px.Chart(chart_type)
    for i in range(series):
        col = "ABCD"[i]
        chart.add_series(values=f"=Sheet1!${col}$1:${col}$5")
    return chart


def save_chart(out: Callable[[str], Path], chart: px.Chart, name: str = "chart.xlsx") -> Path:
    """Save a workbook with the sample data and `chart` inserted at E2."""
    wb, ws = workbook_with_data()
    ws.insert_chart("E2", chart)
    path = out(name)
    wb.save(path)
    return path


def chart_xml(path: Path, n: int = 1) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(f"xl/charts/chart{n}.xml").decode("utf-8")


def plot_area(xml: str) -> str:
    """The <c:plotArea> element (series, axes)."""
    return xml[xml.index("<c:plotArea>") : xml.index("</c:plotArea>")]


def normalized(xml: str) -> str:
    """Chart XML with data caches, axis ids and print settings removed, for
    comparison with Excel-made files."""
    xml = re.sub(r"<c:(num|str)Cache>.*?</c:\1Cache>", "", xml)
    xml = re.sub(r"<c:printSettings>.*?</c:printSettings>", "", xml)
    xml = re.sub(r"<c:printSettings/>", "", xml)
    ids: dict[str, str] = {}

    def axis_id(match: re.Match[str]) -> str:
        value = match.group(2)
        ids.setdefault(value, f"AX{len(ids)}")
        return f'<c:{match.group(1)} val="{ids[value]}"/>'

    xml = re.sub(r'<c:(axId|crossAx) val="(\d+)"/>', axis_id, xml)
    start = xml.index("<c:chart>")
    return xml[start:]
