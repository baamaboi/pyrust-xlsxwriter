# pyrust_xlsxwriter

A fast Python binding for Excel file generation powered by Rust.

## About the Project

I created this as a hobby project to learn PyO3 and Rust. It is essentially a Python binding / port of [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) by [John McNamara](https://github.com/jmcnamara).

This project is still a work in progress. Most of the code has been AI-generated, and I'm actively cleaning it up and tweaking the API to my liking. The API mostly matches John's popular Python [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) library, with some opinionated changes and additions of my own.

My goal with creating this is to get the added performance benefits of Rust.

## Features

- High performance Excel file creation backed by Rust.
- Rich formatting options (`bold`, `italic`, alignment, colors, borders, number formats, rotation, indents, etc.).
- Support for Charts (`Column`, `Bar`, `Line`, `Pie`, `Area`, `Scatter`, etc.) with legend positions and axis titles.
- Sparklines (`Line`, `Column`, `WinLoss`).
- Conditional Formatting (2-Color scale, 3-Color scale, Data Bars).
- VBA Macros (`add_vba_project` and signed projects).
- Memory-efficient workbook handling.
- Built-in type stubs (`.pyi`) for full IDE autocomplete and type checking support.

## Installation & Local Development

Building requires [Rust](https://www.rust-lang.org/) and [maturin](https://github.com/PyO3/maturin).

```bash
# Clone the repository
git clone https://github.com/baamaboi/pyrust_xlsxwriter.git
cd pyrust_xlsxwriter

# Create virtual environment and install dev dependencies
uv sync

# Build and install locally in editable mode
uv run maturin develop
```

## Quick Start

```python
import pyrust_xlsxwriter as pywriter
from pyrust_xlsxwriter import (
    Workbook,
    Format,
    Color,
    FormatAlign,
    FormatBorder,
    Chart,
    ChartType,
)

# Create workbook and worksheet
wb = Workbook()
ws = wb.add_worksheet(name="Demo")

# Create cell formats
header_fmt = Format()
header_fmt.bold = True
header_fmt.background_color = Color.Blue
header_fmt.font_color = Color.White
header_fmt.align = FormatAlign.Center
header_fmt.border = FormatBorder.Thin

# Write headers and data
ws.write(0, 0, "Item", header_fmt)
ws.write(0, 1, "Sales", header_fmt)

ws.write(1, 0, "Apples")
ws.write(1, 1, 120.5)
ws.write(2, 0, "Oranges")
ws.write(2, 1, 85.0)

# Add a chart
chart = Chart(chart_type=ChartType.Column)
chart.add_series(categories="=Demo!$A$2:$A$3", values="=Demo!$B$2:$B$3", name="Sales")
chart.title = "Fruit Sales"
ws.insert_chart(4, 0, chart)

# Save output
wb.save("output.xlsx")
```

## License

This project is licensed under the [MIT License](LICENSE). Third-party software dependencies and their respective licenses/copyright notices are documented in [LICENSE-THIRD-PARTY](LICENSE-THIRD-PARTY).
