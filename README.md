# pyrust_xlsxwriter

Create, read and edit Excel files from Python, powered by Rust.

* **Create** new files with [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter): every feature of version 0.103 is bound, including formats, charts, tables, conditional formats, data validation, sparklines, pivot tables, images, notes, shapes, page setup and protection.
* **Read** xlsx, xlsm, xlsb, xls and ods with [calamine](https://github.com/tafia/calamine), including values, formulas, styles, merged ranges, hyperlinks, tables and defined names.
* **Edit** existing xlsx/xlsm files in place. The editor rewrites only what you change and copies every other part byte-for-byte, so charts, pivot tables, VBA, slicers and images survive.

It started as a hobby project to learn PyO3 and Rust. The writer's API is close to John McNamara's [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), with some opinionated changes.

## Performance

These are measured by the scripts in `benchmarks/`, on a release build on an Apple-silicon Mac.

| Task | pyrust_xlsxwriter | openpyxl |
|---|---|---|
| Read 200,000 × 10 cells (`values()` vs `read_only` + `iter_rows`) | 0.58 s | 5.2 s (9× slower) |
| Load a 100,000-row file, edit 1,000 cells, append 100 rows, save | 0.47 s, 150 MB peak | 8.8 s, 548 MB peak (19× slower) |

Fidelity was tested on 300 Excel-made files (`benchmarks/fidelity.py`), each loaded and saved without changes:

| Engine | Files that lost parts | Parts lost |
|---|---|---|
| pyrust_xlsxwriter editor | 0 | 0 |
| openpyxl 3.1.5 | 131 | 330 |
| umya-spreadsheet 3.1.1 | 44 | 154 (2 outputs also left structurally invalid) |

openpyxl's count includes images it drops because Pillow isn't installed in the test environment.

## Installation

Building requires [Rust](https://www.rust-lang.org/) (1.100 or newer) and [uv](https://docs.astral.sh/uv/).

```bash
git clone https://github.com/baamaboi/pyrust_xlsxwriter.git
cd pyrust_xlsxwriter
uv sync
uv run maturin develop --release
```

## Creating files

```python
import datetime as dt
import pyrust_xlsxwriter as px

with px.Workbook("report.xlsx") as wb:          # saved when the block ends
    ws = wb.add_worksheet("Sales", tab_color="#4472C4")
    header = px.Format(bold=True, background_color="#DDEBF7", border=px.FormatBorder.Thin)

    ws.write_row("A1", ["Region", "Units", "Price", "Date"], header)
    ws.write_rows("A2", [["North", 120, 9.5, dt.date(2026, 1, 5)],
                         ["South", 80, 11.0, dt.date(2026, 1, 6)]])
    ws.write("E2", "=B2*C2")                     # strings starting with "=" are formulas
    ws.set_column_width("A:D", 14)
    ws.freeze_panes("A2")

    ws.add_table("A1:D3", px.Table(style=px.TableStyle.Medium2))
    ws.add_conditional_format("B2:B3", px.ConditionalFormatDataBar())

    chart = px.Chart(px.ChartType.Column)
    series = chart.add_series(values="=Sales!$B$2:$B$3", categories="=Sales!$A$2:$A$3", name="Units")
    series.data_label = px.ChartDataLabel(show_value=True)
    chart.x_axis.name = "Region"                 # nested chart objects can be changed in place
    ws.insert_chart("G2", chart)
```

* **Indexes:** rows and columns are zero-based. Every method that takes a cell or range also accepts `"A1"` / `"A1:B2"`, and column methods accept `"B"` / `"B:D"`.
* **Settings:** every class exposes its settings as properties, and the constructor accepts all of them as keyword arguments (`px.Format(bold=True, num_format="0.00%")`).
* **Value types:** `write()` picks the cell type from the value, accepting `str`, `int`, `float`, `bool`, `Decimal`, `datetime`, `date`, `time`, `timedelta`, `None`, `Formula`, `Url`, and numpy scalars. Dates get a date format automatically.
* **Lists of records:** `ws.write_records("A1", rows)` writes dicts, dataclasses or named tuples with a header row.
* **Large sheets:** `add_worksheet(memory="constant")` keeps memory flat. Such sheets must be written in row order; writing an earlier row raises an error.
* **Enum members named `None`** are spelled `None_` (for example `px.FormatBorder.None_`), because `None` is a Python keyword.

## Reading files

```python
book = px.open_workbook("data.xlsx")             # also bytes or a binary file object
ws = book["Sales"]
rows = ws.values()                               # list of rows, starting at A1
records = ws.records()                           # list of dicts keyed by the header row
ws.cell("B2"), ws.formula("E2"), ws.style("A1")  # value, "=B2*C2", Format | None
ws.merged_ranges, ws.hyperlinks, book.tables, book.defined_names
```

Sheets are parsed on first access, and parsing releases the GIL.

## Editing existing files

```python
book = px.load_workbook("template.xlsx")
ws = book["Report"]

ws.write("B3", 1250)                             # keeps the cell's existing format
ws.write("B4", 980, px.Format(bold=True))        # replaces the format
ws.set_cell_format("B5", px.Format(font_color="red"), merge=True)  # adds to the format
ws.insert_rows(10, 3)                            # formulas, charts, tables, merges and names follow
ws.add_table("H2:J20", px.Table(columns=["A", "B", "C"]))
trend = px.Chart(px.ChartType.Line)              # any writer object can be added
trend.add_series(values="=Report!$B$2:$B$9")
ws.insert_chart("L2", trend)
book["Data"].name = "Raw data"                   # references in other sheets are updated

book.save("report.xlsx")
```

What the editor can do:

* **Cells:** write values, formulas and formats, merge and unmerge ranges, set column widths and row heights, hide rows and columns, freeze panes, set tab colour and zoom.
* **Sheets:** add, rename, hide, reorder and remove sheets; manage defined names and document properties.
* **Structural edits:** `insert_rows`, `delete_rows`, `insert_cols` and `delete_cols` update everything that refers to the moved cells:
  * formulas on every sheet;
  * shared and array formulas;
  * merged ranges, conditional formats and data validation;
  * hyperlinks, autofilters and tables;
  * drawing anchors (charts, images, shapes), comments, chart series ranges, pivot sources and defined names.
* **Adding objects:** the methods `add_conditional_format`, `add_data_validation`, `add_table`, `add_sparkline(_group)`, `insert_image`, `insert_chart`, `insert_shape` and `insert_note` take the same arguments as the writer. The object is rendered by the writer and grafted into the file.
* **Reading back:** pending edits are reflected immediately by `cell()`, `values()`, `formula()` and `style()`.
* **Saving:** saving can be repeated, and it releases the GIL.

## Converting files

```python
px.convert_to_xlsx("old.xls", "new.xlsx")        # xls/xlsb/ods/xlsx to xlsx
```

Conversion copies values, formulas, formats, merged ranges, hyperlinks and names.

## Errors

All exceptions derive from `px.XlsxError`:

* `ParameterError` (also a `ValueError`), with subclasses `SheetNameError`, `OverlapError` and `ImageError`;
* `RangeError` (also an `IndexError`);
* `FileError` (also an `OSError`);
* `ReadError`;
* `EditError`.

## Known limitations

* **Editor file types:** only xlsx/xlsm can be edited, and password-protected files aren't supported.
* **Notes in the editor:** a sheet that already has notes, buttons or VML header images can't have more notes added.
* **Tables in structural edits:** inserting or deleting columns *inside* a table, or deleting a table's header row, is rejected.
* **Recalculation:** formulas aren't calculated. Edited workbooks are marked for full recalculation when Excel opens them.
* **Data frames:** there's no zero-copy Arrow export yet; pass `ws.values()` or `ws.records()` to polars or pandas.
* **Upstream (rust_xlsxwriter 0.103):**
  * custom icon sets lose `reverse_icons` / `icons_only` / `greater_than` in the Excel 2010 format;
  * `ChartFont.strikethrough` isn't written;
  * `set_print_first_page_number` only supports 1.
* **xls and xlsb reading** is untested, because there are no fixtures for those formats.

## Development

```bash
uv run maturin develop                 # debug build
uv run pytest                          # Python tests, including the stub checks below
uv run pyrefly check                   # type-check the package, tests and benchmarks
cargo test --lib                       # Rust unit tests (XML scanner, styles, formula engine, ...)
uv run python benchmarks/bench_read.py
uv run python benchmarks/bench_edit.py
uv run python benchmarks/fidelity.py tests/corpus/*.xls*
```

### Source layout

```
src/common/   values, colours, A1 references, Format, exceptions, binding macros
src/writer/   new files (rust_xlsxwriter): one file per feature area
src/reader/   read-only access (calamine)
src/editor/   in-place editing: XML scanner, package, styles, sheet rewrite,
              formula engine, structural edits, grafting
```

### Tests

* `tests/corpus/` holds Excel-made reference files from rust_xlsxwriter's test suite (MIT licensed).
* `tests/ooxml_check.py` checks every saved file for the structural problems that make Excel offer to repair a workbook.
* The type stubs in `python/pyrust_xlsxwriter/_pyrust_xlsxwriter.pyi` are hand-written. `tests/test_stubs.py` checks them two ways:
  * mypy's stubtest compares them with the compiled module (names, parameters, `@final`). pyrefly has no equivalent, which is the only reason mypy is a dev dependency.
  * pyrefly type-checks the package, tests and benchmarks against them. Tests that pass wrong types on purpose are marked `# pyrefly: ignore[...]`, and unused ignores are errors, so a stub that stops rejecting one of those calls fails the check.

## License

This project is licensed under the [MIT License](LICENSE). Third-party software dependencies and their respective licenses/copyright notices are documented in [LICENSE-THIRD-PARTY](LICENSE-THIRD-PARTY).
