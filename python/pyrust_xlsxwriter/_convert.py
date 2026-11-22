"""Convert any readable spreadsheet (xlsx, xlsm, xlsb, xls, ods) to a new xlsx file."""

from __future__ import annotations

import os
import re
from collections.abc import Sequence
from typing import Any, BinaryIO, overload

from . import _pyrust_xlsxwriter as _px

_PROPERTY_KEYS = ("title", "subject", "author", "manager", "company", "category", "keywords", "comment", "status")


_ODF_REF = re.compile(
    r"\[\$?(?:'((?:[^']|'')+)'|([^.\[\]:']*))\.(\$?[A-Z]+\$?\d+)(?::\$?(?:'(?:[^']|'')+'|[^.\[\]:']*)?\.(\$?[A-Z]+\$?\d+))?\]"
)


def _odf_formula(formula: str) -> str:
    """Translate an OpenDocument formula (``=[.B1]*2``, ``SUM([.A1:.A3];2)``) to Excel syntax."""

    def ref(match: re.Match[str]) -> str:
        sheet = match.group(1) or match.group(2) or ""
        area = match.group(3) + (f":{match.group(4)}" if match.group(4) else "")
        if not sheet:
            return area
        quoted = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else "'" + sheet.replace("'", "''") + "'"
        return f"{quoted}!{area}"

    parts = re.split(r'("(?:[^"]|"")*")', formula)  # keep string literals intact
    return "".join(part if part.startswith('"') else _ODF_REF.sub(ref, part).replace(";", ",") for part in parts)


def _write_row(ws: _px.Worksheet, row: int, values: Sequence[Any]) -> None:
    # Text that looks like a formula must stay text, so write those cells one by one.
    if any(isinstance(v, str) and v.startswith(("=", "{=")) for v in values):
        for col, value in enumerate(values):
            if isinstance(value, str):
                if value:
                    ws.write_string(row, col, value)
            elif value is not None:
                ws.write(row, col, value)
    elif any(v is not None for v in values):
        ws.write_row(row, 0, values)


@overload
def convert_to_xlsx(
    source: str | os.PathLike[str] | bytes | BinaryIO, destination: None = None, *, styles: bool = True
) -> bytes: ...
@overload
def convert_to_xlsx(
    source: str | os.PathLike[str] | bytes | BinaryIO, destination: str | os.PathLike[str], *, styles: bool = True
) -> None: ...
def convert_to_xlsx(
    source: str | os.PathLike[str] | bytes | BinaryIO,
    destination: str | os.PathLike[str] | None = None,
    *,
    styles: bool = True,
) -> bytes | None:
    """Convert a spreadsheet to xlsx.

    Copies values (with the original number formats when ``styles`` is true),
    formulas with their cached results, cell formats, merged ranges,
    hyperlinks, sheet visibility, defined names and document properties.
    Charts, images, conditional formats and other objects are not copied, and
    sheet-scoped defined names become workbook names (duplicates are skipped).

    Returns the xlsx bytes when ``destination`` is None, otherwise writes the
    file and returns None. ``styles=False`` is faster on very large sheets.
    """
    with _px.open_workbook(source) as book:
        wb = _px.Workbook()
        for info in book.sheets:
            if info.kind != "worksheet":
                continue
            sheet = book.worksheet(info.index)
            ws = wb.add_worksheet(info.name)
            if info.visibility == "very_hidden":
                ws.set_very_hidden()
            elif info.visibility == "hidden":
                ws.hidden = True

            values = sheet.values()
            for row, row_values in enumerate(values):
                _write_row(ws, row, row_values)
            for row, row_formulas in enumerate(sheet.formulas()):
                for col, formula in enumerate(row_formulas):
                    if formula and book.format == "ods":
                        formula = _odf_formula(formula)
                    if formula:
                        cached = values[row][col] if row < len(values) and col < len(values[row]) else None
                        ws.write_formula(row, col, formula, value=cached)
            for first_row, first_col, last_row, last_col in sheet.merged_ranges:
                value = (
                    values[first_row][first_col]
                    if first_row < len(values) and first_col < len(values[first_row])
                    else None
                )
                ws.merge_range(first_row, first_col, last_row, last_col, value)
            for link in sheet.hyperlinks:
                row, col = link.cell
                target = link.target or (f"internal:{link.location}" if link.location else None)
                if target:
                    text = link.display or (values[row][col] if row < len(values) and col < len(values[row]) else None)
                    ws.write_url(row, col, target, text=None if text is None else str(text), tip=link.tooltip)
            if styles:
                for row, row_values in enumerate(values):
                    for col in range(len(row_values)):
                        fmt = sheet.style(row, col)
                        if fmt is not None:
                            ws.set_cell_format(row, col, fmt)

        seen: set[str] = set()
        for name, formula in book.defined_names:
            # The reader doesn't report a name's sheet scope, so keep the first of duplicates.
            if name.lower() in seen:
                continue
            seen.add(name.lower())
            try:
                wb.define_name(name, formula)
            except _px.ParameterError, TypeError:
                continue  # names the writer manages itself (e.g. print areas)
        props = {k: v for k, v in book.properties.items() if k in _PROPERTY_KEYS and isinstance(v, str)}
        if props:
            wb.set_properties(props)

    if destination is None:
        return wb.save_to_buffer()
    wb.save(destination)
    return None
