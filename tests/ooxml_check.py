"""Structural checks for xlsx/xlsm packages (stdlib only).

Excel shows "We found a problem with some content" for packages that are
structurally inconsistent. Without Excel available in CI, these checks catch
the usual causes:

* missing content types, dangling relationships, malformed XML;
* worksheet elements out of schema order;
* rows/cells out of order, or cell references that don't match their row;
* style or shared-string indexes that point past the end of their tables;
* inconsistent `count` attributes, overlapping merged ranges;
* defined names scoped to sheets that don't exist.

Usage: ``problems = check_package(path_or_bytes)``; an empty list means OK.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

WORKSHEET_ORDER = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
]

CELL_REF = re.compile(r"^([A-Z]{1,3})([0-9]+)$")


def _q(tag: str, ns: str = NS_MAIN) -> str:
    return f"{{{ns}}}{tag}"


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _parse_ref(ref: str) -> tuple[int, int] | None:
    m = CELL_REF.match(ref.replace("$", ""))
    if not m:
        return None
    return int(m.group(2)) - 1, _col_index(m.group(1))


def _parse_sqref(ref: str) -> tuple[int, int, int, int] | None:
    a, _, b = ref.partition(":")
    first = _parse_ref(a)
    last = _parse_ref(b) if b else first
    if first is None or last is None:
        return None
    return first[0], first[1], last[0], last[1]


def _rels_path(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _resolve(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))


def _read_rels(zf: zipfile.ZipFile, part: str, names: set[str]) -> dict[str, tuple[str, str, str]]:
    """Return {rId: (type, resolved_target, mode)} for a part ('' = package)."""
    rels_name = "_rels/.rels" if part == "" else _rels_path(part)
    if rels_name not in names:
        return {}
    root = ET.fromstring(zf.read(rels_name))
    out = {}
    for rel in root.findall(_q("Relationship", NS_PKG_REL)):
        mode = rel.get("TargetMode", "Internal")
        target = rel.get("Target", "")
        resolved = target if mode == "External" else _resolve(part, target)
        out[rel.get("Id", "")] = (rel.get("Type", ""), resolved, mode)
    return out


def check_package(source: str | Path | bytes) -> list[str]:
    problems: list[str] = []
    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        return [f"not a zip file: {exc}"]

    names = {n for n in zf.namelist() if not n.endswith("/")}

    # 1. Content types.
    if "[Content_Types].xml" not in names:
        return ["missing [Content_Types].xml"]
    ct = ET.fromstring(zf.read("[Content_Types].xml"))
    defaults = {d.get("Extension", "").lower() for d in ct.findall(_q("Default", NS_CT))}
    overrides = {o.get("PartName", "").lstrip("/") for o in ct.findall(_q("Override", NS_CT))}
    for name in names:
        if name == "[Content_Types].xml":
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if name not in overrides and ext not in defaults:
            problems.append(f"no content type for part {name}")
    for name in overrides:
        if name not in names:
            problems.append(f"content type override for missing part {name}")

    # 2. Well-formed XML.
    parsed: dict[str, ET.Element] = {}
    for name in names:
        if name.endswith((".xml", ".rels")):
            try:
                parsed[name] = ET.fromstring(zf.read(name))
            except ET.ParseError as exc:
                problems.append(f"malformed XML in {name}: {exc}")

    # 3. Relationships resolve.
    for name in names:
        if not name.endswith(".rels") or name not in parsed:
            continue
        folder = posixpath.dirname(posixpath.dirname(name))
        source_part = posixpath.join(folder, posixpath.basename(name)[: -len(".rels")])
        if posixpath.basename(name) == ".rels" and folder == "":
            source_part = ""
        for rel in parsed[name].findall(_q("Relationship", NS_PKG_REL)):
            if rel.get("TargetMode") == "External" or rel.get("Target", "").startswith("#"):
                continue  # external links and in-document locations (e.g. "#Sheet1!A1")
            target = _resolve(source_part, rel.get("Target", ""))
            if target not in names:
                problems.append(f"{name}: relationship {rel.get('Id')} targets missing part {target}")

    # 4. Workbook.
    root_rels = _read_rels(zf, "", names)
    wb_part = next((t for (ty, t, _) in root_rels.values() if ty.endswith("/officeDocument")), "xl/workbook.xml")
    if wb_part not in parsed:
        problems.append(f"missing workbook part {wb_part}")
        return problems
    wb = parsed[wb_part]
    wb_rels = _read_rels(zf, wb_part, names)
    sheets = wb.findall(f"{_q('sheets')}/{_q('sheet')}")
    seen_names: set[str] = set()
    sheet_parts: list[tuple[str, str]] = []
    for sheet in sheets:
        name = sheet.get("name", "")
        if name.lower() in seen_names:
            problems.append(f"duplicate sheet name {name!r}")
        seen_names.add(name.lower())
        rid = sheet.get(_q("id", NS_REL))
        if rid not in wb_rels:
            problems.append(f"sheet {name!r} relationship {rid} not found")
            continue
        sheet_parts.append((name, wb_rels[rid][1]))
    for dn in wb.findall(f"{_q('definedNames')}/{_q('definedName')}"):
        local = dn.get("localSheetId")
        if local is not None and int(local) >= len(sheets):
            problems.append(f"defined name {dn.get('name')!r} has localSheetId {local} >= {len(sheets)} sheets")

    # 5. Styles and shared strings.
    styles_part = next((t for (ty, t, _) in wb_rels.values() if ty.endswith("/styles")), None)
    xf_count = None
    if styles_part and styles_part in parsed:
        styles = parsed[styles_part]
        for container, child in [
            ("numFmts", "numFmt"),
            ("fonts", "font"),
            ("fills", "fill"),
            ("borders", "border"),
            ("cellStyleXfs", "xf"),
            ("cellXfs", "xf"),
            ("cellStyles", "cellStyle"),
            ("dxfs", "dxf"),
        ]:
            el = styles.find(_q(container))
            if el is None:
                continue
            n = len(el.findall(_q(child)))
            if el.get("count") is not None and int(el.get("count", "0")) != n:
                problems.append(f"styles {container} count={el.get('count')} but has {n} <{child}>")
            if container == "cellXfs":
                xf_count = n
            if container == "cellXfs":
                fonts = len(styles.findall(f"{_q('fonts')}/{_q('font')}"))
                fills = len(styles.findall(f"{_q('fills')}/{_q('fill')}"))
                borders = len(styles.findall(f"{_q('borders')}/{_q('border')}"))
                for i, xf in enumerate(el.findall(_q("xf"))):
                    for attr, limit in (("fontId", fonts), ("fillId", fills), ("borderId", borders)):
                        if int(xf.get(attr, "0")) >= max(limit, 1):
                            problems.append(f"cellXfs[{i}] {attr}={xf.get(attr)} out of range ({limit})")

    sst_count = None
    sst_part = next((t for (ty, t, _) in wb_rels.values() if ty.endswith("/sharedStrings")), None)
    if sst_part and sst_part in parsed:
        sst = parsed[sst_part]
        sst_count = len(sst.findall(_q("si")))
        if sst.get("uniqueCount") is not None and int(sst.get("uniqueCount", "0")) != sst_count:
            problems.append(f"sharedStrings uniqueCount={sst.get('uniqueCount')} but has {sst_count} <si>")

    # 6. Worksheets.
    for sheet_name, part in sheet_parts:
        if part not in parsed:
            if part in names:
                continue
            problems.append(f"sheet {sheet_name!r} part {part} missing")
            continue
        ws = parsed[part]
        if ws.tag != _q("worksheet"):
            continue  # chartsheet / dialogsheet
        order = [c.tag.split("}")[-1] for c in ws if c.tag.startswith(f"{{{NS_MAIN}}}")]
        positions = [WORKSHEET_ORDER.index(t) for t in order if t in WORKSHEET_ORDER]
        if positions != sorted(positions):
            problems.append(f"{sheet_name!r}: elements out of schema order: {order}")
        for unique in (
            "sheetData",
            "mergeCells",
            "cols",
            "dimension",
            "sheetViews",
            "hyperlinks",
            "dataValidations",
            "tableParts",
        ):
            if order.count(unique) > 1:
                problems.append(f"{sheet_name!r}: duplicate <{unique}>")

        last_row = -1
        for row in ws.iter(_q("row")):
            r = int(row.get("r", "0")) - 1
            if r <= last_row:
                problems.append(f"{sheet_name!r}: row {r + 1} out of order")
            last_row = r
            last_col = -1
            for c in row.findall(_q("c")):
                ref = c.get("r")
                if ref is None:
                    last_col += 1
                    continue
                pos = _parse_ref(ref)
                if pos is None or pos[0] != r:
                    problems.append(f"{sheet_name!r}: cell {ref} in row {r + 1}")
                    continue
                if pos[1] <= last_col:
                    problems.append(f"{sheet_name!r}: cell {ref} out of order")
                last_col = pos[1]
                s = c.get("s")
                if s is not None and xf_count is not None and int(s) >= xf_count:
                    problems.append(f"{sheet_name!r}: cell {ref} style {s} >= {xf_count} cellXfs")
                if c.get("t") == "s":
                    v = c.find(_q("v"))
                    if v is None or v.text is None:
                        problems.append(f"{sheet_name!r}: shared-string cell {ref} has no value")
                    elif sst_count is not None and int(v.text) >= sst_count:
                        problems.append(f"{sheet_name!r}: cell {ref} string index {v.text} >= {sst_count}")

        merges = ws.find(_q("mergeCells"))
        if merges is not None:
            ranges = []
            cells = merges.findall(_q("mergeCell"))
            if merges.get("count") is not None and int(merges.get("count", "0")) != len(cells):
                problems.append(f"{sheet_name!r}: mergeCells count={merges.get('count')} but has {len(cells)}")
            for mc in cells:
                rng = _parse_sqref(mc.get("ref", ""))
                if rng is None:
                    problems.append(f"{sheet_name!r}: bad merge ref {mc.get('ref')}")
                    continue
                for other in ranges:
                    if not (rng[2] < other[0] or rng[0] > other[2] or rng[3] < other[1] or rng[1] > other[3]):
                        problems.append(f"{sheet_name!r}: merged range {mc.get('ref')} overlaps another")
                ranges.append(rng)

        ws_rels = _read_rels(zf, part, names)
        for el in ws.iter():
            rid = el.get(_q("id", NS_REL))
            if rid is not None and rid not in ws_rels:
                problems.append(f"{sheet_name!r}: <{el.tag.split('}')[-1]}> refers to unknown relationship {rid}")

    return problems


def assert_valid(source: str | Path | bytes) -> None:
    problems = check_package(source)
    assert not problems, "invalid package:\n  " + "\n  ".join(problems)
