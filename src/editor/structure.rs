//! Structural edits (insert/delete rows and columns) and sheet renames,
//! applied to package parts.
//!
//! The editor records structural operations in order. Current-state data
//! (pending cell edits, workbook defined names) is transformed when an
//! operation is recorded; parts coming from the original file are
//! transformed at save time by replaying the operations here, before cell
//! edits are applied on top.
//!
//! For a row/column edit on sheet S:
//! * on S: cell and row positions, formulas, shared/array formula ranges,
//!   dimension, columns, merged ranges, conditional formats, data
//!   validations (incl. x14 extensions), hyperlinks, autofilter, tables,
//!   drawing anchors (charts, images, shapes) and comments;
//! * everywhere: formulas referencing S (cells, CF/DV formulas, chart series
//!   ranges, pivot cache sources, hyperlink locations).
//!
//! Shared formulas in rewritten sheets are expanded into ordinary formulas,
//! because an edit between a shared formula's cells changes their meaning.

use std::collections::HashMap;

use crate::common::cellref::{cell_name, parse_cell, MAX_COLS, MAX_ROWS};
use crate::editor::formula::{rename_sheet, shift_formula, shift_range_list, translate_formula, Area, StructuralEdit};
use crate::editor::package::{EResult, Package};
use crate::editor::xml::{attr, attrs, escape_attr, escape_text, local, parse_document, unescape, Child, Node, Scanner, Tok};

/// A recorded operation, in the order the user made it.
#[derive(Clone, Debug)]
pub enum Op {
    /// Rows/columns inserted or deleted on the sheet with this id.
    Structure { sheet_id: u64, edit: StructuralEdit },
    /// A sheet renamed from `old` to `new`.
    Rename { old: String, new: String },
    /// Objects rendered into a scratch workbook, to graft onto a sheet.
    Graft { sheet_id: u64, scratch: crate::common::bytes::SharedBytes },
}

/// Map a position through an edit (None if it was deleted or pushed off the sheet).
pub fn map_cell(edit: &StructuralEdit, row: u32, col: u16) -> Option<(u32, u16)> {
    match edit {
        StructuralEdit::InsertRows { at, count, .. } => {
            if row >= *at {
                let r = row as u64 + *count as u64;
                (r < MAX_ROWS as u64).then_some((r as u32, col))
            } else {
                Some((row, col))
            }
        }
        StructuralEdit::DeleteRows { at, count, .. } => {
            if row < *at {
                Some((row, col))
            } else if (row as u64) < *at as u64 + *count as u64 {
                None
            } else {
                Some((row - count, col))
            }
        }
        StructuralEdit::InsertCols { at, count, .. } => {
            if col >= *at {
                let c = col as u32 + *count as u32;
                (c < MAX_COLS as u32).then_some((row, c as u16))
            } else {
                Some((row, col))
            }
        }
        StructuralEdit::DeleteCols { at, count, .. } => {
            if col < *at {
                Some((row, col))
            } else if (col as u32) < *at as u32 + *count as u32 {
                None
            } else {
                Some((row, col - count))
            }
        }
    }
}

/// The inverse of [`map_cell`]: where a current position came from (None for
/// inserted, empty positions).
pub fn unmap_cell(edit: &StructuralEdit, row: u32, col: u16) -> Option<(u32, u16)> {
    match edit {
        StructuralEdit::InsertRows { at, count, .. } => {
            if row < *at {
                Some((row, col))
            } else if (row as u64) < *at as u64 + *count as u64 {
                None
            } else {
                Some((row - count, col))
            }
        }
        StructuralEdit::DeleteRows { at, count, .. } => Some((if row >= *at { row + count } else { row }, col)),
        StructuralEdit::InsertCols { at, count, .. } => {
            if col < *at {
                Some((row, col))
            } else if (col as u32) < *at as u32 + *count as u32 {
                None
            } else {
                Some((row, col - count))
            }
        }
        StructuralEdit::DeleteCols { at, count, .. } => Some((row, if col >= *at { col + count } else { col })),
    }
}

/// A drawing anchor coordinate: deleted positions clamp to the edit point.
fn map_anchor(edit: &StructuralEdit, value: u32, is_row: bool) -> u32 {
    match (edit, is_row) {
        (StructuralEdit::InsertRows { at, count, .. }, true) if value >= *at => value + count,
        (StructuralEdit::DeleteRows { at, count, .. }, true) if value >= *at => value.saturating_sub(*count).max(*at),
        (StructuralEdit::InsertCols { at, count, .. }, false) if value >= *at as u32 => value + *count as u32,
        (StructuralEdit::DeleteCols { at, count, .. }, false) if value >= *at as u32 => {
            value.saturating_sub(*count as u32).max(*at as u32)
        }
        _ => value,
    }
}

fn is_row_edit(edit: &StructuralEdit) -> bool {
    matches!(edit, StructuralEdit::InsertRows { .. } | StructuralEdit::DeleteRows { .. })
}

/// Shift a single area string (`A1:B5`); None if deleted.
pub fn shift_area(text: &str, edit: &StructuralEdit) -> Option<String> {
    match Area::parse(text) {
        Some(area) => area.shift(edit).map(|a| a.to_string()),
        None => Some(text.to_string()),
    }
}

/// How formulas (and positions) change in one worksheet part.
pub struct SheetTransform<'a> {
    /// Applied to every formula text in the sheet.
    pub formula: &'a dyn Fn(&str) -> EResult<String>,
    /// Set when this sheet is the one whose rows/columns change.
    pub positions: Option<&'a StructuralEdit>,
}

fn rebuild_tag(name: &str, raw_attrs: &str, replace: &[(&str, Option<String>)], self_closing: bool) -> String {
    let mut out = format!("<{name}");
    let mut done: Vec<&str> = Vec::new();
    for (k, v) in attrs(raw_attrs) {
        if let Some((key, new)) = replace.iter().find(|(key, _)| *key == k) {
            done.push(key);
            if let Some(new) = new {
                out.push_str(&format!(" {k}=\"{}\"", escape_attr(new)));
            }
        } else {
            out.push_str(&format!(" {k}=\"{v}\""));
        }
    }
    for (key, new) in replace {
        if !done.contains(key)
            && let Some(new) = new {
                out.push_str(&format!(" {key}=\"{}\"", escape_attr(new)));
            }
    }
    out.push_str(if self_closing { "/>" } else { ">" });
    out
}

/// Rewrite a worksheet part.
pub fn transform_sheet(xml: &str, t: &SheetTransform<'_>) -> EResult<String> {
    let mut scanner = Scanner::new(xml);
    let mut out = String::with_capacity(xml.len() + 1024);
    // Copy up to and including the root start tag.
    let root_name = loop {
        let (tok, span) = scanner.next_tok().map_err(|e| e.to_string())?.ok_or("empty worksheet")?;
        out.push_str(&xml[span.clone()]);
        if let Tok::Start { name, .. } = tok {
            break name.to_string();
        }
        if let Tok::Empty { .. } = tok {
            return Ok(out);
        }
    };
    while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
        match tok {
            Tok::End { name } if name == root_name => {
                out.push_str(&xml[span.start..]);
                return Ok(out);
            }
            Tok::Start { name, attrs: a } | Tok::Empty { name, attrs: a } => {
                let is_start = matches!(tok, Tok::Start { .. });
                let end = if is_start { scanner.skip_element(name).map_err(|e| e.to_string())? } else { span.end };
                let element = &xml[span.start..end];
                match local(name) {
                    "sheetData" if is_start => {
                        let body = &xml[span.end..end - format!("</{name}>").len()];
                        out.push_str(&xml[span.clone()]);
                        transform_sheet_data(body, t, &mut out)?;
                        out.push_str(&format!("</{name}>"));
                    }
                    "dimension" => match t.positions {
                        Some(edit) => {
                            let r = attr(a, "ref").unwrap_or_default();
                            let new = shift_area(&r, edit).unwrap_or_else(|| "A1".into());
                            out.push_str(&rebuild_tag(name, a, &[("ref", Some(new))], true));
                        }
                        None => out.push_str(element),
                    },
                    "cols" => match t.positions {
                        Some(edit) if !is_row_edit(edit) => out.push_str(&transform_cols(element, edit)?),
                        _ => out.push_str(element),
                    },
                    "mergeCells" | "conditionalFormatting" | "dataValidations" | "hyperlinks" | "autoFilter"
                    | "extLst" | "rowBreaks" | "colBreaks" => {
                        let mut doc = parse_document(element).map_err(|e| e.to_string())?;
                        if transform_node(&mut doc.root, t)? {
                            fix_counts(&mut doc.root);
                            if (doc.root.local() == "mergeCells" || doc.root.local() == "dataValidations" || doc.root.local() == "hyperlinks")
                                && doc.root.elements().next().is_none() {
                                    continue;
                                }
                            out.push_str(&doc.root.to_xml());
                        }
                    }
                    _ => out.push_str(element),
                }
            }
            _ => out.push_str(&xml[span]),
        }
    }
    Err("worksheet part ended early".into())
}

/// Update `count` attributes after children were removed.
fn fix_counts(node: &mut Node) {
    if node.get("count").is_some() {
        let n = node.elements().count();
        node.set("count", &n.to_string());
    }
}

/// Transform a sheet-level element tree. Returns false if the element should be dropped.
fn transform_node(node: &mut Node, t: &SheetTransform<'_>) -> EResult<bool> {
    let name = node.local().to_string();
    // Range attributes that move with the sheet.
    if let Some(edit) = t.positions {
        for key in ["sqref", "ref"] {
            if let Some(value) = node.get(key) {
                let shifted = if key == "sqref" || name == "mergeCell" {
                    shift_range_list(&value, true, edit)
                } else {
                    shift_area(&value, edit)
                };
                match shifted {
                    Some(v) if !v.is_empty() => {
                        if v != value {
                            node.set(key, &v);
                        }
                    }
                    _ => {
                        if matches!(name.as_str(), "mergeCell" | "conditionalFormatting" | "dataValidation" | "hyperlink" | "autoFilter" | "selection") {
                            return Ok(false);
                        }
                    }
                }
            }
        }
        if name == "brk"
            && let Some(id) = node.get("id").and_then(|v| v.parse::<u32>().ok()) {
                let (r, c) = if is_row_edit(t.positions.unwrap_or(edit)) { (id, 0) } else { (0, id.min(MAX_COLS as u32 - 1) as u16) };
                match map_cell(edit, r, c) {
                    Some((nr, nc)) => node.set("id", &(if is_row_edit(edit) { nr } else { nc as u32 }).to_string()),
                    None => return Ok(false),
                }
            }
        if name == "sqref" {
            let text = node.text();
            match shift_range_list(&text, true, edit) {
                Some(v) if !v.is_empty() => {
                    if v != text {
                        node.set_text(&v);
                    }
                }
                _ => return Ok(false),
            }
        }
    }
    if matches!(name.as_str(), "formula" | "formula1" | "formula2" | "f") {
        let text = node.text();
        let new = (t.formula)(&text)?;
        if new != text {
            node.set_text(&new);
        }
    }
    if name == "hyperlink"
        && let Some(location) = node.get("location") {
            let new = (t.formula)(&location)?;
            if new != location {
                node.set("location", &new);
            }
        }
    let mut kept = Vec::with_capacity(node.children.len());
    for child in std::mem::take(&mut node.children) {
        match child {
            Child::Element(mut n) => {
                if transform_node(&mut n, t)? {
                    kept.push(Child::Element(n));
                }
            }
            other => kept.push(other),
        }
    }
    node.children = kept;
    fix_counts(node);
    // Containers whose items were all deleted go too.
    Ok(!(matches!(name.as_str(), "conditionalFormatting" | "dataValidations" | "mergeCells") && !node.elements().any(|n| !matches!(n.local(), "sqref"))))
}

fn transform_cols(element: &str, edit: &StructuralEdit) -> EResult<String> {
    let mut doc = parse_document(element).map_err(|e| e.to_string())?;
    let mut out_cols: Vec<Node> = Vec::new();
    for col in doc.root.children_named("col") {
        let min: u32 = col.get("min").and_then(|v| v.parse().ok()).unwrap_or(1);
        let max: u32 = col.get("max").and_then(|v| v.parse().ok()).unwrap_or(min);
        let (lo, hi) = (min - 1, max - 1);
        let new_range = match edit {
            StructuralEdit::InsertCols { at, count, .. } => {
                let (at, count) = (*at as u32, *count as u32);
                if lo >= at {
                    Some((lo + count, hi + count))
                } else if hi >= at {
                    Some((lo, hi + count))
                } else {
                    Some((lo, hi))
                }
            }
            StructuralEdit::DeleteCols { at, count, .. } => {
                let (at, end) = (*at as u32, *at as u32 + *count as u32);
                let keep_lo = if lo >= at && lo < end { end } else { lo };
                let keep_hi = if hi >= at && hi < end { at.saturating_sub(1) } else { hi };
                if lo >= at && hi < end {
                    None
                } else {
                    let shift = |v: u32| if v >= end { v - (*count as u32) } else { v };
                    Some((shift(keep_lo.min(lo.max(keep_lo))), shift(keep_hi.max(keep_lo))))
                }
            }
            _ => Some((lo, hi)),
        };
        if let Some((nlo, nhi)) = new_range {
            let nhi = nhi.min(MAX_COLS as u32 - 1);
            if nlo <= nhi {
                let mut c = col.clone();
                c.set("min", &(nlo + 1).to_string());
                c.set("max", &(nhi + 1).to_string());
                out_cols.push(c);
            }
        }
    }
    if out_cols.is_empty() {
        return Ok(String::new());
    }
    doc.root.children = out_cols.into_iter().map(Child::Element).collect();
    Ok(doc.root.to_xml())
}

struct Shared {
    row: u32,
    col: u16,
    text: String,
}

fn transform_sheet_data(body: &str, t: &SheetTransform<'_>, out: &mut String) -> EResult<()> {
    let mut scanner = Scanner::new(body);
    let mut shared: HashMap<String, Shared> = HashMap::new();
    let mut next_row: u32 = 0;
    let mut rows: Vec<(u32, String)> = Vec::new();
    while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
        match tok {
            Tok::Start { name, attrs: a } | Tok::Empty { name, attrs: a } if local(name) == "row" => {
                let is_start = matches!(tok, Tok::Start { .. });
                let end = if is_start { scanner.skip_element(name).map_err(|e| e.to_string())? } else { span.end };
                let row = attr(a, "r").and_then(|r| r.parse::<u32>().ok()).map(|r| r.saturating_sub(1)).unwrap_or(next_row);
                next_row = row + 1;
                let new_row = match t.positions {
                    Some(edit) if is_row_edit(edit) => map_cell(edit, row, 0).map(|(r, _)| r),
                    _ => Some(row),
                };
                let mut text = String::new();
                let replace: Vec<(&str, Option<String>)> = vec![
                    ("r", new_row.map(|r| (r + 1).to_string())),
                    ("spans", if t.positions.is_some_and(|e| !is_row_edit(e)) { None } else { attr(a, "spans") }),
                ];
                text.push_str(&rebuild_tag(name, a, &replace, !is_start));
                if is_start {
                    let inner = &body[span.end..end - format!("</{name}>").len()];
                    transform_cells(inner, row, new_row, t, &mut shared, &mut text)?;
                    text.push_str(&format!("</{name}>"));
                }
                if let Some(r) = new_row {
                    rows.push((r, text));
                }
            }
            _ => {
                if !body[span.clone()].trim().is_empty() {
                    rows.push((u32::MAX, body[span].to_string()));
                }
            }
        }
    }
    for (_, text) in rows {
        out.push_str(&text);
    }
    Ok(())
}

fn transform_cells(
    inner: &str,
    row: u32,
    new_row: Option<u32>,
    t: &SheetTransform<'_>,
    shared: &mut HashMap<String, Shared>,
    out: &mut String,
) -> EResult<()> {
    let mut scanner = Scanner::new(inner);
    let mut next_col: u16 = 0;
    while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
        match tok {
            Tok::Start { name, attrs: a } | Tok::Empty { name, attrs: a } if local(name) == "c" => {
                let is_start = matches!(tok, Tok::Start { .. });
                let end = if is_start { scanner.skip_element(name).map_err(|e| e.to_string())? } else { span.end };
                let col = attr(a, "r").and_then(|r| parse_cell(&r)).map(|(_, c)| c).unwrap_or(next_col);
                next_col = col.saturating_add(1);
                let target = match (t.positions, new_row) {
                    (_, None) => None,
                    (Some(edit), Some(nr)) if !is_row_edit(edit) => map_cell(edit, nr, col),
                    (_, Some(nr)) => Some((nr, col)),
                };
                let body = if is_start { &inner[span.end..end - format!("</{name}>").len()] } else { "" };
                let new_body = transform_formula_in_cell(body, row, col, t, shared)?;
                if let Some((nr, nc)) = target {
                    out.push_str(&rebuild_tag(name, a, &[("r", Some(cell_name(nr, nc, false, false)))], !is_start));
                    if is_start {
                        out.push_str(&new_body);
                        out.push_str(&format!("</{name}>"));
                    }
                }
            }
            _ => out.push_str(&inner[span]),
        }
    }
    Ok(())
}

/// Rewrite the `<f>` element of a cell body (expanding shared formulas).
fn transform_formula_in_cell(
    body: &str,
    row: u32,
    col: u16,
    t: &SheetTransform<'_>,
    shared: &mut HashMap<String, Shared>,
) -> EResult<String> {
    let Some(start) = body.find("<f") else { return Ok(body.to_string()) };
    // Make sure it's the <f> element, not e.g. <foo>.
    let after = body[start + 2..].chars().next();
    if !matches!(after, Some(' ') | Some('>') | Some('/')) {
        return Ok(body.to_string());
    }
    let tag_end = body[start..].find('>').map(|i| start + i).ok_or("unterminated <f>")?;
    let self_closing = body[..tag_end].ends_with('/');
    let raw_attrs = &body[start + 2..if self_closing { tag_end - 1 } else { tag_end }];
    let (text, end) = if self_closing {
        (String::new(), tag_end + 1)
    } else {
        let close = body[tag_end..].find("</f>").map(|i| tag_end + i).ok_or("unterminated <f>")?;
        (unescape(&body[tag_end + 1..close]), close + 4)
    };
    let kind = attr(raw_attrs, "t");
    let si = attr(raw_attrs, "si");
    let (formula_text, mut replace): (String, Vec<(&str, Option<String>)>) = match (kind.as_deref(), si) {
        (Some("shared"), Some(si)) => {
            let text = if attr(raw_attrs, "ref").is_some() {
                shared.insert(si.clone(), Shared { row, col, text: text.clone() });
                text
            } else {
                match shared.get(&si) {
                    Some(master) => translate_formula(&master.text, row as i64 - master.row as i64, col as i64 - master.col as i64)
                        .map_err(|e| e.to_string())?,
                    None => text,
                }
            };
            (text, vec![("t", None), ("ref", None), ("si", None)])
        }
        (Some("array"), _) => {
            let new_ref = match (t.positions, attr(raw_attrs, "ref")) {
                (Some(edit), Some(r)) => Some(shift_area(&r, edit).unwrap_or_else(|| cell_name(row, col, false, false))),
                (_, r) => r,
            };
            (text, vec![("ref", new_ref)])
        }
        _ => (text, Vec::new()),
    };
    if formula_text.is_empty() {
        return Ok(body.to_string());
    }
    let new_text = (t.formula)(&formula_text)?;
    let mut tag = rebuild_tag("f", raw_attrs, &replace, false);
    replace.clear();
    tag.push_str(&escape_text(&new_text));
    tag.push_str("</f>");
    Ok(format!("{}{}{}", &body[..start], tag, &body[end..]))
}

/// Rewrite every `<c:f>`-style formula in a chart part.
pub fn transform_formulas_in_part(xml: &str, formula: &dyn Fn(&str) -> EResult<String>) -> EResult<String> {
    let mut doc = parse_document(xml).map_err(|e| e.to_string())?;
    fn walk(node: &mut Node, formula: &dyn Fn(&str) -> EResult<String>) -> EResult<()> {
        if node.local() == "f" {
            let text = node.text();
            let new = formula(&text)?;
            if new != text {
                node.set_text(&new);
            }
        }
        for child in node.elements_mut() {
            walk(child, formula)?;
        }
        Ok(())
    }
    walk(&mut doc.root, formula)?;
    Ok(doc.to_xml())
}

/// Shift drawing anchors (`xdr:from`/`xdr:to` rows and columns).
pub fn transform_drawing(xml: &str, edit: &StructuralEdit) -> EResult<String> {
    let mut doc = parse_document(xml).map_err(|e| e.to_string())?;
    fn walk(node: &mut Node, edit: &StructuralEdit) {
        if matches!(node.local(), "from" | "to") {
            for child in node.elements_mut() {
                let is_row = match child.local() {
                    "row" => true,
                    "col" => false,
                    _ => continue,
                };
                if let Ok(v) = child.text().trim().parse::<u32>() {
                    child.set_text(&map_anchor(edit, v, is_row).to_string());
                }
            }
        }
        for child in node.elements_mut() {
            walk(child, edit);
        }
    }
    walk(&mut doc.root, edit);
    Ok(doc.to_xml())
}

/// Shift comment references; returns the new XML.
pub fn transform_comments(xml: &str, edit: &StructuralEdit) -> EResult<String> {
    let mut doc = parse_document(xml).map_err(|e| e.to_string())?;
    if let Some(list) = doc.root.child_mut("commentList") {
        list.children.retain_mut(|c| match c {
            Child::Element(n) if n.local() == "comment" => match n.get("ref").and_then(|r| shift_area(&r, edit)) {
                Some(r) => {
                    n.set("ref", &r);
                    true
                }
                None => false,
            },
            _ => true,
        });
    }
    Ok(doc.to_xml())
}

/// Shift legacy VML note anchors (`x:Row`, `x:Column`, `x:Anchor`).
pub fn transform_vml(xml: &str, edit: &StructuralEdit) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    let row_edit = is_row_edit(edit);
    while let Some(i) = rest.find("<x:") {
        out.push_str(&rest[..i]);
        let after = &rest[i..];
        let tag_end = after.find('>').map(|j| j + 1).unwrap_or(after.len());
        let tag = &after[..tag_end];
        let name = tag[3..].trim_end_matches('>').split_whitespace().next().unwrap_or("");
        let close = format!("</x:{name}>");
        if matches!(name, "Row" | "Column" | "Anchor")
            && let Some(j) = after.find(&close) {
                let text = &after[tag_end..j];
                let new = match name {
                    "Row" if row_edit => text.trim().parse::<u32>().map(|v| map_anchor(edit, v, true).to_string()).unwrap_or(text.to_string()),
                    "Column" if !row_edit => text.trim().parse::<u32>().map(|v| map_anchor(edit, v, false).to_string()).unwrap_or(text.to_string()),
                    "Anchor" => {
                        let parts: Vec<&str> = text.split(',').map(str::trim).collect();
                        if parts.len() == 8 {
                            parts
                                .iter()
                                .enumerate()
                                .map(|(k, p)| match (k, p.parse::<u32>()) {
                                    (0 | 4, Ok(v)) if !row_edit => map_anchor(edit, v, false).to_string(),
                                    (2 | 6, Ok(v)) if row_edit => map_anchor(edit, v, true).to_string(),
                                    _ => p.to_string(),
                                })
                                .collect::<Vec<_>>()
                                .join(", ")
                        } else {
                            text.to_string()
                        }
                    }
                    _ => text.to_string(),
                };
                out.push_str(tag);
                out.push_str(&new);
                out.push_str(&close);
                rest = &after[j + close.len()..];
                continue;
            }
        out.push_str(tag);
        rest = &after[tag_end..];
    }
    out.push_str(rest);
    out
}

/// Shift a table part. Column edits that cut through a table aren't supported.
pub fn transform_table(xml: &str, edit: &StructuralEdit) -> EResult<String> {
    let mut doc = parse_document(xml).map_err(|e| e.to_string())?;
    let name = doc.root.get("displayName").or_else(|| doc.root.get("name")).unwrap_or_default();
    let Some(range) = doc.root.get("ref") else { return Ok(xml.to_string()) };
    let Some((r1, c1, r2, c2)) = Area::parse(&range).and_then(|a| a.bounds()) else { return Ok(xml.to_string()) };
    match edit {
        StructuralEdit::InsertCols { at, .. } if *at > c1 && *at <= c2 => {
            return Err(format!("inserting columns inside table {name:?} is not supported"));
        }
        StructuralEdit::DeleteCols { at, count, .. } if (*at as u32) <= c2 as u32 && (*at as u32 + *count as u32) > c1 as u32 => {
            return Err(format!("deleting columns of table {name:?} is not supported"));
        }
        StructuralEdit::DeleteRows { at, count, .. } if *at <= r1 && *at + count > r1 => {
            if *at + count > r2 {
                return Err(format!("deleting every row of table {name:?} is not supported; remove the table first"));
            }
            return Err(format!("deleting the header row of table {name:?} is not supported"));
        }
        _ => {}
    }
    let new = shift_area(&range, edit).ok_or_else(|| format!("table {name:?} would be deleted"))?;
    doc.root.set("ref", &new);
    if let Some(filter) = doc.root.child_mut("autoFilter")
        && let Some(r) = filter.get("ref").and_then(|r| shift_area(&r, edit)) {
            filter.set("ref", &r);
        }
    Ok(doc.to_xml())
}

/// Context for replaying operations over a package.
pub struct Replay<'a> {
    /// (sheet id, worksheet part) for every worksheet in the original file.
    pub sheets: &'a [(u64, String, String)],
    /// Chart, pivot cache and pivot table parts.
    pub charts: Vec<String>,
    pub pivot_caches: Vec<String>,
}

/// Apply one recorded operation to the original-file parts in `package`.
/// `names` maps sheet id -> its name at this point of the replay.
pub fn replay(
    package: &mut Package,
    styles: &mut crate::editor::styles::Styles,
    replay: &Replay<'_>,
    names: &mut HashMap<u64, String>,
    op: &Op,
) -> EResult<()> {
    match op {
        Op::Graft { sheet_id, scratch } => {
            let part = replay
                .sheets
                .iter()
                .find(|(id, _, _)| id == sheet_id)
                .map(|(_, part, _)| part.clone())
                .ok_or("grafted sheet not found")?;
            if package.exists(&part) {
                crate::editor::graft::apply(package, styles, &part, scratch)?;
            }
        }
        Op::Rename { old, new } => {
            let f = |text: &str| rename_sheet(text, old, new).map_err(|e| e.to_string());
            for (id, part, _) in replay.sheets {
                if !package.exists(part) {
                    continue;
                }
                let xml = package.read_string(part)?;
                if mentions(&xml, old) {
                    let out = transform_sheet(&xml, &SheetTransform { formula: &f, positions: None })?;
                    if out != xml {
                        package.write(part, out.into_bytes());
                    }
                }
                if names.get(id).is_some_and(|n| n.eq_ignore_ascii_case(old)) {
                    names.insert(*id, new.clone());
                }
            }
            for part in &replay.charts {
                rewrite_formulas(package, part, old, &f)?;
            }
            for part in &replay.pivot_caches {
                let mut doc = package.read_xml(part)?;
                let mut changed = false;
                if let Some(source) = doc.root.child_mut("cacheSource").and_then(|c| c.child_mut("worksheetSource"))
                    && source.get("sheet").is_some_and(|s| s.eq_ignore_ascii_case(old)) {
                        source.set("sheet", new);
                        changed = true;
                    }
                if changed {
                    package.write_xml(part, &doc);
                }
            }
        }
        Op::Structure { sheet_id, edit } => {
            let edited_name = names.get(sheet_id).cloned().unwrap_or_default();
            for (id, part, _) in replay.sheets {
                if !package.exists(part) {
                    continue;
                }
                let name = names.get(id).cloned().unwrap_or_default();
                let xml = package.read_string(part)?;
                let is_edited = id == sheet_id;
                if !is_edited && !mentions(&xml, &edited_name) {
                    continue;
                }
                let f = |text: &str| shift_formula(text, &name, edit).map_err(|e| e.to_string());
                let positions = if is_edited { Some(edit) } else { None };
                let out = transform_sheet(&xml, &SheetTransform { formula: &f, positions })?;
                if out != xml {
                    package.write(part, out.into_bytes());
                }
                if is_edited {
                    transform_sheet_children(package, part, edit)?;
                }
            }
            let f = |text: &str| shift_formula(text, "", edit).map_err(|e| e.to_string());
            for part in &replay.charts {
                rewrite_formulas(package, part, &edited_name, &f)?;
            }
            for part in &replay.pivot_caches {
                let mut doc = package.read_xml(part)?;
                let mut changed = false;
                if let Some(source) = doc.root.child_mut("cacheSource").and_then(|c| c.child_mut("worksheetSource"))
                    && source.get("sheet").is_some_and(|s| s.eq_ignore_ascii_case(&edited_name))
                        && let Some(r) = source.get("ref").and_then(|r| shift_area(&r, edit)) {
                            source.set("ref", &r);
                            changed = true;
                        }
                if changed {
                    package.write_xml(part, &doc);
                }
            }
        }
    }
    Ok(())
}

/// Whether `haystack` contains `needle`, ignoring case the way sheet names
/// compare, without lowercasing a copy of `haystack` (which can be a whole
/// sheet). A cheap check before rewriting a part.
fn mentions(haystack: &str, needle: &str) -> bool {
    let needle: Vec<char> = needle.chars().flat_map(char::to_lowercase).collect();
    let Some(&first) = needle.first() else { return true };
    haystack.char_indices().any(|(i, c)| {
        c.to_lowercase().next() == Some(first)
            && haystack[i..].chars().flat_map(char::to_lowercase).take(needle.len()).eq(needle.iter().copied())
    })
}

fn rewrite_formulas(package: &mut Package, part: &str, needle: &str, f: &dyn Fn(&str) -> EResult<String>) -> EResult<()> {
    if !package.exists(part) {
        return Ok(());
    }
    let xml = package.read_string(part)?;
    if !mentions(&xml, needle) {
        return Ok(());
    }
    let out = transform_formulas_in_part(&xml, f)?;
    if out != xml {
        package.write(part, out.into_bytes());
    }
    Ok(())
}

/// Tables, drawings, comments and pivot tables that belong to the edited sheet.
fn transform_sheet_children(package: &mut Package, sheet_part: &str, edit: &StructuralEdit) -> EResult<()> {
    let rels = package.rels(sheet_part)?;
    for rel in rels.list() {
        if rel.external {
            continue;
        }
        let part = crate::editor::package::resolve_target(sheet_part, &rel.target);
        if !package.exists(&part) {
            continue;
        }
        let kind = rel.rel_type.rsplit('/').next().unwrap_or("").to_string();
        if !matches!(kind.as_str(), "table" | "drawing" | "comments" | "vmlDrawing" | "pivotTable") {
            continue;
        }
        let xml = package.read_string(&part)?;
        let out = match kind.as_str() {
            "table" => transform_table(&xml, edit)?,
            "drawing" => transform_drawing(&xml, edit)?,
            "comments" => transform_comments(&xml, edit)?,
            "vmlDrawing" => transform_vml(&xml, edit),
            "pivotTable" => {
                let mut doc = parse_document(&xml).map_err(|e| e.to_string())?;
                if let Some(location) = doc.root.child_mut("location")
                    && let Some(r) = location.get("ref").and_then(|r| shift_area(&r, edit)) {
                        location.set("ref", &r);
                    }
                doc.to_xml()
            }
            _ => continue,
        };
        if out != xml {
            package.write(&part, out.into_bytes());
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn mentions_ignores_case() {
        assert!(mentions("<f>'Data Sheet'!A1</f>", "data SHEET"));
        assert!(mentions("<f>ÄRGER!A1</f>", "ärger"));
        assert!(!mentions("<f>Sheet1!A1</f>", "Sheet2"));
        assert!(!mentions("<f>She", "Sheet"));
        assert!(mentions("anything", ""));
    }

    fn rows(at: u32, count: u32) -> StructuralEdit {
        StructuralEdit::InsertRows { sheet: "S".into(), at, count }
    }

    #[test]
    fn maps_and_unmaps() {
        let ins = rows(2, 3);
        assert_eq!(map_cell(&ins, 1, 0), Some((1, 0)));
        assert_eq!(map_cell(&ins, 2, 0), Some((5, 0)));
        assert_eq!(unmap_cell(&ins, 3, 0), None);
        assert_eq!(unmap_cell(&ins, 5, 0), Some((2, 0)));
        let del = StructuralEdit::DeleteCols { sheet: "S".into(), at: 1, count: 2 };
        assert_eq!(map_cell(&del, 0, 1), None);
        assert_eq!(map_cell(&del, 0, 4), Some((0, 2)));
        assert_eq!(unmap_cell(&del, 0, 2), Some((0, 4)));
    }

    #[test]
    fn sheet_rows_and_formulas_shift() {
        let xml = "<worksheet><dimension ref=\"A1:B3\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><f>SUM(A1:A1)+B3</f><v>1</v></c></row><row r=\"3\"><c r=\"B3\"><v>5</v></c></row></sheetData><mergeCells count=\"1\"><mergeCell ref=\"A3:B3\"/></mergeCells></worksheet>";
        let edit = rows(1, 2);
        let f = |t: &str| shift_formula(t, "S", &edit).map_err(|e| e.to_string());
        let out = transform_sheet(xml, &SheetTransform { formula: &f, positions: Some(&edit) }).unwrap();
        assert!(out.contains("<dimension ref=\"A1:B5\"/>"));
        assert!(out.contains("<row r=\"4\"><c r=\"A4\"><f>SUM(A1:A1)+B5</f><v>1</v></c></row>"));
        assert!(out.contains("<row r=\"5\"><c r=\"B5\"><v>5</v></c></row>"));
        assert!(out.contains("<mergeCell ref=\"A5:B5\"/>"));
    }

    #[test]
    fn shared_formulas_are_expanded() {
        let xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A3\" si=\"0\">B1*2</f><v>2</v></c></row><row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row><row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row></sheetData></worksheet>";
        let edit = rows(1, 1);
        let f = |t: &str| shift_formula(t, "S", &edit).map_err(|e| e.to_string());
        let out = transform_sheet(xml, &SheetTransform { formula: &f, positions: Some(&edit) }).unwrap();
        assert!(out.contains("<c r=\"A1\"><f>B1*2</f>"));
        assert!(out.contains("<c r=\"A3\"><f>B3*2</f>"));
        assert!(out.contains("<c r=\"A4\"><f>B4*2</f>"));
    }

    #[test]
    fn deleting_rows_drops_cells_and_merges() {
        let xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row><row r=\"3\"><c r=\"A3\"><f>A2+A1</f></c></row></sheetData><mergeCells count=\"1\"><mergeCell ref=\"A2:B2\"/></mergeCells></worksheet>";
        let edit = StructuralEdit::DeleteRows { sheet: "S".into(), at: 1, count: 1 };
        let f = |t: &str| shift_formula(t, "S", &edit).map_err(|e| e.to_string());
        let out = transform_sheet(xml, &SheetTransform { formula: &f, positions: Some(&edit) }).unwrap();
        assert!(!out.contains("<v>2</v>"));
        assert!(out.contains("<row r=\"2\"><c r=\"A2\"><f>#REF!+A1</f></c></row>"));
        assert!(!out.contains("mergeCell"));
    }

    #[test]
    fn columns_shift() {
        let xml = "<worksheet><cols><col min=\"1\" max=\"1\" width=\"5\"/><col min=\"2\" max=\"4\" width=\"9\"/></cols><sheetData><row r=\"1\" spans=\"1:3\"><c r=\"A1\"/><c r=\"C1\"><f>A1</f></c></row></sheetData></worksheet>";
        let edit = StructuralEdit::InsertCols { sheet: "S".into(), at: 1, count: 2 };
        let f = |t: &str| shift_formula(t, "S", &edit).map_err(|e| e.to_string());
        let out = transform_sheet(xml, &SheetTransform { formula: &f, positions: Some(&edit) }).unwrap();
        assert!(out.contains("<col min=\"1\" max=\"1\" width=\"5\"/><col min=\"4\" max=\"6\" width=\"9\"/>"));
        assert!(out.contains("<c r=\"E1\"><f>A1</f></c>"));
        assert!(!out.contains("spans"));
    }

    #[test]
    fn anchors_and_vml() {
        let edit = rows(2, 1);
        assert_eq!(map_anchor(&edit, 1, true), 1);
        assert_eq!(map_anchor(&edit, 2, true), 3);
        let vml = "<x:ClientData><x:Anchor>1, 15, 3, 10, 3, 15, 7, 4</x:Anchor><x:Row>3</x:Row><x:Column>1</x:Column></x:ClientData>";
        let out = transform_vml(vml, &edit);
        assert!(out.contains("<x:Anchor>1, 15, 4, 10, 3, 15, 8, 4</x:Anchor><x:Row>4</x:Row><x:Column>1</x:Column>"));
    }
}
