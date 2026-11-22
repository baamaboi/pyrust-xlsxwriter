//! Worksheet edits and the streaming rewrite of a worksheet part.
//!
//! Edits are recorded in [`SheetEdits`] (a sparse map; nothing is rendered
//! until save). [`rewrite`] then streams the original sheet XML:
//!
//! * rows without edits are copied byte-for-byte;
//! * in edited rows, untouched cells are copied byte-for-byte and edited
//!   cells are re-rendered (keeping or adapting their original style);
//! * a few top-level elements are rebuilt when edited (`<cols>`,
//!   `<sheetViews>`, `<mergeCells>`, `<hyperlinks>`, `<sheetPr>`,
//!   `<dimension>`), inserted in schema order if missing;
//! * everything else is copied verbatim.
//!
//! The output is built in one `String` sized for the whole sheet: rows and
//! cells are appended to it directly, never assembled in a second buffer.
//! Writing to a `String` cannot fail, so `write!` results are ignored.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt::Write;
use std::sync::Arc;

use crate::common::cellref::{cell_name, parse_cell, parse_range, range_name, CellRange};
use crate::common::color::Color;
use crate::common::format::FormatSpec;
use crate::common::value::Url;
use crate::editor::package::{EResult, Relationships};
use crate::editor::sst::SharedStrings;
use crate::editor::styles::{format_number, Styles};
use crate::editor::xml::{attr, attrs, escape_attr, escape_text, local, parse_document, Child, Node, Scanner, Tok};

pub const WORKSHEET_ORDER: &[&str] = &[
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
    "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState",
    "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
    "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
    "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags",
    "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
    "controls", "webPublishItems", "tableParts", "extLst",
];

/// How an edit affects a cell's style.
///
/// Formats are behind an `Arc`: a format applied to a whole range is stored
/// once and shared by every cell edit, and the op stays pointer-sized.
#[derive(Clone, Debug, PartialEq)]
pub enum StyleOp {
    /// Keep the cell's current style.
    Keep,
    /// Use exactly this format (on top of the workbook default).
    Replace(Arc<FormatSpec>),
    /// Lay this format's set properties over the current style.
    Merge(Arc<FormatSpec>),
    /// Keep the current style if it already shows dates, otherwise add this number format.
    DefaultNumFormat(&'static str),
}

impl StyleOp {
    /// `self` followed by `next`.
    pub fn then(self, next: StyleOp) -> StyleOp {
        match (self, next) {
            (_, StyleOp::Replace(r)) => StyleOp::Replace(r),
            (current, StyleOp::Keep) => current,
            (StyleOp::Keep, next) => next,
            (StyleOp::Merge(a), StyleOp::Merge(b)) => StyleOp::Merge(Arc::new(a.overlay(&b))),
            (StyleOp::Replace(a), StyleOp::Merge(b)) => StyleOp::Replace(Arc::new(a.overlay(&b))),
            (StyleOp::DefaultNumFormat(code), StyleOp::Merge(b)) => StyleOp::Merge(Arc::new(
                FormatSpec { num_format: Some(code.to_string()), ..Default::default() }.overlay(&b),
            )),
            (current, StyleOp::DefaultNumFormat(_)) => current,
        }
    }
}

/// A new cell value. Formulas are boxed so the common values keep an edit small.
#[derive(Clone, Debug, PartialEq)]
pub enum EditValue {
    Blank,
    Number(f64),
    Str(String),
    Bool(bool),
    Error(String),
    Formula(Box<FormulaEdit>),
}

/// A formula written to a cell.
#[derive(Clone, Debug, PartialEq)]
pub struct FormulaEdit {
    /// The formula without the leading '='.
    pub text: String,
    /// The cached result to store with it, if any.
    pub result: Option<String>,
    /// The range of an array formula.
    pub array_ref: Option<String>,
}

impl EditValue {
    /// A formula value; `text` has no leading '='.
    pub fn formula(text: String, result: Option<String>, array_ref: Option<String>) -> EditValue {
        EditValue::Formula(Box::new(FormulaEdit { text, result, array_ref }))
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum CellEdit {
    Set { value: EditValue, style: StyleOp },
    Style(StyleOp),
    Clear,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct RowEdit {
    pub height: Option<Option<f64>>,
    pub hidden: Option<bool>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ColEdit {
    pub first: u16,
    pub last: u16,
    pub width: Option<f64>,
    pub hidden: Option<bool>,
    pub style: Option<StyleOp>,
}

#[derive(Clone, Debug, Default)]
pub struct SheetEdits {
    pub cells: BTreeMap<(u32, u16), CellEdit>,
    pub rows: BTreeMap<u32, RowEdit>,
    pub cols: Vec<ColEdit>,
    pub merges_add: Vec<CellRange>,
    pub merges_remove: Vec<CellRange>,
    pub freeze: Option<Option<(u32, u16)>>,
    pub tab_color: Option<Option<Color>>,
    pub zoom: Option<u16>,
    pub hyperlinks: Vec<((u32, u16), Url)>,
}

impl SheetEdits {
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
            && self.rows.is_empty()
            && self.cols.is_empty()
            && self.merges_add.is_empty()
            && self.merges_remove.is_empty()
            && self.freeze.is_none()
            && self.tab_color.is_none()
            && self.zoom.is_none()
            && self.hyperlinks.is_empty()
    }

    pub fn has_cell_edits(&self) -> bool {
        !self.cells.is_empty()
    }

    /// Whether the rewrite may add shared strings: a text value, or any value
    /// set on a table header cell (headers are always stored as text).
    pub fn adds_strings(&self, table_headers: &HashMap<(u32, u16), (String, usize)>) -> bool {
        self.cells.iter().any(|(pos, edit)| match edit {
            CellEdit::Set { value, .. } => matches!(value, EditValue::Str(_)) || table_headers.contains_key(pos),
            _ => false,
        })
    }

    /// Record a cell edit, combining it with an earlier edit of the same cell.
    pub fn set_cell(&mut self, row: u32, col: u16, edit: CellEdit) {
        let key = (row, col);
        let combined = match (self.cells.remove(&key), edit) {
            (_, CellEdit::Clear) => CellEdit::Clear,
            (Some(CellEdit::Set { style: old, .. }) | Some(CellEdit::Style(old)), CellEdit::Set { value, style }) => {
                CellEdit::Set { value, style: old.then(style) }
            }
            (Some(CellEdit::Set { value, style: old }), CellEdit::Style(op)) => CellEdit::Set { value, style: old.then(op) },
            (Some(CellEdit::Style(old)), CellEdit::Style(op)) => CellEdit::Style(old.then(op)),
            (Some(CellEdit::Clear), CellEdit::Style(op)) => {
                let op = match op {
                    StyleOp::Merge(spec) => StyleOp::Replace(spec),
                    other => other,
                };
                CellEdit::Set { value: EditValue::Blank, style: op }
            }
            (Some(CellEdit::Clear), CellEdit::Set { value, style }) => {
                let style = match style {
                    StyleOp::Merge(spec) => StyleOp::Replace(spec),
                    other => other,
                };
                CellEdit::Set { value, style }
            }
            (None, edit) => edit,
        };
        self.cells.insert(key, combined);
    }

    /// Bounding box of cells that will hold something after the edits.
    fn edit_bounds(&self) -> Option<CellRange> {
        let mut bounds: Option<CellRange> = None;
        for ((r, c), edit) in &self.cells {
            if matches!(edit, CellEdit::Clear) {
                continue;
            }
            bounds = Some(match bounds {
                None => (*r, *c, *r, *c),
                Some((r1, c1, r2, c2)) => (r1.min(*r), c1.min(*c), r2.max(*r), c2.max(*c)),
            });
        }
        bounds
    }
}

/// Everything the rewrite needs besides the sheet XML.
pub struct RewriteContext<'a> {
    pub styles: &'a mut Styles,
    /// The shared string table; None when the edits add no strings (see
    /// [`SheetEdits::adds_strings`]), so a large table isn't read for nothing.
    pub sst: Option<&'a mut SharedStrings>,
    /// The formula of a cell of the original sheet, expanded from its shared
    /// formula (loaded on first use).
    pub formula_at: &'a mut dyn FnMut(u32, u16) -> EResult<Option<String>>,
    /// The sheet's relationships (for hyperlinks); marked dirty if changed.
    pub rels: &'a mut Relationships,
    pub rels_dirty: bool,
    /// Table header cells: (row, col) -> (table part, column index).
    pub table_headers: HashMap<(u32, u16), (String, usize)>,
    /// New header names to write into table parts: (part, column index, name).
    pub table_header_updates: Vec<(String, usize, String)>,
}

fn resolve_style(op: &StyleOp, original: Option<u32>, ctx: &mut RewriteContext<'_>) -> EResult<Option<u32>> {
    Ok(match op {
        StyleOp::Keep => original,
        StyleOp::Replace(spec) => Some(ctx.styles.apply(spec)?),
        StyleOp::Merge(spec) => Some(ctx.styles.merge(original.unwrap_or(0), spec)?),
        StyleOp::DefaultNumFormat(code) => match original {
            Some(xf) if ctx.styles.is_date_xf(xf) => Some(xf),
            _ => Some(ctx.styles.merge(
                original.unwrap_or(0),
                &FormatSpec { num_format: Some(code.to_string()), ..Default::default() },
            )?),
        },
    }
    .filter(|xf| *xf != 0 || original.is_some_and(|o| o != 0)))
}

fn style_attr(xf: Option<u32>) -> String {
    match xf {
        Some(xf) if xf != 0 => format!(" s=\"{xf}\""),
        _ => String::new(),
    }
}

/// Append an edited cell to `out`, styled with `xf`.
fn render_cell(
    out: &mut String,
    row: u32,
    col: u16,
    value: &EditValue,
    xf: Option<u32>,
    ctx: &mut RewriteContext<'_>,
) -> EResult<()> {
    let r = cell_name(row, col, false, false);
    let s = style_attr(xf);
    match value {
        EditValue::Blank => {
            if !s.is_empty() {
                let _ = write!(out, "<c r=\"{r}\"{s}/>");
            }
        }
        EditValue::Number(n) if n.is_nan() => {
            let _ = write!(out, "<c r=\"{r}\"{s} t=\"e\"><v>#NUM!</v></c>");
        }
        EditValue::Number(n) if n.is_infinite() => {
            let _ = write!(out, "<c r=\"{r}\"{s} t=\"e\"><v>#DIV/0!</v></c>");
        }
        EditValue::Number(n) => {
            let _ = write!(out, "<c r=\"{r}\"{s}><v>{}</v></c>", format_number(*n));
        }
        EditValue::Str(text) => {
            let sst = ctx.sst.as_deref_mut().ok_or("internal error: the shared string table was not loaded")?;
            let _ = write!(out, "<c r=\"{r}\"{s} t=\"s\"><v>{}</v></c>", sst.add(text));
        }
        EditValue::Bool(b) => {
            let _ = write!(out, "<c r=\"{r}\"{s} t=\"b\"><v>{}</v></c>", u8::from(*b));
        }
        EditValue::Error(e) => {
            let _ = write!(out, "<c r=\"{r}\"{s} t=\"e\"><v>{}</v></c>", escape_text(e));
        }
        EditValue::Formula(formula) => {
            let FormulaEdit { text, result, array_ref } = &**formula;
            let numeric = result.as_deref().is_some_and(|v| v.parse::<f64>().is_ok());
            let kind = match result.as_deref() {
                Some(_) if numeric => "",
                Some("TRUE" | "FALSE") => " t=\"b\"",
                Some(_) => " t=\"str\"",
                None => "",
            };
            let _ = write!(out, "<c r=\"{r}\"{s}{kind}>");
            let _ = match array_ref {
                Some(range) => write!(out, "<f t=\"array\" ref=\"{range}\">{}</f>", escape_text(text)),
                None => write!(out, "<f>{}</f>", escape_text(text)),
            };
            let _ = match result.as_deref() {
                None => Ok(()),
                Some(v) if numeric => write!(out, "<v>{v}</v>"),
                Some("TRUE") => write!(out, "<v>1</v>"),
                Some("FALSE") => write!(out, "<v>0</v>"),
                Some(v) => write!(out, "<v>{}</v>", escape_text(v)),
            };
            out.push_str("</c>");
        }
    }
    Ok(())
}

/// Append a raw `<c ...>` element with its `s` attribute replaced (or added/removed).
fn restyle_raw_cell(out: &mut String, raw: &str, xf: Option<u32>) {
    let close = raw.find('>').unwrap_or(raw.len());
    let tag_end = if raw[..close].ends_with('/') { close - 1 } else { close };
    let head = &raw[..tag_end];
    let name_end = head.find(|c: char| c.is_ascii_whitespace()).unwrap_or(head.len());
    out.push_str(&head[..name_end]);
    for (k, v) in attrs(&head[name_end..]) {
        if k != "s" {
            let _ = write!(out, " {k}=\"{v}\"");
        }
    }
    out.push_str(&style_attr(xf));
    out.push_str(&raw[tag_end..]);
}

/// A raw cell from the original row.
struct RawCell {
    xf: Option<u32>,
    shared_master: Option<(String, String)>,
    shared_dependent: Option<String>,
    array_ref: Option<String>,
}

fn scan_cell(raw: &str, start_attrs: &str) -> RawCell {
    let xf = attr(start_attrs, "s").and_then(|s| s.parse().ok());
    let mut cell = RawCell { xf, shared_master: None, shared_dependent: None, array_ref: None };
    if let Some(i) = raw.find("<f") {
        let tag_end = raw[i..].find('>').map(|j| i + j).unwrap_or(raw.len());
        let f_attrs = &raw[i + 2..tag_end];
        match attr(f_attrs, "t").as_deref() {
            Some("shared") => {
                let si = attr(f_attrs, "si").unwrap_or_default();
                match attr(f_attrs, "ref") {
                    Some(range) => cell.shared_master = Some((si, range)),
                    None => cell.shared_dependent = Some(si),
                }
            }
            Some("array") => cell.array_ref = attr(f_attrs, "ref"),
            _ => {}
        }
    }
    cell
}

fn range_contains(range: &str, row: u32, col: u16) -> bool {
    parse_range(range).is_some_and(|(r1, c1, r2, c2)| row >= r1 && row <= r2 && col >= c1 && col <= c2)
}

struct RowState {
    broken_shared: HashSet<String>,
    arrays: Vec<(u32, u16, String)>,
}

/// Rewrite one row that has edits, appending it to `out`. `start_attrs` is
/// the raw attribute text of `<row ...>`, `body` the XML between `<row>` and
/// `</row>` (None for `<row/>`).
fn rewrite_row(
    row: u32,
    start_name: &str,
    start_attrs: Option<&str>,
    body: Option<&str>,
    edits: &SheetEdits,
    state: &mut RowState,
    ctx: &mut RewriteContext<'_>,
    out: &mut String,
) -> EResult<()> {
    // Row start tag.
    let row_edit = edits.rows.get(&row);
    out.push('<');
    out.push_str(start_name);
    let mut has_r = false;
    if let Some(raw) = start_attrs {
        for (k, v) in attrs(raw) {
            match k {
                "spans" => continue,
                "ht" | "customHeight" if row_edit.is_some_and(|e| e.height.is_some()) => continue,
                "hidden" if row_edit.is_some_and(|e| e.hidden.is_some()) => continue,
                "r" => has_r = true,
                _ => {}
            }
            let _ = write!(out, " {k}=\"{v}\"");
        }
    }
    if !has_r {
        let _ = write!(out, " r=\"{}\"", row + 1);
    }
    if let Some(edit) = row_edit {
        if let Some(Some(height)) = edit.height {
            let _ = write!(out, " ht=\"{}\" customHeight=\"1\"", format_number(height));
        }
        if edit.hidden == Some(true) {
            out.push_str(" hidden=\"1\"");
        }
    }
    let row_xf: Option<u32> = start_attrs.and_then(|raw| {
        if attr(raw, "customFormat").as_deref() == Some("1") { attr(raw, "s").and_then(|s| s.parse().ok()) } else { None }
    });

    // Cells, written right after the start tag: original cells merged with edits by column.
    out.push('>');
    let cells_start = out.len();
    let mut cell_edits = edits.cells.range((row, 0)..=(row, u16::MAX)).peekable();
    let flush_new_cells = |before: Option<u16>, out: &mut String, ctx: &mut RewriteContext<'_>, cell_edits: &mut std::iter::Peekable<std::collections::btree_map::Range<'_, (u32, u16), CellEdit>>| -> EResult<()> {
        while let Some(((_, c), edit)) = cell_edits.peek() {
            if before.is_some_and(|b| *c >= b) {
                break;
            }
            let c = *c;
            match edit {
                CellEdit::Clear => {}
                CellEdit::Style(op) => {
                    let xf = resolve_style(op, row_xf, ctx)?;
                    if xf.is_some() {
                        let _ = write!(out, "<c r=\"{}\"{}/>", cell_name(row, c, false, false), style_attr(xf));
                    }
                }
                CellEdit::Set { value, style } => {
                    let xf = resolve_style(style, row_xf, ctx)?;
                    let value = header_value(row, c, value, ctx);
                    render_cell(out, row, c, &value, xf, ctx)?;
                }
            }
            cell_edits.next();
        }
        Ok(())
    };

    if let Some(body) = body {
        let mut scanner = Scanner::new(body);
        let mut next_col: u16 = 0;
        while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
            match tok {
                Tok::Start { name, attrs: a } | Tok::Empty { name, attrs: a } if local(name) == "c" => {
                    let end = if matches!(tok, Tok::Start { .. }) {
                        scanner.skip_element(name).map_err(|e| e.to_string())?
                    } else {
                        span.end
                    };
                    let raw = &body[span.start..end];
                    let col = attr(a, "r").and_then(|r| parse_cell(&r)).map(|(_, c)| c).unwrap_or(next_col);
                    next_col = col.saturating_add(1);
                    let cell = scan_cell(raw, a);
                    flush_new_cells(Some(col), out, ctx, &mut cell_edits)?;

                    if let Some(range) = &cell.array_ref {
                        state.arrays.push((row, col, range.clone()));
                    }
                    let edit = match cell_edits.peek() {
                        Some(((_, c), _)) if *c == col => cell_edits.next().map(|(_, e)| e),
                        _ => None,
                    };
                    match edit {
                        None => {
                            // Untouched cell; a dependent of a broken shared formula
                            // gets its own formula text.
                            match &cell.shared_dependent {
                                Some(si) if state.broken_shared.contains(si) => {
                                    let text = (ctx.formula_at)(row, col)?;
                                    expand_shared(out, raw, text.as_deref());
                                }
                                _ => out.push_str(raw),
                            }
                        }
                        Some(edit) => {
                            check_array_membership(row, col, &cell, state)?;
                            if let Some((si, _)) = &cell.shared_master {
                                state.broken_shared.insert(si.clone());
                            }
                            match edit {
                                CellEdit::Clear => {}
                                CellEdit::Style(op) => {
                                    let xf = resolve_style(op, cell.xf, ctx)?;
                                    restyle_raw_cell(out, raw, xf);
                                }
                                CellEdit::Set { value, style } => {
                                    let xf = resolve_style(style, cell.xf, ctx)?;
                                    let value = header_value(row, col, value, ctx);
                                    render_cell(out, row, col, &value, xf, ctx)?;
                                }
                            }
                        }
                    }
                }
                _ => out.push_str(&body[span]),
            }
        }
    }
    flush_new_cells(None, out, ctx, &mut cell_edits)?;

    if out.len() == cells_start {
        // No cells left: close the start tag as `<row .../>`.
        out.pop();
        out.push_str("/>");
    } else {
        let _ = write!(out, "</{start_name}>");
    }
    Ok(())
}

/// Table header cells must hold text that matches the table's column name.
/// Other cells keep their value (borrowed, not copied).
fn header_value<'v>(row: u32, col: u16, value: &'v EditValue, ctx: &mut RewriteContext<'_>) -> Cow<'v, EditValue> {
    let Some((part, index)) = ctx.table_headers.get(&(row, col)).cloned() else {
        return Cow::Borrowed(value);
    };
    let text = match value {
        EditValue::Str(s) => s.clone(),
        EditValue::Number(n) => format_number(*n),
        EditValue::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        EditValue::Error(e) => e.clone(),
        EditValue::Blank => format!("Column{}", index + 1),
        EditValue::Formula(f) => f.text.clone(),
    };
    ctx.table_header_updates.push((part, index, text.clone()));
    Cow::Owned(EditValue::Str(text))
}

fn check_array_membership(row: u32, col: u16, cell: &RawCell, state: &RowState) -> EResult<()> {
    for (mr, mc, range) in &state.arrays {
        if (*mr, *mc) != (row, col) && range_contains(range, row, col) && cell.array_ref.is_none() {
            return Err(format!(
                "cell {} is part of the array formula in {} ({}); write the formula's top-left cell instead",
                cell_name(row, col, false, false),
                cell_name(*mr, *mc, false, false),
                range
            ));
        }
    }
    Ok(())
}

/// Append a dependent cell of a broken shared formula, with its
/// `<f t="shared" si=".."/>` replaced by an explicit formula.
fn expand_shared(out: &mut String, raw: &str, text: Option<&str>) {
    let Some(text) = text else {
        // Unknown expansion: drop the formula, keep the cached value.
        return remove_formula(out, raw);
    };
    let Some(start) = raw.find("<f") else { return out.push_str(raw) };
    let tag_end = match raw[start..].find('>') {
        Some(i) => start + i,
        None => return out.push_str(raw),
    };
    let end = if raw[..tag_end].ends_with('/') {
        tag_end + 1
    } else {
        match raw[tag_end..].find("</f>") {
            Some(i) => tag_end + i + 4,
            None => return out.push_str(raw),
        }
    };
    let body = text.strip_prefix('=').unwrap_or(text);
    let _ = write!(out, "{}<f>{}</f>{}", &raw[..start], escape_text(body), &raw[end..]);
}

/// Append a raw cell without its `<f>` element.
fn remove_formula(out: &mut String, raw: &str) {
    let Some(start) = raw.find("<f") else { return out.push_str(raw) };
    let tag_end = match raw[start..].find('>') {
        Some(i) => start + i,
        None => return out.push_str(raw),
    };
    let end = if raw[..tag_end].ends_with('/') {
        tag_end + 1
    } else {
        raw[tag_end..].find("</f>").map(|i| tag_end + i + 4).unwrap_or(tag_end + 1)
    };
    out.push_str(&raw[..start]);
    out.push_str(&raw[end..]);
}

fn color_node(name: &str, color: &Color) -> Option<Node> {
    use rust_xlsxwriter::Color as RColor;
    match color.0 {
        RColor::Default => None,
        RColor::Automatic => Some(Node::new(name).with_attr("auto", "1")),
        RColor::Theme(c, _) => Some(Node::new(name).with_attr("theme", &c.to_string())),
        _ => color.rgb_value().map(|v| Node::new(name).with_attr("rgb", &format!("FF{v:06X}"))),
    }
}

/// Rebuild `<cols>` with column edits applied.
fn rewrite_cols(existing: Option<&str>, edits: &[ColEdit], ctx: &mut RewriteContext<'_>, prefix: &str) -> EResult<String> {
    let mut columns: BTreeMap<u16, Vec<(String, String)>> = BTreeMap::new();
    if let Some(xml) = existing {
        let doc = parse_document(xml).map_err(|e| e.to_string())?;
        for col in doc.root.children_named("col") {
            let min: u16 = col.get("min").and_then(|v| v.parse().ok()).unwrap_or(1);
            let max: u16 = col.get("max").and_then(|v| v.parse().ok()).unwrap_or(min);
            let attrs: Vec<(String, String)> = col.attrs.iter().filter(|(k, _)| k != "min" && k != "max").cloned().collect();
            for c in min.max(1)..=max.min(16_384) {
                columns.insert(c, attrs.clone());
            }
        }
    }
    for edit in edits {
        for c in (edit.first + 1)..=(edit.last + 1) {
            let attrs = columns.entry(c).or_default();
            let mut set = |k: &str, v: Option<String>| {
                attrs.retain(|(name, _)| name != k);
                if let Some(v) = v {
                    attrs.push((k.to_string(), v));
                }
            };
            if let Some(width) = edit.width {
                set("width", Some(format_number(width)));
                set("customWidth", Some("1".into()));
            }
            if let Some(hidden) = edit.hidden {
                set("hidden", hidden.then(|| "1".into()));
            }
            if let Some(op) = &edit.style {
                let current = attrs.iter().find(|(k, _)| k == "style").and_then(|(_, v)| v.parse().ok());
                let xf = resolve_style(op, current, ctx)?;
                attrs.retain(|(k, _)| k != "style");
                if let Some(xf) = xf {
                    attrs.push(("style".into(), xf.to_string()));
                }
            }
            if !attrs.iter().any(|(k, _)| k == "width") {
                attrs.push(("width".into(), "8.7109375".into()));
            }
        }
    }
    let mut out = format!("<{prefix}cols>");
    let mut iter = columns.into_iter().peekable();
    while let Some((start, attrs)) = iter.next() {
        let mut end = start;
        while let Some((next, next_attrs)) = iter.peek() {
            if *next == end + 1 && *next_attrs == attrs {
                end = *next;
                iter.next();
            } else {
                break;
            }
        }
        out.push_str(&format!("<{prefix}col min=\"{start}\" max=\"{end}\""));
        for (k, v) in &attrs {
            out.push_str(&format!(" {k}=\"{v}\""));
        }
        out.push_str("/>");
    }
    out.push_str(&format!("</{prefix}cols>"));
    Ok(out)
}

fn rewrite_sheet_views(existing: Option<&str>, edits: &SheetEdits, prefix: &str) -> EResult<String> {
    let mut root = match existing {
        Some(xml) => parse_document(xml).map_err(|e| e.to_string())?.root,
        None => Node::new(&format!("{prefix}sheetViews")),
    };
    if root.child("sheetView").is_none() {
        root.push(Node::new(&format!("{prefix}sheetView")).with_attr("workbookViewId", "0"));
    }
    let view = root.child_mut("sheetView").ok_or("missing sheetView")?;
    if let Some(zoom) = edits.zoom {
        view.set("zoomScale", &zoom.to_string());
        view.set("zoomScaleNormal", &zoom.to_string());
    }
    if let Some(freeze) = edits.freeze {
        view.remove_children("pane");
        view.remove_children("selection");
        if let Some((row, col)) = freeze.filter(|(r, c)| *r > 0 || *c > 0) {
            let pane_name = match (row > 0, col > 0) {
                (true, true) => "bottomRight",
                (true, false) => "bottomLeft",
                _ => "topRight",
            };
            let mut pane = Node::new(&format!("{prefix}pane"));
            if col > 0 {
                pane.set("xSplit", &col.to_string());
            }
            if row > 0 {
                pane.set("ySplit", &row.to_string());
            }
            let top_left = cell_name(row, col, false, false);
            pane.set("topLeftCell", &top_left);
            pane.set("activePane", pane_name);
            pane.set("state", "frozen");
            view.children.insert(0, Child::Element(pane));
            view.self_closing = false;
            view.push(
                Node::new(&format!("{prefix}selection"))
                    .with_attr("pane", pane_name)
                    .with_attr("activeCell", &top_left)
                    .with_attr("sqref", &top_left),
            );
        }
    }
    Ok(root.to_xml())
}

fn rewrite_merges(existing: Option<&str>, edits: &SheetEdits, prefix: &str) -> EResult<String> {
    let mut ranges: Vec<String> = Vec::new();
    if let Some(xml) = existing {
        let doc = parse_document(xml).map_err(|e| e.to_string())?;
        ranges = doc.root.children_named("mergeCell").filter_map(|m| m.get("ref")).collect();
    }
    let removed: HashSet<String> = edits.merges_remove.iter().map(|r| range_name(*r, false)).collect();
    ranges.retain(|r| !removed.contains(&r.replace('$', "")));
    for range in &edits.merges_add {
        let name = range_name(*range, false);
        if !ranges.contains(&name) {
            ranges.push(name);
        }
    }
    if ranges.is_empty() {
        return Ok(String::new());
    }
    let mut out = format!("<{prefix}mergeCells count=\"{}\">", ranges.len());
    for r in ranges {
        out.push_str(&format!("<{prefix}mergeCell ref=\"{}\"/>", escape_attr(&r)));
    }
    out.push_str(&format!("</{prefix}mergeCells>"));
    Ok(out)
}

fn rewrite_hyperlinks(existing: Option<&str>, edits: &SheetEdits, ctx: &mut RewriteContext<'_>, prefix: &str) -> EResult<String> {
    let mut root = match existing {
        Some(xml) => parse_document(xml).map_err(|e| e.to_string())?.root,
        None => Node::new(&format!("{prefix}hyperlinks")),
    };
    for ((row, col), url) in &edits.hyperlinks {
        let cell = cell_name(*row, *col, false, false);
        // Drop an existing link on the same cell (and its relationship).
        let old: Vec<String> = root
            .children_named("hyperlink")
            .filter(|h| h.get("ref").as_deref() == Some(cell.as_str()))
            .filter_map(|h| h.get("r:id"))
            .collect();
        for id in old {
            ctx.rels.remove(&id);
            ctx.rels_dirty = true;
        }
        root.children.retain(|c| !matches!(c, Child::Element(h) if h.local() == "hyperlink" && h.get("ref").as_deref() == Some(cell.as_str())));
        let mut link = Node::new(&format!("{prefix}hyperlink")).with_attr("ref", &cell);
        if let Some(location) = url.link.strip_prefix("internal:") {
            link.set("location", location);
        } else {
            let (target, location) = match url.link.split_once('#') {
                Some((t, l)) => (t.to_string(), Some(l.to_string())),
                None => (url.link.clone(), None),
            };
            let target = target.strip_prefix("external:").unwrap_or(&target).to_string();
            let id = ctx.rels.add("hyperlink", &target, true);
            ctx.rels_dirty = true;
            link.set("r:id", &id);
            if let Some(location) = location {
                link.set("location", &location);
            }
        }
        if let Some(tip) = &url.tip {
            link.set("tooltip", tip);
        }
        if let Some(text) = &url.text {
            link.set("display", text);
        }
        root.push(link);
    }
    if root.elements().next().is_none() {
        return Ok(String::new());
    }
    Ok(root.to_xml())
}

fn rewrite_sheet_pr(existing: Option<&str>, color: Option<Color>, prefix: &str) -> EResult<String> {
    let mut root = match existing {
        Some(xml) => parse_document(xml).map_err(|e| e.to_string())?.root,
        None => Node::new(&format!("{prefix}sheetPr")),
    };
    root.remove_children("tabColor");
    if let Some(node) = color.as_ref().and_then(|c| color_node(&format!("{prefix}tabColor"), c)) {
        root.children.insert(0, Child::Element(node));
        root.self_closing = false;
    }
    if root.children.is_empty() {
        root.self_closing = true;
    }
    Ok(root.to_xml())
}

/// Rewrite a worksheet's XML with `edits` applied.
pub fn rewrite(xml: &str, edits: &SheetEdits, ctx: &mut RewriteContext<'_>) -> EResult<String> {
    let mut scanner = Scanner::new(xml);
    let mut out = String::with_capacity(xml.len() + edits.cells.len() * 40 + 1024);

    // Prolog and the root start tag.
    let (root_name, prefix) = loop {
        let (tok, span) = scanner.next_tok().map_err(|e| e.to_string())?.ok_or("empty worksheet part")?;
        match tok {
            Tok::Start { name, attrs: a } => {
                let prefix = match name.rfind(':') {
                    Some(i) => name[..=i].to_string(),
                    None => String::new(),
                };
                out.push_str(&xml[..span.start]);
                out.push('<');
                out.push_str(name);
                out.push_str(a);
                let needs_r = !edits.hyperlinks.is_empty()
                    && !attrs(a).any(|(k, v)| k.starts_with("xmlns:") && v == crate::editor::package::NS_REL);
                if needs_r {
                    out.push_str(&format!(" xmlns:r=\"{}\"", crate::editor::package::NS_REL));
                }
                out.push('>');
                break (name.to_string(), prefix);
            }
            Tok::Empty { .. } => return Err("worksheet root element is empty".into()),
            _ => {}
        }
    };

    // Elements we may need to create, with their rendering.
    let mut wanted: HashSet<&str> = HashSet::new();
    if edits.tab_color.is_some() {
        wanted.insert("sheetPr");
    }
    if edits.freeze.is_some() || edits.zoom.is_some() {
        wanted.insert("sheetViews");
    }
    if !edits.cols.is_empty() {
        wanted.insert("cols");
    }
    if !edits.merges_add.is_empty() || !edits.merges_remove.is_empty() {
        wanted.insert("mergeCells");
    }
    if !edits.hyperlinks.is_empty() {
        wanted.insert("hyperlinks");
    }
    let mut seen: HashSet<String> = HashSet::new();
    let bounds = edits.edit_bounds();
    let mut row_state = RowState { broken_shared: HashSet::new(), arrays: Vec::new() };

    let render_missing = |name: &str, ctx: &mut RewriteContext<'_>| -> EResult<String> {
        Ok(match name {
            "sheetPr" => rewrite_sheet_pr(None, edits.tab_color.flatten(), &prefix)?,
            "sheetViews" => rewrite_sheet_views(None, edits, &prefix)?,
            "cols" => rewrite_cols(None, &edits.cols, ctx, &prefix)?,
            "mergeCells" => rewrite_merges(None, edits, &prefix)?,
            "hyperlinks" => rewrite_hyperlinks(None, edits, ctx, &prefix)?,
            _ => String::new(),
        })
    };
    // Insert any wanted-but-missing elements that come before `rank`.
    let flush_missing = |rank: usize, out: &mut String, seen: &mut HashSet<String>, ctx: &mut RewriteContext<'_>| -> EResult<()> {
        for name in WORKSHEET_ORDER.iter().take(rank) {
            if wanted.contains(name) && !seen.contains(*name) {
                seen.insert(name.to_string());
                out.push_str(&render_missing(name, ctx)?);
            }
        }
        Ok(())
    };

    loop {
        let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? else {
            return Err("worksheet part ended early".into());
        };
        match tok {
            Tok::End { name } if name == root_name => {
                flush_missing(WORKSHEET_ORDER.len(), &mut out, &mut seen, ctx)?;
                out.push_str(&xml[span.start..]);
                return Ok(out);
            }
            Tok::Start { name, attrs: a } | Tok::Empty { name, attrs: a } => {
                let is_start = matches!(tok, Tok::Start { .. });
                let lname = local(name);
                let rank = WORKSHEET_ORDER.iter().position(|n| *n == lname).unwrap_or(WORKSHEET_ORDER.len());
                flush_missing(rank, &mut out, &mut seen, ctx)?;
                seen.insert(lname.to_string());
                let end = if is_start { scanner.skip_element(name).map_err(|e| e.to_string())? } else { span.end };
                let element = &xml[span.start..end];
                match lname {
                    "sheetData" => {
                        let body = if is_start { Some(&xml[span.end..end - format!("</{name}>").len()]) } else { None };
                        rewrite_sheet_data(name, a, body, edits, &mut row_state, ctx, &mut out)?;
                    }
                    "dimension" => {
                        let original = attr(a, "ref").and_then(|r| parse_range(&r));
                        let merged = match (original, bounds) {
                            (Some((r1, c1, r2, c2)), Some((b1, b2, b3, b4))) => Some((r1.min(b1), c1.min(b2), r2.max(b3), c2.max(b4))),
                            (o, b) => o.or(b),
                        };
                        match merged {
                            Some(range) => out.push_str(&format!("<{name} ref=\"{}\"/>", range_name(range, false))),
                            None => out.push_str(element),
                        }
                    }
                    "sheetPr" if wanted.contains("sheetPr") => out.push_str(&rewrite_sheet_pr(Some(element), edits.tab_color.flatten(), &prefix)?),
                    "sheetViews" if wanted.contains("sheetViews") => out.push_str(&rewrite_sheet_views(Some(element), edits, &prefix)?),
                    "cols" if wanted.contains("cols") => out.push_str(&rewrite_cols(Some(element), &edits.cols, ctx, &prefix)?),
                    "mergeCells" if wanted.contains("mergeCells") => out.push_str(&rewrite_merges(Some(element), edits, &prefix)?),
                    "hyperlinks" if wanted.contains("hyperlinks") => out.push_str(&rewrite_hyperlinks(Some(element), edits, ctx, &prefix)?),
                    _ => out.push_str(element),
                }
            }
            _ => out.push_str(&xml[span]),
        }
    }
}

/// Rewrite `<sheetData>`, appending it to `out`. Rows without edits are
/// copied as they are; the rows go straight into `out` (no second buffer).
fn rewrite_sheet_data(
    name: &str,
    start_attrs: &str,
    body: Option<&str>,
    edits: &SheetEdits,
    state: &mut RowState,
    ctx: &mut RewriteContext<'_>,
    out: &mut String,
) -> EResult<()> {
    let edit_rows: BTreeSet<u32> = edits.cells.keys().map(|(r, _)| *r).chain(edits.rows.keys().copied()).collect();
    let mut pending = edit_rows.iter().copied().peekable();
    let row_prefix = match name.rfind(':') {
        Some(i) => &name[..=i],
        None => "",
    };
    let row_tag = format!("{row_prefix}row");
    out.push('<');
    out.push_str(name);
    out.push_str(start_attrs);
    out.push('>');
    let rows_start = out.len();
    if let Some(body) = body {
        let mut scanner = Scanner::new(body);
        let mut next_row: u32 = 0;
        while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
            match tok {
                Tok::Start { name: rn, attrs: a } | Tok::Empty { name: rn, attrs: a } if local(rn) == "row" => {
                    let is_start = matches!(tok, Tok::Start { .. });
                    let end = if is_start { scanner.skip_element(rn).map_err(|e| e.to_string())? } else { span.end };
                    // The row's content, between `<row ...>` and `</row>`.
                    let inner = is_start.then(|| &body[span.end..end - (rn.len() + 3)]);
                    let row = attr(a, "r").and_then(|r| r.parse::<u32>().ok()).map(|r| r.saturating_sub(1)).unwrap_or(next_row);
                    next_row = row + 1;
                    while let Some(new_row) = pending.next_if(|r| *r < row) {
                        rewrite_row(new_row, &row_tag, None, None, edits, state, ctx, out)?;
                    }
                    if pending.next_if(|r| *r == row).is_some() {
                        rewrite_row(row, rn, Some(a), inner, edits, state, ctx, out)?;
                    } else if state.broken_shared.is_empty() || !body[span.start..end].contains("t=\"shared\"") {
                        let raw_row = &body[span.start..end];
                        if raw_row.contains("t=\"array\"") {
                            record_arrays(row, raw_row, state)?;
                        }
                        out.push_str(raw_row);
                    } else {
                        // Untouched row that may hold dependents of a broken shared formula.
                        rewrite_row(row, rn, Some(a), inner, edits, state, ctx, out)?;
                    }
                }
                _ => out.push_str(&body[span]),
            }
        }
    }
    for new_row in pending {
        rewrite_row(new_row, &row_tag, None, None, edits, state, ctx, out)?;
    }
    if out.len() == rows_start {
        // No rows: close the start tag as `<sheetData/>`.
        out.pop();
        out.push_str("/>");
    } else {
        let _ = write!(out, "</{name}>");
    }
    Ok(())
}

/// Remember the array formula ranges of an untouched row.
fn record_arrays(row: u32, raw_row: &str, state: &mut RowState) -> EResult<()> {
    let mut scanner = Scanner::new(raw_row);
    let mut next_col: u16 = 0;
    while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
        if let Tok::Start { name, attrs: a } = tok
            && local(name) == "c" {
                let end = scanner.skip_element(name).map_err(|e| e.to_string())?;
                let col = attr(a, "r").and_then(|r| parse_cell(&r)).map(|(_, c)| c).unwrap_or(next_col);
                next_col = col.saturating_add(1);
                if let Some(range) = scan_cell(&raw_row[span.start..end], a).array_ref {
                    state.arrays.push((row, col, range));
                }
            }
    }
    Ok(())
}

/// A minimal worksheet part for new sheets.
pub fn blank_sheet_xml() -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{}\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>",
        crate::editor::package::NS_REL
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::editor::styles::Styles;
    use crate::editor::xml::parse_document;

    fn ctx_parts() -> (Styles, SharedStrings, Relationships) {
        let styles = Styles::new("xl/styles.xml", parse_document(Styles::minimal_xml()).unwrap());
        let sst = SharedStrings::empty("xl/sharedStrings.xml");
        let rels = crate::editor::package::Relationships::empty("xl/worksheets/sheet1.xml");
        (styles, sst, rels)
    }

    fn run(xml: &str, edits: &SheetEdits) -> (String, SharedStrings) {
        let (mut styles, mut sst, mut rels) = ctx_parts();
        let mut loader = |_, _| Ok(None);
        let mut ctx = RewriteContext {
            styles: &mut styles,
            sst: Some(&mut sst),
            formula_at: &mut loader,
            rels: &mut rels,
            rels_dirty: false,
            table_headers: HashMap::new(),
            table_header_updates: Vec::new(),
        };
        let out = rewrite(xml, edits, &mut ctx).unwrap();
        (out, sst)
    }

    const SHEET: &str = "<?xml version=\"1.0\"?><worksheet xmlns=\"m\"><dimension ref=\"A1:B2\"/><sheetData><row r=\"1\" spans=\"1:2\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>0</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row></sheetData><pageMargins left=\"0.7\"/></worksheet>";

    #[test]
    fn untouched_rows_are_verbatim() {
        let mut edits = SheetEdits::default();
        edits.set_cell(0, 2, CellEdit::Set { value: EditValue::Number(5.0), style: StyleOp::Keep });
        let (out, _) = run(SHEET, &edits);
        assert!(out.contains("<row r=\"3\"><c r=\"A3\"><v>3</v></c></row>"));
        assert!(out.contains("<c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>0</v></c><c r=\"C1\"><v>5</v></c>"));
        assert!(out.contains("<dimension ref=\"A1:C2\"/>"));
        assert!(!out.contains("spans"));
    }

    #[test]
    fn new_rows_are_inserted_in_order() {
        let mut edits = SheetEdits::default();
        edits.set_cell(1, 0, CellEdit::Set { value: EditValue::Str("x".into()), style: StyleOp::Keep });
        edits.set_cell(9, 1, CellEdit::Set { value: EditValue::Bool(true), style: StyleOp::Keep });
        let (out, sst) = run(SHEET, &edits);
        let i1 = out.find("<row r=\"1\"").unwrap();
        let i2 = out.find("<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>").unwrap();
        let i3 = out.find("<row r=\"3\"").unwrap();
        let i10 = out.find("<row r=\"10\"><c r=\"B10\" t=\"b\"><v>1</v></c></row>").unwrap();
        assert!(i1 < i2 && i2 < i3 && i3 < i10);
        assert!(sst.is_dirty());
    }

    #[test]
    fn keeps_style_and_clears() {
        let mut edits = SheetEdits::default();
        edits.set_cell(0, 0, CellEdit::Set { value: EditValue::Number(2.0), style: StyleOp::Keep });
        edits.set_cell(0, 1, CellEdit::Clear);
        let (out, _) = run(SHEET, &edits);
        assert!(out.contains("<row r=\"1\"><c r=\"A1\" s=\"1\"><v>2</v></c></row>"));
    }

    #[test]
    fn merges_cols_and_panes_are_inserted_in_schema_order() {
        let mut edits = SheetEdits::default();
        edits.merges_add.push((0, 0, 0, 3));
        edits.cols.push(ColEdit { first: 1, last: 2, width: Some(20.0), hidden: None, style: None });
        edits.freeze = Some(Some((1, 0)));
        let (out, _) = run(SHEET, &edits);
        let views = out.find("<sheetViews>").unwrap();
        let cols = out.find("<cols><col min=\"2\" max=\"3\" width=\"20\" customWidth=\"1\"/></cols>").unwrap();
        let data = out.find("<sheetData>").unwrap();
        let merges = out.find("<mergeCells count=\"1\"><mergeCell ref=\"A1:D1\"/></mergeCells>").unwrap();
        let margins = out.find("<pageMargins").unwrap();
        assert!(views < cols && cols < data && data < merges && merges < margins);
        assert!(out.contains("ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\""));
    }

    #[test]
    fn shared_formula_master_edit_expands_dependents() {
        let xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A2\" si=\"0\">B1*2</f><v>2</v></c></row><row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row></sheetData></worksheet>";
        let mut edits = SheetEdits::default();
        edits.set_cell(0, 0, CellEdit::Set { value: EditValue::Number(7.0), style: StyleOp::Keep });
        let (mut styles, mut sst, mut rels) = ctx_parts();
        let expanded = HashMap::from([((1, 0), "B2*2".to_string())]);
        let mut loader = |row, col| Ok(expanded.get(&(row, col)).cloned());
        let mut ctx = RewriteContext { styles: &mut styles, sst: Some(&mut sst), formula_at: &mut loader, rels: &mut rels, rels_dirty: false, table_headers: HashMap::new(), table_header_updates: Vec::new() };
        let out = rewrite(xml, &edits, &mut ctx).unwrap();
        assert!(out.contains("<c r=\"A1\"><v>7</v></c>"));
        assert!(out.contains("<c r=\"A2\"><f>B2*2</f><v>4</v></c>"));
    }

    #[test]
    fn editing_inside_an_array_formula_is_rejected() {
        let xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">B1:B2*2</f><v>2</v></c></row><row r=\"2\"><c r=\"A2\"><v>4</v></c></row></sheetData></worksheet>";
        let mut edits = SheetEdits::default();
        edits.set_cell(1, 0, CellEdit::Set { value: EditValue::Number(7.0), style: StyleOp::Keep });
        let (mut styles, mut sst, mut rels) = ctx_parts();
        let mut loader = |_, _| Ok(None);
        let mut ctx = RewriteContext { styles: &mut styles, sst: Some(&mut sst), formula_at: &mut loader, rels: &mut rels, rels_dirty: false, table_headers: HashMap::new(), table_header_updates: Vec::new() };
        assert!(rewrite(xml, &edits, &mut ctx).unwrap_err().contains("array formula"));
    }

    #[test]
    fn cell_edits_stay_small() {
        // Pending edits can number in the millions: formats are shared and
        // formulas boxed so each edit stays this small (was 248 bytes).
        assert!(std::mem::size_of::<CellEdit>() <= 56);
    }

    #[test]
    fn style_ops_compose() {
        let bold = FormatSpec { bold: true, ..Default::default() };
        let italic = FormatSpec { italic: true, ..Default::default() };
        let combined = StyleOp::Merge(Arc::new(bold.clone())).then(StyleOp::Merge(Arc::new(italic)));
        assert_eq!(combined, StyleOp::Merge(Arc::new(FormatSpec { bold: true, italic: true, ..Default::default() })));
        let bold = Arc::new(bold);
        assert_eq!(StyleOp::Replace(bold.clone()).then(StyleOp::Keep), StyleOp::Replace(bold));
    }
}
