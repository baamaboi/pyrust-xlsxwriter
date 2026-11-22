//! Workbook-level editor state and the save pipeline.
//!
//! `BookState` never modifies the original package, styles or calamine
//! reader. Edits are recorded (cells per sheet, structure in the workbook
//! DOM); every save starts from clones of the originals and re-applies all
//! edits, so saving repeatedly (or to bytes) gives consistent results.

use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::sync::Arc;

use calamine::{Data, Range, Reader, Xlsx};

use crate::common::bytes::SharedBytes;
use crate::common::cellref::{parse_range, CellRange};
use crate::common::format::FormatSpec;
use crate::editor::package::{content_type, rel_type, relative_target, rels_part_name, EResult, Package, Relationships};
use crate::editor::formula::{self as formula, StructuralEdit};
use crate::editor::sheet::{self, CellEdit, EditValue, RewriteContext, SheetEdits};
use crate::editor::structure::{self, map_cell, unmap_cell, Op};
use crate::editor::sst::SharedStrings;
use crate::editor::styles::Styles;
use crate::editor::xml::{attr, local, parse_document, Child, Document, Node, Scanner, Tok};

const WORKBOOK_ORDER: &[&str] = &[
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
    "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize",
    "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
    "fileRecoveryPr", "webPublishObjects", "extLst",
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Visibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl Visibility {
    pub fn as_str(self) -> &'static str {
        match self {
            Visibility::Visible => "visible",
            Visibility::Hidden => "hidden",
            Visibility::VeryHidden => "very_hidden",
        }
    }
}

#[derive(Default)]
pub struct SheetCache {
    pub values: Option<Range<Data>>,
    pub formulas: Option<HashMap<(u32, u16), String>>,
    pub xfs: Option<HashMap<(u32, u16), u32>>,
    pub merges: Option<Vec<CellRange>>,
}

pub struct SheetEntry {
    /// Stable handle used by Python worksheet objects.
    pub id: u64,
    pub name: String,
    /// Name in the original file (for reading original values); None for new sheets.
    pub original_name: Option<String>,
    pub part: String,
    pub rel_id: String,
    pub is_worksheet: bool,
    pub is_new: bool,
    pub edits: SheetEdits,
    pub cache: SheetCache,
}

pub struct BookState {
    package: Package,
    workbook_part: String,
    workbook: Document,
    workbook_dirty: bool,
    wb_rels: Relationships,
    wb_rels_dirty: bool,
    pub sheets: Vec<SheetEntry>,
    deleted_parts: Vec<String>,
    next_id: u64,
    pub styles: Styles,
    styles_is_new: bool,
    sst_part: Option<String>,
    calamine: Option<Xlsx<Cursor<SharedBytes>>>,
    core_updates: Vec<(String, String)>,
    pub is_1904: bool,
    /// Structural operations and renames, in order (replayed on save).
    ops: Vec<Op>,
    /// (id, part, original name) of every worksheet in the original file.
    original_sheets: Vec<(u64, String, String)>,
}

/// Excel's worksheet name rules.
pub fn check_sheet_name(name: &str) -> EResult<()> {
    if name.is_empty() {
        return Err("worksheet name cannot be blank".into());
    }
    if name.chars().count() > 31 {
        return Err(format!("worksheet name {name:?} is longer than 31 characters"));
    }
    if name.contains(['[', ']', ':', '*', '?', '/', '\\']) {
        return Err(format!("worksheet name {name:?} contains one of the invalid characters []:*?/\\"));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(format!("worksheet name {name:?} cannot start or end with an apostrophe"));
    }
    if name.eq_ignore_ascii_case("History") {
        return Err("\"History\" is a reserved worksheet name".into());
    }
    Ok(())
}

/// Replace sheet references in a formula (used for defined names on rename).
pub fn rename_in_formula(formula: &str, old: &str, new: &str) -> String {
    let quote = |n: &str| rust_xlsxwriter::utility::quote_sheet_name(n);
    let old_quoted = format!("'{}'!", old.replace('\'', "''"));
    let new_ref = format!("{}!", quote(new));
    let mut out = formula.replace(&old_quoted, &new_ref);
    let bare = format!("{old}!");
    if !old.contains([' ', '\'', '-', '(', ')']) {
        let mut result = String::with_capacity(out.len());
        let mut rest = out.as_str();
        while let Some(i) = rest.find(&bare) {
            let before_ok = rest[..i].chars().last().is_none_or(|c| !(c.is_alphanumeric() || c == '_' || c == '.' || c == '\''));
            result.push_str(&rest[..i]);
            result.push_str(if before_ok { &new_ref } else { &bare });
            rest = &rest[i + bare.len()..];
        }
        result.push_str(rest);
        out = result;
    }
    out
}

impl BookState {
    pub fn open(data: SharedBytes) -> EResult<Self> {
        let mut package = Package::open(data)?;
        let root_rels = package.rels("")?;
        let workbook_part = root_rels
            .find_type("officeDocument")
            .first()
            .map(|r| crate::editor::package::resolve_target("", &r.target))
            .ok_or("not an Excel workbook: the package has no main document")?;
        let workbook = package.read_xml(&workbook_part)?;
        if workbook.root.local() != "workbook" {
            return Err(format!("{workbook_part} is not a SpreadsheetML workbook"));
        }
        let wb_rels = package.rels(&workbook_part)?;

        let mut sheets = Vec::new();
        let mut next_id = 1;
        if let Some(list) = workbook.root.child("sheets") {
            for node in list.children_named("sheet") {
                let name = node.get("name").unwrap_or_default();
                let rel_id = node.get("r:id").or_else(|| node.get("id")).unwrap_or_default();
                let rel = wb_rels.get(&rel_id).ok_or_else(|| format!("sheet {name:?} has no relationship {rel_id}"))?;
                let part = crate::editor::package::resolve_target(&workbook_part, &rel.target);
                sheets.push(SheetEntry {
                    id: next_id,
                    original_name: Some(name.clone()),
                    name,
                    part,
                    rel_id,
                    is_worksheet: rel_type::is(&rel.rel_type, "worksheet"),
                    is_new: false,
                    edits: SheetEdits::default(),
                    cache: SheetCache::default(),
                });
                next_id += 1;
            }
        }

        let styles_rel = wb_rels.find_type("styles").into_iter().next();
        let (styles, styles_is_new) = match styles_rel {
            Some(rel) => {
                let part = crate::editor::package::resolve_target(&workbook_part, &rel.target);
                let doc = package.read_xml(&part)?;
                (Styles::new(&part, doc), false)
            }
            None => {
                let doc = parse_document(Styles::minimal_xml()).map_err(|e| e.to_string())?;
                (Styles::new("xl/styles.xml", doc), true)
            }
        };
        let sst_part = wb_rels
            .find_type("sharedStrings")
            .into_iter()
            .next()
            .map(|r| crate::editor::package::resolve_target(&workbook_part, &r.target));
        let is_1904 = workbook
            .root
            .child("workbookPr")
            .and_then(|p| p.get("date1904"))
            .is_some_and(|v| v == "1" || v == "true");

        Ok(BookState {
            package,
            workbook_part,
            workbook,
            workbook_dirty: false,
            wb_rels,
            wb_rels_dirty: false,
            sheets,
            deleted_parts: Vec::new(),
            next_id,
            styles,
            styles_is_new,
            sst_part,
            calamine: None,
            core_updates: Vec::new(),
            is_1904,
            ops: Vec::new(),
            original_sheets: Vec::new(),
        })
        .map(|mut state| {
            state.original_sheets = state
                .sheets
                .iter()
                .filter(|s| s.is_worksheet)
                .map(|s| (s.id, s.part.clone(), s.name.clone()))
                .collect();
            state
        })
    }

    pub fn index_of(&self, id: u64) -> EResult<usize> {
        self.sheets
            .iter()
            .position(|s| s.id == id)
            .ok_or_else(|| "this worksheet was removed from the workbook".to_string())
    }

    pub fn index_by_name(&self, name: &str) -> Option<usize> {
        self.sheets.iter().position(|s| s.name == name)
    }

    fn sheet_nodes(&mut self) -> Vec<&mut Node> {
        match self.workbook.root.child_mut("sheets") {
            Some(list) => list.elements_mut().filter(|n| n.local() == "sheet").collect(),
            None => Vec::new(),
        }
    }

    pub fn visibility(&self, index: usize) -> Visibility {
        let node = self
            .workbook
            .root
            .child("sheets")
            .and_then(|l| l.children_named("sheet").nth(index).cloned());
        match node.and_then(|n| n.get("state")).as_deref() {
            Some("hidden") => Visibility::Hidden,
            Some("veryHidden") => Visibility::VeryHidden,
            _ => Visibility::Visible,
        }
    }

    pub fn set_visibility(&mut self, index: usize, state: Visibility) -> EResult<()> {
        if state != Visibility::Visible {
            let others_visible = (0..self.sheets.len())
                .any(|i| i != index && self.visibility(i) == Visibility::Visible);
            if !others_visible {
                return Err("a workbook must have at least one visible worksheet".into());
            }
            if self.active_index() == index {
                let next = (0..self.sheets.len())
                    .find(|i| *i != index && self.visibility(*i) == Visibility::Visible)
                    .unwrap_or(0);
                self.set_active_index(next)?;
            }
        }
        let mut nodes = self.sheet_nodes();
        let node = nodes.get_mut(index).ok_or("sheet index out of range")?;
        match state {
            Visibility::Visible => node.remove_attr("state"),
            Visibility::Hidden => node.set("state", "hidden"),
            Visibility::VeryHidden => node.set("state", "veryHidden"),
        }
        self.workbook_dirty = true;
        Ok(())
    }

    pub fn active_index(&self) -> usize {
        self.workbook
            .root
            .child("bookViews")
            .and_then(|b| b.child("workbookView"))
            .and_then(|v| v.get("activeTab"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0)
    }

    pub fn set_active_index(&mut self, index: usize) -> EResult<()> {
        if index >= self.sheets.len() {
            return Err("sheet index out of range".into());
        }
        if self.workbook.root.child("bookViews").is_none() {
            let mut views = Node::new(&format!("{}bookViews", self.workbook.root.prefix()));
            views.push(Node::new(&format!("{}workbookView", self.workbook.root.prefix())));
            self.workbook.root.insert_ordered(views, WORKBOOK_ORDER);
        }
        let views = self.workbook.root.child_mut("bookViews").ok_or("missing bookViews")?;
        if views.child("workbookView").is_none() {
            views.push(Node::new("workbookView"));
        }
        let view = views.child_mut("workbookView").ok_or("missing workbookView")?;
        if index == 0 {
            view.remove_attr("activeTab");
        } else {
            view.set("activeTab", &index.to_string());
        }
        if view.get("firstSheet").and_then(|v| v.parse::<usize>().ok()).is_some_and(|f| f > index) {
            view.set("firstSheet", &index.to_string());
        }
        self.workbook_dirty = true;
        Ok(())
    }

    fn check_unique(&self, name: &str, except: Option<usize>) -> EResult<()> {
        if self
            .sheets
            .iter()
            .enumerate()
            .any(|(i, s)| Some(i) != except && s.name.eq_ignore_ascii_case(name))
        {
            return Err(format!("worksheet name {name:?} is already used in this workbook"));
        }
        Ok(())
    }

    pub fn rename(&mut self, index: usize, new: &str) -> EResult<()> {
        check_sheet_name(new)?;
        self.check_unique(new, Some(index))?;
        let old = self.sheets[index].name.clone();
        if old == new {
            return Ok(());
        }
        {
            let mut nodes = self.sheet_nodes();
            let node = nodes.get_mut(index).ok_or("sheet index out of range")?;
            node.set("name", new);
        }
        let rename = |text: &str| formula::rename_sheet(text, &old, new).unwrap_or_else(|_| rename_in_formula(text, &old, new));
        if let Some(names) = self.workbook.root.child_mut("definedNames") {
            for dn in names.elements_mut().filter(|n| n.local() == "definedName") {
                let text = dn.text();
                let renamed = rename(&text);
                if renamed != text {
                    dn.set_text(&renamed);
                }
            }
        }
        for entry in &mut self.sheets {
            transform_pending_formulas(&mut entry.edits, &rename);
            for (_, url) in &mut entry.edits.hyperlinks {
                if let Some(location) = url.link.strip_prefix("internal:") {
                    url.link = format!("internal:{}", rename(location));
                }
            }
        }
        self.sheets[index].name = new.to_string();
        self.ops.push(Op::Rename { old, new: new.to_string() });
        self.workbook_dirty = true;
        Ok(())
    }

    /// Insert/delete rows or columns on a sheet.
    pub fn structural_edit(&mut self, index: usize, edit: StructuralEdit) -> EResult<()> {
        let sheet_id = self.sheets[index].id;
        let local_names: Vec<String> = self.sheets.iter().map(|s| s.name.clone()).collect();
        if let Some(names) = self.workbook.root.child_mut("definedNames") {
            for dn in names.elements_mut().filter(|n| n.local() == "definedName") {
                let scope = dn.get("localSheetId").and_then(|v| v.parse::<usize>().ok()).and_then(|i| local_names.get(i).cloned()).unwrap_or_default();
                let text = dn.text();
                if let Ok(shifted) = formula::shift_formula(&text, &scope, &edit)
                    && shifted != text {
                        dn.set_text(&shifted);
                    }
            }
        }
        for entry in &mut self.sheets {
            let name = entry.name.clone();
            transform_pending_formulas(&mut entry.edits, &|text: &str| {
                formula::shift_formula(text, &name, &edit).unwrap_or_else(|_| text.to_string())
            });
            if entry.id == sheet_id {
                remap_edits(&mut entry.edits, &edit);
            }
        }
        self.ops.push(Op::Structure { sheet_id, edit });
        self.workbook_dirty = true;
        Ok(())
    }

    /// Record objects rendered by the writer into `scratch` (see `graft`):
    /// cells the writer filled in become edits, the rest is grafted on save.
    pub fn record_graft(&mut self, index: usize, scratch: Vec<u8>) -> EResult<()> {
        let scratch = SharedBytes::from_vec(scratch);
        let mut reader: Xlsx<_> = calamine::open_workbook_from_rs(Cursor::new(scratch.clone()))
            .map_err(|e| format!("reading rendered objects: {e}"))?;
        let name = reader.sheet_names().first().cloned().ok_or("rendered workbook has no sheets")?;
        let values = reader.worksheet_range(&name).map_err(|e| e.to_string())?;
        let formulas = load_formulas(&mut reader, &name)?;
        let (xfs, scratch_styles) = crate::editor::graft::scratch_cell_styles(&scratch)?;
        let mut cells: Vec<((u32, u16), EditValue, Option<FormatSpec>)> = Vec::new();
        if let Some((r0, c0)) = values.start() {
            for (r, c, v) in values.used_cells() {
                let pos = (r0 + r as u32, (c0 + c as u32) as u16);
                let value = match formulas.get(&pos) {
                    Some(f) => EditValue::formula(f.clone(), data_text(v), None),
                    None => match v {
                        Data::String(s) => EditValue::Str(s.clone()),
                        Data::Float(f) => EditValue::Number(*f),
                        Data::Int(i) => EditValue::Number(*i as f64),
                        Data::Bool(b) => EditValue::Bool(*b),
                        Data::Empty => EditValue::Blank,
                        other => EditValue::Str(other.to_string()),
                    },
                };
                cells.push((pos, value, xfs.get(&pos).map(|xf| scratch_styles.spec(*xf))));
            }
        }
        for pos in formulas.keys() {
            if !cells.iter().any(|(p, _, _)| p == pos) {
                let text = formulas[pos].clone();
                cells.push((*pos, EditValue::formula(text, None, None), xfs.get(pos).map(|xf| scratch_styles.spec(*xf))));
            }
        }
        let entry = &mut self.sheets[index];
        for ((r, c), value, spec) in cells {
            let style = spec.map(|spec| sheet::StyleOp::Replace(Arc::new(spec))).unwrap_or(sheet::StyleOp::Keep);
            entry.edits.set_cell(r, c, CellEdit::Set { value, style });
        }
        self.ops.push(Op::Graft { sheet_id: entry.id, scratch });
        Ok(())
    }

    /// Where a current cell came from in the original file (None for new or inserted cells).
    pub fn to_original(&self, index: usize, row: u32, col: u16) -> Option<(u32, u16)> {
        let id = self.sheets[index].id;
        let mut pos = (row, col);
        for op in self.ops.iter().rev() {
            if let Op::Structure { sheet_id, edit } = op
                && *sheet_id == id {
                    pos = unmap_cell(edit, pos.0, pos.1)?;
                }
        }
        Some(pos)
    }

    /// Where an original cell is now (None if it was deleted).
    pub fn from_original(&self, index: usize, row: u32, col: u16) -> Option<(u32, u16)> {
        let id = self.sheets[index].id;
        let mut pos = (row, col);
        for op in &self.ops {
            if let Op::Structure { sheet_id, edit } = op
                && *sheet_id == id {
                    pos = map_cell(edit, pos.0, pos.1)?;
                }
        }
        Some(pos)
    }

    /// Bring an original formula of a sheet up to date with the recorded operations.
    pub fn replay_formula(&self, index: usize, text: &str) -> String {
        let mut name = self.sheets[index].original_name.clone().unwrap_or_default();
        let mut text = text.to_string();
        for op in &self.ops {
            match op {
                Op::Rename { old, new } => {
                    text = formula::rename_sheet(&text, old, new).unwrap_or(text);
                    if name.eq_ignore_ascii_case(old) {
                        name = new.clone();
                    }
                }
                Op::Structure { edit, .. } => {
                    text = formula::shift_formula(&text, &name, edit).unwrap_or(text);
                }
                Op::Graft { .. } => {}
            }
        }
        text
    }

    /// Move an original range through the recorded operations (None if deleted).
    pub fn current_range(&self, index: usize, range: CellRange) -> Option<CellRange> {
        let id = self.sheets[index].id;
        let mut text = crate::common::cellref::range_name(range, false);
        for op in &self.ops {
            if let Op::Structure { sheet_id, edit } = op
                && *sheet_id == id {
                    match formula::shift_range_list(&text, true, edit) {
                        Some(t) if !t.is_empty() => text = t,
                        _ => return None,
                    }
                }
        }
        parse_range(&text)
    }

    /// Original merged ranges moved through the recorded operations.
    pub fn current_merges(&mut self, index: usize) -> EResult<Vec<CellRange>> {
        let merges = self.original_merges(index)?;
        Ok(merges.into_iter().filter_map(|m| self.current_range(index, m)).collect())
    }

    pub fn add_sheet(&mut self, name: Option<&str>, index: Option<usize>) -> EResult<u64> {
        let name = match name {
            Some(n) => n.to_string(),
            None => (1..)
                .map(|n| format!("Sheet{n}"))
                .find(|n| self.check_unique(n, None).is_ok())
                .unwrap_or_default(),
        };
        check_sheet_name(&name)?;
        self.check_unique(&name, None)?;
        let reserved: HashSet<&str> = self.sheets.iter().map(|s| s.part.as_str()).collect();
        let part = (1..)
            .map(|n| format!("xl/worksheets/sheet{n}.xml"))
            .find(|p| !self.package.exists(p) && !reserved.contains(p.as_str()))
            .unwrap_or_default();
        let target = relative_target(&self.workbook_part, &part);
        let rel_id = self.wb_rels.add("worksheet", &target, false);
        self.wb_rels_dirty = true;
        let sheet_id = self
            .workbook
            .root
            .child("sheets")
            .map(|l| l.children_named("sheet").filter_map(|n| n.get("sheetId")?.parse::<u32>().ok()).max().unwrap_or(0))
            .unwrap_or(0)
            + 1;
        let prefix = self.workbook.root.prefix().to_string();
        let rel_prefix = self.relationship_prefix();
        let node = Node::new(&format!("{prefix}sheet"))
            .with_attr("name", &name)
            .with_attr("sheetId", &sheet_id.to_string())
            .with_attr(&format!("{rel_prefix}id"), &rel_id);
        if self.workbook.root.child("sheets").is_none() {
            let list = Node::new(&format!("{prefix}sheets"));
            self.workbook.root.insert_ordered(list, WORKBOOK_ORDER);
        }
        let position = index.unwrap_or(self.sheets.len()).min(self.sheets.len());
        {
            let list = self.workbook.root.child_mut("sheets").ok_or("missing sheets")?;
            let element_positions: Vec<usize> = list
                .children
                .iter()
                .enumerate()
                .filter(|(_, c)| matches!(c, Child::Element(n) if n.local() == "sheet"))
                .map(|(i, _)| i)
                .collect();
            let insert_at = element_positions.get(position).copied().unwrap_or(list.children.len());
            list.children.insert(insert_at, Child::Element(node));
            list.self_closing = false;
        }
        let id = self.next_id;
        self.next_id += 1;
        self.shift_local_ids(|i| if i >= position { i + 1 } else { i });
        self.original_sheets.push((id, part.clone(), name.clone()));
        self.sheets.insert(
            position,
            SheetEntry {
                id,
                name,
                original_name: None,
                part,
                rel_id,
                is_worksheet: true,
                is_new: true,
                edits: SheetEdits::default(),
                cache: SheetCache::default(),
            },
        );
        self.workbook_dirty = true;
        Ok(id)
    }

    fn relationship_prefix(&self) -> String {
        for (k, v) in &self.workbook.root.attrs {
            if v == crate::editor::package::NS_REL
                && let Some(p) = k.strip_prefix("xmlns:") {
                    return format!("{p}:");
                }
        }
        "r:".to_string()
    }

    /// Remap `localSheetId` of sheet-scoped defined names (None removes the name).
    fn remap_local_ids(&mut self, map: impl Fn(usize) -> Option<usize>) {
        if let Some(names) = self.workbook.root.child_mut("definedNames") {
            names.children.retain_mut(|c| {
                if let Child::Element(n) = c
                    && let Some(local) = n.get("localSheetId").and_then(|v| v.parse::<usize>().ok()) {
                        match map(local) {
                            Some(new) => n.set("localSheetId", &new.to_string()),
                            None => return false,
                        }
                    }
                true
            });
            if names.elements().next().is_none() {
                self.workbook.root.remove_children("definedNames");
            }
        }
    }

    fn shift_local_ids(&mut self, map: impl Fn(usize) -> usize) {
        self.remap_local_ids(|i| Some(map(i)));
    }

    pub fn remove_sheet(&mut self, index: usize) -> EResult<()> {
        if self.sheets.len() == 1 {
            return Err("cannot remove the only worksheet in a workbook".into());
        }
        let visible_left = (0..self.sheets.len()).any(|i| i != index && self.visibility(i) == Visibility::Visible);
        if !visible_left {
            return Err("cannot remove the last visible worksheet".into());
        }
        let active = self.active_index();
        let entry = self.sheets.remove(index);
        if let Some(list) = self.workbook.root.child_mut("sheets") {
            let mut seen = 0usize;
            list.children.retain(|c| {
                if let Child::Element(n) = c
                    && n.local() == "sheet" {
                        seen += 1;
                        return seen - 1 != index;
                    }
                true
            });
        }
        self.wb_rels.remove(&entry.rel_id);
        self.wb_rels_dirty = true;
        if !entry.is_new {
            self.deleted_parts.push(entry.part);
        }
        self.remap_local_ids(|i| match i.cmp(&index) {
            std::cmp::Ordering::Less => Some(i),
            std::cmp::Ordering::Equal => None,
            std::cmp::Ordering::Greater => Some(i - 1),
        });
        let new_active = if active > index || active >= self.sheets.len() { active.saturating_sub(1) } else { active };
        let new_active = (new_active..self.sheets.len())
            .chain(0..new_active)
            .find(|i| self.visibility(*i) == Visibility::Visible)
            .unwrap_or(0);
        self.set_active_index(new_active)?;
        if let Some(view) = self.workbook.root.child_mut("bookViews").and_then(|b| b.child_mut("workbookView")) {
            view.remove_attr("firstSheet");
        }
        self.workbook_dirty = true;
        Ok(())
    }

    pub fn move_sheet(&mut self, from: usize, to: usize) -> EResult<()> {
        let len = self.sheets.len();
        if from >= len || to >= len {
            return Err("sheet index out of range".into());
        }
        if from == to {
            return Ok(());
        }
        let active = self.active_index();
        let entry = self.sheets.remove(from);
        self.sheets.insert(to, entry);
        if let Some(list) = self.workbook.root.child_mut("sheets") {
            let positions: Vec<usize> = list
                .children
                .iter()
                .enumerate()
                .filter(|(_, c)| matches!(c, Child::Element(n) if n.local() == "sheet"))
                .map(|(i, _)| i)
                .collect();
            let mut nodes: Vec<Child> = positions.iter().map(|i| list.children[*i].clone()).collect();
            let node = nodes.remove(from);
            nodes.insert(to, node);
            for (slot, node) in positions.iter().zip(nodes) {
                list.children[*slot] = node;
            }
        }
        let map = move |i: usize| -> usize {
            if i == from {
                to
            } else if from < to && i > from && i <= to {
                i - 1
            } else if to < from && i >= to && i < from {
                i + 1
            } else {
                i
            }
        };
        self.shift_local_ids(map);
        self.set_active_index(map(active))?;
        self.workbook_dirty = true;
        Ok(())
    }

    pub fn defined_names(&self) -> Vec<(String, String, Option<usize>)> {
        self.workbook
            .root
            .child("definedNames")
            .map(|names| {
                names
                    .children_named("definedName")
                    .map(|n| {
                        (
                            n.get("name").unwrap_or_default(),
                            n.text(),
                            n.get("localSheetId").and_then(|v| v.parse().ok()),
                        )
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    pub fn define_name(&mut self, name: &str, formula: &str, local: Option<usize>) -> EResult<()> {
        if name.is_empty() || name.contains(char::is_whitespace) {
            return Err(format!("invalid defined name {name:?}"));
        }
        let formula = formula.strip_prefix('=').unwrap_or(formula);
        let prefix = self.workbook.root.prefix().to_string();
        if self.workbook.root.child("definedNames").is_none() {
            self.workbook.root.insert_ordered(Node::new(&format!("{prefix}definedNames")), WORKBOOK_ORDER);
        }
        let names = self.workbook.root.child_mut("definedNames").ok_or("missing definedNames")?;
        let local_text = local.map(|l| l.to_string());
        names.children.retain(|c| {
            !matches!(c, Child::Element(n) if n.local() == "definedName"
                && n.get("name").is_some_and(|v| v.eq_ignore_ascii_case(name))
                && n.get("localSheetId") == local_text)
        });
        let mut node = Node::new(&format!("{prefix}definedName")).with_attr("name", name);
        if let Some(local) = local {
            node.set("localSheetId", &local.to_string());
        }
        node.set_text(formula);
        names.push(node);
        self.workbook_dirty = true;
        Ok(())
    }

    pub fn remove_name(&mut self, name: &str, local: Option<usize>) -> EResult<bool> {
        let local_text = local.map(|l| l.to_string());
        let Some(names) = self.workbook.root.child_mut("definedNames") else { return Ok(false) };
        let before = names.children.len();
        names.children.retain(|c| {
            !matches!(c, Child::Element(n) if n.local() == "definedName"
                && n.get("name").is_some_and(|v| v.eq_ignore_ascii_case(name))
                && n.get("localSheetId") == local_text)
        });
        let removed = names.children.len() != before;
        if names.elements().next().is_none() {
            self.workbook.root.remove_children("definedNames");
        }
        self.workbook_dirty |= removed;
        Ok(removed)
    }

    pub fn set_property(&mut self, key: &str, value: &str) {
        self.core_updates.retain(|(k, _)| k != key);
        self.core_updates.push((key.to_string(), value.to_string()));
    }

    pub fn properties(&mut self) -> EResult<Vec<(String, String)>> {
        let mut out = Vec::new();
        if let Some(part) = self.core_part()? {
            let doc = self.package.read_xml(&part)?;
            for node in doc.root.elements() {
                let text = node.text();
                if !text.is_empty() {
                    out.push((core_key(node.local()).to_string(), text));
                }
            }
        }
        for (k, v) in &self.core_updates {
            out.retain(|(key, _)| key != k);
            out.push((k.clone(), v.clone()));
        }
        Ok(out)
    }

    fn core_part(&mut self) -> EResult<Option<String>> {
        Ok(self
            .package
            .rels("")?
            .find_type("core-properties")
            .first()
            .map(|r| crate::editor::package::resolve_target("", &r.target))
            .or_else(|| {
                self.package
                    .rels("")
                    .ok()?
                    .list()
                    .into_iter()
                    .find(|r| r.rel_type.ends_with("/core-properties"))
                    .map(|r| crate::editor::package::resolve_target("", &r.target))
            }))
    }

    fn reader(&mut self) -> EResult<&mut Xlsx<Cursor<SharedBytes>>> {
        if self.calamine.is_none() {
            let reader: Xlsx<_> = calamine::open_workbook_from_rs(Cursor::new(self.package.original_bytes()))
                .map_err(|e| format!("reading workbook: {e}"))?;
            self.calamine = Some(reader);
        }
        Ok(self.calamine.as_mut().expect("reader just created"))
    }

    /// The original values of a sheet (empty for new sheets).
    pub fn values(&mut self, index: usize) -> EResult<&Range<Data>> {
        if self.sheets[index].cache.values.is_none() {
            let range = match self.sheets[index].original_name.clone() {
                Some(name) if self.sheets[index].is_worksheet => {
                    self.reader()?.worksheet_range(&name).map_err(|e| format!("reading {name:?}: {e}"))?
                }
                _ => Range::empty(),
            };
            self.sheets[index].cache.values = Some(range);
        }
        Ok(self.sheets[index].cache.values.as_ref().expect("values just loaded"))
    }

    /// The original value behind current cell (row, col): None if the cell
    /// is new or empty, or if the values haven't been loaded with
    /// [`BookState::values`]. Borrowed, so reading never copies the value.
    pub fn original_value(&self, index: usize, row: u32, col: u16) -> Option<&Data> {
        let (row, col) = self.to_original(index, row, col)?;
        self.sheets[index].cache.values.as_ref()?.get_value((row, col as u32))
    }

    /// Original formulas of a sheet, by absolute (row, col), without '='.
    pub fn formulas(&mut self, index: usize) -> EResult<&HashMap<(u32, u16), String>> {
        if self.sheets[index].cache.formulas.is_none() {
            let map = match self.sheets[index].original_name.clone() {
                Some(name) if self.sheets[index].is_worksheet => load_formulas(self.reader()?, &name)?,
                _ => HashMap::new(),
            };
            self.sheets[index].cache.formulas = Some(map);
        }
        Ok(self.sheets[index].cache.formulas.as_ref().expect("formulas just loaded"))
    }

    /// Original style index of each cell (from the sheet XML).
    pub fn cell_styles(&mut self, index: usize) -> EResult<&HashMap<(u32, u16), u32>> {
        if self.sheets[index].cache.xfs.is_none() {
            let entry = &self.sheets[index];
            let (map, merges) = if entry.is_new || !entry.is_worksheet {
                (HashMap::new(), Vec::new())
            } else {
                let part = entry.part.clone();
                let xml = self.package.read_string(&part)?;
                scan_styles_and_merges(&xml)?
            };
            self.sheets[index].cache.xfs = Some(map);
            self.sheets[index].cache.merges = Some(merges);
        }
        Ok(self.sheets[index].cache.xfs.as_ref().expect("styles just loaded"))
    }

    pub fn original_merges(&mut self, index: usize) -> EResult<Vec<CellRange>> {
        if let Some(merges) = &self.sheets[index].cache.merges {
            return Ok(merges.clone());
        }
        let entry = &self.sheets[index];
        let merges = if entry.is_new || !entry.is_worksheet {
            Vec::new()
        } else {
            // <mergeCells> follows <sheetData>, so only the tail needs scanning.
            let part = entry.part.clone();
            let xml = self.package.read_string(&part)?;
            let tail_start = xml.rfind("sheetData>").unwrap_or(0);
            let tail = &xml[tail_start..];
            let mut merges = Vec::new();
            if tail.contains("mergeCell") {
                let mut scanner = Scanner::new(tail);
                while let Some((tok, _)) = scanner.next_tok().map_err(|e| e.to_string())? {
                    if let Tok::Empty { name, attrs } | Tok::Start { name, attrs } = tok {
                        if local(name) == "mergeCell" {
                            if let Some(range) = attr(attrs, "ref").and_then(|r| parse_range(&r)) {
                                merges.push(range);
                            }
                        }
                    }
                }
            }
            merges
        };
        self.sheets[index].cache.merges = Some(merges.clone());
        Ok(merges)
    }

    /// Table header cells of a sheet: (row, col) -> (table part, column index).
    fn table_headers(package: &mut Package, sheet_part: &str, edits: &SheetEdits) -> EResult<HashMap<(u32, u16), (String, usize)>> {
        let mut out = HashMap::new();
        if edits.cells.is_empty() {
            return Ok(out);
        }
        let rels = package.rels(sheet_part)?;
        for rel in rels.find_type("table") {
            let part = crate::editor::package::resolve_target(sheet_part, &rel.target);
            let doc = package.read_xml(&part)?;
            if doc.root.get("headerRowCount").as_deref() == Some("0") {
                continue;
            }
            let Some((r1, c1, _, c2)) = doc.root.get("ref").and_then(|r| parse_range(&r)) else { continue };
            for col in c1..=c2 {
                out.insert((r1, col), (part.clone(), (col - c1) as usize));
            }
        }
        Ok(out)
    }

    /// Build the edited package; write it with [`Package::save`] (to a file)
    /// or [`Package::save_to_vec`]. The original package is cloned cheaply
    /// (untouched parts stay compressed in the shared original bytes), so
    /// building can be repeated.
    pub fn build(&mut self) -> EResult<Package> {
        let mut package = self.package.clone();
        let mut styles = self.styles.clone();
        let mut workbook = self.workbook.clone();
        let mut wb_rels = self.wb_rels.clone();
        let mut workbook_dirty = self.workbook_dirty;
        let mut wb_rels_dirty = self.wb_rels_dirty;
        let mut any_cells = !self.ops.is_empty();

        // New sheets get their (blank) part first, so grafts and edits can target them.
        for entry in &self.sheets {
            if entry.is_new {
                package.write(&entry.part, sheet::blank_sheet_xml().into_bytes());
                package.set_override(&entry.part, content_type::WORKSHEET);
            }
        }

        if !self.ops.is_empty() {
            let mut names: HashMap<u64, String> =
                self.original_sheets.iter().map(|(id, _, name)| (*id, name.clone())).collect();
            let all_parts: Vec<String> = package.part_names();
            let replay = structure::Replay {
                sheets: &self.original_sheets,
                charts: all_parts.iter().filter(|n| n.starts_with("xl/charts/chart") && n.ends_with(".xml")).cloned().collect(),
                pivot_caches: all_parts.iter().filter(|n| n.starts_with("xl/pivotCache/pivotCacheDefinition")).cloned().collect(),
            };
            for op in &self.ops {
                structure::replay(&mut package, &mut styles, &replay, &mut names, op)?;
            }
        }

        // The sheets to rewrite, with their table header cells. The shared
        // string table can be large, so it is only read if a rewrite may add
        // a string to it.
        let mut rewrites = Vec::new();
        for (index, entry) in self.sheets.iter().enumerate() {
            if entry.is_worksheet && (!entry.edits.is_empty() || entry.is_new) {
                rewrites.push((index, Self::table_headers(&mut package, &entry.part, &entry.edits)?));
            }
        }
        let needs_strings = rewrites.iter().any(|(index, headers)| self.sheets[*index].edits.adds_strings(headers));
        let mut sst = match &self.sst_part {
            _ if !needs_strings => None,
            Some(part) => Some(SharedStrings::parse(part, package.read_string(part)?.into_owned())?),
            None => Some(SharedStrings::empty("xl/sharedStrings.xml")),
        };

        for (index, table_headers) in rewrites {
            let entry = &self.sheets[index];
            any_cells |= entry.edits.has_cell_edits();
            let part = entry.part.clone();
            let mut rels = package.rels(&part)?;
            // Formulas of cells whose shared formula breaks; read from the
            // cache (borrowed) or, the first time one is needed, from calamine.
            let cached = entry.cache.formulas.as_ref();
            let original_name = entry.original_name.as_deref();
            let calamine = &mut self.calamine;
            let original_bytes = package.original_bytes();
            let mut loaded: Option<HashMap<(u32, u16), String>> = None;
            let mut formula_at = |row: u32, col: u16| -> EResult<Option<String>> {
                if cached.is_none() && loaded.is_none() {
                    loaded = Some(match original_name {
                        Some(name) => {
                            if calamine.is_none() {
                                let reader: Xlsx<_> = calamine::open_workbook_from_rs(Cursor::new(original_bytes.clone()))
                                    .map_err(|e| format!("reading workbook: {e}"))?;
                                *calamine = Some(reader);
                            }
                            load_formulas(calamine.as_mut().expect("reader just created"), name)?
                        }
                        None => HashMap::new(),
                    });
                }
                let map = cached.or(loaded.as_ref()).expect("formulas loaded above");
                Ok(map.get(&(row, col)).cloned())
            };
            let xml = package.read_string(&part)?;
            let mut ctx = RewriteContext {
                styles: &mut styles,
                sst: sst.as_mut(),
                formula_at: &mut formula_at,
                rels: &mut rels,
                rels_dirty: false,
                table_headers,
                table_header_updates: Vec::new(),
            };
            let out = sheet::rewrite(&xml, &entry.edits, &mut ctx)
                .map_err(|e| format!("worksheet {:?}: {e}", entry.name))?;
            let rels_dirty = ctx.rels_dirty;
            let updates = std::mem::take(&mut ctx.table_header_updates);
            drop(ctx);
            drop(xml);
            package.write(&part, out.into_bytes());
            if rels_dirty {
                package.set_rels(&rels);
            }
            let mut by_part: HashMap<String, Vec<(usize, String)>> = HashMap::new();
            for (table, column, name) in updates {
                by_part.entry(table).or_default().push((column, name));
            }
            for (table, changes) in by_part {
                let mut doc = package.read_xml(&table)?;
                if let Some(columns) = doc.root.child_mut("tableColumns") {
                    let mut nodes: Vec<&mut Node> = columns.elements_mut().filter(|n| n.local() == "tableColumn").collect();
                    for (column, name) in changes {
                        if let Some(node) = nodes.get_mut(column) {
                            node.set("name", &name);
                        }
                    }
                }
                package.write_xml(&table, &doc);
            }
        }

        for part in &self.deleted_parts {
            package.delete(part);
            package.delete(&rels_part_name(part));
        }

        let structural = workbook_dirty || !self.deleted_parts.is_empty();
        if any_cells || structural {
            // Excel rebuilds the calculation chain; a stale one causes repair prompts.
            for rel in wb_rels.find_type("calcChain") {
                let part = crate::editor::package::resolve_target(&self.workbook_part, &rel.target);
                package.delete(&part);
                wb_rels.remove(&rel.id);
                wb_rels_dirty = true;
            }
        }
        if any_cells {
            let prefix = workbook.root.prefix().to_string();
            if workbook.root.child("calcPr").is_none() {
                workbook.root.insert_ordered(Node::new(&format!("{prefix}calcPr")), WORKBOOK_ORDER);
            }
            if let Some(calc) = workbook.root.child_mut("calcPr") {
                calc.set("fullCalcOnLoad", "1");
            }
            workbook_dirty = true;
        }

        if styles.dirty || self.styles_is_new && any_cells {
            package.write(&styles.part, styles.to_xml().into_bytes());
            if self.styles_is_new {
                let target = relative_target(&self.workbook_part, &styles.part);
                wb_rels.add("styles", &target, false);
                wb_rels_dirty = true;
                package.set_override(&styles.part, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            }
        }
        if let Some(sst) = sst.filter(SharedStrings::is_dirty) {
            let part = sst.part.clone();
            let is_new = sst.is_new_part();
            package.write(&part, sst.into_xml()?.into_bytes());
            if is_new {
                let target = relative_target(&self.workbook_part, &part);
                wb_rels.add("sharedStrings", &target, false);
                wb_rels_dirty = true;
                package.set_override(&part, content_type::SHARED_STRINGS);
            }
        }
        if workbook_dirty {
            package.write_xml(&self.workbook_part, &workbook);
        }
        if wb_rels_dirty {
            package.set_rels(&wb_rels);
        }
        if !self.core_updates.is_empty() {
            self.write_core_properties(&mut package)?;
        }
        Ok(package)
    }

    fn write_core_properties(&mut self, package: &mut Package) -> EResult<()> {
        const CORE_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        let part = self.core_part()?;
        let (part, mut doc, is_new) = match part {
            Some(part) => {
                let doc = package.read_xml(&part)?;
                (part, doc, false)
            }
            None => {
                let xml = format!(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<cp:coreProperties xmlns:cp=\"{CORE_NS}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"></cp:coreProperties>"
                );
                ("docProps/core.xml".to_string(), parse_document(&xml).map_err(|e| e.to_string())?, true)
            }
        };
        for (key, value) in &self.core_updates {
            let element = core_element(key).ok_or_else(|| format!("unknown document property {key:?}"))?;
            let lname = local(element);
            doc.root.remove_children(lname);
            let mut node = Node::new(element);
            if lname == "created" || lname == "modified" {
                node.set("xsi:type", "dcterms:W3CDTF");
                if !doc.root.attrs.iter().any(|(k, _)| k == "xmlns:xsi") {
                    doc.root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                }
            }
            node.set_text(value);
            doc.root.push(node);
        }
        package.write_xml(&part, &doc);
        if is_new {
            let mut root_rels = package.rels("")?;
            root_rels.add_with_type(
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                &part,
            );
            package.set_rels(&root_rels);
            package.set_override(&part, "application/vnd.openxmlformats-package.core-properties+xml");
        }
        Ok(())
    }
}

fn data_text(value: &Data) -> Option<String> {
    match value {
        Data::Empty => None,
        Data::Bool(b) => Some(if *b { "TRUE".into() } else { "FALSE".into() }),
        other => Some(other.to_string()),
    }
}

/// Apply `f` to every formula text in pending edits.
fn transform_pending_formulas(edits: &mut SheetEdits, f: &dyn Fn(&str) -> String) {
    for edit in edits.cells.values_mut() {
        if let CellEdit::Set { value: EditValue::Formula(formula), .. } = edit {
            let new = f(&formula.text);
            if new != formula.text {
                formula.text = new;
            }
        }
    }
}

/// Move pending edits of the edited sheet to their new positions.
fn remap_edits(edits: &mut SheetEdits, edit: &StructuralEdit) {
    let cells = std::mem::take(&mut edits.cells);
    for ((r, c), mut cell) in cells {
        if let Some((nr, nc)) = map_cell(edit, r, c) {
            if let CellEdit::Set { value: EditValue::Formula(formula), .. } = &mut cell
                && let Some(range) = &mut formula.array_ref
                && let Some(new) = structure::shift_area(range, edit) {
                    *range = new;
                }
            edits.cells.insert((nr, nc), cell);
        }
    }
    let rows = std::mem::take(&mut edits.rows);
    for (r, row) in rows {
        if let Some((nr, _)) = map_cell(edit, r, 0) {
            edits.rows.insert(nr, row);
        }
    }
    let shift_ranges = |list: &mut Vec<CellRange>| {
        let old = std::mem::take(list);
        for range in old {
            let text = crate::common::cellref::range_name(range, false);
            if let Some(new) = formula::shift_range_list(&text, true, edit).and_then(|t| parse_range(&t)) {
                list.push(new);
            }
        }
    };
    shift_ranges(&mut edits.merges_add);
    shift_ranges(&mut edits.merges_remove);
    edits.hyperlinks = std::mem::take(&mut edits.hyperlinks)
        .into_iter()
        .filter_map(|((r, c), url)| map_cell(edit, r, c).map(|p| (p, url)))
        .collect();
    for col in &mut edits.cols {
        let first = map_cell(edit, 0, col.first).map(|(_, c)| c);
        let last = map_cell(edit, 0, col.last).map(|(_, c)| c);
        if let (Some(f), Some(l)) = (first, last) {
            col.first = f;
            col.last = l;
        }
    }
}

fn load_formulas(reader: &mut Xlsx<Cursor<SharedBytes>>, name: &str) -> EResult<HashMap<(u32, u16), String>> {
    let range = reader.worksheet_formula(name).map_err(|e| format!("reading formulas of {name:?}: {e}"))?;
    let mut map = HashMap::new();
    if let Some((r0, c0)) = range.start() {
        for (r, c, f) in range.used_cells() {
            if !f.is_empty() {
                map.insert((r0 + r as u32, (c0 + c as u32) as u16), f.clone());
            }
        }
    }
    Ok(map)
}

/// Map the core.xml element for a Python property name.
fn core_element(key: &str) -> Option<&'static str> {
    Some(match key {
        "title" => "dc:title",
        "subject" => "dc:subject",
        "author" | "creator" => "dc:creator",
        "keywords" => "cp:keywords",
        "comments" | "description" => "dc:description",
        "last_modified_by" => "cp:lastModifiedBy",
        "category" => "cp:category",
        "status" => "cp:contentStatus",
        "created" => "dcterms:created",
        "modified" => "dcterms:modified",
        _ => return None,
    })
}

fn core_key(element: &str) -> &str {
    match element {
        "creator" => "author",
        "description" => "comments",
        "lastModifiedBy" => "last_modified_by",
        "contentStatus" => "status",
        other => other,
    }
}

/// Scan a sheet for cell style indexes and merged ranges.
fn scan_styles_and_merges(xml: &str) -> EResult<(HashMap<(u32, u16), u32>, Vec<CellRange>)> {
    let mut styles = HashMap::new();
    let mut merges = Vec::new();
    let mut scanner = Scanner::new(xml);
    let mut row: u32 = 0;
    let mut next_row: u32 = 0;
    let mut next_col: u16 = 0;
    while let Some((tok, _)) = scanner.next_tok().map_err(|e| e.to_string())? {
        match tok {
            Tok::Start { name, attrs } | Tok::Empty { name, attrs } => match local(name) {
                "row" => {
                    row = attr(attrs, "r").and_then(|r| r.parse::<u32>().ok()).map(|r| r - 1).unwrap_or(next_row);
                    next_row = row + 1;
                    next_col = 0;
                }
                "c" => {
                    let col = attr(attrs, "r")
                        .and_then(|r| crate::common::cellref::parse_cell(&r))
                        .map(|(_, c)| c)
                        .unwrap_or(next_col);
                    next_col = col.saturating_add(1);
                    if let Some(s) = attr(attrs, "s").and_then(|s| s.parse().ok()) {
                        styles.insert((row, col), s);
                    }
                }
                "mergeCell" => {
                    if let Some(range) = attr(attrs, "ref").and_then(|r| parse_range(&r)) {
                        merges.push(range);
                    }
                }
                _ => {}
            },
            _ => {}
        }
    }
    Ok((styles, merges))
}

impl Relationships {
    /// Add a relationship with a full type URI and an absolute part target.
    pub fn add_with_type(&mut self, full_type: &str, part: &str) -> String {
        let kind = full_type.rsplit('/').next().unwrap_or(full_type).to_string();
        let target = relative_target(&self.source, part);
        let id = self.add(&kind, &target, false);
        // `add` builds an officeDocument-relative type; fix it to the given URI.
        self.retype(&id, full_type);
        id
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn renames_sheet_references() {
        assert_eq!(rename_in_formula("Sheet1!$A$1:$B$2", "Sheet1", "Data"), "Data!$A$1:$B$2");
        assert_eq!(rename_in_formula("'My Sheet'!A1", "My Sheet", "Other"), "Other!A1");
        assert_eq!(rename_in_formula("Sheet1!A1", "Sheet1", "New Name"), "'New Name'!A1");
        assert_eq!(rename_in_formula("XSheet1!A1+Sheet1!B2", "Sheet1", "S"), "XSheet1!A1+S!B2");
    }

    #[test]
    fn sheet_name_rules() {
        assert!(check_sheet_name("Ok name").is_ok());
        assert!(check_sheet_name("").is_err());
        assert!(check_sheet_name("a/b").is_err());
        assert!(check_sheet_name("'x").is_err());
        assert!(check_sheet_name(&"x".repeat(32)).is_err());
    }
}
