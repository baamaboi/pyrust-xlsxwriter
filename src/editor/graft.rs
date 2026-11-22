//! Adding objects to existing sheets by "grafting".
//!
//! An object (conditional format, data validation, table, sparkline, image,
//! chart, shape or note) is first rendered by the writer into a scratch
//! workbook whose first sheet has the target sheet's name (other sheet names
//! exist as empty placeholders so chart ranges validate). This module then
//! transplants the result into the real package:
//!
//! * `<conditionalFormatting>` elements, with their `dxf` styles appended to
//!   the target stylesheet, priorities placed after the sheet's existing ones,
//!   and x14 extension ids made unique;
//! * `<dataValidation>` items, sparkline groups and other x14 extensions;
//! * drawing anchors (images, charts, shapes) merged into the sheet's drawing,
//!   with media/chart parts copied under fresh names;
//! * table parts with a workbook-unique id and name;
//! * notes (comments + VML) for sheets that don't have notes yet.
//!
//! Cells the writer filled in (table headers and totals) are applied as normal
//! cell edits when the graft is recorded.

use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicU64, Ordering};

use crate::common::bytes::SharedBytes;
use crate::editor::package::{relative_target, resolve_target, EResult, Package, Relationships, NS_REL};
use crate::editor::sheet::WORKSHEET_ORDER;
use crate::editor::styles::Styles;
use crate::editor::xml::{attr, local, parse_document, Child, Node, Scanner, Tok};

/// The scratch workbook's sheet that holds the rendered objects.
pub const SCRATCH_SHEET: &str = "xl/worksheets/sheet1.xml";

const URI_X14_CF: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

static GUID_COUNTER: AtomicU64 = AtomicU64::new(1);

/// A fresh GUID-shaped id for x14 extension links.
fn fresh_guid() -> String {
    let n = GUID_COUNTER.fetch_add(1, Ordering::Relaxed);
    let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_nanos() as u64).unwrap_or(0);
    let mixed = t ^ n.wrapping_mul(0x9E37_79B9_7F4A_7C15);
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:04X}-{:012X}}}",
        (mixed >> 32) as u32,
        (mixed >> 16) as u16,
        mixed as u16 | 0x4000,
        (n as u16) | 0x8000,
        n & 0xFFFF_FFFF_FFFF
    )
}

fn element_rank(name: &str) -> usize {
    WORKSHEET_ORDER.iter().position(|n| *n == name).unwrap_or(WORKSHEET_ORDER.len())
}

/// What to splice into the target sheet's XML.
#[derive(Default)]
struct Splice {
    conditional_formats: Vec<String>,
    data_validations: Vec<Node>,
    drawing_rid: Option<String>,
    legacy_rid: Option<String>,
    table_rids: Vec<String>,
    exts: Vec<Node>,
}

/// Insert/merge the splice into a worksheet's XML.
fn splice_sheet(xml: &str, sp: &Splice) -> EResult<String> {
    let mut scanner = Scanner::new(xml);
    let mut out = String::with_capacity(xml.len() + 4096);
    let (root_name, prefix) = loop {
        let (tok, span) = scanner.next_tok().map_err(|e| e.to_string())?.ok_or("empty worksheet")?;
        if let Tok::Start { name, attrs } = tok {
            out.push_str(&xml[..span.start]);
            out.push('<');
            out.push_str(name);
            out.push_str(attrs);
            let has_r = crate::editor::xml::attrs(attrs).any(|(k, v)| k == "xmlns:r" && v == NS_REL);
            if !has_r && crate::editor::xml::attrs(attrs).all(|(k, _)| k != "xmlns:r") {
                out.push_str(&format!(" xmlns:r=\"{NS_REL}\""));
            }
            out.push('>');
            let prefix = name.rfind(':').map(|i| name[..=i].to_string()).unwrap_or_default();
            break (name.to_string(), prefix);
        }
        if let Tok::Empty { .. } = tok {
            return Err("worksheet root is empty".into());
        }
    };
    let mut done: HashSet<&str> = HashSet::new();
    let render_new = |name: &str, prefix: &str| -> String {
        match name {
            "conditionalFormatting" => sp.conditional_formats.concat(),
            "dataValidations" if !sp.data_validations.is_empty() => {
                let mut node = Node::new(&format!("{prefix}dataValidations"));
                node.set("count", &sp.data_validations.len().to_string());
                for dv in &sp.data_validations {
                    node.push(dv.clone());
                }
                node.to_xml()
            }
            "drawing" => sp.drawing_rid.as_ref().map(|id| format!("<{prefix}drawing r:id=\"{id}\"/>")).unwrap_or_default(),
            "legacyDrawing" => sp.legacy_rid.as_ref().map(|id| format!("<{prefix}legacyDrawing r:id=\"{id}\"/>")).unwrap_or_default(),
            "tableParts" if !sp.table_rids.is_empty() => {
                let parts: String = sp.table_rids.iter().map(|id| format!("<{prefix}tablePart r:id=\"{id}\"/>")).collect();
                format!("<{prefix}tableParts count=\"{}\">{parts}</{prefix}tableParts>", sp.table_rids.len())
            }
            "extLst" if !sp.exts.is_empty() => {
                let mut node = Node::new(&format!("{prefix}extLst"));
                for ext in &sp.exts {
                    node.push(ext.clone());
                }
                node.to_xml()
            }
            _ => String::new(),
        }
    };
    const SPLICED: [&str; 6] = ["conditionalFormatting", "dataValidations", "drawing", "legacyDrawing", "tableParts", "extLst"];
    let flush = |rank: usize, out: &mut String, done: &mut HashSet<&str>| {
        for name in SPLICED {
            if element_rank(name) < rank && !done.contains(name) {
                done.insert(name);
                out.push_str(&render_new(name, &prefix));
            }
        }
    };
    while let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? {
        match tok {
            Tok::End { name } if name == root_name => {
                flush(WORKSHEET_ORDER.len() + 1, &mut out, &mut done);
                out.push_str(&xml[span.start..]);
                return Ok(out);
            }
            Tok::Start { name, .. } | Tok::Empty { name, .. } => {
                let is_start = matches!(tok, Tok::Start { .. });
                let end = if is_start { scanner.skip_element(name).map_err(|e| e.to_string())? } else { span.end };
                let element = &xml[span.start..end];
                let lname = local(name);
                flush(element_rank(lname), &mut out, &mut done);
                match lname {
                    "dataValidations" if !sp.data_validations.is_empty() => {
                        let mut doc = parse_document(element).map_err(|e| e.to_string())?;
                        for dv in &sp.data_validations {
                            doc.root.push(dv.clone());
                        }
                        let n = doc.root.children_named("dataValidation").count();
                        doc.root.set("count", &n.to_string());
                        out.push_str(&doc.root.to_xml());
                        done.insert("dataValidations");
                    }
                    "tableParts" if !sp.table_rids.is_empty() => {
                        let mut doc = parse_document(element).map_err(|e| e.to_string())?;
                        for id in &sp.table_rids {
                            doc.root.push(Node::new(&format!("{prefix}tablePart")).with_attr("r:id", id));
                        }
                        let n = doc.root.children_named("tablePart").count();
                        doc.root.set("count", &n.to_string());
                        out.push_str(&doc.root.to_xml());
                        done.insert("tableParts");
                    }
                    "extLst" if !sp.exts.is_empty() => {
                        let mut doc = parse_document(element).map_err(|e| e.to_string())?;
                        merge_exts(&mut doc.root, &sp.exts);
                        out.push_str(&doc.root.to_xml());
                        done.insert("extLst");
                    }
                    "drawing" => {
                        done.insert("drawing");
                        out.push_str(element);
                    }
                    "legacyDrawing" => {
                        done.insert("legacyDrawing");
                        out.push_str(element);
                    }
                    _ => out.push_str(element),
                }
            }
            _ => out.push_str(&xml[span]),
        }
    }
    Err("worksheet ended early".into())
}

/// Merge `<ext>` nodes into an `<extLst>`: same-uri extensions get their
/// inner items appended, others are added whole.
fn merge_exts(ext_list: &mut Node, exts: &[Node]) {
    for ext in exts {
        let uri = ext.get("uri");
        let existing = ext_list.elements_mut().find(|n| n.local() == "ext" && n.get("uri") == uri);
        match existing {
            Some(target) => {
                // <ext><x14:conditionalFormattings>items</...></ext>
                let source_container = ext.elements().next();
                let target_container = target.elements_mut().next();
                match (source_container, target_container) {
                    (Some(src), Some(dst)) if src.local() == dst.local() => {
                        for item in src.elements() {
                            dst.push(item.clone());
                        }
                    }
                    _ => ext_list.push(ext.clone()),
                }
            }
            None => ext_list.push(ext.clone()),
        }
    }
}

/// Copy a part from the scratch package under a fresh name; returns the new name.
fn copy_part(target: &mut Package, scratch: &mut Package, part: &str, taken: &mut HashSet<String>) -> EResult<String> {
    let (stem, ext) = match part.rfind('.') {
        Some(i) => (&part[..i], &part[i..]),
        None => (part, ""),
    };
    let base = stem.trim_end_matches(|c: char| c.is_ascii_digit());
    let name = (1..)
        .map(|n| format!("{base}{n}{ext}"))
        .find(|n| !target.exists(n) && !taken.contains(n))
        .unwrap_or_default();
    taken.insert(name.clone());
    let bytes = scratch.read(part)?.into_owned();
    target.write(&name, bytes);
    Ok(name)
}

fn content_type_for(part: &str) -> Option<&'static str> {
    if part.contains("/charts/chart") {
        Some("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
    } else if part.contains("/drawings/drawing") && part.ends_with(".xml") {
        Some("application/vnd.openxmlformats-officedocument.drawing+xml")
    } else if part.contains("/tables/table") {
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    } else if part.contains("/comments") {
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
    } else {
        None
    }
}

fn default_type_for(ext: &str) -> Option<&'static str> {
    Some(match ext {
        "png" => "image/png",
        "jpeg" | "jpg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "vml" => "application/vnd.openxmlformats-officedocument.vmlDrawing",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => return None,
    })
}

fn register_part(target: &mut Package, part: &str) {
    if let Some(ct) = content_type_for(part) {
        target.set_override(part, ct);
    } else if let Some(ext) = part.rsplit('.').next()
        && let Some(ct) = default_type_for(&ext.to_ascii_lowercase()) {
            target.ensure_default(&ext.to_ascii_lowercase(), ct);
        }
}

/// Rewrite relationship-id attributes (`r:id`, `r:embed`, `r:link`, ...) via `map`.
fn remap_rids(node: &mut Node, map: &HashMap<String, String>) {
    for (k, v) in node.attrs.iter_mut() {
        if k.starts_with("r:")
            && let Some(new) = map.get(v.as_str()) {
                *v = new.clone();
            }
    }
    for child in node.elements_mut() {
        remap_rids(child, map);
    }
}

fn max_drawing_id(node: &Node) -> u32 {
    let own = if node.local() == "cNvPr" { node.get("id").and_then(|v| v.parse().ok()).unwrap_or(0) } else { 0 };
    node.elements().map(max_drawing_id).fold(own, u32::max)
}

fn renumber_drawing_ids(node: &mut Node, next: &mut u32) {
    if node.local() == "cNvPr" {
        node.set("id", &next.to_string());
        *next += 1;
    }
    for child in node.elements_mut() {
        renumber_drawing_ids(child, next);
    }
}

/// Graft the objects rendered in `scratch` into `sheet_part` of `target`.
pub fn apply(target: &mut Package, styles: &mut Styles, sheet_part: &str, scratch_bytes: &SharedBytes) -> EResult<()> {
    let mut scratch = Package::open(scratch_bytes.clone())?;
    let scratch_sheet = scratch.read_xml(SCRATCH_SHEET)?;
    let scratch_rels = scratch.rels(SCRATCH_SHEET)?;
    let mut sheet_rels = target.rels(sheet_part)?;
    let mut taken: HashSet<String> = HashSet::new();
    let mut splice = Splice::default();
    // Owned: the package changes while the objects are grafted.
    let target_xml = target.read_string(sheet_part)?.into_owned();

    // Styles referenced by conditional formats.
    let scratch_dxfs: Vec<Node> = {
        let styles_doc = scratch.read_xml("xl/styles.xml")?;
        styles_doc.root.child("dxfs").map(|d| d.children_named("dxf").cloned().collect()).unwrap_or_default()
    };
    let mut dxf_map: HashMap<String, String> = HashMap::new();
    let mut guid_map: HashMap<String, String> = HashMap::new();
    let max_priority = target_xml
        .match_indices("priority=\"")
        .filter_map(|(i, m)| target_xml[i + m.len()..].split('"').next()?.parse::<u32>().ok())
        .max()
        .unwrap_or(0);

    let fix_cf = |node: &mut Node, styles: &mut Styles, dxf_map: &mut HashMap<String, String>, guid_map: &mut HashMap<String, String>| -> EResult<()> {
        fn walk(node: &mut Node, styles: &mut Styles, dxfs: &[Node], dxf_map: &mut HashMap<String, String>, guid_map: &mut HashMap<String, String>, max_priority: u32) -> EResult<()> {
            if let Some(dxf) = node.get("dxfId") {
                let new = match dxf_map.get(&dxf) {
                    Some(n) => n.clone(),
                    None => {
                        let source = dxf.parse::<usize>().ok().and_then(|i| dxfs.get(i)).ok_or("conditional format refers to a missing style")?;
                        let n = styles.add_dxf(source.clone()).to_string();
                        dxf_map.insert(dxf.clone(), n.clone());
                        n
                    }
                };
                node.set("dxfId", &new);
            }
            if let Some(p) = node.get("priority").and_then(|p| p.parse::<u32>().ok()) {
                node.set("priority", &(p + max_priority).to_string());
            }
            if let Some(id) = node.get("id").filter(|v| v.starts_with('{')) {
                let new = guid_map.entry(id).or_insert_with(fresh_guid).clone();
                node.set("id", &new);
            }
            if node.local() == "id" {
                let text = node.text();
                if text.starts_with('{') {
                    let new = guid_map.entry(text).or_insert_with(fresh_guid).clone();
                    node.set_text(&new);
                }
            }
            for child in node.elements_mut() {
                walk(child, styles, dxfs, dxf_map, guid_map, max_priority)?;
            }
            Ok(())
        }
        walk(node, styles, &scratch_dxfs, dxf_map, guid_map, max_priority)
    };

    for element in scratch_sheet.root.elements() {
        match element.local() {
            "conditionalFormatting" => {
                let mut node = element.clone();
                fix_cf(&mut node, styles, &mut dxf_map, &mut guid_map)?;
                splice.conditional_formats.push(node.to_xml());
            }
            "dataValidations" => {
                splice.data_validations.extend(element.children_named("dataValidation").cloned());
            }
            "extLst" => {
                for ext in element.children_named("ext") {
                    let mut ext = ext.clone();
                    if ext.get("uri").as_deref() == Some(URI_X14_CF) {
                        fix_cf(&mut ext, styles, &mut dxf_map, &mut guid_map)?;
                    }
                    splice.exts.push(ext);
                }
            }
            "drawing" => {
                let rid = element.get("r:id").ok_or("drawing without relationship")?;
                let drawing_part = scratch_rels.target_part(&rid).ok_or("missing scratch drawing")?;
                graft_drawing(target, &mut scratch, &drawing_part, sheet_part, &mut sheet_rels, &mut splice, &mut taken)?;
            }
            "legacyDrawing" => {
                graft_notes(target, &mut scratch, &scratch_rels, sheet_part, &mut sheet_rels, &mut splice, &mut taken)?;
            }
            "tableParts" => {
                for part in element.children_named("tablePart") {
                    let rid = part.get("r:id").ok_or("tablePart without relationship")?;
                    let table_part = scratch_rels.target_part(&rid).ok_or("missing scratch table")?;
                    let new_rid = graft_table(target, &mut scratch, &table_part, sheet_part, &mut sheet_rels, &mut taken)?;
                    splice.table_rids.push(new_rid);
                }
            }
            "autoFilter" | "hyperlinks" | "mergeCells" | "protectedRanges" | "sheetProtection" => {
                return Err(format!("adding <{}> to an existing sheet is not supported", element.local()));
            }
            _ => {}
        }
    }

    let out = splice_sheet(&target_xml, &splice)?;
    target.write(sheet_part, out.into_bytes());
    target.set_rels(&sheet_rels);
    Ok(())
}

fn graft_drawing(
    target: &mut Package,
    scratch: &mut Package,
    scratch_drawing: &str,
    sheet_part: &str,
    sheet_rels: &mut Relationships,
    splice: &mut Splice,
    taken: &mut HashSet<String>,
) -> EResult<()> {
    let source = scratch.read_xml(scratch_drawing)?;
    let source_rels = scratch.rels(scratch_drawing)?;
    // Existing drawing on the target sheet, or a new one.
    let existing = sheet_rels.find_type("drawing").into_iter().next().map(|r| resolve_target(sheet_part, &r.target));
    let (drawing_part, mut doc, is_new) = match existing {
        Some(part) => {
            let doc = target.read_xml(&part)?;
            (part, doc, false)
        }
        None => {
            let name = (1..)
                .map(|n| format!("xl/drawings/drawing{n}.xml"))
                .find(|n| !target.exists(n) && !taken.contains(n))
                .unwrap_or_default();
            taken.insert(name.clone());
            let mut doc = source.clone();
            doc.root.children.retain(|c| !matches!(c, Child::Element(_)));
            (name, doc, true)
        }
    };
    let mut drawing_rels = target.rels(&drawing_part)?;
    let mut rid_map: HashMap<String, String> = HashMap::new();
    for rel in source_rels.list() {
        if rel.external {
            let kind = rel.rel_type.rsplit('/').next().unwrap_or("").to_string();
            let id = drawing_rels.add(&kind, &rel.target, true);
            drawing_rels.retype(&id, &rel.rel_type);
            rid_map.insert(rel.id, id);
            continue;
        }
        let part = resolve_target(scratch_drawing, &rel.target);
        let new_part = copy_part(target, scratch, &part, taken)?;
        register_part(target, &new_part);
        // Charts may have their own relationships (styles, colours, images).
        let chart_rels = scratch.rels(&part)?;
        if !chart_rels.is_empty() {
            let mut new_rels = Relationships::empty(&new_part);
            for sub in chart_rels.list() {
                if sub.external {
                    new_rels.push_raw(&sub.id, &sub.rel_type, &sub.target, true);
                    continue;
                }
                let sub_part = resolve_target(&part, &sub.target);
                let new_sub = copy_part(target, scratch, &sub_part, taken)?;
                register_part(target, &new_sub);
                // Keep the original id so references inside the chart still match.
                new_rels.push_raw(&sub.id, &sub.rel_type, &relative_target(&new_part, &new_sub), false);
            }
            target.set_rels(&new_rels);
        }
        let id = drawing_rels.add("x", &relative_target(&drawing_part, &new_part), false);
        drawing_rels.retype(&id, &rel.rel_type);
        rid_map.insert(rel.id, id);
    }
    let mut next_id = max_drawing_id(&doc.root) + 1;
    for anchor in source.root.elements() {
        let mut anchor = anchor.clone();
        remap_rids(&mut anchor, &rid_map);
        renumber_drawing_ids(&mut anchor, &mut next_id);
        doc.root.push(anchor);
    }
    target.write_xml(&drawing_part, &doc);
    target.set_rels(&drawing_rels);
    if is_new {
        register_part(target, &drawing_part);
        let id = sheet_rels.add("drawing", &relative_target(sheet_part, &drawing_part), false);
        splice.drawing_rid = Some(id);
    }
    Ok(())
}

fn graft_notes(
    target: &mut Package,
    scratch: &mut Package,
    scratch_rels: &Relationships,
    sheet_part: &str,
    sheet_rels: &mut Relationships,
    splice: &mut Splice,
    taken: &mut HashSet<String>,
) -> EResult<()> {
    if !sheet_rels.find_type("comments").is_empty() || !sheet_rels.find_type("vmlDrawing").is_empty() {
        return Err("adding notes to a sheet that already has notes, buttons or header images is not supported yet".into());
    }
    let comments = scratch_rels.find_type("comments").into_iter().next().ok_or("scratch notes without comments")?;
    let vml = scratch_rels.find_type("vmlDrawing").into_iter().next().ok_or("scratch notes without VML")?;
    let comments_part = resolve_target(SCRATCH_SHEET, &comments.target);
    let vml_part = resolve_target(SCRATCH_SHEET, &vml.target);
    let new_comments = copy_part(target, scratch, &comments_part, taken)?;
    register_part(target, &new_comments);
    // VML shape ids come in blocks of 1024 per drawing; use an unused block.
    let mut used = 0;
    for name in target.part_names().iter().filter(|n| n.ends_with(".vml")) {
        let Ok(text) = target.read_string(name) else { continue };
        let blocks = text
            .match_indices("data=\"")
            .filter_map(|(i, m)| text[i + m.len()..].split('"').next().map(|s| s.split(',').filter_map(|v| v.trim().parse::<u32>().ok()).max().unwrap_or(0)));
        used = used.max(blocks.max().unwrap_or(0));
    }
    let block = used + 1;
    let vml_text = scratch.read_string(&vml_part)?;
    let mut renumbered = vml_text.replace("data=\"1\"", &format!("data=\"{block}\""));
    for n in (1025..1025 + 1024).rev() {
        let old = format!("_x0000_s{n}");
        if renumbered.contains(&old) {
            renumbered = renumbered.replace(&old, &format!("_x0000_s{}", block * 1024 + (n - 1024)));
        }
    }
    let new_vml = (1..)
        .map(|n| format!("xl/drawings/vmlDrawing{n}.vml"))
        .find(|n| !target.exists(n) && !taken.contains(n))
        .unwrap_or_default();
    taken.insert(new_vml.clone());
    target.write(&new_vml, renumbered.into_bytes());
    register_part(target, &new_vml);
    sheet_rels.add("comments", &relative_target(sheet_part, &new_comments), false);
    let id = sheet_rels.add("vmlDrawing", &relative_target(sheet_part, &new_vml), false);
    splice.legacy_rid = Some(id);
    Ok(())
}

/// Copy a table part with a workbook-unique id and name; returns the sheet relationship id.
fn graft_table(
    target: &mut Package,
    scratch: &mut Package,
    table_part: &str,
    sheet_part: &str,
    sheet_rels: &mut Relationships,
    taken: &mut HashSet<String>,
) -> EResult<String> {
    let mut existing_ids = HashSet::new();
    let mut existing_names = HashSet::new();
    for part in target.part_names().into_iter().filter(|n| n.starts_with("xl/tables/table")) {
        if let Ok(doc) = target.read_xml(&part) {
            if let Some(id) = doc.root.get("id").and_then(|v| v.parse::<u32>().ok()) {
                existing_ids.insert(id);
            }
            for key in ["name", "displayName"] {
                if let Some(v) = doc.root.get(key) {
                    existing_names.insert(v.to_lowercase());
                }
            }
        }
    }
    let mut doc = scratch.read_xml(table_part)?;
    let id = (1..).find(|i| !existing_ids.contains(i)).unwrap_or(1);
    doc.root.set("id", &id.to_string());
    let name = doc.root.get("displayName").unwrap_or_else(|| "Table1".into());
    if existing_names.contains(&name.to_lowercase()) {
        let is_default = name.strip_prefix("Table").is_some_and(|n| n.chars().all(|c| c.is_ascii_digit()));
        if !is_default {
            return Err(format!("a table named {name:?} already exists in this workbook"));
        }
        let fresh = (1..).map(|n| format!("Table{n}")).find(|n| !existing_names.contains(&n.to_lowercase())).unwrap_or_default();
        doc.root.set("name", &fresh);
        doc.root.set("displayName", &fresh);
    }
    let new_part = (1..)
        .map(|n| format!("xl/tables/table{n}.xml"))
        .find(|n| !target.exists(n) && !taken.contains(n))
        .unwrap_or_default();
    taken.insert(new_part.clone());
    target.write_xml(&new_part, &doc);
    register_part(target, &new_part);
    Ok(sheet_rels.add("table", &relative_target(sheet_part, &new_part), false))
}

/// Style indexes of the cells the writer filled in, with the scratch stylesheet.
pub fn scratch_cell_styles(scratch: &SharedBytes) -> EResult<(HashMap<(u32, u16), u32>, Styles)> {
    let mut package = Package::open(scratch.clone())?;
    let styles = Styles::new("xl/styles.xml", package.read_xml("xl/styles.xml")?);
    let xml = package.read_string(SCRATCH_SHEET)?;
    let mut map = HashMap::new();
    let mut scanner = Scanner::new(&xml);
    while let Some((tok, _)) = scanner.next_tok().map_err(|e| e.to_string())? {
        if let Tok::Start { name, attrs } | Tok::Empty { name, attrs } = tok
            && local(name) == "c"
                && let (Some(r), Some(s)) = (attr(attrs, "r"), attr(attrs, "s"))
                    && let (Some(pos), Ok(s)) = (crate::common::cellref::parse_cell(&r), s.parse::<u32>())
                        && s != 0 {
                            map.insert(pos, s);
                        }
    }
    Ok((map, styles))
}

