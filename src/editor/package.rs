//! The OPC package (zip) of an xlsx file.
//!
//! Parts are read on demand. Changed and new parts are kept in memory; on
//! save, every untouched entry is copied with `raw_copy_file`, i.e. its
//! compressed bytes are reused without decompressing, so parts the editor
//! doesn't understand (charts, pivots, VBA, slicers, ...) survive exactly.
//!
//! The original file is held once, as [`SharedBytes`], and shared with
//! clones of the package and with calamine. Changed parts are lent out by
//! [`Package::read`] rather than copied, and the result can be written
//! straight into a file by [`Package::save`].

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::common::bytes::SharedBytes;
use crate::editor::xml::{parse_document, Document, Node, XML_DECL};

pub type EResult<T> = Result<T, String>;

pub const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

pub mod rel_type {
    const BASE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    pub fn of(kind: &str) -> String {
        format!("{BASE}{kind}")
    }
    pub fn is(rel_type: &str, kind: &str) -> bool {
        rel_type.rsplit('/').next() == Some(kind)
    }
}

pub mod content_type {
    pub const WORKSHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const SHARED_STRINGS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    pub const CALC_CHAIN: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
}

/// Resolve a relationship target relative to its source part.
pub fn resolve_target(source_part: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut parts: Vec<&str> = match source_part.rfind('/') {
        Some(i) => source_part[..i].split('/').collect(),
        None => Vec::new(),
    };
    for segment in target.split('/') {
        match segment {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// A target for `source_part`'s relationships pointing at `part` (relative).
pub fn relative_target(source_part: &str, part: &str) -> String {
    let from: Vec<&str> = match source_part.rfind('/') {
        Some(i) => source_part[..i].split('/').collect(),
        None => Vec::new(),
    };
    let to: Vec<&str> = part.split('/').collect();
    let common = from.iter().zip(&to).take_while(|(a, b)| a == b).count();
    let mut out: Vec<&str> = vec![".."; from.len() - common];
    out.extend(&to[common..]);
    out.join("/")
}

pub fn rels_part_name(part: &str) -> String {
    match part.rfind('/') {
        Some(i) => format!("{}/_rels/{}.rels", &part[..i], &part[i + 1..]),
        None => format!("_rels/{part}.rels"),
    }
}

/// One relationship.
#[derive(Debug, Clone, PartialEq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub external: bool,
}

/// A `.rels` part.
#[derive(Debug, Clone)]
pub struct Relationships {
    pub source: String,
    doc: Document,
}

impl Relationships {
    pub fn empty(source: &str) -> Self {
        let root = Node::new("Relationships").with_attr("xmlns", NS_PKG_REL);
        Relationships {
            source: source.to_string(),
            doc: Document { prolog: XML_DECL.to_string(), root, epilog: String::new() },
        }
    }

    pub fn list(&self) -> Vec<Relationship> {
        self.doc
            .root
            .children_named("Relationship")
            .map(|n| Relationship {
                id: n.get("Id").unwrap_or_default(),
                rel_type: n.get("Type").unwrap_or_default(),
                target: n.get("Target").unwrap_or_default(),
                external: n.get("TargetMode").as_deref() == Some("External"),
            })
            .collect()
    }

    pub fn get(&self, id: &str) -> Option<Relationship> {
        self.list().into_iter().find(|r| r.id == id)
    }

    /// The resolved part a relationship points at (None for external targets).
    pub fn target_part(&self, id: &str) -> Option<String> {
        let rel = self.get(id)?;
        if rel.external { None } else { Some(resolve_target(&self.source, &rel.target)) }
    }

    pub fn find_type(&self, kind: &str) -> Vec<Relationship> {
        self.list().into_iter().filter(|r| rel_type::is(&r.rel_type, kind)).collect()
    }

    fn next_id(&self) -> String {
        let used: HashSet<String> = self.list().into_iter().map(|r| r.id).collect();
        (1..).map(|n| format!("rId{n}")).find(|id| !used.contains(id)).unwrap_or_default()
    }

    /// Add a relationship and return its id.
    pub fn add(&mut self, kind: &str, target: &str, external: bool) -> String {
        let id = self.next_id();
        let mut node = Node::new("Relationship")
            .with_attr("Id", &id)
            .with_attr("Type", &rel_type::of(kind))
            .with_attr("Target", target);
        if external {
            node.set("TargetMode", "External");
        }
        self.doc.root.push(node);
        id
    }

    /// Add a relationship with an exact id and full type URI.
    pub fn push_raw(&mut self, id: &str, full_type: &str, target: &str, external: bool) {
        let mut node = Node::new("Relationship")
            .with_attr("Id", id)
            .with_attr("Type", full_type)
            .with_attr("Target", target);
        if external {
            node.set("TargetMode", "External");
        }
        self.doc.root.push(node);
    }

    /// Change the full type URI of a relationship.
    pub fn retype(&mut self, id: &str, full_type: &str) {
        for node in self.doc.root.elements_mut() {
            if node.get("Id").as_deref() == Some(id) {
                node.set("Type", full_type);
            }
        }
    }

    pub fn remove(&mut self, id: &str) {
        self.doc
            .root
            .children
            .retain(|c| !matches!(c, crate::editor::xml::Child::Element(n) if n.get("Id").as_deref() == Some(id)));
    }

    pub fn is_empty(&self) -> bool {
        self.doc.root.elements().next().is_none()
    }

    pub fn to_xml(&self) -> String {
        self.doc.to_xml()
    }
}

#[derive(Clone)]
pub struct Package {
    archive: ZipArchive<Cursor<SharedBytes>>,
    data: SharedBytes,
    /// Original entry names, in order.
    entries: Vec<String>,
    original: HashSet<String>,
    modified: HashMap<String, Vec<u8>>,
    added: Vec<String>,
    deleted: HashSet<String>,
    content_types: Document,
    content_types_dirty: bool,
}

impl Package {
    pub fn open(data: SharedBytes) -> EResult<Self> {
        let archive = ZipArchive::new(Cursor::new(data.clone()))
            .map_err(|e| format!("not an xlsx/zip file: {e}"))?;
        let entries: Vec<String> = archive
            .file_names()
            .filter_map(|n| n.ok().map(|n| n.into_owned()))
            .collect();
        let original = entries.iter().cloned().collect();
        let mut package = Package {
            archive,
            data,
            entries,
            original,
            modified: HashMap::new(),
            added: Vec::new(),
            deleted: HashSet::new(),
            content_types: Document {
                prolog: String::new(),
                root: Node::new("Types"),
                epilog: String::new(),
            },
            content_types_dirty: false,
        };
        package.content_types = package.read_xml("[Content_Types].xml")?;
        Ok(package)
    }

    /// The bytes of the original file (shared, not copied).
    pub fn original_bytes(&self) -> SharedBytes {
        self.data.clone()
    }

    /// Names of all parts currently in the package.
    pub fn part_names(&self) -> Vec<String> {
        self.entries
            .iter()
            .chain(self.added.iter())
            .filter(|n| self.exists(n))
            .cloned()
            .collect()
    }

    pub fn exists(&self, name: &str) -> bool {
        !self.deleted.contains(name) && (self.modified.contains_key(name) || self.original.contains(name))
    }

    pub fn is_modified(&self, name: &str) -> bool {
        self.modified.contains_key(name)
    }

    /// The bytes of a part. A part already changed in this package is lent
    /// out as it is (no copy); an original part is decompressed into a new
    /// buffer. Callers that need to keep the bytes while changing the
    /// package call `.into_owned()`, which copies only in the lent case.
    pub fn read(&mut self, name: &str) -> EResult<Cow<'_, [u8]>> {
        if self.deleted.contains(name) {
            return Err(format!("part {name} was deleted"));
        }
        if self.modified.contains_key(name) {
            return Ok(Cow::Borrowed(&self.modified[name]));
        }
        let mut file = self.archive.by_name(name).map_err(|e| format!("missing part {name}: {e}"))?;
        let mut out = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut out).map_err(|e| format!("reading {name}: {e}"))?;
        Ok(Cow::Owned(out))
    }

    /// A part as text (see [`Package::read`] for when it is lent or decompressed).
    pub fn read_string(&mut self, name: &str) -> EResult<Cow<'_, str>> {
        let not_utf8 = || format!("part {name} is not UTF-8");
        Ok(match self.read(name)? {
            Cow::Borrowed(bytes) => Cow::Borrowed(std::str::from_utf8(bytes).map_err(|_| not_utf8())?),
            Cow::Owned(bytes) => Cow::Owned(String::from_utf8(bytes).map_err(|_| not_utf8())?),
        })
    }

    pub fn read_xml(&mut self, name: &str) -> EResult<Document> {
        let text = self.read_string(name)?;
        parse_document(&text).map_err(|e| format!("{name}: {e}"))
    }

    pub fn write(&mut self, name: &str, bytes: Vec<u8>) {
        self.deleted.remove(name);
        if !self.original.contains(name) && !self.added.iter().any(|n| n == name) {
            self.added.push(name.to_string());
        }
        self.modified.insert(name.to_string(), bytes);
    }

    pub fn write_xml(&mut self, name: &str, doc: &Document) {
        self.write(name, doc.to_xml().into_bytes());
    }

    pub fn delete(&mut self, name: &str) {
        self.modified.remove(name);
        self.added.retain(|n| n != name);
        if self.original.contains(name) {
            self.deleted.insert(name.to_string());
        }
        self.remove_override(name);
    }

    /// The relationships of `part` ("" for the package root).
    pub fn rels(&mut self, part: &str) -> EResult<Relationships> {
        let rels_name = if part.is_empty() { "_rels/.rels".to_string() } else { rels_part_name(part) };
        if !self.exists(&rels_name) {
            return Ok(Relationships::empty(part));
        }
        let doc = self.read_xml(&rels_name)?;
        Ok(Relationships { source: part.to_string(), doc })
    }

    pub fn set_rels(&mut self, rels: &Relationships) {
        let rels_name = if rels.source.is_empty() {
            "_rels/.rels".to_string()
        } else {
            rels_part_name(&rels.source)
        };
        if rels.is_empty() {
            self.delete(&rels_name);
        } else {
            self.write(&rels_name, rels.to_xml().into_bytes());
        }
    }

    pub fn set_override(&mut self, part: &str, content_type: &str) {
        let part_name = format!("/{part}");
        let root = &mut self.content_types.root;
        if let Some(node) = root
            .elements_mut()
            .find(|n| n.local() == "Override" && n.get("PartName").as_deref() == Some(part_name.as_str()))
        {
            node.set("ContentType", content_type);
        } else {
            root.push(
                Node::new("Override")
                    .with_attr("PartName", &part_name)
                    .with_attr("ContentType", content_type),
            );
        }
        self.content_types_dirty = true;
    }

    /// Add a `<Default>` content type for an extension if there is none.
    pub fn ensure_default(&mut self, extension: &str, content_type: &str) {
        let root = &mut self.content_types.root;
        let exists = root
            .elements()
            .any(|n| n.local() == "Default" && n.get("Extension").is_some_and(|e| e.eq_ignore_ascii_case(extension)));
        if !exists {
            let node = Node::new("Default").with_attr("Extension", extension).with_attr("ContentType", content_type);
            root.children.insert(0, crate::editor::xml::Child::Element(node));
            root.self_closing = false;
            self.content_types_dirty = true;
        }
    }

    pub fn remove_override(&mut self, part: &str) {
        let part_name = format!("/{part}");
        let before = self.content_types.root.children.len();
        self.content_types.root.children.retain(|c| {
            !matches!(c, crate::editor::xml::Child::Element(n)
                if n.local() == "Override" && n.get("PartName").as_deref() == Some(part_name.as_str()))
        });
        if self.content_types.root.children.len() != before {
            self.content_types_dirty = true;
        }
    }

    /// The smallest unused `{prefix}{n}{suffix}` part name.
    pub fn next_part_name(&self, prefix: &str, suffix: &str) -> String {
        (1..)
            .map(|n| format!("{prefix}{n}{suffix}"))
            .find(|name| !self.exists(name))
            .unwrap_or_default()
    }

    /// Write the package to an in-memory buffer sized like the original file.
    pub fn save_to_vec(&mut self) -> EResult<Vec<u8>> {
        let buffer = Cursor::new(Vec::with_capacity(self.data.len() + 4096));
        Ok(self.save(buffer)?.into_inner())
    }

    /// Write the package as a zip into `out` (a file or a buffer) and return `out`.
    /// Writing to a file directly avoids holding the whole result in memory.
    pub fn save<W: Write + Seek>(&mut self, out: W) -> EResult<W> {
        let mut writer = ZipWriter::new(out);
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .large_file(self.data.len() > u32::MAX as usize / 2);
        let content_types = if self.content_types_dirty {
            Some(self.content_types.to_xml().into_bytes())
        } else {
            None
        };

        for (index, name) in self.entries.iter().enumerate() {
            if self.deleted.contains(name) || name.ends_with('/') {
                continue;
            }
            let replacement = if name == "[Content_Types].xml" {
                content_types.as_ref()
            } else {
                self.modified.get(name)
            };
            match replacement {
                Some(bytes) => {
                    writer.start_file(name, options).map_err(|e| e.to_string())?;
                    writer.write_all(bytes).map_err(|e| e.to_string())?;
                }
                None => {
                    let file = self.archive.by_index_raw(index).map_err(|e| e.to_string())?;
                    writer.raw_copy_file(file).map_err(|e| e.to_string())?;
                }
            }
        }
        for name in &self.added {
            let bytes = &self.modified[name];
            writer.start_file(name, options).map_err(|e| e.to_string())?;
            writer.write_all(bytes).map_err(|e| e.to_string())?;
        }
        writer.finish().map_err(|e| e.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets() {
        assert_eq!(resolve_target("xl/workbook.xml", "worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml");
        assert_eq!(resolve_target("xl/worksheets/sheet1.xml", "../drawings/d1.xml"), "xl/drawings/d1.xml");
        assert_eq!(resolve_target("xl/workbook.xml", "/xl/styles.xml"), "xl/styles.xml");
        assert_eq!(relative_target("xl/workbook.xml", "xl/worksheets/sheet2.xml"), "worksheets/sheet2.xml");
        assert_eq!(relative_target("xl/worksheets/sheet1.xml", "xl/sharedStrings.xml"), "../sharedStrings.xml");
        assert_eq!(rels_part_name("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
    }
}
