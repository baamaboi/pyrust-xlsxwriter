//! The shared string table, append-only.
//!
//! Untouched sheets keep referring to existing string indexes, so existing
//! `<si>` items are never changed or reordered; new strings are appended.
//! The original XML is edited in place on save (new counts on the root tag,
//! new items before its end tag) instead of being copied into a new string.

use std::collections::HashMap;
use std::fmt::Write;

use crate::editor::package::EResult;
use crate::editor::xml::{attr, escape_attr, escape_text_into, local, Scanner, Tok};

pub struct SharedStrings {
    pub part: String,
    /// The original part's XML; None when the workbook has no shared strings yet.
    original: Option<String>,
    existing: u32,
    original_count: Option<u64>,
    /// New strings and their indexes (each string is stored once).
    added: HashMap<String, u32>,
    references: u64,
}

impl SharedStrings {
    pub fn empty(part: &str) -> Self {
        SharedStrings {
            part: part.to_string(),
            original: None,
            existing: 0,
            original_count: None,
            added: HashMap::new(),
            references: 0,
        }
    }

    pub fn parse(part: &str, xml: String) -> EResult<Self> {
        let mut scanner = Scanner::new(&xml);
        let mut existing = 0u32;
        let mut count = None;
        let mut depth = 0usize;
        while let Some((tok, _)) = scanner.next_tok().map_err(|e| e.to_string())? {
            match tok {
                Tok::Start { name, attrs } => {
                    if depth == 0 {
                        count = attr(attrs, "count").and_then(|c| c.parse().ok());
                    } else if depth == 1 && local(name) == "si" {
                        existing += 1;
                    }
                    depth += 1;
                }
                Tok::Empty { name, .. } if depth == 1 && local(name) == "si" => existing += 1,
                Tok::End { .. } => depth = depth.saturating_sub(1),
                _ => {}
            }
        }
        let mut table = SharedStrings::empty(part);
        table.original = Some(xml);
        table.existing = existing;
        table.original_count = count;
        Ok(table)
    }

    pub fn is_dirty(&self) -> bool {
        !self.added.is_empty()
    }

    pub fn is_new_part(&self) -> bool {
        self.original.is_none()
    }

    /// Index of a string, appending it if it's new.
    pub fn add(&mut self, text: &str) -> u32 {
        self.references += 1;
        if let Some(i) = self.added.get(text) {
            return *i;
        }
        let i = self.existing + self.added.len() as u32;
        self.added.insert(text.to_string(), i);
        i
    }

    /// The new `<si>` items, in index order.
    fn new_items(&self) -> String {
        let mut items: Vec<(&str, u32)> = self.added.iter().map(|(text, i)| (text.as_str(), *i)).collect();
        items.sort_unstable_by_key(|(_, i)| *i);
        let mut out = String::with_capacity(items.iter().map(|(text, _)| text.len() + 40).sum());
        for (text, _) in items {
            let preserve = text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace) || text.contains('\n');
            out.push_str(if preserve { "<si><t xml:space=\"preserve\">" } else { "<si><t>" });
            escape_text_into(&mut out, text);
            out.push_str("</t></si>");
        }
        out
    }

    /// The updated part. The original XML is edited in place, so saving
    /// never holds two copies of a large table.
    pub fn into_xml(self) -> EResult<String> {
        let unique = self.existing as u64 + self.added.len() as u64;
        let count = self.original_count.unwrap_or(self.existing as u64) + self.references;
        let new_items = self.new_items();
        let Some(mut xml) = self.original else {
            return Ok(format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{count}\" uniqueCount=\"{unique}\">{new_items}</sst>"
            ));
        };
        // The root start tag, rebuilt with the new counts.
        let mut scanner = Scanner::new(&xml);
        let (span, name, empty, mut tag) = loop {
            let Some((tok, span)) = scanner.next_tok().map_err(|e| e.to_string())? else {
                return Err("sharedStrings.xml has no root element".into());
            };
            if let Tok::Start { name, attrs } | Tok::Empty { name, attrs } = tok {
                let mut tag = format!("<{name}");
                for (k, v) in crate::editor::xml::attrs(attrs) {
                    if k != "count" && k != "uniqueCount" {
                        let _ = write!(tag, " {k}=\"{v}\"");
                    }
                }
                let _ = write!(tag, " count=\"{count}\" uniqueCount=\"{unique}\">");
                break (span, name.to_string(), matches!(tok, Tok::Empty { .. }), tag);
            }
        };
        xml.reserve(new_items.len() + tag.len() + name.len() + 3);
        if empty {
            // `<sst .../>` becomes `<sst ...>items</sst>`.
            tag.push_str(&new_items);
            let _ = write!(tag, "</{name}>");
        } else {
            // Insert the items first: they go after the start tag, so its span stays valid.
            let close = xml.rfind(&format!("</{name}>")).ok_or("sharedStrings.xml has no end tag")?;
            xml.insert_str(close, &new_items);
        }
        xml.replace_range(span, &tag);
        Ok(xml)
    }
}

/// Escape helper re-exported for sheet rendering.
pub fn escape_attribute(text: &str) -> String {
    escape_attr(text)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn appends_after_existing() {
        let xml = "<?xml version=\"1.0\"?><sst xmlns=\"x\" count=\"3\" uniqueCount=\"2\"><si><t>a</t></si><si><r><t>b</t></r></si></sst>".to_string();
        let mut sst = SharedStrings::parse("xl/sharedStrings.xml", xml).unwrap();
        assert_eq!(sst.add("new"), 2);
        assert_eq!(sst.add(" pad"), 3);
        assert_eq!(sst.add("new"), 2);
        let out = sst.into_xml().unwrap();
        assert!(out.contains("<si><r><t>b</t></r></si><si><t>new</t></si><si><t xml:space=\"preserve\"> pad</t></si></sst>"));
        assert!(out.contains("count=\"6\" uniqueCount=\"4\""));
    }
}
