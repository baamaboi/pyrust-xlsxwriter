//! A small XML scanner and DOM for OOXML parts.
//!
//! The editor must leave everything it does not change byte-for-byte
//! identical, so instead of a full XML library it uses:
//!
//! * [`Scanner`]: a tokenizer that reports each tag/text with its byte span,
//!   so unchanged regions can be copied verbatim (used for large sheets);
//! * [`Node`]: a minimal DOM for small parts (workbook.xml, styles.xml, rels,
//!   content types). Text and attribute values are kept in their original
//!   escaped form, so an unmodified node serialises back exactly.
//!
//! OOXML parts are machine-generated, well-formed, DTD-free XML, which is
//! what this scanner supports.

use std::ops::Range;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlError(pub String);

impl std::fmt::Display for XmlError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "XML error: {}", self.0)
    }
}

pub type XmlResult<T> = Result<T, XmlError>;

/// A token and the byte range it occupies in the source.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Tok<'a> {
    /// `<name attrs>`; `attrs` is the raw attribute text.
    Start { name: &'a str, attrs: &'a str },
    /// `<name attrs/>`.
    Empty { name: &'a str, attrs: &'a str },
    /// `</name>`.
    End { name: &'a str },
    /// Character data between tags (raw, still escaped).
    Text(&'a str),
    /// Declaration, comment, processing instruction, CDATA or doctype.
    Other(&'a str),
}

pub struct Scanner<'a> {
    src: &'a str,
    pos: usize,
}

impl<'a> Scanner<'a> {
    pub fn new(src: &'a str) -> Self {
        Scanner { src, pos: 0 }
    }

    pub fn position(&self) -> usize {
        self.pos
    }

    pub fn set_position(&mut self, pos: usize) {
        self.pos = pos;
    }

    pub fn source(&self) -> &'a str {
        self.src
    }

    fn find(&self, needle: &str, from: usize) -> XmlResult<usize> {
        self.src[from..]
            .find(needle)
            .map(|i| from + i)
            .ok_or_else(|| XmlError(format!("unterminated construct at byte {from}")))
    }

    /// The next token, or None at the end of input.
    pub fn next_tok(&mut self) -> XmlResult<Option<(Tok<'a>, Range<usize>)>> {
        let src = self.src;
        let start = self.pos;
        if start >= src.len() {
            return Ok(None);
        }
        let bytes = src.as_bytes();
        if bytes[start] != b'<' {
            let end = src[start..].find('<').map(|i| start + i).unwrap_or(src.len());
            self.pos = end;
            return Ok(Some((Tok::Text(&src[start..end]), start..end)));
        }
        let rest = &src[start..];
        let end = if rest.starts_with("<!--") {
            self.find("-->", start + 4)? + 3
        } else if rest.starts_with("<![CDATA[") {
            self.find("]]>", start + 9)? + 3
        } else if rest.starts_with("<?") {
            self.find("?>", start + 2)? + 2
        } else if rest.starts_with("<!") {
            self.find(">", start + 2)? + 1
        } else {
            // A tag: find the closing '>' outside quoted attribute values.
            let mut i = start + 1;
            let mut quote: Option<u8> = None;
            loop {
                if i >= bytes.len() {
                    return Err(XmlError(format!("unterminated tag at byte {start}")));
                }
                let b = bytes[i];
                match quote {
                    Some(q) if b == q => quote = None,
                    Some(_) => {}
                    None if b == b'"' || b == b'\'' => quote = Some(b),
                    None if b == b'>' => break,
                    None => {}
                }
                i += 1;
            }
            let end = i + 1;
            self.pos = end;
            let inner = &src[start + 1..i];
            if let Some(name) = inner.strip_prefix('/') {
                return Ok(Some((Tok::End { name: name.trim() }, start..end)));
            }
            let (inner, empty) = match inner.strip_suffix('/') {
                Some(stripped) => (stripped, true),
                None => (inner, false),
            };
            let name_end = inner
                .find(|c: char| c.is_ascii_whitespace())
                .unwrap_or(inner.len());
            let name = &inner[..name_end];
            let attrs = &inner[name_end..];
            let tok = if empty { Tok::Empty { name, attrs } } else { Tok::Start { name, attrs } };
            return Ok(Some((tok, start..end)));
        };
        self.pos = end;
        Ok(Some((Tok::Other(&src[start..end]), start..end)))
    }

    /// Skip to just after the end tag matching an already-consumed start tag.
    /// Returns the end offset.
    pub fn skip_element(&mut self, name: &str) -> XmlResult<usize> {
        let mut depth = 1usize;
        while let Some((tok, span)) = self.next_tok()? {
            match tok {
                Tok::Start { name: n, .. } if n == name => depth += 1,
                Tok::End { name: n } if n == name => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(span.end);
                    }
                }
                _ => {}
            }
        }
        Err(XmlError(format!("missing </{name}>")))
    }
}

/// The local part of a possibly prefixed name: "x:row" -> "row".
pub fn local(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

/// Iterate over raw `(name, value)` attribute pairs (values still escaped).
pub fn attrs(raw: &str) -> AttrIter<'_> {
    AttrIter { raw, pos: 0 }
}

pub struct AttrIter<'a> {
    raw: &'a str,
    pos: usize,
}

impl<'a> Iterator for AttrIter<'a> {
    type Item = (&'a str, &'a str);

    fn next(&mut self) -> Option<Self::Item> {
        let raw = self.raw;
        let rest = &raw[self.pos..];
        let offset = rest.find(|c: char| !c.is_ascii_whitespace())?;
        let rest = &rest[offset..];
        let eq = rest.find('=')?;
        let name = rest[..eq].trim();
        let after = rest[eq + 1..].trim_start();
        let quote = after.chars().next()?;
        if quote != '"' && quote != '\'' {
            return None;
        }
        let close = after[1..].find(quote)?;
        let value = &after[1..1 + close];
        let consumed = (after.as_ptr() as usize - raw.as_ptr() as usize) + 1 + close + 1;
        self.pos = consumed;
        Some((name, value))
    }
}

/// The unescaped value of an attribute, by local or full name.
pub fn attr(raw: &str, name: &str) -> Option<String> {
    attrs(raw).find(|(n, _)| *n == name || local(n) == name).map(|(_, v)| unescape(v))
}

/// Unescape XML entities (predefined and numeric).
pub fn unescape(text: &str) -> String {
    if !text.contains('&') {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(i) = rest.find('&') {
        out.push_str(&rest[..i]);
        let after = &rest[i + 1..];
        match after.find(';') {
            Some(j) => {
                let entity = &after[..j];
                let decoded = match entity {
                    "amp" => Some('&'),
                    "lt" => Some('<'),
                    "gt" => Some('>'),
                    "quot" => Some('"'),
                    "apos" => Some('\''),
                    _ if entity.starts_with("#x") || entity.starts_with("#X") => {
                        u32::from_str_radix(&entity[2..], 16).ok().and_then(char::from_u32)
                    }
                    _ if entity.starts_with('#') => entity[1..].parse().ok().and_then(char::from_u32),
                    _ => None,
                };
                match decoded {
                    Some(c) => out.push(c),
                    None => {
                        out.push('&');
                        out.push_str(entity);
                        out.push(';');
                    }
                }
                rest = &after[j + 1..];
            }
            None => {
                out.push('&');
                rest = after;
            }
        }
    }
    out.push_str(rest);
    out
}

/// Escape text content.
pub fn escape_text(text: &str) -> String {
    escape(text, false)
}

/// Escape an attribute value (also escapes quotes and line breaks).
pub fn escape_attr(text: &str) -> String {
    escape(text, true)
}

fn escape(text: &str, attribute: bool) -> String {
    let mut out = String::with_capacity(text.len());
    escape_into(&mut out, text, attribute);
    out
}

/// Append escaped text content to `out` (no intermediate string).
pub fn escape_text_into(out: &mut String, text: &str) {
    escape_into(out, text, false);
}

/// Append `text` to `out`, escaped as text content or as an attribute value.
fn escape_into(out: &mut String, text: &str, attribute: bool) {
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attribute => out.push_str("&quot;"),
            '\n' if attribute => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' if attribute => out.push_str("&#9;"),
            // Control characters are not allowed in XML 1.0: use Excel's _xHHHH_ form.
            c if (c as u32) < 0x20 && c != '\n' && c != '\t' => {
                let _ = std::fmt::Write::write_fmt(out, format_args!("_x{:04X}_", c as u32));
            }
            c => out.push(c),
        }
    }
}

/// A child of a [`Node`].
#[derive(Debug, Clone, PartialEq)]
pub enum Child {
    Element(Node),
    /// Raw (escaped) text.
    Text(String),
    /// Raw comment / PI / CDATA.
    Raw(String),
}

/// A DOM element. Attribute values and text are stored escaped, exactly as read.
#[derive(Debug, Clone, PartialEq)]
pub struct Node {
    pub name: String,
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Child>,
    /// Serialise as `<name/>` when there are no children.
    pub self_closing: bool,
}

/// A parsed document: the prolog (declaration etc.) and the root element.
#[derive(Debug, Clone, PartialEq)]
pub struct Document {
    pub prolog: String,
    pub root: Node,
    pub epilog: String,
}

impl Node {
    pub fn new(name: &str) -> Self {
        Node { name: name.to_string(), attrs: Vec::new(), children: Vec::new(), self_closing: true }
    }

    pub fn local(&self) -> &str {
        local(&self.name)
    }

    /// Unescaped attribute value by local or full name.
    pub fn get(&self, name: &str) -> Option<String> {
        self.attrs
            .iter()
            .find(|(n, _)| n == name || local(n) == name)
            .map(|(_, v)| unescape(v))
    }

    /// Set (or add) an attribute; `value` is unescaped.
    pub fn set(&mut self, name: &str, value: &str) {
        let escaped = escape_attr(value);
        match self.attrs.iter_mut().find(|(n, _)| n == name) {
            Some(slot) => slot.1 = escaped,
            None => self.attrs.push((name.to_string(), escaped)),
        }
    }

    pub fn remove_attr(&mut self, name: &str) {
        self.attrs.retain(|(n, _)| n != name && local(n) != name);
    }

    pub fn with_attr(mut self, name: &str, value: &str) -> Self {
        self.set(name, value);
        self
    }

    pub fn elements(&self) -> impl Iterator<Item = &Node> {
        self.children.iter().filter_map(|c| match c {
            Child::Element(n) => Some(n),
            _ => None,
        })
    }

    pub fn elements_mut(&mut self) -> impl Iterator<Item = &mut Node> {
        self.children.iter_mut().filter_map(|c| match c {
            Child::Element(n) => Some(n),
            _ => None,
        })
    }

    /// First child element with the given local name.
    pub fn child(&self, name: &str) -> Option<&Node> {
        self.elements().find(|n| n.local() == name)
    }

    pub fn child_mut(&mut self, name: &str) -> Option<&mut Node> {
        self.elements_mut().find(|n| n.local() == name)
    }

    pub fn children_named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Node> + 'a {
        self.elements().filter(move |n| n.local() == name)
    }

    /// Remove all child elements with the given local name.
    pub fn remove_children(&mut self, name: &str) {
        self.children.retain(|c| !matches!(c, Child::Element(n) if n.local() == name));
    }

    pub fn push(&mut self, node: Node) {
        self.self_closing = false;
        self.children.push(Child::Element(node));
    }

    /// Insert a child so that children stay in the schema `order` (local names).
    /// Unknown elements keep their positions.
    pub fn insert_ordered(&mut self, node: Node, order: &[&str]) {
        self.self_closing = false;
        let rank = order.iter().position(|n| *n == node.local()).unwrap_or(order.len());
        let index = self
            .children
            .iter()
            .position(|c| match c {
                Child::Element(n) => order.iter().position(|o| *o == n.local()).is_some_and(|r| r > rank),
                _ => false,
            })
            .unwrap_or(self.children.len());
        self.children.insert(index, Child::Element(node));
    }

    /// Unescaped concatenated text of this element (not descendants).
    pub fn text(&self) -> String {
        let mut out = String::new();
        for c in &self.children {
            if let Child::Text(t) = c {
                out.push_str(&unescape(t));
            }
        }
        out
    }

    pub fn set_text(&mut self, text: &str) {
        self.children.retain(|c| !matches!(c, Child::Text(_)));
        self.children.push(Child::Text(escape_text(text)));
        self.self_closing = false;
    }

    /// Prefix (with colon) used by this element, e.g. "x:" or "".
    pub fn prefix(&self) -> &str {
        match self.name.rfind(':') {
            Some(i) => &self.name[..=i],
            None => "",
        }
    }

    pub fn write(&self, out: &mut String) {
        out.push('<');
        out.push_str(&self.name);
        for (k, v) in &self.attrs {
            out.push(' ');
            out.push_str(k);
            out.push_str("=\"");
            // Values are stored escaped; re-escape a bare quote if one slipped in.
            out.push_str(&v.replace('"', "&quot;"));
            out.push('"');
        }
        if self.children.is_empty() && self.self_closing {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for c in &self.children {
            match c {
                Child::Element(n) => n.write(out),
                Child::Text(t) | Child::Raw(t) => out.push_str(t),
            }
        }
        out.push_str("</");
        out.push_str(&self.name);
        out.push('>');
    }

    pub fn to_xml(&self) -> String {
        let mut out = String::new();
        self.write(&mut out);
        out
    }
}

fn parse_attrs(raw: &str) -> Vec<(String, String)> {
    attrs(raw).map(|(k, v)| (k.to_string(), v.to_string())).collect()
}

/// Parse an element subtree that starts with the token just read.
fn parse_node(scanner: &mut Scanner<'_>, name: &str, raw_attrs: &str) -> XmlResult<Node> {
    let mut node = Node {
        name: name.to_string(),
        attrs: parse_attrs(raw_attrs),
        children: Vec::new(),
        self_closing: false,
    };
    while let Some((tok, _)) = scanner.next_tok()? {
        match tok {
            Tok::Start { name, attrs } => node.children.push(Child::Element(parse_node(scanner, name, attrs)?)),
            Tok::Empty { name, attrs } => node.children.push(Child::Element(Node {
                name: name.to_string(),
                attrs: parse_attrs(attrs),
                children: Vec::new(),
                self_closing: true,
            })),
            Tok::End { name: end } => {
                if end != node.name {
                    return Err(XmlError(format!("expected </{}>, found </{end}>", node.name)));
                }
                return Ok(node);
            }
            Tok::Text(t) => node.children.push(Child::Text(t.to_string())),
            Tok::Other(t) => node.children.push(Child::Raw(t.to_string())),
        }
    }
    Err(XmlError(format!("missing </{}>", node.name)))
}

/// Parse a whole XML document.
pub fn parse_document(src: &str) -> XmlResult<Document> {
    let src = src.strip_prefix('\u{feff}').unwrap_or(src);
    let mut scanner = Scanner::new(src);
    let mut prolog = String::new();
    while let Some((tok, span)) = scanner.next_tok()? {
        match tok {
            Tok::Start { name, attrs } => {
                let root = parse_node(&mut scanner, name, attrs)?;
                let epilog = src[scanner.position()..].to_string();
                return Ok(Document { prolog, root, epilog });
            }
            Tok::Empty { name, attrs } => {
                let root = Node { name: name.to_string(), attrs: parse_attrs(attrs), children: Vec::new(), self_closing: true };
                let epilog = src[scanner.position()..].to_string();
                return Ok(Document { prolog, root, epilog });
            }
            _ => prolog.push_str(&src[span]),
        }
    }
    Err(XmlError("document has no root element".into()))
}

impl Document {
    pub fn to_xml(&self) -> String {
        let mut out = String::with_capacity(4096);
        out.push_str(&self.prolog);
        self.root.write(&mut out);
        out.push_str(&self.epilog);
        out
    }
}

pub const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_exactly() {
        let src = "<?xml version=\"1.0\"?>\r\n<a xmlns:x=\"urn:x\" k='v &amp; w'><!-- c --><x:b c=\"1\"/>t&lt;x<c>&#65;</c></a>";
        let doc = parse_document(src).unwrap();
        assert_eq!(doc.to_xml(), src.replace("k='v &amp; w'", "k=\"v &amp; w\""));
        assert_eq!(doc.root.get("k").unwrap(), "v & w");
        assert_eq!(doc.root.child("b").unwrap().get("c").unwrap(), "1");
        assert_eq!(doc.root.child("c").unwrap().text(), "A");
    }

    #[test]
    fn scanner_spans() {
        let src = "<r><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\"/></r>";
        let mut s = Scanner::new(src);
        let mut names = vec![];
        while let Some((tok, span)) = s.next_tok().unwrap() {
            if let Tok::Start { name, .. } | Tok::Empty { name, .. } = tok {
                names.push((name, &src[span]));
            }
        }
        assert_eq!(names[1], ("c", "<c r=\"A1\" t=\"s\">"));
        assert_eq!(names[3], ("c", "<c r=\"B1\"/>"));
    }

    #[test]
    fn escaping() {
        assert_eq!(escape_attr("a\"<&\n"), "a&quot;&lt;&amp;&#10;");
        assert_eq!(escape_text("x\u{1}y"), "x_x0001_y");
        assert_eq!(unescape("&lt;&#x41;&#66;&unknown;"), "<AB&unknown;");
        assert_eq!(attr(" a=\"1\" x:b='2'", "b").as_deref(), Some("2"));
    }

    #[test]
    fn insert_ordered_respects_schema() {
        let mut n = parse_document("<w><a/><c/><zz/></w>").unwrap().root;
        n.insert_ordered(Node::new("b"), &["a", "b", "c"]);
        assert_eq!(n.to_xml(), "<w><a/><b/><c/><zz/></w>");
    }
}
