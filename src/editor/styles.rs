//! Editing `styles.xml`.
//!
//! A cell's style is an index into `cellXfs`. Each `xf` points at a font, a
//! fill, a border and a number format, and may contain alignment and
//! protection. To apply a `Format` the editor starts from the cell's current
//! `xf` (or the default `xf 0`) and patches only the components the format
//! sets, keeping everything else (theme tints, font effects, gradient fills,
//! named cell styles) untouched. New fonts/fills/borders/xfs are
//! de-duplicated against existing ones, so repeated edits don't bloat the file.

use std::collections::HashMap;

use rust_xlsxwriter::Color as RColor;

use crate::common::color::Color;
use crate::common::format::{
    FontScheme, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern, FormatScript,
    FormatSpec, FormatUnderline,
};
use crate::editor::package::EResult;
use crate::editor::xml::{Child, Document, Node};

const FONT_ORDER: &[&str] = &[
    "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color",
    "name", "family", "charset", "scheme",
];
const BORDER_ORDER: &[&str] = &["start", "left", "end", "right", "top", "bottom", "diagonal", "vertical", "horizontal"];
const XF_CHILD_ORDER: &[&str] = &["alignment", "protection", "extLst"];
const STYLESHEET_ORDER: &[&str] = &[
    "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
    "tableStyles", "colors", "extLst",
];

/// Built-in number formats whose meaning doesn't depend on the locale.
const BUILTIN_NUM_FORMATS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mm:ss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// Codes of the remaining built-in formats (for reading styles back).
const BUILTIN_OTHER: &[(u32, &str)] = &[
    (5, "($#,##0_);($#,##0)"),
    (6, "($#,##0_);[Red]($#,##0)"),
    (7, "($#,##0.00_);($#,##0.00)"),
    (8, "($#,##0.00_);[Red]($#,##0.00)"),
    (14, "m/d/yyyy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yyyy h:mm"),
    (37, "#,##0 ;(#,##0)"),
    (38, "#,##0 ;[Red](#,##0)"),
    (39, "#,##0.00;(#,##0.00)"),
    (40, "#,##0.00;[Red](#,##0.00)"),
    (41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
    (42, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"),
    (43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
    (44, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
];

/// Excel's tint values for `Color.theme(color, shade)` (same table as rust_xlsxwriter).
fn theme_tint(color: u8, shade: u8) -> Option<&'static str> {
    let table: [&str; 5] = match color {
        0 => ["-4.9989318521683403E-2", "-0.14999847407452621", "-0.249977111117893", "-0.34998626667073579", "-0.499984740745262"],
        1 => ["0.499984740745262", "0.34998626667073579", "0.249977111117893", "0.14999847407452621", "4.9989318521683403E-2"],
        2 => ["-9.9978637043366805E-2", "-0.249977111117893", "-0.499984740745262", "-0.749992370372631", "-0.89999084444715716"],
        _ => ["0.79998168889431442", "0.59999389629810485", "0.39997558519241921", "-0.249977111117893", "-0.499984740745262"],
    };
    if shade == 0 { None } else { table.get(shade as usize - 1).copied() }
}

fn theme_shade(color: u8, tint: f64) -> u8 {
    (1..=5u8)
        .find(|&shade| {
            theme_tint(color, shade).and_then(|t| t.parse::<f64>().ok()).is_some_and(|t| (t - tint).abs() < 1e-3)
        })
        .unwrap_or(0)
}

/// The legacy indexed palette (used by older files).
const INDEXED: [u32; 64] = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000,
    0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000,
    0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC,
    0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF,
    0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF,
    0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600,
    0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399,
    0x333333,
];

/// A `<color>`-like element (`name` = "color", "fgColor", ...) for a colour.
fn color_node(name: &str, color: &Color) -> Option<Node> {
    let node = Node::new(name);
    match color.0 {
        RColor::Default => None,
        RColor::Automatic => Some(node.with_attr("auto", "1")),
        RColor::Theme(c, shade) => {
            let node = node.with_attr("theme", &c.to_string());
            Some(match theme_tint(c, shade) {
                Some(tint) => node.with_attr("tint", tint),
                None => node,
            })
        }
        _ => color.rgb_value().map(|v| node.with_attr("rgb", &format!("FF{v:06X}"))),
    }
}

/// Read a colour element back into a `Color`.
pub fn read_color(node: &Node) -> Option<Color> {
    if node.get("auto").as_deref() == Some("1") {
        return Some(Color(RColor::Automatic));
    }
    if let Some(rgb) = node.get("rgb") {
        let hex = if rgb.len() == 8 { &rgb[2..] } else { rgb.as_str() };
        return u32::from_str_radix(hex, 16).ok().map(|v| Color(RColor::RGB(v)));
    }
    if let Some(theme) = node.get("theme").and_then(|t| t.parse::<u8>().ok()) {
        if theme > 9 {
            return None;
        }
        let tint = node.get("tint").and_then(|t| t.parse::<f64>().ok()).unwrap_or(0.0);
        return Some(Color(RColor::Theme(theme, theme_shade(theme, tint))));
    }
    if let Some(index) = node.get("indexed").and_then(|i| i.parse::<usize>().ok()) {
        return INDEXED.get(index).map(|v| Color(RColor::RGB(*v)));
    }
    None
}

fn border_style_name(border: FormatBorder) -> &'static str {
    match border {
        FormatBorder::None => "none",
        FormatBorder::Thin => "thin",
        FormatBorder::Medium => "medium",
        FormatBorder::Dashed => "dashed",
        FormatBorder::Dotted => "dotted",
        FormatBorder::Thick => "thick",
        FormatBorder::Double => "double",
        FormatBorder::Hair => "hair",
        FormatBorder::MediumDashed => "mediumDashed",
        FormatBorder::DashDot => "dashDot",
        FormatBorder::MediumDashDot => "mediumDashDot",
        FormatBorder::DashDotDot => "dashDotDot",
        FormatBorder::MediumDashDotDot => "mediumDashDotDot",
        FormatBorder::SlantDashDot => "slantDashDot",
    }
}

fn border_from_name(name: &str) -> Option<FormatBorder> {
    use FormatBorder::*;
    [None, Thin, Medium, Dashed, Dotted, Thick, Double, Hair, MediumDashed, DashDot, MediumDashDot, DashDotDot, MediumDashDotDot, SlantDashDot]
        .into_iter()
        .find(|b| border_style_name(*b) == name)
}

fn pattern_name(pattern: FormatPattern) -> &'static str {
    match pattern {
        FormatPattern::None => "none",
        FormatPattern::Solid => "solid",
        FormatPattern::MediumGray => "mediumGray",
        FormatPattern::DarkGray => "darkGray",
        FormatPattern::LightGray => "lightGray",
        FormatPattern::DarkHorizontal => "darkHorizontal",
        FormatPattern::DarkVertical => "darkVertical",
        FormatPattern::DarkDown => "darkDown",
        FormatPattern::DarkUp => "darkUp",
        FormatPattern::DarkGrid => "darkGrid",
        FormatPattern::DarkTrellis => "darkTrellis",
        FormatPattern::LightHorizontal => "lightHorizontal",
        FormatPattern::LightVertical => "lightVertical",
        FormatPattern::LightDown => "lightDown",
        FormatPattern::LightUp => "lightUp",
        FormatPattern::LightGrid => "lightGrid",
        FormatPattern::LightTrellis => "lightTrellis",
        FormatPattern::Gray125 => "gray125",
        FormatPattern::Gray0625 => "gray0625",
    }
}

fn pattern_from_name(name: &str) -> Option<FormatPattern> {
    use FormatPattern::*;
    [None, Solid, MediumGray, DarkGray, LightGray, DarkHorizontal, DarkVertical, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightHorizontal, LightVertical, LightDown, LightUp, LightGrid, LightTrellis, Gray125, Gray0625]
        .into_iter()
        .find(|p| pattern_name(*p) == name)
}

fn align_name(align: FormatAlign) -> &'static str {
    match align {
        FormatAlign::General => "general",
        FormatAlign::Left => "left",
        FormatAlign::Center | FormatAlign::VerticalCenter => "center",
        FormatAlign::Right => "right",
        FormatAlign::Fill => "fill",
        FormatAlign::Justify | FormatAlign::VerticalJustify => "justify",
        FormatAlign::CenterAcross => "centerContinuous",
        FormatAlign::Distributed | FormatAlign::VerticalDistributed => "distributed",
        FormatAlign::Top => "top",
        FormatAlign::Bottom => "bottom",
    }
}

fn underline_name(u: FormatUnderline) -> Option<&'static str> {
    match u {
        FormatUnderline::None => None,
        FormatUnderline::Single => Some("single"),
        FormatUnderline::Double => Some("double"),
        FormatUnderline::SingleAccounting => Some("singleAccounting"),
        FormatUnderline::DoubleAccounting => Some("doubleAccounting"),
    }
}

/// Is a number format a date/time format?
pub fn is_date_format(id: u32, code: &str) -> bool {
    if matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58) {
        return true;
    }
    if id < 164 && code.is_empty() {
        return false;
    }
    let mut in_quote = false;
    let mut bracket: Option<String> = None;
    let mut chars = code.chars().peekable();
    // Only the first section matters for positive values.
    while let Some(c) = chars.next() {
        if let Some(content) = bracket.as_mut() {
            if c == ']' {
                // [h], [mm], [ss] are elapsed-time codes; [Red], [$-409] are not.
                let lower = content.to_ascii_lowercase();
                if !lower.is_empty() && lower.chars().all(|ch| matches!(ch, 'h' | 'm' | 's')) {
                    return true;
                }
                bracket = None;
            } else {
                content.push(c);
            }
            continue;
        }
        match c {
            '"' => in_quote = !in_quote,
            _ if in_quote => {}
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => bracket = Some(String::new()),
            ';' => break,
            'd' | 'D' | 'm' | 'M' | 'y' | 'Y' | 'h' | 'H' | 's' | 'S' => return true,
            _ => {}
        }
    }
    false
}

#[derive(Clone)]
pub struct Styles {
    pub part: String,
    doc: Document,
    pub dirty: bool,
    font_index: HashMap<String, u32>,
    fill_index: HashMap<String, u32>,
    border_index: HashMap<String, u32>,
    xf_index: HashMap<String, u32>,
    num_fmt_index: HashMap<String, u32>,
    merge_cache: HashMap<(u32, String), u32>,
    pub date_xf_cache: HashMap<(u32, &'static str), u32>,
}

fn ensure_container<'a>(root: &'a mut Node, name: &str) -> &'a mut Node {
    if root.child(name).is_none() {
        let mut node = Node::new(name);
        node.set("count", "0");
        root.insert_ordered(node, STYLESHEET_ORDER);
    }
    root.child_mut(name).expect("container just ensured")
}

fn canonical(node: &Node) -> String {
    node.to_xml()
}

impl Styles {
    pub fn new(part: &str, doc: Document) -> Self {
        let mut styles = Styles {
            part: part.to_string(),
            doc,
            dirty: false,
            font_index: HashMap::new(),
            fill_index: HashMap::new(),
            border_index: HashMap::new(),
            xf_index: HashMap::new(),
            num_fmt_index: HashMap::new(),
            merge_cache: HashMap::new(),
            date_xf_cache: HashMap::new(),
        };
        for (container, child, which) in [("fonts", "font", 0), ("fills", "fill", 1), ("borders", "border", 2), ("cellXfs", "xf", 3)] {
            let entries: Vec<String> = styles
                .doc
                .root
                .child(container)
                .map(|c| c.children_named(child).map(canonical).collect())
                .unwrap_or_default();
            let map = match which {
                0 => &mut styles.font_index,
                1 => &mut styles.fill_index,
                2 => &mut styles.border_index,
                _ => &mut styles.xf_index,
            };
            for (i, xml) in entries.into_iter().enumerate() {
                map.entry(xml).or_insert(i as u32);
            }
        }
        let custom: Vec<(String, u32)> = styles
            .doc
            .root
            .child("numFmts")
            .map(|c| {
                c.children_named("numFmt")
                    .filter_map(|n| Some((n.get("formatCode")?, n.get("numFmtId")?.parse().ok()?)))
                    .collect()
            })
            .unwrap_or_default();
        for (code, id) in custom {
            styles.num_fmt_index.entry(code).or_insert(id);
        }
        styles
    }

    /// A minimal stylesheet for packages that lack one.
    pub fn minimal_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>"
    }

    pub fn to_xml(&self) -> String {
        self.doc.to_xml()
    }

    fn item(&self, container: &str, child: &str, index: u32) -> Option<Node> {
        self.doc.root.child(container)?.children_named(child).nth(index as usize).cloned()
    }

    pub fn xf_count(&self) -> u32 {
        self.doc.root.child("cellXfs").map(|c| c.children_named("xf").count() as u32).unwrap_or(0)
    }

    fn xf(&self, index: u32) -> Node {
        self.item("cellXfs", "xf", index)
            .or_else(|| self.item("cellXfs", "xf", 0))
            .unwrap_or_else(|| Node::new("xf").with_attr("numFmtId", "0").with_attr("fontId", "0").with_attr("fillId", "0").with_attr("borderId", "0").with_attr("xfId", "0"))
    }

    fn attr_u32(node: &Node, name: &str) -> u32 {
        node.get(name).and_then(|v| v.parse().ok()).unwrap_or(0)
    }

    /// Append (or find) an entry in a container, returning its index.
    fn intern(&mut self, container: &str, child: &str, node: Node) -> u32 {
        let key = canonical(&node);
        let map = match container {
            "fonts" => &self.font_index,
            "fills" => &self.fill_index,
            "borders" => &self.border_index,
            _ => &self.xf_index,
        };
        if let Some(index) = map.get(&key) {
            return *index;
        }
        let parent = ensure_container(&mut self.doc.root, container);
        let index = parent.children_named(child).count() as u32;
        parent.push(node);
        parent.set("count", &(index + 1).to_string());
        self.dirty = true;
        let map = match container {
            "fonts" => &mut self.font_index,
            "fills" => &mut self.fill_index,
            "borders" => &mut self.border_index,
            _ => &mut self.xf_index,
        };
        map.insert(key, index);
        index
    }

    /// Append a differential (`dxf`) style, used by conditional formats.
    pub fn add_dxf(&mut self, dxf: Node) -> u32 {
        let dxfs = ensure_container(&mut self.doc.root, "dxfs");
        let index = dxfs.children_named("dxf").count() as u32;
        dxfs.push(dxf);
        dxfs.set("count", &(index + 1).to_string());
        self.dirty = true;
        index
    }

    /// The numFmtId for a format code (built-in or custom, adding it if new).
    fn num_fmt_id(&mut self, code: &str) -> u32 {
        if let Some((id, _)) = BUILTIN_NUM_FORMATS.iter().find(|(_, c)| *c == code) {
            return *id;
        }
        if let Some(id) = self.num_fmt_index.get(code) {
            return *id;
        }
        let id = self.num_fmt_index.values().copied().max().map_or(164, |m| m.max(163) + 1);
        let numfmts = ensure_container(&mut self.doc.root, "numFmts");
        numfmts.push(Node::new("numFmt").with_attr("numFmtId", &id.to_string()).with_attr("formatCode", code));
        let count = numfmts.children_named("numFmt").count();
        numfmts.set("count", &count.to_string());
        self.num_fmt_index.insert(code.to_string(), id);
        self.dirty = true;
        id
    }

    /// The format code of a numFmtId.
    pub fn num_format_code(&self, id: u32) -> String {
        if let Some((_, code)) = BUILTIN_NUM_FORMATS.iter().chain(BUILTIN_OTHER).find(|(i, _)| *i == id) {
            return code.to_string();
        }
        self.num_fmt_index
            .iter()
            .find(|(_, i)| **i == id)
            .map(|(code, _)| code.clone())
            .unwrap_or_default()
    }

    /// Whether a cell style shows numbers as dates/times.
    pub fn is_date_xf(&self, xf: u32) -> bool {
        let id = Self::attr_u32(&self.xf(xf), "numFmtId");
        is_date_format(id, &self.num_format_code(id))
    }

    /// Return the xf index for `base` with `overlay`'s set properties applied.
    pub fn merge(&mut self, base: u32, overlay: &FormatSpec) -> EResult<u32> {
        let key = (base, format!("{overlay:?}"));
        if let Some(index) = self.merge_cache.get(&key) {
            return Ok(*index);
        }
        let index = self.build_merged(base, overlay)?;
        self.merge_cache.insert(key, index);
        Ok(index)
    }

    /// Return the xf index for exactly `spec` (on top of the default style).
    pub fn apply(&mut self, spec: &FormatSpec) -> EResult<u32> {
        self.merge(0, spec)
    }

    fn build_merged(&mut self, base: u32, o: &FormatSpec) -> EResult<u32> {
        if o.checkbox {
            return Err("checkbox formats are not supported when editing existing files".into());
        }
        let d = FormatSpec::default();
        let mut xf = self.xf(base);

        // Number format.
        if let Some(code) = &o.num_format {
            let id = self.num_fmt_id(code);
            xf.set("numFmtId", &id.to_string());
            xf.set("applyNumberFormat", "1");
        } else if let Some(id) = o.num_format_index {
            xf.set("numFmtId", &id.to_string());
            xf.set("applyNumberFormat", "1");
        }

        // Font.
        let font_changed = o.bold || o.italic || o.strikethrough || o.font_name.is_some() || o.font_size.is_some()
            || o.font_color.is_some() || o.font_family.is_some() || o.font_charset.is_some()
            || o.font_scheme.is_some() || o.font_script.is_some() || o.underline.is_some() || o.hyperlink;
        if font_changed {
            let mut font = self.item("fonts", "font", Self::attr_u32(&xf, "fontId")).unwrap_or_else(|| Node::new("font"));
            font.self_closing = false;
            let put = |font: &mut Node, node: Node| {
                font.remove_children(node.local());
                font.insert_ordered(node, FONT_ORDER);
            };
            if o.bold {
                put(&mut font, Node::new("b"));
            }
            if o.italic {
                put(&mut font, Node::new("i"));
            }
            if o.strikethrough {
                put(&mut font, Node::new("strike"));
            }
            let underline = o.underline.or(if o.hyperlink { Some(FormatUnderline::Single) } else { None });
            if let Some(u) = underline {
                font.remove_children("u");
                match underline_name(u) {
                    Some("single") => font.insert_ordered(Node::new("u"), FONT_ORDER),
                    Some(name) => font.insert_ordered(Node::new("u").with_attr("val", name), FONT_ORDER),
                    None => {}
                }
            }
            if let Some(script) = o.font_script {
                font.remove_children("vertAlign");
                match script {
                    FormatScript::Superscript => font.insert_ordered(Node::new("vertAlign").with_attr("val", "superscript"), FONT_ORDER),
                    FormatScript::Subscript => font.insert_ordered(Node::new("vertAlign").with_attr("val", "subscript"), FONT_ORDER),
                    FormatScript::None => {}
                }
            }
            if let Some(size) = o.font_size {
                put(&mut font, Node::new("sz").with_attr("val", &format_number(size)));
            }
            let color = o.font_color.or(if o.hyperlink { Some(Color(RColor::Theme(10, 0))) } else { None });
            if let Some(color) = color {
                font.remove_children("color");
                if let Some(node) = color_node("color", &color) {
                    font.insert_ordered(node, FONT_ORDER);
                }
            }
            if let Some(name) = &o.font_name {
                put(&mut font, Node::new("name").with_attr("val", name));
                // A theme scheme overrides the name, so drop it unless one was asked for.
                if o.font_scheme.is_none() {
                    font.remove_children("scheme");
                }
            }
            if let Some(family) = o.font_family {
                put(&mut font, Node::new("family").with_attr("val", &family.to_string()));
            }
            if let Some(charset) = o.font_charset {
                put(&mut font, Node::new("charset").with_attr("val", &charset.to_string()));
            }
            if let Some(scheme) = o.font_scheme {
                let value = match scheme {
                    FontScheme::Body => "minor",
                    FontScheme::Headings => "major",
                    FontScheme::None => "none",
                };
                put(&mut font, Node::new("scheme").with_attr("val", value));
            }
            let id = self.intern("fonts", "font", font);
            xf.set("fontId", &id.to_string());
            xf.set("applyFont", "1");
        }

        // Fill.
        if o.pattern.is_some() || o.background_color.is_some() || o.foreground_color.is_some() {
            let base_fill = self.item("fills", "fill", Self::attr_u32(&xf, "fillId"));
            let base_pattern = base_fill.as_ref().and_then(|f| f.child("patternFill").cloned());
            let pattern = o.pattern.unwrap_or_else(|| {
                base_pattern
                    .as_ref()
                    .and_then(|p| p.get("patternType"))
                    .and_then(|p| pattern_from_name(&p))
                    .filter(|p| *p != FormatPattern::None)
                    .unwrap_or(FormatPattern::Solid)
            });
            let mut pattern_fill = Node::new("patternFill").with_attr("patternType", pattern_name(pattern));
            // Excel stores a solid fill's colour in fgColor.
            let (fg, bg) = if pattern == FormatPattern::Solid {
                (o.foreground_color.or(o.background_color), None)
            } else {
                (o.foreground_color, o.background_color)
            };
            let fg = fg.or_else(|| base_pattern.as_ref().and_then(|p| p.child("fgColor")).and_then(read_color));
            if let Some(node) = fg.as_ref().and_then(|c| color_node("fgColor", c)) {
                pattern_fill.push(node);
            }
            match bg.as_ref().and_then(|c| color_node("bgColor", c)) {
                Some(node) => pattern_fill.push(node),
                None if pattern == FormatPattern::Solid => {
                    pattern_fill.push(Node::new("bgColor").with_attr("indexed", "64"));
                }
                None => {
                    if let Some(old) = base_pattern.as_ref().and_then(|p| p.child("bgColor")) {
                        pattern_fill.push(old.clone());
                    }
                }
            }
            let mut fill = Node::new("fill");
            fill.push(pattern_fill);
            let id = self.intern("fills", "fill", fill);
            xf.set("fillId", &id.to_string());
            xf.set("applyFill", "1");
        }

        // Borders.
        let sides: [(&str, Option<FormatBorder>, Option<Color>); 4] = [
            ("left", o.border_left, o.border_left_color),
            ("right", o.border_right, o.border_right_color),
            ("top", o.border_top, o.border_top_color),
            ("bottom", o.border_bottom, o.border_bottom_color),
        ];
        let border_changed = sides.iter().any(|(_, s, c)| s.is_some() || c.is_some())
            || o.border_diagonal.is_some() || o.border_diagonal_color.is_some() || o.border_diagonal_type.is_some();
        if border_changed {
            let mut border = self.item("borders", "border", Self::attr_u32(&xf, "borderId")).unwrap_or_else(|| Node::new("border"));
            border.self_closing = false;
            let all_sides = sides
                .into_iter()
                .chain(std::iter::once(("diagonal", o.border_diagonal, o.border_diagonal_color)));
            for (side, style, color) in all_sides {
                if style.is_none() && color.is_none() {
                    if border.child(side).is_none() {
                        border.insert_ordered(Node::new(side), BORDER_ORDER);
                    }
                    continue;
                }
                let mut node = border.child(side).cloned().unwrap_or_else(|| Node::new(side));
                if let Some(style) = style {
                    if style == FormatBorder::None {
                        node = Node::new(side);
                    } else {
                        node.set("style", border_style_name(style));
                    }
                }
                if let Some(color) = color {
                    node.remove_children("color");
                    if let Some(c) = color_node("color", &color) {
                        node.push(c);
                    }
                }
                border.remove_children(side);
                border.insert_ordered(node, BORDER_ORDER);
            }
            if let Some(kind) = o.border_diagonal_type {
                border.remove_attr("diagonalUp");
                border.remove_attr("diagonalDown");
                match kind {
                    FormatDiagonalBorder::BorderUp => border.set("diagonalUp", "1"),
                    FormatDiagonalBorder::BorderDown => border.set("diagonalDown", "1"),
                    FormatDiagonalBorder::BorderUpDown => {
                        border.set("diagonalUp", "1");
                        border.set("diagonalDown", "1");
                    }
                    FormatDiagonalBorder::None => {}
                }
            }
            let id = self.intern("borders", "border", border);
            xf.set("borderId", &id.to_string());
            xf.set("applyBorder", "1");
        }

        // Alignment.
        if o.align.is_some() || o.valign.is_some() || o.text_wrap || o.shrink || o.indent != d.indent
            || o.rotation != d.rotation || o.reading_direction.is_some()
        {
            let mut alignment = xf.child("alignment").cloned().unwrap_or_else(|| Node::new("alignment"));
            if let Some(a) = o.align {
                alignment.set("horizontal", align_name(a));
            }
            if let Some(v) = o.valign {
                alignment.set("vertical", align_name(v));
            }
            if o.text_wrap {
                alignment.set("wrapText", "1");
            }
            if o.shrink {
                alignment.set("shrinkToFit", "1");
            }
            if o.indent != 0 {
                alignment.set("indent", &o.indent.to_string());
            }
            if o.rotation != 0 {
                let rotation = match o.rotation {
                    270 => 255,
                    r if r < 0 => 90 - r as i32,
                    r => r as i32,
                };
                alignment.set("textRotation", &rotation.to_string());
            }
            if let Some(direction) = o.reading_direction {
                alignment.set("readingOrder", &direction.to_string());
            }
            xf.remove_children("alignment");
            xf.insert_ordered(alignment, XF_CHILD_ORDER);
            xf.set("applyAlignment", "1");
        }

        // Protection.
        if !o.locked || o.hidden {
            let mut protection = xf.child("protection").cloned().unwrap_or_else(|| Node::new("protection"));
            if !o.locked {
                protection.set("locked", "0");
            }
            if o.hidden {
                protection.set("hidden", "1");
            }
            xf.remove_children("protection");
            xf.insert_ordered(protection, XF_CHILD_ORDER);
            xf.set("applyProtection", "1");
        }

        if o.quote_prefix {
            xf.set("quotePrefix", "1");
        }
        if xf.children.iter().all(|c| !matches!(c, Child::Element(_))) {
            xf.children.clear();
            xf.self_closing = true;
        }
        Ok(self.intern("cellXfs", "xf", xf))
    }

    /// Describe a cell style as a `FormatSpec` (for reading styles back).
    pub fn spec(&self, index: u32) -> FormatSpec {
        let xf = self.xf(index);
        let mut s = FormatSpec::default();
        let num_fmt = Self::attr_u32(&xf, "numFmtId");
        if num_fmt != 0 {
            let code = self.num_format_code(num_fmt);
            if code.is_empty() {
                s.num_format_index = u8::try_from(num_fmt).ok();
            } else {
                s.num_format = Some(code);
            }
        }
        if let Some(font) = self.item("fonts", "font", Self::attr_u32(&xf, "fontId")) {
            let flag = |name: &str| font.child(name).is_some_and(|n| n.get("val").as_deref() != Some("0") && n.get("val").as_deref() != Some("false"));
            s.bold = flag("b");
            s.italic = flag("i");
            s.strikethrough = flag("strike");
            s.font_name = font.child("name").and_then(|n| n.get("val"));
            s.font_size = font.child("sz").and_then(|n| n.get("val")).and_then(|v| v.parse().ok());
            s.font_color = font.child("color").and_then(read_color);
            s.font_family = font.child("family").and_then(|n| n.get("val")).and_then(|v| v.parse().ok());
            s.font_charset = font.child("charset").and_then(|n| n.get("val")).and_then(|v| v.parse().ok());
            s.font_scheme = font.child("scheme").and_then(|n| n.get("val")).map(|v| match v.as_str() {
                "minor" => FontScheme::Body,
                "major" => FontScheme::Headings,
                _ => FontScheme::None,
            });
            s.underline = font.child("u").map(|n| match n.get("val").as_deref() {
                Some("double") => FormatUnderline::Double,
                Some("singleAccounting") => FormatUnderline::SingleAccounting,
                Some("doubleAccounting") => FormatUnderline::DoubleAccounting,
                Some("none") => FormatUnderline::None,
                _ => FormatUnderline::Single,
            });
            s.font_script = font.child("vertAlign").and_then(|n| n.get("val")).and_then(|v| match v.as_str() {
                "superscript" => Some(FormatScript::Superscript),
                "subscript" => Some(FormatScript::Subscript),
                _ => None,
            });
        }
        if let Some(fill) = self.item("fills", "fill", Self::attr_u32(&xf, "fillId"))
            && let Some(p) = fill.child("patternFill") {
                let pattern = p.get("patternType").and_then(|n| pattern_from_name(&n));
                if pattern.is_some_and(|p| p != FormatPattern::None) {
                    s.pattern = pattern;
                    let fg = p.child("fgColor").and_then(read_color);
                    let bg = p.child("bgColor").and_then(read_color);
                    if pattern == Some(FormatPattern::Solid) {
                        s.background_color = fg;
                    } else {
                        s.foreground_color = fg;
                        s.background_color = bg;
                    }
                }
            }
        if let Some(border) = self.item("borders", "border", Self::attr_u32(&xf, "borderId")) {
            let side = |name: &str| -> (Option<FormatBorder>, Option<Color>) {
                match border.child(name) {
                    Some(n) => (
                        n.get("style").and_then(|v| border_from_name(&v)).filter(|b| *b != FormatBorder::None),
                        n.child("color").and_then(read_color),
                    ),
                    None => (None, None),
                }
            };
            (s.border_left, s.border_left_color) = side("left");
            (s.border_right, s.border_right_color) = side("right");
            (s.border_top, s.border_top_color) = side("top");
            (s.border_bottom, s.border_bottom_color) = side("bottom");
            (s.border_diagonal, s.border_diagonal_color) = side("diagonal");
            let up = border.get("diagonalUp").as_deref() == Some("1");
            let down = border.get("diagonalDown").as_deref() == Some("1");
            s.border_diagonal_type = match (up, down) {
                (true, true) => Some(FormatDiagonalBorder::BorderUpDown),
                (true, false) => Some(FormatDiagonalBorder::BorderUp),
                (false, true) => Some(FormatDiagonalBorder::BorderDown),
                _ => None,
            };
        }
        if let Some(a) = xf.child("alignment") {
            s.align = a.get("horizontal").and_then(|v| match v.as_str() {
                "general" => Some(FormatAlign::General),
                "left" => Some(FormatAlign::Left),
                "center" => Some(FormatAlign::Center),
                "right" => Some(FormatAlign::Right),
                "fill" => Some(FormatAlign::Fill),
                "justify" => Some(FormatAlign::Justify),
                "centerContinuous" => Some(FormatAlign::CenterAcross),
                "distributed" => Some(FormatAlign::Distributed),
                _ => None,
            });
            s.valign = a.get("vertical").and_then(|v| match v.as_str() {
                "top" => Some(FormatAlign::Top),
                "bottom" => Some(FormatAlign::Bottom),
                "center" => Some(FormatAlign::VerticalCenter),
                "justify" => Some(FormatAlign::VerticalJustify),
                "distributed" => Some(FormatAlign::VerticalDistributed),
                _ => None,
            });
            s.text_wrap = a.get("wrapText").as_deref() == Some("1");
            s.shrink = a.get("shrinkToFit").as_deref() == Some("1");
            s.indent = a.get("indent").and_then(|v| v.parse().ok()).unwrap_or(0);
            s.rotation = a
                .get("textRotation")
                .and_then(|v| v.parse::<i32>().ok())
                .map(|r| match r {
                    255 => 270,
                    r if r > 90 => (90 - r) as i16,
                    r => r as i16,
                })
                .unwrap_or(0);
            s.reading_direction = a.get("readingOrder").and_then(|v| v.parse().ok());
        }
        if let Some(p) = xf.child("protection") {
            s.locked = p.get("locked").as_deref() != Some("0");
            s.hidden = p.get("hidden").as_deref() == Some("1");
        }
        s.quote_prefix = xf.get("quotePrefix").as_deref() == Some("1");
        s
    }
}

/// Shortest round-trip text for a number, without exponent for normal magnitudes.
pub fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        let text = format!("{n}");
        if text.len() > 20 { format!("{n:E}") } else { text }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::editor::xml::parse_document;

    fn styles() -> Styles {
        Styles::new("xl/styles.xml", parse_document(Styles::minimal_xml()).unwrap())
    }

    #[test]
    fn dedupes_and_merges() {
        let mut st = styles();
        let bold = FormatSpec { bold: true, ..Default::default() };
        let a = st.apply(&bold).unwrap();
        let b = st.apply(&bold).unwrap();
        assert_eq!(a, b);
        assert_eq!(a, 1);
        let red = FormatSpec { font_color: Some(Color(RColor::RGB(0xFF0000))), ..Default::default() };
        let c = st.merge(a, &red).unwrap();
        let spec = st.spec(c);
        assert!(spec.bold);
        assert_eq!(spec.font_color, Some(Color(RColor::RGB(0xFF0000))));
        assert_eq!(st.xf_count(), 3);
    }

    #[test]
    fn number_formats_and_fills() {
        let mut st = styles();
        let money = FormatSpec { num_format: Some("$#,##0.00".into()), background_color: Some(Color(RColor::Yellow)), ..Default::default() };
        let x = st.apply(&money).unwrap();
        let spec = st.spec(x);
        assert_eq!(spec.num_format.as_deref(), Some("$#,##0.00"));
        assert_eq!(spec.pattern, Some(FormatPattern::Solid));
        assert_eq!(spec.background_color.and_then(|c| c.rgb_value()), Some(0xFFFF00));
        let pct = st.apply(&FormatSpec { num_format: Some("0%".into()), ..Default::default() }).unwrap();
        assert_eq!(st.spec(pct).num_format.as_deref(), Some("0%"));
        assert!(st.to_xml().contains("numFmtId=\"164\""));
    }

    #[test]
    fn date_detection() {
        assert!(is_date_format(14, ""));
        assert!(is_date_format(164, "yyyy-mm-dd"));
        assert!(is_date_format(164, "[h]:mm"));
        assert!(!is_date_format(164, "0.00"));
        assert!(!is_date_format(164, "\"days\" 0"));
        assert!(!is_date_format(164, "[Red]0.00"));
        assert!(!is_date_format(164, "[Magenta]#,##0"));
        assert!(is_date_format(164, "[$-409]mmmm d, yyyy"));
        assert!(!is_date_format(0, "General"));
    }

    #[test]
    fn themes_round_trip() {
        let node = color_node("color", &Color(RColor::Theme(4, 2))).unwrap();
        assert_eq!(read_color(&node), Some(Color(RColor::Theme(4, 2))));
    }
}
