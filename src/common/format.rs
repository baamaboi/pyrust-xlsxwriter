//! `Format`: cell formatting shared by the writer, reader and editor.
//!
//! `FormatSpec` is the plain description of a format. The writer turns it
//! into a `rust_xlsxwriter::Format`; the reader builds one from a cell's
//! style; the editor turns it into `styles.xml` records.

use pyo3::prelude::*;
use pyo3::types::PyDict;
use rust_xlsxwriter::{
    FontScheme as RFontScheme, Format as RFormat, FormatAlign as RFormatAlign,
    FormatBorder as RFormatBorder, FormatDiagonalBorder as RFormatDiagonalBorder,
    FormatPattern as RFormatPattern, FormatScript as RFormatScript,
    FormatUnderline as RFormatUnderline,
};

use crate::common::color::Color;
use crate::common::errors::param_err;
use pyo3::IntoPyObjectExt;

py_enum!(
    /// Horizontal and vertical cell alignment.
    FormatAlign => RFormatAlign {
        General, Left, Center, Right, Fill, Justify, CenterAcross, Distributed,
        Top, Bottom, VerticalCenter, VerticalJustify, VerticalDistributed,
    }
);

impl FormatAlign {
    pub fn is_vertical(self) -> bool {
        matches!(
            self,
            FormatAlign::Top
                | FormatAlign::Bottom
                | FormatAlign::VerticalCenter
                | FormatAlign::VerticalJustify
                | FormatAlign::VerticalDistributed
        )
    }
}

py_enum!(
    /// Cell border styles.
    FormatBorder => RFormatBorder {
        #[pyo3(name = "None_")] None, Thin, Medium, Dashed, Dotted, Thick, Double, Hair, MediumDashed, DashDot,
        MediumDashDot, DashDotDot, MediumDashDotDot, SlantDashDot,
    }
);

py_enum!(
    /// Diagonal border direction.
    FormatDiagonalBorder => RFormatDiagonalBorder { #[pyo3(name = "None_")] None, BorderUp, BorderDown, BorderUpDown }
);

py_enum!(
    /// Font underline styles.
    FormatUnderline => RFormatUnderline { #[pyo3(name = "None_")] None, Single, Double, SingleAccounting, DoubleAccounting }
);

py_enum!(
    /// Superscript / subscript.
    FormatScript => RFormatScript { #[pyo3(name = "None_")] None, Superscript, Subscript }
);

py_enum!(
    /// Cell fill patterns.
    FormatPattern => RFormatPattern {
        #[pyo3(name = "None_")] None, Solid, MediumGray, DarkGray, LightGray, DarkHorizontal, DarkVertical, DarkDown,
        DarkUp, DarkGrid, DarkTrellis, LightHorizontal, LightVertical, LightDown, LightUp,
        LightGrid, LightTrellis, Gray125, Gray0625,
    }
);

py_enum!(
    /// Font scheme (theme font) of a format.
    FontScheme => RFontScheme { Body, Headings, #[pyo3(name = "None_")] None }
);

/// Plain description of a cell format. `Default` is Excel's default format.
#[derive(Clone, Debug, PartialEq)]
pub struct FormatSpec {
    pub num_format: Option<String>,
    pub num_format_index: Option<u8>,
    pub bold: bool,
    pub italic: bool,
    pub strikethrough: bool,
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<Color>,
    pub font_family: Option<u8>,
    pub font_charset: Option<u8>,
    pub font_scheme: Option<FontScheme>,
    pub font_script: Option<FormatScript>,
    pub underline: Option<FormatUnderline>,
    pub align: Option<FormatAlign>,
    pub valign: Option<FormatAlign>,
    pub text_wrap: bool,
    pub shrink: bool,
    pub indent: u8,
    pub rotation: i16,
    pub reading_direction: Option<u8>,
    pub pattern: Option<FormatPattern>,
    pub background_color: Option<Color>,
    pub foreground_color: Option<Color>,
    pub border_top: Option<FormatBorder>,
    pub border_bottom: Option<FormatBorder>,
    pub border_left: Option<FormatBorder>,
    pub border_right: Option<FormatBorder>,
    pub border_top_color: Option<Color>,
    pub border_bottom_color: Option<Color>,
    pub border_left_color: Option<Color>,
    pub border_right_color: Option<Color>,
    pub border_diagonal: Option<FormatBorder>,
    pub border_diagonal_color: Option<Color>,
    pub border_diagonal_type: Option<FormatDiagonalBorder>,
    pub hyperlink: bool,
    pub quote_prefix: bool,
    pub checkbox: bool,
    pub locked: bool,
    pub hidden: bool,
}

impl Default for FormatSpec {
    fn default() -> Self {
        FormatSpec {
            num_format: None,
            num_format_index: None,
            bold: false,
            italic: false,
            strikethrough: false,
            font_name: None,
            font_size: None,
            font_color: None,
            font_family: None,
            font_charset: None,
            font_scheme: None,
            font_script: None,
            underline: None,
            align: None,
            valign: None,
            text_wrap: false,
            shrink: false,
            indent: 0,
            rotation: 0,
            reading_direction: None,
            pattern: None,
            background_color: None,
            foreground_color: None,
            border_top: None,
            border_bottom: None,
            border_left: None,
            border_right: None,
            border_top_color: None,
            border_bottom_color: None,
            border_left_color: None,
            border_right_color: None,
            border_diagonal: None,
            border_diagonal_color: None,
            border_diagonal_type: None,
            hyperlink: false,
            quote_prefix: false,
            checkbox: false,
            locked: true,
            hidden: false,
        }
    }
}

impl FormatSpec {
    /// Build the equivalent `rust_xlsxwriter::Format`.
    pub fn to_rust(&self) -> RFormat {
        let mut f = RFormat::new();
        if let Some(v) = &self.num_format {
            f = f.set_num_format(v);
        }
        if let Some(v) = self.num_format_index {
            f = f.set_num_format_index(v);
        }
        if self.bold {
            f = f.set_bold();
        }
        if self.italic {
            f = f.set_italic();
        }
        if self.strikethrough {
            f = f.set_font_strikethrough();
        }
        if let Some(v) = &self.font_name {
            f = f.set_font_name(v);
        }
        if let Some(v) = self.font_size {
            f = f.set_font_size(v);
        }
        if let Some(v) = self.font_color {
            f = f.set_font_color(v);
        }
        if let Some(v) = self.font_family {
            f = f.set_font_family(v);
        }
        if let Some(v) = self.font_charset {
            f = f.set_font_charset(v);
        }
        if let Some(v) = self.font_scheme {
            f = f.set_font_scheme(v.into());
        }
        if let Some(v) = self.font_script {
            f = f.set_font_script(v.into());
        }
        if let Some(v) = self.underline {
            f = f.set_underline(v.into());
        }
        if let Some(v) = self.align {
            f = f.set_align(v.into());
        }
        if let Some(v) = self.valign {
            f = f.set_align(v.into());
        }
        if self.text_wrap {
            f = f.set_text_wrap();
        }
        if self.shrink {
            f = f.set_shrink();
        }
        if self.indent != 0 {
            f = f.set_indent(self.indent);
        }
        if self.rotation != 0 {
            f = f.set_rotation(self.rotation);
        }
        if let Some(v) = self.reading_direction {
            f = f.set_reading_direction(v);
        }
        if let Some(v) = self.pattern {
            f = f.set_pattern(v.into());
        }
        if let Some(v) = self.background_color {
            f = f.set_background_color(v);
        }
        if let Some(v) = self.foreground_color {
            f = f.set_foreground_color(v);
        }
        if let Some(v) = self.border_top {
            f = f.set_border_top(v.into());
        }
        if let Some(v) = self.border_bottom {
            f = f.set_border_bottom(v.into());
        }
        if let Some(v) = self.border_left {
            f = f.set_border_left(v.into());
        }
        if let Some(v) = self.border_right {
            f = f.set_border_right(v.into());
        }
        if let Some(v) = self.border_top_color {
            f = f.set_border_top_color(v);
        }
        if let Some(v) = self.border_bottom_color {
            f = f.set_border_bottom_color(v);
        }
        if let Some(v) = self.border_left_color {
            f = f.set_border_left_color(v);
        }
        if let Some(v) = self.border_right_color {
            f = f.set_border_right_color(v);
        }
        if let Some(v) = self.border_diagonal {
            f = f.set_border_diagonal(v.into());
        }
        if let Some(v) = self.border_diagonal_color {
            f = f.set_border_diagonal_color(v);
        }
        if let Some(v) = self.border_diagonal_type {
            f = f.set_border_diagonal_type(v.into());
        }
        if self.hyperlink {
            f = f.set_hyperlink();
        }
        if self.quote_prefix {
            f = f.set_quote_prefix();
        }
        if self.checkbox {
            f = f.set_checkbox();
        }
        if !self.locked {
            f = f.set_unlocked();
        }
        if self.hidden {
            f = f.set_hidden();
        }
        f
    }

    /// Overlay: every property set in `other` replaces the one in `self`.
    /// Used by "merge with the existing style" in the editor and by `Format.merge()`.
    pub fn overlay(&self, other: &FormatSpec) -> FormatSpec {
        let d = FormatSpec::default();
        let mut out = self.clone();
        macro_rules! take {
            ($($field:ident),*) => { $( if other.$field != d.$field { out.$field = other.$field.clone(); } )* };
        }
        take!(
            num_format, num_format_index, bold, italic, strikethrough, font_name, font_size,
            font_color, font_family, font_charset, font_scheme, font_script, underline, align,
            valign, text_wrap, shrink, indent, rotation, reading_direction, pattern,
            background_color, foreground_color, border_top, border_bottom, border_left,
            border_right, border_top_color, border_bottom_color, border_left_color,
            border_right_color, border_diagonal, border_diagonal_color, border_diagonal_type,
            hyperlink, quote_prefix, checkbox, locked, hidden
        );
        out
    }

    /// Property names that differ from the default, with their Python reprs.
    fn changed(&self, py: Python<'_>) -> PyResult<Vec<(String, String)>> {
        let d = FormatSpec::default();
        let mut out = Vec::new();
        macro_rules! diff {
            ($($field:ident),*) => {
                $( if self.$field != d.$field {
                    let value = self.$field.clone().into_bound_py_any(py)?;
                    out.push((stringify!($field).to_string(), value.repr()?.to_string()));
                } )*
            };
        }
        diff!(
            num_format, num_format_index, bold, italic, strikethrough, font_name, font_size,
            font_color, font_family, font_charset, font_scheme, font_script, underline, align,
            valign, text_wrap, shrink, indent, rotation, reading_direction, pattern,
            background_color, foreground_color, border_top, border_bottom, border_left,
            border_right, border_top_color, border_bottom_color, border_left_color,
            border_right_color, border_diagonal, border_diagonal_color, border_diagonal_type,
            hyperlink, quote_prefix, checkbox, locked, hidden
        );
        Ok(out)
    }
}

/// A cell format. Every property can also be passed to the constructor:
/// `Format(bold=True, font_color="#9C0006", background_color="#FFC7CE")`.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct Format {
    pub spec: FormatSpec,
    pub inner: RFormat,
}

impl Format {
    pub fn from_spec(spec: FormatSpec) -> Self {
        let inner = spec.to_rust();
        Format { spec, inner }
    }

    fn on_change(&mut self) -> PyResult<()> {
        self.inner = self.spec.to_rust();
        Ok(())
    }

    fn get_align(&self) -> Option<FormatAlign> {
        self.spec.align
    }

    /// `align` accepts horizontal and vertical values; vertical ones are
    /// stored in `valign` (as XlsxWriter does).
    fn set_align(&mut self, value: Option<FormatAlign>) -> PyResult<()> {
        match value {
            Some(v) if v.is_vertical() => self.spec.valign = Some(v),
            other => self.spec.align = other,
        }
        Ok(())
    }

    fn get_valign(&self) -> Option<FormatAlign> {
        self.spec.valign
    }

    fn set_valign(&mut self, value: Option<FormatAlign>) -> PyResult<()> {
        if let Some(v) = value
            && !v.is_vertical() {
                return Err(param_err(format!(
                    "valign must be a vertical alignment (Top, Bottom, VerticalCenter, ...), not {v:?}"
                )));
            }
        self.spec.valign = value;
        Ok(())
    }

    /// All four borders; reads back as None unless all four are equal.
    fn get_border(&self) -> Option<FormatBorder> {
        let s = &self.spec;
        if s.border_top == s.border_bottom && s.border_top == s.border_left && s.border_top == s.border_right {
            s.border_top
        } else {
            None
        }
    }

    fn set_border(&mut self, value: Option<FormatBorder>) -> PyResult<()> {
        let s = &mut self.spec;
        s.border_top = value;
        s.border_bottom = value;
        s.border_left = value;
        s.border_right = value;
        Ok(())
    }

    fn get_border_color(&self) -> Option<Color> {
        let s = &self.spec;
        if s.border_top_color == s.border_bottom_color
            && s.border_top_color == s.border_left_color
            && s.border_top_color == s.border_right_color
        {
            s.border_top_color
        } else {
            None
        }
    }

    fn set_border_color(&mut self, value: Option<Color>) -> PyResult<()> {
        let s = &mut self.spec;
        s.border_top_color = value;
        s.border_bottom_color = value;
        s.border_left_color = value;
        s.border_right_color = value;
        Ok(())
    }

    /// Inverse of `locked`, kept for backwards compatibility.
    fn get_unlocked(&self) -> bool {
        !self.spec.locked
    }

    fn set_unlocked(&mut self, value: bool) -> PyResult<()> {
        self.spec.locked = !value;
        Ok(())
    }
}

py_props!(Format
    fields {
        num_format: Option<String>,
        num_format_index: Option<u8>,
        bold: bool,
        italic: bool,
        strikethrough: bool,
        font_name: Option<String>,
        font_size: Option<f64>,
        font_color: Option<Color>,
        font_family: Option<u8>,
        font_charset: Option<u8>,
        font_scheme: Option<FontScheme>,
        font_script: Option<FormatScript>,
        underline: Option<FormatUnderline>,
        text_wrap: bool,
        shrink: bool,
        indent: u8,
        rotation: i16,
        reading_direction: Option<u8>,
        pattern: Option<FormatPattern>,
        background_color: Option<Color>,
        foreground_color: Option<Color>,
        border_top: Option<FormatBorder>,
        border_bottom: Option<FormatBorder>,
        border_left: Option<FormatBorder>,
        border_right: Option<FormatBorder>,
        border_top_color: Option<Color>,
        border_bottom_color: Option<Color>,
        border_left_color: Option<Color>,
        border_right_color: Option<Color>,
        border_diagonal: Option<FormatBorder>,
        border_diagonal_color: Option<Color>,
        border_diagonal_type: Option<FormatDiagonalBorder>,
        hyperlink: bool,
        quote_prefix: bool,
        checkbox: bool,
        locked: bool,
        hidden: bool,
    }
    custom {
        align: Option<FormatAlign>,
        valign: Option<FormatAlign>,
        border: Option<FormatBorder>,
        border_color: Option<Color>,
        unlocked: bool,
    }
);

#[pymethods]
impl Format {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut format = Format::from_spec(FormatSpec::default());
        format.apply_kwargs(kwargs)?;
        Ok(format)
    }

    /// Set several properties at once and return self (for chaining).
    #[pyo3(signature = (**kwargs))]
    fn set<'py>(
        mut slf: PyRefMut<'py, Self>,
        kwargs: Option<&Bound<'py, PyDict>>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.apply_kwargs(kwargs)?;
        Ok(slf)
    }

    /// A copy of this format.
    fn copy(&self) -> Self {
        self.clone()
    }

    fn __copy__(&self) -> Self {
        self.clone()
    }

    fn __deepcopy__(&self, _memo: &Bound<'_, PyAny>) -> Self {
        self.clone()
    }

    /// A new format with `other`'s set properties laid over this one's.
    fn merge(&self, other: PyRef<'_, Format>) -> Self {
        Format::from_spec(self.spec.overlay(&other.spec))
    }

    /// The properties that differ from Excel's default format.
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        let defaults = FormatSpec::default();
        macro_rules! put {
            ($($field:ident),*) => { $( if self.spec.$field != defaults.$field { dict.set_item(stringify!($field), self.spec.$field.clone())?; } )* };
        }
        put!(
            num_format, num_format_index, bold, italic, strikethrough, font_name, font_size,
            font_color, font_family, font_charset, font_scheme, font_script, underline, align,
            valign, text_wrap, shrink, indent, rotation, reading_direction, pattern,
            background_color, foreground_color, border_top, border_bottom, border_left,
            border_right, border_top_color, border_bottom_color, border_left_color,
            border_right_color, border_diagonal, border_diagonal_color, border_diagonal_type,
            hyperlink, quote_prefix, checkbox, locked, hidden
        );
        Ok(dict)
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<Format>() {
            Ok(other) => self.spec == other.borrow().spec,
            Err(_) => false,
        }
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let parts: Vec<String> = self
            .spec
            .changed(py)?
            .into_iter()
            .map(|(k, v)| format!("{k}={v}"))
            .collect();
        Ok(format!("Format({})", parts.join(", ")))
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for Format {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        Ok(obj.cast::<Format>()?.borrow().clone())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Format>()?;
    m.add_class::<FormatAlign>()?;
    m.add_class::<FormatBorder>()?;
    m.add_class::<FormatDiagonalBorder>()?;
    m.add_class::<FormatUnderline>()?;
    m.add_class::<FormatScript>()?;
    m.add_class::<FormatPattern>()?;
    m.add_class::<FontScheme>()?;
    Ok(())
}
