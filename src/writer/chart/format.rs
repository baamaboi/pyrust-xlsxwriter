//! Chart formatting: lines, fills, `ChartFormat`, `ChartFont`, `ChartLayout`.

use pyo3::IntoPyObjectExt;
use pyo3::basic::CompareOp;
use pyo3::prelude::*;
use pyo3::types::PyDict;
use rust_xlsxwriter::{
    ChartFont as RChartFont, ChartFormat as RChartFormat, ChartGradientFill as RChartGradientFill,
    ChartGradientStop as RChartGradientStop, ChartLayout as RChartLayout, ChartLine as RChartLine,
    ChartPatternFill as RChartPatternFill, ChartSolidFill as RChartSolidFill,
};

use super::enums::{ChartGradientFillType, ChartLineDashType, ChartPatternFillType};
use super::node::{
    Build, Node, Props, Slot, check_angle, check_font_rotation, check_font_size,
    check_fraction_pair, check_line_width, check_percent, coerce_default, new_with, repr_of,
};
use crate::common::color::Color;
use crate::common::errors::param_err;

// ---------------------------------------------------------------------------
// ChartLine
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartLineSpec {
    pub color: Option<Color>,
    pub width: Option<f64>,
    pub transparency: u8,
    pub dash_type: Option<ChartLineDashType>,
    pub hidden: bool,
}

/// A line or border: `ChartLine(color="#FF0000", width=1.5, dash_type=ChartLineDashType.Dash)`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartLine {
    pub spec: ChartLineSpec,
}

chart_props!(ChartLine
    fields {
        /// Line colour.
        color: Option<Color>,
        /// Width in points (0 to 1584).
        width: Option<f64> => check_line_width,
        /// Transparency in percent (0 to 100).
        transparency: u8 => check_percent,
        /// Dash style.
        dash_type: Option<ChartLineDashType>,
        /// Hide the line.
        hidden: bool,
    }
    children {}
    custom {}
);

impl Node for ChartLine {}

impl Build for ChartLine {
    type Rust = RChartLine;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartLine> {
        let s = &self.spec;
        let mut line = RChartLine::new();
        if let Some(color) = s.color {
            line.set_color(color.0);
        }
        if let Some(width) = s.width {
            line.set_width(width);
        }
        if s.transparency != 0 {
            line.set_transparency(s.transparency);
        }
        if let Some(dash) = s.dash_type {
            line.set_dash_type(dash.into());
        }
        if s.hidden {
            line.set_hidden(true);
        }
        Ok(line)
    }
}

#[pymethods]
impl ChartLine {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartLine", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartSolidFill
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartSolidFillSpec {
    pub color: Option<Color>,
    pub transparency: u8,
}

/// A solid fill: `ChartSolidFill(color="#4472C4", transparency=20)`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartSolidFill {
    pub spec: ChartSolidFillSpec,
}

chart_props!(ChartSolidFill
    fields {
        /// Fill colour.
        color: Option<Color>,
        /// Transparency in percent (0 to 100).
        transparency: u8 => check_percent,
    }
    children {}
    custom {}
);

impl Node for ChartSolidFill {}

impl Build for ChartSolidFill {
    type Rust = RChartSolidFill;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartSolidFill> {
        let mut fill = RChartSolidFill::new();
        if let Some(color) = self.spec.color {
            fill.set_color(color.0);
        }
        if self.spec.transparency != 0 {
            fill.set_transparency(self.spec.transparency);
        }
        Ok(fill)
    }
}

#[pymethods]
impl ChartSolidFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartSolidFill", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartPatternFill
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartPatternFillSpec {
    pub pattern: Option<ChartPatternFillType>,
    pub background_color: Option<Color>,
    pub foreground_color: Option<Color>,
}

/// A pattern fill: `ChartPatternFill(pattern=ChartPatternFillType.Weave, foreground_color="red")`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartPatternFill {
    pub spec: ChartPatternFillSpec,
}

chart_props!(ChartPatternFill
    fields {
        /// The pattern (default Dotted5Percent).
        pattern: Option<ChartPatternFillType>,
        /// Background colour.
        background_color: Option<Color>,
        /// Foreground colour.
        foreground_color: Option<Color>,
    }
    children {}
    custom {}
);

impl Node for ChartPatternFill {}

impl Build for ChartPatternFill {
    type Rust = RChartPatternFill;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartPatternFill> {
        let s = &self.spec;
        let mut fill = RChartPatternFill::new();
        if let Some(pattern) = s.pattern {
            fill.set_pattern(pattern.into());
        }
        if let Some(color) = s.background_color {
            fill.set_background_color(color.0);
        }
        if let Some(color) = s.foreground_color {
            fill.set_foreground_color(color.0);
        }
        Ok(fill)
    }
}

#[pymethods]
impl ChartPatternFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartPatternFill", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartGradientStop, ChartGradientFill
// ---------------------------------------------------------------------------

/// One colour stop of a gradient: `ChartGradientStop("#FF0000", 50)`
/// (position in percent, 0 to 100). Immutable.
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ChartGradientStop {
    pub color: Color,
    pub position: u8,
}

#[pymethods]
impl ChartGradientStop {
    #[new]
    fn py_new(color: Color, position: u8) -> PyResult<Self> {
        check_percent("position", &position)?;
        Ok(ChartGradientStop { color, position })
    }

    /// The stop colour.
    #[getter]
    fn color(&self) -> Color {
        self.color
    }

    /// The stop position in percent.
    #[getter]
    fn position(&self) -> u8 {
        self.position
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let color = self.color.into_bound_py_any(py)?;
        Ok(format!(
            "ChartGradientStop({}, {})",
            color.repr()?,
            self.position
        ))
    }

    fn __richcmp__(&self, other: &Bound<'_, PyAny>, op: CompareOp) -> PyResult<Py<PyAny>> {
        let py = other.py();
        let Ok(other) = other.cast::<ChartGradientStop>() else {
            return Ok(py.NotImplemented());
        };
        let equal = self == other.get();
        match op {
            CompareOp::Eq => equal.into_py_any(py),
            CompareOp::Ne => (!equal).into_py_any(py),
            _ => Ok(py.NotImplemented()),
        }
    }

    fn __hash__(&self) -> u64 {
        use std::hash::{Hash, Hasher};
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        (self.color, self.position).hash(&mut hasher);
        hasher.finish()
    }
}

#[allow(clippy::ptr_arg)] // the validator signature is `fn(&str, &FieldType)`
fn check_gradient_stops(name: &str, value: &Vec<ChartGradientStop>) -> PyResult<()> {
    if !value.is_empty() && !(2..=10).contains(&value.len()) {
        return Err(param_err(format!(
            "{name} must contain between 2 and 10 stops, got {}",
            value.len()
        )));
    }
    Ok(())
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartGradientFillSpec {
    pub type_: Option<ChartGradientFillType>,
    pub gradient_stops: Vec<ChartGradientStop>,
    pub angle: Option<u16>,
}

/// A gradient fill:
/// `ChartGradientFill(gradient_stops=[ChartGradientStop("#DDEBCF", 0), ChartGradientStop("#156B13", 100)])`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartGradientFill {
    pub spec: ChartGradientFillSpec,
}

chart_props!(ChartGradientFill
    fields {
        /// Gradient type (default Linear).
        type_ as r#type: Option<ChartGradientFillType>,
        /// 2 to 10 colour stops (an empty list uses Excel's default gradient).
        gradient_stops: Vec<ChartGradientStop> => check_gradient_stops,
        /// Angle of a linear gradient in degrees (0 to 359, default 90).
        angle: Option<u16> => check_angle,
    }
    children {}
    custom {}
);

impl Node for ChartGradientFill {}

impl Build for ChartGradientFill {
    type Rust = RChartGradientFill;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartGradientFill> {
        let s = &self.spec;
        let mut fill = RChartGradientFill::new();
        if let Some(kind) = s.type_ {
            fill.set_type(kind.into());
        }
        if !s.gradient_stops.is_empty() {
            let stops: Vec<RChartGradientStop> = s
                .gradient_stops
                .iter()
                .map(|stop| RChartGradientStop::new(stop.color.0, stop.position))
                .collect();
            fill.set_gradient_stops(&stops);
        }
        if let Some(angle) = s.angle {
            fill.set_angle(angle);
        }
        Ok(fill)
    }
}

#[pymethods]
impl ChartGradientFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartGradientFill", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartFormat
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartFormatSpec {
    pub no_line: bool,
    pub no_fill: bool,
}

/// Formatting of a chart element: line/border and fill.
///
/// ```python
/// series.format = ChartFormat(line=ChartLine(color="red"), solid_fill=ChartSolidFill(color="#FFC7CE"))
/// series.format.line.width = 2.5            # nested objects are live
/// series.format = ChartSolidFill(color="red")  # a line or fill alone is accepted too
/// ```
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartFormat {
    pub spec: ChartFormatSpec,
    pub line: Slot<ChartLine>,
    pub solid_fill: Slot<ChartSolidFill>,
    pub pattern_fill: Slot<ChartPatternFill>,
    pub gradient_fill: Slot<ChartGradientFill>,
}

chart_props!(ChartFormat
    fields {
        /// Turn off the line/border.
        no_line: bool,
        /// Turn off the fill.
        no_fill: bool,
    }
    children {
        /// The line (or border) properties.
        line: ChartLine,
        /// Solid fill.
        solid_fill: ChartSolidFill,
        /// Pattern fill.
        pattern_fill: ChartPatternFill,
        /// Gradient fill.
        gradient_fill: ChartGradientFill,
    }
    custom {
        /// Alias of `line` (Excel calls the line of a filled shape its border).
        border,
        /// Alias of `no_line`.
        no_border,
    }
);

impl ChartFormat {
    fn get_border(&self, py: Python<'_>) -> PyResult<Py<ChartLine>> {
        self.line.get(py)
    }

    fn set_border(&mut self, _py: Python<'_>, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        self.line.assign(value.as_ref())
    }

    fn get_no_border(&self, _py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.no_line)
    }

    fn set_no_border(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.no_line = value;
        Ok(())
    }
}

impl Node for ChartFormat {
    /// Accepts a ChartFormat, a dict, or a ChartLine / ChartSolidFill /
    /// ChartPatternFill / ChartGradientFill (wrapped in a new format).
    fn coerce(value: &Bound<'_, PyAny>) -> PyResult<Option<Py<Self>>> {
        let py = value.py();
        let mut format = ChartFormat::default();
        if value.cast::<ChartLine>().is_ok() {
            format.line.assign(Some(value))?;
        } else if value.cast::<ChartSolidFill>().is_ok() {
            format.solid_fill.assign(Some(value))?;
        } else if value.cast::<ChartPatternFill>().is_ok() {
            format.pattern_fill.assign(Some(value))?;
        } else if value.cast::<ChartGradientFill>().is_ok() {
            format.gradient_fill.assign(Some(value))?;
        } else {
            return coerce_default::<Self>(value);
        }
        Ok(Some(Py::new(py, format)?))
    }
}

impl Build for ChartFormat {
    type Rust = RChartFormat;

    fn build(&self, py: Python<'_>) -> PyResult<RChartFormat> {
        let mut format = RChartFormat::new();
        if let Some(line) = self.line.build(py)? {
            format.set_line(&line);
        }
        if self.spec.no_line {
            format.set_no_line();
        }
        if self.spec.no_fill {
            format.set_no_fill();
        }
        if let Some(fill) = self.solid_fill.build(py)? {
            format.set_solid_fill(&fill);
        }
        if let Some(fill) = self.pattern_fill.build(py)? {
            format.set_pattern_fill(&fill);
        }
        if let Some(fill) = self.gradient_fill.build(py)? {
            format.set_gradient_fill(&fill);
        }
        Ok(format)
    }
}

#[pymethods]
impl ChartFormat {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartFormat", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartFont
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartFontSpec {
    pub bold: Option<bool>,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<Color>,
    pub rotation: Option<i16>,
    pub right_to_left: Option<bool>,
    pub pitch_family: u8,
    pub character_set: u8,
}

/// Font of chart text: `ChartFont(bold=True, size=12, color="#1F4E79", name="Calibri")`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartFont {
    pub spec: ChartFontSpec,
}

chart_props!(ChartFont
    fields {
        /// True/False to force bold on/off; None keeps the element's default
        /// (titles are bold by default).
        bold: Option<bool>,
        /// Italic text.
        italic: bool,
        /// Underlined text.
        underline: bool,
        /// Strikethrough text.
        strikethrough: bool,
        /// Font name, e.g. "Calibri".
        name: Option<String>,
        /// Font size in points (1 to 409).
        size: Option<f64> => check_font_size,
        /// Font colour.
        color: Option<Color>,
        /// Text rotation in degrees (-90 to 90; 270 or 271 for stacked text).
        rotation: Option<i16> => check_font_rotation,
        /// Right-to-left text.
        right_to_left: Option<bool>,
        /// Font pitch family (advanced).
        pitch_family: u8,
        /// Font character set (advanced).
        character_set: u8,
    }
    children {}
    custom {}
);

impl Node for ChartFont {}

impl Build for ChartFont {
    type Rust = RChartFont;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartFont> {
        let s = &self.spec;
        let mut font = RChartFont::new();
        match s.bold {
            Some(true) => {
                font.set_bold();
            }
            Some(false) => {
                font.unset_bold();
            }
            None => {}
        }
        if s.italic {
            font.set_italic();
        }
        if s.underline {
            font.set_underline();
        }
        if s.strikethrough {
            font.set_strikethrough();
        }
        if let Some(name) = &s.name {
            font.set_name(name);
        }
        if let Some(size) = s.size {
            font.set_size(size);
        }
        if let Some(color) = s.color {
            font.set_color(color.0);
        }
        if let Some(rotation) = s.rotation {
            font.set_rotation(rotation);
        }
        if let Some(rtl) = s.right_to_left {
            font.set_right_to_left(rtl);
        }
        if s.pitch_family != 0 {
            font.set_pitch_family(s.pitch_family);
        }
        if s.character_set != 0 {
            font.set_character_set(s.character_set);
        }
        Ok(font)
    }
}

#[pymethods]
impl ChartFont {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartFont", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartLayout
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartLayoutSpec {
    pub offset: Option<(f64, f64)>,
    pub dimensions: Option<(f64, f64)>,
}

/// Manual position/size of a chart element, as fractions (0 to 1) of the
/// chart area: `ChartLayout(offset=(0.1, 0.2), dimensions=(0.6, 0.5))`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartLayout {
    pub spec: ChartLayoutSpec,
}

chart_props!(ChartLayout
    fields {
        /// `(x, y)` of the top-left corner, as fractions of the chart area.
        offset: Option<(f64, f64)> => check_fraction_pair,
        /// `(width, height)`, as fractions of the chart area.
        dimensions: Option<(f64, f64)> => check_fraction_pair,
    }
    children {}
    custom {}
);

impl Node for ChartLayout {}

impl Build for ChartLayout {
    type Rust = RChartLayout;

    fn build(&self, _py: Python<'_>) -> PyResult<RChartLayout> {
        let mut layout = RChartLayout::new();
        if let Some((x, y)) = self.spec.offset {
            layout = layout.set_offset(x, y);
        }
        if let Some((width, height)) = self.spec.dimensions {
            layout = layout.set_dimensions(width, height);
        }
        Ok(layout)
    }
}

#[pymethods]
impl ChartLayout {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartLayout", self.props_repr(py)?))
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ChartLine>()?;
    m.add_class::<ChartSolidFill>()?;
    m.add_class::<ChartPatternFill>()?;
    m.add_class::<ChartGradientStop>()?;
    m.add_class::<ChartGradientFill>()?;
    m.add_class::<ChartFormat>()?;
    m.add_class::<ChartFont>()?;
    m.add_class::<ChartLayout>()?;
    Ok(())
}
