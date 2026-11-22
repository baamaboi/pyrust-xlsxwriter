//! `Shape` (text boxes) and their formatting: `ShapeFormat`, `ShapeLine`,
//! `ShapeSolidFill`, `ShapePatternFill`, `ShapeGradientFill`,
//! `ShapeGradientStop`, `ShapeFont` and `ShapeText`.
//!
//! Nested objects are live: a `Shape` keeps references to its `ShapeFormat`,
//! `ShapeFont` and `ShapeText`, a `ShapeFormat` to its line and fill, and so
//! on, so `shape.format.line.color = "red"` changes the shape. The
//! `rust_xlsxwriter` objects are built from the whole tree when the shape is
//! inserted.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use pyo3::{IntoPyObjectExt, PyClass};
use rust_xlsxwriter::{
    Shape as RShape, ShapeFont as RShapeFont, ShapeFormat as RShapeFormat,
    ShapeGradientFill as RShapeGradientFill, ShapeGradientFillType as RShapeGradientFillType,
    ShapeGradientStop as RShapeGradientStop, ShapeLine as RShapeLine,
    ShapeLineDashType as RShapeLineDashType, ShapePatternFill as RShapePatternFill,
    ShapePatternFillType as RShapePatternFillType, ShapeSolidFill as RShapeSolidFill,
    ShapeText as RShapeText, ShapeTextDirection as RShapeTextDirection,
    ShapeTextHorizontalAlignment as RShapeTextHorizontalAlignment,
    ShapeTextVerticalAlignment as RShapeTextVerticalAlignment,
};

use crate::common::cellref::{split_cell, Params};
use crate::common::color::Color;
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::value::Formula;
use crate::py_props;
use crate::writer::image::{
    check_alt_text, check_object_url, offset_arg, positive_size, read_child, repr_str, UrlArg,
};
use crate::writer::object_movement::ObjectMovement;
use crate::writer::worksheet::Worksheet;

// ---------------------------------------------------------------------------
// Enums
// ---------------------------------------------------------------------------

crate::py_enum!(
    /// Line dash styles for `ShapeLine.dash_type`.
    ShapeLineDashType => RShapeLineDashType {
        Solid, RoundDot, SquareDot, Dash, DashDot, LongDash, LongDashDot, LongDashDotDot,
    }
);

crate::py_enum!(
    /// Patterns for `ShapePatternFill.pattern`.
    ShapePatternFillType => RShapePatternFillType {
        Dotted5Percent, Dotted10Percent, Dotted20Percent, Dotted25Percent, Dotted30Percent,
        Dotted40Percent, Dotted50Percent, Dotted60Percent, Dotted70Percent, Dotted75Percent,
        Dotted80Percent, Dotted90Percent, DiagonalStripesLightDownwards,
        DiagonalStripesLightUpwards, DiagonalStripesDarkDownwards, DiagonalStripesDarkUpwards,
        DiagonalStripesWideDownwards, DiagonalStripesWideUpwards, VerticalStripesLight,
        HorizontalStripesLight, VerticalStripesNarrow, HorizontalStripesNarrow,
        VerticalStripesDark, HorizontalStripesDark, StripesBackslashes, StripesForwardSlashes,
        HorizontalStripesAlternating, VerticalStripesAlternating, SmallConfetti, LargeConfetti,
        Zigzag, Wave, DiagonalBrick, HorizontalBrick, Weave, Plaid, Divot, DottedGrid,
        DottedDiamond, Shingle, Trellis, Sphere, SmallGrid, LargeGrid, SmallCheckerboard,
        LargeCheckerboard, OutlinedDiamondGrid, SolidDiamondGrid,
    }
);

crate::py_enum!(
    /// Gradient types for `ShapeGradientFill.type`.
    ShapeGradientFillType => RShapeGradientFillType { Linear, Radial, Rectangular, Path }
);

crate::py_enum!(
    /// Text direction for `ShapeText.direction`.
    ShapeTextDirection => RShapeTextDirection {
        Horizontal, Rotate270, Rotate90, Rotate90EastAsian, Stacked,
    }
);

crate::py_enum!(
    /// Horizontal text alignment for `ShapeText.horizontal_alignment`.
    ShapeTextHorizontalAlignment => RShapeTextHorizontalAlignment { Default, Left, Center, Right }
);

crate::py_enum!(
    /// Vertical text alignment for `ShapeText.vertical_alignment`.
    ShapeTextVerticalAlignment => RShapeTextVerticalAlignment {
        Top, Middle, Bottom, TopCentered, MiddleCentered, BottomCentered,
    }
);

// ---------------------------------------------------------------------------
// Shared plumbing
// ---------------------------------------------------------------------------

/// Comparison and copying for objects that may hold live children.
trait Node: PyClass + Into<pyo3::PyClassInitializer<Self>> {
    fn node_eq(&self, py: Python<'_>, other: &Self) -> PyResult<bool>;
    /// A deep copy (children are copied too).
    fn node_copy(&self, py: Python<'_>) -> PyResult<Self>;
}

fn child_eq<T: Node>(py: Python<'_>, a: &Py<T>, b: &Py<T>) -> PyResult<bool> {
    if a.is(b) {
        return Ok(true);
    }
    read_child(py, a)?.node_eq(py, &*read_child(py, b)?)
}

fn opt_child_eq<T: Node>(py: Python<'_>, a: &Option<Py<T>>, b: &Option<Py<T>>) -> PyResult<bool> {
    match (a, b) {
        (None, None) => Ok(true),
        (Some(a), Some(b)) => child_eq(py, a, b),
        _ => Ok(false),
    }
}

fn child_copy<T: Node>(py: Python<'_>, obj: &Py<T>) -> PyResult<Py<T>> {
    let copy = read_child(py, obj)?.node_copy(py)?;
    Py::new(py, copy)
}

fn opt_child_copy<T: Node>(py: Python<'_>, obj: &Option<Py<T>>) -> PyResult<Option<Py<T>>> {
    obj.as_ref().map(|o| child_copy(py, o)).transpose()
}

fn clone_opt<T>(obj: &Option<Py<T>>) -> Option<Py<T>> {
    Python::attach(|py| obj.as_ref().map(|o| o.clone_ref(py)))
}

/// `__eq__` for a `Node` class.
fn node_richeq<T: Node>(this: &T, other: &Bound<'_, PyAny>) -> PyResult<bool> {
    match other.cast::<T>() {
        Ok(other) => {
            let other = other
                .try_borrow()
                .map_err(|_| param_err(format!("a {} object is in use elsewhere", <T as PyClass>::NAME)))?;
            this.node_eq(other.py(), &*other)
        }
        Err(_) => Ok(false),
    }
}

/// `Name(prop=value, ...)` listing the properties that differ from the default.
macro_rules! changed_repr {
    ($py:expr, $name:literal, $spec:expr, $default:expr, [$($field:ident),* $(,)?]) => {{
        let mut parts: Vec<String> = Vec::new();
        $(
            if $spec.$field != $default.$field {
                let value = $spec.$field.clone().into_bound_py_any($py)?;
                parts.push(format!("{}={}", stringify!($field), value.repr()?));
            }
        )*
        Ok(format!("{}({})", $name, parts.join(", ")))
    }};
}

fn check_transparency(value: i64) -> PyResult<u8> {
    if !(0..=100).contains(&value) {
        return Err(param_err(format!("transparency must be 0-100 (percent), not {value}")));
    }
    Ok(value as u8)
}

fn color_or_default(color: Option<Color>) -> rust_xlsxwriter::Color {
    color.map(|c| c.0).unwrap_or(rust_xlsxwriter::Color::Default)
}

/// Generates `copy()`, `__copy__()`, `__deepcopy__()` and `__eq__()`.
macro_rules! node_methods {
    ($ty:ident) => {
        #[pymethods]
        impl $ty {
            /// A deep copy (nested objects are copied too).
            fn copy(&self, py: Python<'_>) -> PyResult<Self> {
                self.node_copy(py)
            }

            fn __copy__(&self, py: Python<'_>) -> PyResult<Self> {
                self.node_copy(py)
            }

            fn __deepcopy__(&self, py: Python<'_>, _memo: &Bound<'_, PyAny>) -> PyResult<Self> {
                self.node_copy(py)
            }

            fn __eq__(&self, other: &Bound<'_, PyAny>) -> PyResult<bool> {
                node_richeq(self, other)
            }
        }
    };
}

/// A leaf class: a plain `spec` that derives `Clone` and `PartialEq`.
macro_rules! leaf_node {
    ($ty:ident) => {
        impl Node for $ty {
            fn node_eq(&self, _py: Python<'_>, other: &Self) -> PyResult<bool> {
                Ok(self.spec == other.spec)
            }

            fn node_copy(&self, _py: Python<'_>) -> PyResult<Self> {
                Ok($ty { spec: self.spec.clone() })
            }
        }

        impl $ty {
            fn on_change(&mut self) -> PyResult<()> {
                Ok(())
            }
        }

        node_methods!($ty);
    };
}

// ---------------------------------------------------------------------------
// ShapeLine
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ShapeLineSpec {
    pub color: Option<Color>,
    pub width: f64,
    pub dash_type: ShapeLineDashType,
    pub transparency: u8,
    pub hidden: bool,
}

impl Default for ShapeLineSpec {
    fn default() -> Self {
        ShapeLineSpec {
            color: None,
            width: 0.75,
            dash_type: ShapeLineDashType::Solid,
            transparency: 0,
            hidden: false,
        }
    }
}

/// The border line of a shape: `ShapeLine(color="#FF0000", width=2.25)`.
/// `width` is in points (0-1584); `transparency` is 0-100 percent.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeLine {
    pub spec: ShapeLineSpec,
}

impl ShapeLine {
    pub fn to_rust(&self) -> RShapeLine {
        let s = &self.spec;
        RShapeLine::new()
            .set_color(color_or_default(s.color))
            .set_width(s.width)
            .set_dash_type(s.dash_type.into())
            .set_transparency(s.transparency)
            .set_hidden(s.hidden)
    }

    fn get_width(&self) -> f64 {
        self.spec.width
    }

    fn set_width(&mut self, value: f64) -> PyResult<()> {
        if !(0.0..=1584.0).contains(&value) {
            return Err(param_err(format!("line width must be 0-1584 points, not {value}")));
        }
        self.spec.width = value;
        Ok(())
    }

    fn get_transparency(&self) -> i64 {
        self.spec.transparency as i64
    }

    fn set_transparency(&mut self, value: i64) -> PyResult<()> {
        self.spec.transparency = check_transparency(value)?;
        Ok(())
    }
}

leaf_node!(ShapeLine);

py_props!(ShapeLine
    fields {
        color: Option<Color>,
        dash_type: ShapeLineDashType,
        hidden: bool,
    }
    custom {
        width: f64,
        transparency: i64,
    }
);

#[pymethods]
impl ShapeLine {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut line = ShapeLine { spec: ShapeLineSpec::default() };
        line.apply_kwargs(kwargs)?;
        Ok(line)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        changed_repr!(py, "ShapeLine", self.spec, ShapeLineSpec::default(),
            [color, width, dash_type, transparency, hidden])
    }
}

// ---------------------------------------------------------------------------
// ShapeSolidFill
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq, Default)]
pub struct ShapeSolidFillSpec {
    pub color: Option<Color>,
    pub transparency: u8,
}

/// A solid fill: `ShapeSolidFill(color="#8ED6F5", transparency=50)`.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeSolidFill {
    pub spec: ShapeSolidFillSpec,
}

impl ShapeSolidFill {
    pub fn to_rust(&self) -> RShapeSolidFill {
        RShapeSolidFill::new()
            .set_color(color_or_default(self.spec.color))
            .set_transparency(self.spec.transparency)
    }

    fn get_transparency(&self) -> i64 {
        self.spec.transparency as i64
    }

    fn set_transparency(&mut self, value: i64) -> PyResult<()> {
        self.spec.transparency = check_transparency(value)?;
        Ok(())
    }
}

leaf_node!(ShapeSolidFill);

py_props!(ShapeSolidFill
    fields { color: Option<Color> }
    custom { transparency: i64 }
);

#[pymethods]
impl ShapeSolidFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut fill = ShapeSolidFill { spec: ShapeSolidFillSpec::default() };
        fill.apply_kwargs(kwargs)?;
        Ok(fill)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        changed_repr!(py, "ShapeSolidFill", self.spec, ShapeSolidFillSpec::default(), [color, transparency])
    }
}

// ---------------------------------------------------------------------------
// ShapePatternFill
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ShapePatternFillSpec {
    pub pattern: ShapePatternFillType,
    pub background_color: Option<Color>,
    pub foreground_color: Option<Color>,
}

impl Default for ShapePatternFillSpec {
    fn default() -> Self {
        ShapePatternFillSpec {
            pattern: ShapePatternFillType::Dotted5Percent,
            background_color: None,
            foreground_color: None,
        }
    }
}

/// A pattern fill: `ShapePatternFill(pattern=ShapePatternFillType.Plaid,
/// foreground_color="red", background_color="yellow")`.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapePatternFill {
    pub spec: ShapePatternFillSpec,
}

impl ShapePatternFill {
    pub fn to_rust(&self) -> RShapePatternFill {
        let s = &self.spec;
        let mut fill = RShapePatternFill::new().set_pattern(s.pattern.into());
        if let Some(color) = s.background_color {
            fill = fill.set_background_color(color.0);
        }
        if let Some(color) = s.foreground_color {
            fill = fill.set_foreground_color(color.0);
        }
        fill
    }
}

leaf_node!(ShapePatternFill);

py_props!(ShapePatternFill
    fields {
        pattern: ShapePatternFillType,
        background_color: Option<Color>,
        foreground_color: Option<Color>,
    }
);

#[pymethods]
impl ShapePatternFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut fill = ShapePatternFill { spec: ShapePatternFillSpec::default() };
        fill.apply_kwargs(kwargs)?;
        Ok(fill)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        changed_repr!(py, "ShapePatternFill", self.spec, ShapePatternFillSpec::default(),
            [pattern, background_color, foreground_color])
    }
}

// ---------------------------------------------------------------------------
// ShapeGradientStop
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ShapeGradientStopSpec {
    pub color: Color,
    pub position: u8,
}

/// One colour stop of a gradient: `ShapeGradientStop("#DDEBCF", 0)`.
/// `position` is 0-100 (percent along the gradient).
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeGradientStop {
    pub spec: ShapeGradientStopSpec,
}

fn check_position(value: i64) -> PyResult<u8> {
    if !(0..=100).contains(&value) {
        return Err(param_err(format!("gradient stop position must be 0-100, not {value}")));
    }
    Ok(value as u8)
}

impl ShapeGradientStop {
    pub fn to_rust(&self) -> RShapeGradientStop {
        RShapeGradientStop::new(self.spec.color.0, self.spec.position)
    }

    fn get_position(&self) -> i64 {
        self.spec.position as i64
    }

    fn set_position(&mut self, value: i64) -> PyResult<()> {
        self.spec.position = check_position(value)?;
        Ok(())
    }
}

leaf_node!(ShapeGradientStop);

py_props!(ShapeGradientStop
    fields { color: Color }
    custom { position: i64 }
);

#[pymethods]
impl ShapeGradientStop {
    #[new]
    fn py_new(color: Color, position: i64) -> PyResult<Self> {
        Ok(ShapeGradientStop { spec: ShapeGradientStopSpec { color, position: check_position(position)? } })
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!(
            "ShapeGradientStop({}, {})",
            self.spec.color.into_bound_py_any(py)?.repr()?,
            self.spec.position
        ))
    }
}

/// A gradient stop argument: a `ShapeGradientStop` or a `(color, position)` pair.
pub struct StopArg(Py<ShapeGradientStop>);

impl<'a, 'py> FromPyObject<'a, 'py> for StopArg {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(stop) = obj.cast::<ShapeGradientStop>() {
            return Ok(StopArg(stop.to_owned().unbind()));
        }
        let (color, position): (Color, i64) = obj.extract().map_err(|_| {
            PyTypeError::new_err("gradient stops must be ShapeGradientStop objects or (color, position) pairs")
        })?;
        let stop = ShapeGradientStop { spec: ShapeGradientStopSpec { color, position: check_position(position)? } };
        Ok(StopArg(Py::new(obj.py(), stop)?))
    }
}

impl<'py> IntoPyObject<'py> for StopArg {
    type Target = ShapeGradientStop;
    type Output = Bound<'py, ShapeGradientStop>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(self.0.into_bound(py))
    }
}

// ---------------------------------------------------------------------------
// ShapeGradientFill
// ---------------------------------------------------------------------------

pub struct ShapeGradientFillSpec {
    pub gradient_type: ShapeGradientFillType,
    pub gradient_stops: Vec<Py<ShapeGradientStop>>,
    pub angle: u16,
}

/// A gradient fill with 2-10 stops:
/// `ShapeGradientFill(gradient_stops=[("#DDEBCF", 0), ("#156B13", 100)])`.
/// `angle` (0-359 degrees) applies to linear gradients.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeGradientFill {
    pub spec: ShapeGradientFillSpec,
}

impl ShapeGradientFill {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    pub fn to_rust(&self, py: Python<'_>) -> PyResult<RShapeGradientFill> {
        let s = &self.spec;
        let count = s.gradient_stops.len();
        if !(2..=10).contains(&count) {
            return Err(param_err(format!(
                "a gradient fill needs 2 to 10 gradient_stops, it has {count}"
            )));
        }
        let stops = s
            .gradient_stops
            .iter()
            .map(|stop| read_child(py, stop).map(|stop| stop.to_rust()))
            .collect::<PyResult<Vec<_>>>()?;
        Ok(RShapeGradientFill::new()
            .set_type(s.gradient_type.into())
            .set_gradient_stops(&stops)
            .set_angle(s.angle))
    }

    fn get_type(&self) -> ShapeGradientFillType {
        self.spec.gradient_type
    }

    fn set_type(&mut self, value: ShapeGradientFillType) -> PyResult<()> {
        self.spec.gradient_type = value;
        Ok(())
    }

    fn get_gradient_stops(&self) -> Vec<StopArg> {
        Python::attach(|py| self.spec.gradient_stops.iter().map(|s| StopArg(s.clone_ref(py))).collect())
    }

    fn set_gradient_stops(&mut self, value: Vec<StopArg>) -> PyResult<()> {
        if !value.is_empty() && !(2..=10).contains(&value.len()) {
            return Err(param_err(format!(
                "a gradient fill needs 2 to 10 gradient_stops, not {}",
                value.len()
            )));
        }
        self.spec.gradient_stops = value.into_iter().map(|s| s.0).collect();
        Ok(())
    }

    fn get_angle(&self) -> i64 {
        self.spec.angle as i64
    }

    fn set_angle(&mut self, value: i64) -> PyResult<()> {
        if !(0..360).contains(&value) {
            return Err(param_err(format!("gradient angle must be 0-359 degrees, not {value}")));
        }
        self.spec.angle = value as u16;
        Ok(())
    }
}

impl Node for ShapeGradientFill {
    fn node_eq(&self, py: Python<'_>, other: &Self) -> PyResult<bool> {
        let (a, b) = (&self.spec, &other.spec);
        if a.gradient_type != b.gradient_type
            || a.angle != b.angle
            || a.gradient_stops.len() != b.gradient_stops.len()
        {
            return Ok(false);
        }
        for (x, y) in a.gradient_stops.iter().zip(&b.gradient_stops) {
            if !child_eq(py, x, y)? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    fn node_copy(&self, py: Python<'_>) -> PyResult<Self> {
        Ok(ShapeGradientFill {
            spec: ShapeGradientFillSpec {
                gradient_type: self.spec.gradient_type,
                gradient_stops: self
                    .spec
                    .gradient_stops
                    .iter()
                    .map(|s| child_copy(py, s))
                    .collect::<PyResult<_>>()?,
                angle: self.spec.angle,
            },
        })
    }
}

node_methods!(ShapeGradientFill);

py_props!(ShapeGradientFill
    custom {
        gradient_stops: Vec<StopArg>,
        angle: i64,
    }
);

#[pymethods]
impl ShapeGradientFill {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut fill = ShapeGradientFill {
            spec: ShapeGradientFillSpec {
                gradient_type: ShapeGradientFillType::Linear,
                gradient_stops: Vec::new(),
                angle: 90,
            },
        };
        // `type` is a Rust keyword, so it is handled here instead of by `py_props!`.
        let kwargs = match kwargs {
            Some(kwargs) => {
                let kwargs = kwargs.copy()?;
                if let Some(value) = kwargs.get_item("type")? {
                    fill.set_type(value.extract()?)?;
                    kwargs.del_item("type")?;
                }
                Some(kwargs)
            }
            None => None,
        };
        fill.apply_kwargs(kwargs.as_ref())?;
        Ok(fill)
    }

    /// The gradient type (default `ShapeGradientFillType.Linear`).
    #[getter(r#type)]
    fn py_get_type(&self) -> ShapeGradientFillType {
        self.get_type()
    }

    #[setter(r#type)]
    fn py_set_type(&mut self, value: ShapeGradientFillType) -> PyResult<()> {
        self.set_type(value)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let stops = self
            .spec
            .gradient_stops
            .iter()
            .map(|s| Ok(s.bind(py).repr()?.to_string()))
            .collect::<PyResult<Vec<_>>>()?;
        Ok(format!(
            "ShapeGradientFill(type={}, angle={}, gradient_stops=[{}])",
            self.spec.gradient_type.into_bound_py_any(py)?.repr()?,
            self.spec.angle,
            stops.join(", ")
        ))
    }
}

// ---------------------------------------------------------------------------
// ShapeFormat
// ---------------------------------------------------------------------------

pub struct ShapeFormatSpec {
    pub line: Py<ShapeLine>,
    pub no_line: bool,
    pub solid_fill: Option<Py<ShapeSolidFill>>,
    pub no_fill: bool,
    pub pattern_fill: Option<Py<ShapePatternFill>>,
    pub gradient_fill: Option<Py<ShapeGradientFill>>,
}

impl ShapeFormatSpec {
    fn new(py: Python<'_>) -> PyResult<Self> {
        Ok(ShapeFormatSpec {
            line: Py::new(py, ShapeLine { spec: ShapeLineSpec::default() })?,
            no_line: false,
            solid_fill: None,
            no_fill: false,
            pattern_fill: None,
            gradient_fill: None,
        })
    }

    fn clear_fills(&mut self) {
        self.solid_fill = None;
        self.pattern_fill = None;
        self.gradient_fill = None;
        self.no_fill = false;
    }
}

/// The line and fill of a shape. `line` is always present (a default
/// `ShapeLine`); at most one fill is active: setting `solid_fill`,
/// `pattern_fill`, `gradient_fill` or `no_fill` replaces the others.
///
/// ```python
/// shape.format.line.color = "red"          # nested objects are live
/// shape.format.solid_fill = ShapeSolidFill(color="#8ED6F5")
/// ```
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeFormat {
    pub spec: ShapeFormatSpec,
}

impl ShapeFormat {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    pub fn to_rust(&self, py: Python<'_>) -> PyResult<RShapeFormat> {
        let s = &self.spec;
        let mut format = RShapeFormat::new();
        if s.no_line {
            format = format.set_no_line();
        } else {
            format = format.set_line(&read_child(py, &s.line)?.to_rust());
        }
        if s.no_fill {
            format = format.set_no_fill();
        } else if let Some(fill) = &s.solid_fill {
            format = format.set_solid_fill(&read_child(py, fill)?.to_rust());
        } else if let Some(fill) = &s.pattern_fill {
            format = format.set_pattern_fill(&read_child(py, fill)?.to_rust());
        } else if let Some(fill) = &s.gradient_fill {
            format = format.set_gradient_fill(&read_child(py, fill)?.to_rust(py)?);
        }
        Ok(format)
    }

    fn get_line(&self) -> Option<Py<ShapeLine>> {
        Python::attach(|py| Some(self.spec.line.clone_ref(py)))
    }

    fn set_line(&mut self, value: Option<Py<ShapeLine>>) -> PyResult<()> {
        self.spec.line = match value {
            Some(line) => {
                self.spec.no_line = false;
                line
            }
            None => Python::attach(|py| Py::new(py, ShapeLine { spec: ShapeLineSpec::default() }))?,
        };
        Ok(())
    }

    fn get_no_line(&self) -> bool {
        self.spec.no_line
    }

    fn set_no_line(&mut self, value: bool) -> PyResult<()> {
        self.spec.no_line = value;
        Ok(())
    }

    fn get_no_fill(&self) -> bool {
        self.spec.no_fill
    }

    fn set_no_fill(&mut self, value: bool) -> PyResult<()> {
        if value {
            self.spec.clear_fills();
        }
        self.spec.no_fill = value;
        Ok(())
    }

    fn get_solid_fill(&self) -> Option<Py<ShapeSolidFill>> {
        clone_opt(&self.spec.solid_fill)
    }

    fn set_solid_fill(&mut self, value: Option<Py<ShapeSolidFill>>) -> PyResult<()> {
        if value.is_some() {
            self.spec.clear_fills();
        }
        self.spec.solid_fill = value;
        Ok(())
    }

    fn get_pattern_fill(&self) -> Option<Py<ShapePatternFill>> {
        clone_opt(&self.spec.pattern_fill)
    }

    fn set_pattern_fill(&mut self, value: Option<Py<ShapePatternFill>>) -> PyResult<()> {
        if value.is_some() {
            self.spec.clear_fills();
        }
        self.spec.pattern_fill = value;
        Ok(())
    }

    fn get_gradient_fill(&self) -> Option<Py<ShapeGradientFill>> {
        clone_opt(&self.spec.gradient_fill)
    }

    fn set_gradient_fill(&mut self, value: Option<Py<ShapeGradientFill>>) -> PyResult<()> {
        if value.is_some() {
            self.spec.clear_fills();
        }
        self.spec.gradient_fill = value;
        Ok(())
    }
}

impl Node for ShapeFormat {
    fn node_eq(&self, py: Python<'_>, other: &Self) -> PyResult<bool> {
        let (a, b) = (&self.spec, &other.spec);
        Ok(a.no_line == b.no_line
            && a.no_fill == b.no_fill
            && child_eq(py, &a.line, &b.line)?
            && opt_child_eq(py, &a.solid_fill, &b.solid_fill)?
            && opt_child_eq(py, &a.pattern_fill, &b.pattern_fill)?
            && opt_child_eq(py, &a.gradient_fill, &b.gradient_fill)?)
    }

    fn node_copy(&self, py: Python<'_>) -> PyResult<Self> {
        let s = &self.spec;
        Ok(ShapeFormat {
            spec: ShapeFormatSpec {
                line: child_copy(py, &s.line)?,
                no_line: s.no_line,
                solid_fill: opt_child_copy(py, &s.solid_fill)?,
                no_fill: s.no_fill,
                pattern_fill: opt_child_copy(py, &s.pattern_fill)?,
                gradient_fill: opt_child_copy(py, &s.gradient_fill)?,
            },
        })
    }
}

node_methods!(ShapeFormat);

py_props!(ShapeFormat
    custom {
        line: Option<Py<ShapeLine>>,
        no_line: bool,
        solid_fill: Option<Py<ShapeSolidFill>>,
        no_fill: bool,
        pattern_fill: Option<Py<ShapePatternFill>>,
        gradient_fill: Option<Py<ShapeGradientFill>>,
    }
);

#[pymethods]
impl ShapeFormat {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut format = ShapeFormat { spec: ShapeFormatSpec::new(py)? };
        format.apply_kwargs(kwargs)?;
        Ok(format)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let s = &self.spec;
        let mut parts = Vec::new();
        if s.no_line {
            parts.push("no_line=True".to_string());
        } else {
            parts.push(format!("line={}", s.line.bind(py).repr()?));
        }
        if s.no_fill {
            parts.push("no_fill=True".to_string());
        }
        if let Some(fill) = &s.solid_fill {
            parts.push(format!("solid_fill={}", fill.bind(py).repr()?));
        }
        if let Some(fill) = &s.pattern_fill {
            parts.push(format!("pattern_fill={}", fill.bind(py).repr()?));
        }
        if let Some(fill) = &s.gradient_fill {
            parts.push(format!("gradient_fill={}", fill.bind(py).repr()?));
        }
        Ok(format!("ShapeFormat({})", parts.join(", ")))
    }
}

/// A shape format argument: a `ShapeFormat`, or a single `ShapeLine` /
/// `ShapeSolidFill` / `ShapePatternFill` / `ShapeGradientFill` that is
/// wrapped in a new `ShapeFormat`.
pub struct ShapeFormatArg(Py<ShapeFormat>);

impl<'a, 'py> FromPyObject<'a, 'py> for ShapeFormatArg {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        let py = obj.py();
        if let Ok(format) = obj.cast::<ShapeFormat>() {
            return Ok(ShapeFormatArg(format.to_owned().unbind()));
        }
        let mut spec = ShapeFormatSpec::new(py)?;
        if let Ok(line) = obj.cast::<ShapeLine>() {
            spec.line = line.to_owned().unbind();
        } else if let Ok(fill) = obj.cast::<ShapeSolidFill>() {
            spec.solid_fill = Some(fill.to_owned().unbind());
        } else if let Ok(fill) = obj.cast::<ShapePatternFill>() {
            spec.pattern_fill = Some(fill.to_owned().unbind());
        } else if let Ok(fill) = obj.cast::<ShapeGradientFill>() {
            spec.gradient_fill = Some(fill.to_owned().unbind());
        } else {
            return Err(PyTypeError::new_err(format!(
                "expected a ShapeFormat, ShapeLine, ShapeSolidFill, ShapePatternFill or ShapeGradientFill, not {}",
                obj.get_type().name()?
            )));
        }
        Ok(ShapeFormatArg(Py::new(py, ShapeFormat { spec })?))
    }
}

impl<'py> IntoPyObject<'py> for ShapeFormatArg {
    type Target = ShapeFormat;
    type Output = Bound<'py, ShapeFormat>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(self.0.into_bound(py))
    }
}

// ---------------------------------------------------------------------------
// ShapeFont
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq, Default)]
pub struct ShapeFontSpec {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<Color>,
    pub right_to_left: Option<bool>,
    pub pitch_family: u8,
    pub character_set: u8,
}

/// The font of a shape's text: `ShapeFont(bold=True, color="red", size=14)`.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeFont {
    pub spec: ShapeFontSpec,
}

impl ShapeFont {
    pub fn to_rust(&self) -> RShapeFont {
        let s = &self.spec;
        let mut font = RShapeFont::new()
            .set_pitch_family(s.pitch_family)
            .set_character_set(s.character_set);
        if s.bold {
            font = font.set_bold();
        }
        if s.italic {
            font = font.set_italic();
        }
        if s.underline {
            font = font.set_underline();
        }
        if s.strikethrough {
            font = font.set_strikethrough();
        }
        if let Some(name) = &s.name {
            font = font.set_name(name.as_str());
        }
        if let Some(size) = s.size {
            font = font.set_size(size);
        }
        if let Some(color) = s.color {
            font = font.set_color(color.0);
        }
        if let Some(rtl) = s.right_to_left {
            font = font.set_right_to_left(rtl);
        }
        font
    }

    fn get_size(&self) -> Option<f64> {
        self.spec.size
    }

    fn set_size(&mut self, value: Option<f64>) -> PyResult<()> {
        if let Some(size) = value
            && !(size.is_finite() && size > 0.0 && size <= 4000.0)
        {
            return Err(param_err(format!("font size must be 0-4000 points, not {size}")));
        }
        self.spec.size = value;
        Ok(())
    }
}

leaf_node!(ShapeFont);

py_props!(ShapeFont
    fields {
        bold: bool,
        italic: bool,
        underline: bool,
        strikethrough: bool,
        name: Option<String>,
        color: Option<Color>,
        right_to_left: Option<bool>,
        pitch_family: u8,
        character_set: u8,
    }
    custom { size: Option<f64> }
);

#[pymethods]
impl ShapeFont {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut font = ShapeFont { spec: ShapeFontSpec::default() };
        font.apply_kwargs(kwargs)?;
        Ok(font)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        changed_repr!(py, "ShapeFont", self.spec, ShapeFontSpec::default(),
            [bold, italic, underline, strikethrough, name, size, color, right_to_left,
             pitch_family, character_set])
    }
}

// ---------------------------------------------------------------------------
// ShapeText
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ShapeTextSpec {
    pub horizontal_alignment: ShapeTextHorizontalAlignment,
    pub vertical_alignment: ShapeTextVerticalAlignment,
    pub direction: ShapeTextDirection,
}

impl Default for ShapeTextSpec {
    fn default() -> Self {
        ShapeTextSpec {
            horizontal_alignment: ShapeTextHorizontalAlignment::Default,
            vertical_alignment: ShapeTextVerticalAlignment::Top,
            direction: ShapeTextDirection::Horizontal,
        }
    }
}

/// Text layout options of a shape: alignment and direction.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct ShapeText {
    pub spec: ShapeTextSpec,
}

impl ShapeText {
    pub fn to_rust(&self) -> RShapeText {
        RShapeText::new()
            .set_horizontal_alignment(self.spec.horizontal_alignment.into())
            .set_vertical_alignment(self.spec.vertical_alignment.into())
            .set_direction(self.spec.direction.into())
    }
}

leaf_node!(ShapeText);

py_props!(ShapeText
    fields {
        horizontal_alignment: ShapeTextHorizontalAlignment,
        vertical_alignment: ShapeTextVerticalAlignment,
        direction: ShapeTextDirection,
    }
);

#[pymethods]
impl ShapeText {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut text = ShapeText { spec: ShapeTextSpec::default() };
        text.apply_kwargs(kwargs)?;
        Ok(text)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        changed_repr!(py, "ShapeText", self.spec, ShapeTextSpec::default(),
            [horizontal_alignment, vertical_alignment, direction])
    }
}

// ---------------------------------------------------------------------------
// Shape
// ---------------------------------------------------------------------------

/// A text-link argument: a cell reference string or a `Formula`.
#[derive(Clone, Debug, PartialEq)]
pub struct TextLinkArg(String);

impl<'a, 'py> FromPyObject<'a, 'py> for TextLinkArg {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        let text = if let Ok(formula) = obj.cast::<Formula>() {
            formula.get().formula.clone()
        } else if let Ok(text) = obj.cast::<PyString>() {
            text.to_str()?.to_string()
        } else {
            return Err(PyTypeError::new_err(format!(
                "text_link must be a cell reference such as \"=A1\" or a Formula, not {}",
                obj.get_type().name()?
            )));
        };
        if text.trim_start_matches('=').trim().is_empty() {
            return Err(param_err("text_link must not be empty"));
        }
        Ok(TextLinkArg(text))
    }
}

impl<'py> IntoPyObject<'py> for TextLinkArg {
    type Target = PyString;
    type Output = Bound<'py, PyString>;
    type Error = std::convert::Infallible;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(PyString::new(py, &self.0))
    }
}

pub struct ShapeSpec {
    pub text: String,
    pub text_link: Option<TextLinkArg>,
    pub width: u32,
    pub height: u32,
    pub format: Py<ShapeFormat>,
    pub font: Py<ShapeFont>,
    pub text_options: Py<ShapeText>,
    pub url: Option<UrlArg>,
    pub alt_text: String,
    pub object_movement: ObjectMovement,
}

impl ShapeSpec {
    fn new(py: Python<'_>) -> PyResult<Self> {
        Ok(ShapeSpec {
            text: String::new(),
            text_link: None,
            width: 192,
            height: 120,
            format: Py::new(py, ShapeFormat { spec: ShapeFormatSpec::new(py)? })?,
            font: Py::new(py, ShapeFont { spec: ShapeFontSpec::default() })?,
            text_options: Py::new(py, ShapeText { spec: ShapeTextSpec::default() })?,
            url: None,
            alt_text: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
        })
    }
}

/// A text box shape.
///
/// ```python
/// box = Shape.textbox(text="Hello", width=240, height=80)
/// box.font.bold = True
/// box.format.solid_fill = ShapeSolidFill(color="#DAEEF3")
/// ws.insert_shape("B2", box)
/// ```
///
/// `format`, `font` and `text_options` are always present and live (change
/// them in place); assigning `None` resets them to defaults. Every property
/// can be passed to the constructor as a keyword argument.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct Shape {
    pub spec: ShapeSpec,
}

impl Shape {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    /// Build the `rust_xlsxwriter::Shape` for insertion.
    pub fn to_rust(&self, py: Python<'_>) -> PyResult<RShape> {
        let s = &self.spec;
        let format = read_child(py, &s.format)?.to_rust(py)?;
        let font = read_child(py, &s.font)?.to_rust();
        let text_options = read_child(py, &s.text_options)?.to_rust();
        let mut shape = RShape::textbox()
            .set_text(s.text.as_str())
            .set_width(s.width)
            .set_height(s.height)
            .set_format(&format)
            .set_font(&font)
            .set_text_options(&text_options)
            .set_object_movement(s.object_movement.into());
        if let Some(link) = &s.text_link {
            shape = shape.set_text_link(link.0.as_str());
        }
        if !s.alt_text.is_empty() {
            shape = shape.set_alt_text(s.alt_text.as_str());
        }
        if let Some(url) = &s.url {
            shape = shape.set_url(url.0.to_rust()).py()?;
        }
        Ok(shape)
    }

    fn get_text_link(&self) -> Option<TextLinkArg> {
        self.spec.text_link.clone()
    }

    fn set_text_link(&mut self, value: Option<TextLinkArg>) -> PyResult<()> {
        self.spec.text_link = value;
        Ok(())
    }

    fn get_width(&self) -> i64 {
        self.spec.width as i64
    }

    fn set_width(&mut self, value: i64) -> PyResult<()> {
        self.spec.width = positive_size("width", value, 5)?;
        Ok(())
    }

    fn get_height(&self) -> i64 {
        self.spec.height as i64
    }

    fn set_height(&mut self, value: i64) -> PyResult<()> {
        self.spec.height = positive_size("height", value, 5)?;
        Ok(())
    }

    fn get_format(&self) -> Option<ShapeFormatArg> {
        Python::attach(|py| Some(ShapeFormatArg(self.spec.format.clone_ref(py))))
    }

    fn set_format(&mut self, value: Option<ShapeFormatArg>) -> PyResult<()> {
        self.spec.format = match value {
            Some(format) => format.0,
            None => Python::attach(|py| -> PyResult<Py<ShapeFormat>> {
                Py::new(py, ShapeFormat { spec: ShapeFormatSpec::new(py)? })
            })?,
        };
        Ok(())
    }

    fn get_font(&self) -> Option<Py<ShapeFont>> {
        Python::attach(|py| Some(self.spec.font.clone_ref(py)))
    }

    fn set_font(&mut self, value: Option<Py<ShapeFont>>) -> PyResult<()> {
        self.spec.font = match value {
            Some(font) => font,
            None => Python::attach(|py| Py::new(py, ShapeFont { spec: ShapeFontSpec::default() }))?,
        };
        Ok(())
    }

    fn get_text_options(&self) -> Option<Py<ShapeText>> {
        Python::attach(|py| Some(self.spec.text_options.clone_ref(py)))
    }

    fn set_text_options(&mut self, value: Option<Py<ShapeText>>) -> PyResult<()> {
        self.spec.text_options = match value {
            Some(text) => text,
            None => Python::attach(|py| Py::new(py, ShapeText { spec: ShapeTextSpec::default() }))?,
        };
        Ok(())
    }

    fn get_url(&self) -> Option<UrlArg> {
        self.spec.url.clone()
    }

    fn set_url(&mut self, value: Option<UrlArg>) -> PyResult<()> {
        if let Some(url) = &value {
            check_object_url(&url.0)?;
        }
        self.spec.url = value;
        Ok(())
    }

    fn get_alt_text(&self) -> String {
        self.spec.alt_text.clone()
    }

    fn set_alt_text(&mut self, value: String) -> PyResult<()> {
        check_alt_text(&value)?;
        self.spec.alt_text = value;
        Ok(())
    }

    fn create(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut shape = Shape { spec: ShapeSpec::new(py)? };
        shape.apply_kwargs(kwargs)?;
        Ok(shape)
    }
}

impl Node for Shape {
    fn node_eq(&self, py: Python<'_>, other: &Self) -> PyResult<bool> {
        let (a, b) = (&self.spec, &other.spec);
        Ok(a.text == b.text
            && a.text_link == b.text_link
            && a.width == b.width
            && a.height == b.height
            && a.url == b.url
            && a.alt_text == b.alt_text
            && a.object_movement == b.object_movement
            && child_eq(py, &a.format, &b.format)?
            && child_eq(py, &a.font, &b.font)?
            && child_eq(py, &a.text_options, &b.text_options)?)
    }

    fn node_copy(&self, py: Python<'_>) -> PyResult<Self> {
        let s = &self.spec;
        Ok(Shape {
            spec: ShapeSpec {
                text: s.text.clone(),
                text_link: s.text_link.clone(),
                width: s.width,
                height: s.height,
                format: child_copy(py, &s.format)?,
                font: child_copy(py, &s.font)?,
                text_options: child_copy(py, &s.text_options)?,
                url: s.url.clone(),
                alt_text: s.alt_text.clone(),
                object_movement: s.object_movement,
            },
        })
    }
}

node_methods!(Shape);

py_props!(Shape
    fields {
        text: String,
        object_movement: ObjectMovement,
    }
    custom {
        text_link: Option<TextLinkArg>,
        width: i64,
        height: i64,
        format: Option<ShapeFormatArg>,
        font: Option<Py<ShapeFont>>,
        text_options: Option<Py<ShapeText>>,
        url: Option<UrlArg>,
        alt_text: String,
    }
);

#[pymethods]
impl Shape {
    /// A text box (the only shape type Excel files support here); same as
    /// `Shape.textbox(**properties)`.
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Shape::create(py, kwargs)
    }

    /// Create a text box: `Shape.textbox(text="Hello", width=240)`.
    #[staticmethod]
    #[pyo3(signature = (**kwargs))]
    fn textbox(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Shape::create(py, kwargs)
    }

    fn __repr__(&self) -> String {
        format!("Shape.textbox(text={}, width={}, height={})", repr_str(&self.spec.text), self.spec.width, self.spec.height)
    }
}

#[pymethods]
impl Worksheet {
    // ---- area: shapes ----

    /// Insert a shape (text box) with its top-left corner in the cell, plus
    /// optional pixel offsets: `insert_shape("B2", shape, x_offset=10)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_shape(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_shape", rest, kwargs, &["shape", "x_offset", "y_offset"], 1)?;
        let shape: PyRef<'_, Shape> = p.required(0)?;
        let shape = shape.to_rust(args.py())?;
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner.insert_shape_with_offset(row, col, &shape, x_offset, y_offset).py()?;
        Ok(())
    }

    /// Insert a shape with pixel offsets from the top-left corner of the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_shape_with_offset(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_shape_with_offset", rest, kwargs, &["shape", "x_offset", "y_offset"], 3)?;
        let shape: PyRef<'_, Shape> = p.required(0)?;
        let shape = shape.to_rust(args.py())?;
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner.insert_shape_with_offset(row, col, &shape, x_offset, y_offset).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Shape>()?;
    m.add_class::<ShapeFormat>()?;
    m.add_class::<ShapeLine>()?;
    m.add_class::<ShapeSolidFill>()?;
    m.add_class::<ShapePatternFill>()?;
    m.add_class::<ShapeGradientFill>()?;
    m.add_class::<ShapeGradientStop>()?;
    m.add_class::<ShapeFont>()?;
    m.add_class::<ShapeText>()?;
    m.add_class::<ShapeLineDashType>()?;
    m.add_class::<ShapePatternFillType>()?;
    m.add_class::<ShapeGradientFillType>()?;
    m.add_class::<ShapeTextDirection>()?;
    m.add_class::<ShapeTextHorizontalAlignment>()?;
    m.add_class::<ShapeTextVerticalAlignment>()?;
    Ok(())
}
