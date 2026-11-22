//! `Color`: named colours, RGB and theme colours.
//!
//! Anywhere a colour is accepted, Python callers may pass a `Color`, a name
//! (`"red"`), a hex string (`"#FF0000"` / `"FF0000"`), an int (`0xFF0000`) or
//! an `(r, g, b)` tuple.

use pyo3::basic::CompareOp;
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyInt, PyList, PyString, PyTuple};
use rust_xlsxwriter::Color as RColor;

use crate::common::errors::param_err;

const NAMED: [(&str, RColor); 18] = [
    ("Default", RColor::Default),
    ("Automatic", RColor::Automatic),
    ("Black", RColor::Black),
    ("Blue", RColor::Blue),
    ("Brown", RColor::Brown),
    ("Cyan", RColor::Cyan),
    ("Gray", RColor::Gray),
    ("Green", RColor::Green),
    ("Lime", RColor::Lime),
    ("Magenta", RColor::Magenta),
    ("Navy", RColor::Navy),
    ("Orange", RColor::Orange),
    ("Pink", RColor::Pink),
    ("Purple", RColor::Purple),
    ("Red", RColor::Red),
    ("Silver", RColor::Silver),
    ("White", RColor::White),
    ("Yellow", RColor::Yellow),
];

/// An Excel colour.
///
/// Use a named constant (`Color.Red`), `Color("#4472C4")`, `Color(0x4472C4)`,
/// `Color((68, 114, 196))`, `Color.rgb(68, 114, 196)` or a theme colour
/// `Color.theme(color, shade)`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, skip_from_py_object)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Color(pub RColor);

impl From<Color> for RColor {
    fn from(value: Color) -> Self {
        value.0
    }
}

impl From<RColor> for Color {
    fn from(value: RColor) -> Self {
        Color(value)
    }
}

impl Color {
    /// The 24-bit RGB value if this colour has one.
    pub fn rgb_value(&self) -> Option<u32> {
        match self.0 {
            RColor::RGB(value) => Some(value),
            RColor::Default | RColor::Automatic | RColor::Theme(_, _) => None,
            RColor::Black => Some(0x000000),
            RColor::Blue => Some(0x0000FF),
            RColor::Brown => Some(0x800000),
            RColor::Cyan => Some(0x00FFFF),
            RColor::Gray => Some(0x808080),
            RColor::Green => Some(0x008000),
            RColor::Lime => Some(0x00FF00),
            RColor::Magenta => Some(0xFF00FF),
            RColor::Navy => Some(0x000080),
            RColor::Orange => Some(0xFF6600),
            RColor::Pink => Some(0xFFC0CB),
            RColor::Purple => Some(0x800080),
            RColor::Red => Some(0xFF0000),
            RColor::Silver => Some(0xC0C0C0),
            RColor::White => Some(0xFFFFFF),
            RColor::Yellow => Some(0xFFFF00),
        }
    }

    fn named(&self) -> Option<&'static str> {
        NAMED.iter().find(|(_, c)| *c == self.0).map(|(name, _)| *name)
    }
}

fn parse_hex(text: &str) -> Option<u32> {
    let hex = text.strip_prefix('#').unwrap_or(text);
    if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        u32::from_str_radix(hex, 16).ok()
    } else {
        None
    }
}

fn check_rgb(value: i64) -> PyResult<u32> {
    if (0..=0xFF_FFFF).contains(&value) {
        Ok(value as u32)
    } else {
        Err(param_err(format!("RGB colour value {value:#X} is outside 0x000000..=0xFFFFFF")))
    }
}

fn rgb_from_sequence(items: &[Bound<'_, PyAny>]) -> PyResult<RColor> {
    if items.len() != 3 {
        return Err(param_err("an RGB colour tuple must have exactly 3 items (r, g, b)"));
    }
    let mut value: u32 = 0;
    for item in items {
        let channel: i64 = item.extract()?;
        if !(0..=255).contains(&channel) {
            return Err(param_err(format!("RGB channel {channel} is outside 0..=255")));
        }
        value = (value << 8) | channel as u32;
    }
    Ok(RColor::RGB(value))
}

/// Parse any supported Python colour representation.
pub fn parse_color(value: &Bound<'_, PyAny>) -> PyResult<RColor> {
    if let Ok(color) = value.cast::<Color>() {
        return Ok(color.get().0);
    }
    if let Ok(text) = value.cast::<PyString>() {
        let text = text.to_str()?.trim();
        if let Some((_, color)) = NAMED.iter().find(|(name, _)| name.eq_ignore_ascii_case(text)) {
            return Ok(*color);
        }
        if let Some(rgb) = parse_hex(text) {
            return Ok(RColor::RGB(rgb));
        }
        return Err(param_err(format!(
            "invalid colour {text:?}: expected a colour name or '#RRGGBB'"
        )));
    }
    if value.is_instance_of::<PyInt>() && !value.is_instance_of::<PyBool>() {
        return Ok(RColor::RGB(check_rgb(value.extract()?)?));
    }
    if let Ok(tuple) = value.cast::<PyTuple>() {
        return rgb_from_sequence(&tuple.iter().collect::<Vec<_>>());
    }
    if let Ok(list) = value.cast::<PyList>() {
        return rgb_from_sequence(&list.iter().collect::<Vec<_>>());
    }
    Err(PyTypeError::new_err(format!(
        "expected a Color, colour name, '#RRGGBB' string, int or (r, g, b) tuple, not {}",
        value.get_type().name()?
    )))
}

impl<'a, 'py> FromPyObject<'a, 'py> for Color {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        parse_color(&obj.to_owned()).map(Color)
    }
}

macro_rules! named_colors {
    ($( $name:literal => $fn:ident => $variant:ident ),* $(,)?) => {
        #[pymethods]
        impl Color {
            $(
                #[classattr]
                #[pyo3(name = $name)]
                fn $fn() -> Color {
                    Color(RColor::$variant)
                }
            )*
        }
    };
}

named_colors!(
    "Default" => default_ => Default,
    "Automatic" => automatic => Automatic,
    "Black" => black => Black,
    "Blue" => blue => Blue,
    "Brown" => brown => Brown,
    "Cyan" => cyan => Cyan,
    "Gray" => gray => Gray,
    "Green" => green => Green,
    "Lime" => lime => Lime,
    "Magenta" => magenta => Magenta,
    "Navy" => navy => Navy,
    "Orange" => orange => Orange,
    "Pink" => pink => Pink,
    "Purple" => purple => Purple,
    "Red" => red => Red,
    "Silver" => silver => Silver,
    "White" => white => White,
    "Yellow" => yellow => Yellow,
);

#[pymethods]
impl Color {
    #[new]
    fn py_new(value: &Bound<'_, PyAny>) -> PyResult<Self> {
        parse_color(value).map(Color)
    }

    /// `Color.rgb(0x4472C4)` or `Color.rgb(68, 114, 196)`.
    #[staticmethod]
    #[pyo3(signature = (red_or_value, green=None, blue=None))]
    fn rgb(red_or_value: i64, green: Option<i64>, blue: Option<i64>) -> PyResult<Self> {
        match (green, blue) {
            (None, None) => Ok(Color(RColor::RGB(check_rgb(red_or_value)?))),
            (Some(green), Some(blue)) => {
                for channel in [red_or_value, green, blue] {
                    if !(0..=255).contains(&channel) {
                        return Err(param_err(format!("RGB channel {channel} is outside 0..=255")));
                    }
                }
                Ok(Color(RColor::RGB(
                    ((red_or_value as u32) << 16) | ((green as u32) << 8) | blue as u32,
                )))
            }
            _ => Err(param_err("Color.rgb() takes either one value or three channels")),
        }
    }

    /// A theme colour: `color` 0-9 (column in Excel's palette), `shade` 0-5 (row).
    #[staticmethod]
    fn theme(color: u8, shade: u8) -> PyResult<Self> {
        if color > 9 || shade > 5 {
            return Err(param_err(format!(
                "theme colour ({color}, {shade}) is outside color 0..=9, shade 0..=5"
            )));
        }
        Ok(Color(RColor::Theme(color, shade)))
    }

    /// "named", "rgb", "theme", "default" or "automatic".
    #[getter]
    fn kind(&self) -> &'static str {
        match self.0 {
            RColor::RGB(_) => "rgb",
            RColor::Theme(_, _) => "theme",
            RColor::Default => "default",
            RColor::Automatic => "automatic",
            _ => "named",
        }
    }

    /// The colour as "#RRGGBB" (None for default, automatic and theme colours).
    #[getter]
    fn hex(&self) -> Option<String> {
        self.rgb_value().map(|v| format!("#{v:06X}"))
    }

    /// The colour as a 24-bit int (None for default, automatic and theme colours).
    #[getter]
    fn value(&self) -> Option<u32> {
        self.rgb_value()
    }

    /// `(color, shade)` for theme colours, otherwise None.
    #[getter]
    fn theme_index(&self) -> Option<(u8, u8)> {
        match self.0 {
            RColor::Theme(color, shade) => Some((color, shade)),
            _ => None,
        }
    }

    fn __repr__(&self) -> String {
        match self.0 {
            RColor::RGB(v) => format!("Color('#{v:06X}')"),
            RColor::Theme(c, s) => format!("Color.theme({c}, {s})"),
            _ => format!("Color.{}", self.named().unwrap_or("Default")),
        }
    }

    fn __richcmp__(&self, other: &Bound<'_, PyAny>, op: CompareOp) -> PyResult<Py<PyAny>> {
        let py = other.py();
        let other = match other.cast::<Color>() {
            Ok(c) => c.get().0,
            Err(_) => return Ok(py.NotImplemented()),
        };
        Ok(match op {
            CompareOp::Eq => (self.0 == other).into_pyobject(py)?.to_owned().into_any().unbind(),
            CompareOp::Ne => (self.0 != other).into_pyobject(py)?.to_owned().into_any().unbind(),
            _ => py.NotImplemented(),
        })
    }

    fn __hash__(&self) -> u64 {
        use std::hash::{Hash, Hasher};
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        self.0.hash(&mut hasher);
        hasher.finish()
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Color>()
}
