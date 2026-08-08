use pyo3::prelude::*;
use rust_xlsxwriter::{
    Color as RColor, Format as RFormat, FormatAlign as RFormatAlign, FormatBorder as RFormatBorder,
    FormatPattern as RFormatPattern, FormatScript as RFormatScript,
    FormatUnderline as RFormatUnderline,
};

use crate::workbook::Color;

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormatAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterAcross,
    Distributed,
    Top,
    Bottom,
    VerticalCenter,
    VerticalJustify,
    VerticalDistributed,
}

impl From<FormatAlign> for RFormatAlign {
    fn from(fa: FormatAlign) -> Self {
        match fa {
            FormatAlign::General => RFormatAlign::General,
            FormatAlign::Left => RFormatAlign::Left,
            FormatAlign::Center => RFormatAlign::Center,
            FormatAlign::Right => RFormatAlign::Right,
            FormatAlign::Fill => RFormatAlign::Fill,
            FormatAlign::Justify => RFormatAlign::Justify,
            FormatAlign::CenterAcross => RFormatAlign::CenterAcross,
            FormatAlign::Distributed => RFormatAlign::Distributed,
            FormatAlign::Top => RFormatAlign::Top,
            FormatAlign::Bottom => RFormatAlign::Bottom,
            FormatAlign::VerticalCenter => RFormatAlign::VerticalCenter,
            FormatAlign::VerticalJustify => RFormatAlign::VerticalJustify,
            FormatAlign::VerticalDistributed => RFormatAlign::VerticalDistributed,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormatBorder {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl From<FormatBorder> for RFormatBorder {
    fn from(fb: FormatBorder) -> Self {
        match fb {
            FormatBorder::None => RFormatBorder::None,
            FormatBorder::Thin => RFormatBorder::Thin,
            FormatBorder::Medium => RFormatBorder::Medium,
            FormatBorder::Dashed => RFormatBorder::Dashed,
            FormatBorder::Dotted => RFormatBorder::Dotted,
            FormatBorder::Thick => RFormatBorder::Thick,
            FormatBorder::Double => RFormatBorder::Double,
            FormatBorder::Hair => RFormatBorder::Hair,
            FormatBorder::MediumDashed => RFormatBorder::MediumDashed,
            FormatBorder::DashDot => RFormatBorder::DashDot,
            FormatBorder::MediumDashDot => RFormatBorder::MediumDashDot,
            FormatBorder::DashDotDot => RFormatBorder::DashDotDot,
            FormatBorder::MediumDashDotDot => RFormatBorder::MediumDashDotDot,
            FormatBorder::SlantDashDot => RFormatBorder::SlantDashDot,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormatUnderline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

impl From<FormatUnderline> for RFormatUnderline {
    fn from(fu: FormatUnderline) -> Self {
        match fu {
            FormatUnderline::None => RFormatUnderline::None,
            FormatUnderline::Single => RFormatUnderline::Single,
            FormatUnderline::Double => RFormatUnderline::Double,
            FormatUnderline::SingleAccounting => RFormatUnderline::SingleAccounting,
            FormatUnderline::DoubleAccounting => RFormatUnderline::DoubleAccounting,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormatScript {
    None,
    Superscript,
    Subscript,
}

impl From<FormatScript> for RFormatScript {
    fn from(fs: FormatScript) -> Self {
        match fs {
            FormatScript::None => RFormatScript::None,
            FormatScript::Superscript => RFormatScript::Superscript,
            FormatScript::Subscript => RFormatScript::Subscript,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormatPattern {
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

impl From<FormatPattern> for RFormatPattern {
    fn from(fp: FormatPattern) -> Self {
        match fp {
            FormatPattern::None => RFormatPattern::None,
            FormatPattern::Solid => RFormatPattern::Solid,
            FormatPattern::MediumGray => RFormatPattern::MediumGray,
            FormatPattern::DarkGray => RFormatPattern::DarkGray,
            FormatPattern::LightGray => RFormatPattern::LightGray,
            FormatPattern::DarkHorizontal => RFormatPattern::DarkHorizontal,
            FormatPattern::DarkVertical => RFormatPattern::DarkVertical,
            FormatPattern::DarkDown => RFormatPattern::DarkDown,
            FormatPattern::DarkUp => RFormatPattern::DarkUp,
            FormatPattern::DarkGrid => RFormatPattern::DarkGrid,
            FormatPattern::DarkTrellis => RFormatPattern::DarkTrellis,
            FormatPattern::LightHorizontal => RFormatPattern::LightHorizontal,
            FormatPattern::LightVertical => RFormatPattern::LightVertical,
            FormatPattern::LightDown => RFormatPattern::LightDown,
            FormatPattern::LightUp => RFormatPattern::LightUp,
            FormatPattern::LightGrid => RFormatPattern::LightGrid,
            FormatPattern::LightTrellis => RFormatPattern::LightTrellis,
            FormatPattern::Gray125 => RFormatPattern::Gray125,
            FormatPattern::Gray0625 => RFormatPattern::Gray0625,
        }
    }
}

fn extract_color(val: &Bound<'_, PyAny>) -> Option<RColor> {
    if let Ok(c) = val.extract::<Color>() {
        return Some(RColor::from(c));
    }
    if let Ok(s) = val.extract::<String>() {
        return Some(RColor::from(s.as_str()));
    }
    None
}

#[pyclass]
pub struct Format {
    pub(crate) inner: RFormat,
    bold_val: bool,
    italic_val: bool,
    font_name_val: String,
    font_size_val: f64,
    font_color_val: Option<Py<PyAny>>,
    num_format_val: String,
    align_val: Option<Py<PyAny>>,
    text_wrap_val: bool,
    background_color_val: Option<Py<PyAny>>,
    foreground_color_val: Option<Py<PyAny>>,
    border_val: Option<Py<PyAny>>,
    border_color_val: Option<Py<PyAny>>,
    underline_val: Option<Py<PyAny>>,
    rotation_val: i16,
    indent_val: u8,
    shrink_val: bool,
    unlocked_val: bool,
    hidden_val: bool,
}

impl Clone for Format {
    fn clone(&self) -> Self {
        Python::attach(|py| Format {
            inner: self.inner.clone(),
            bold_val: self.bold_val,
            italic_val: self.italic_val,
            font_name_val: self.font_name_val.clone(),
            font_size_val: self.font_size_val,
            font_color_val: self.font_color_val.as_ref().map(|c| c.clone_ref(py)),
            num_format_val: self.num_format_val.clone(),
            align_val: self.align_val.as_ref().map(|a| a.clone_ref(py)),
            text_wrap_val: self.text_wrap_val,
            background_color_val: self.background_color_val.as_ref().map(|c| c.clone_ref(py)),
            foreground_color_val: self.foreground_color_val.as_ref().map(|c| c.clone_ref(py)),
            border_val: self.border_val.as_ref().map(|b| b.clone_ref(py)),
            border_color_val: self.border_color_val.as_ref().map(|c| c.clone_ref(py)),
            underline_val: self.underline_val.as_ref().map(|u| u.clone_ref(py)),
            rotation_val: self.rotation_val,
            indent_val: self.indent_val,
            shrink_val: self.shrink_val,
            unlocked_val: self.unlocked_val,
            hidden_val: self.hidden_val,
        })
    }
}

#[pymethods]
impl Format {
    #[new]
    pub fn new() -> Self {
        Format {
            inner: RFormat::new(),
            bold_val: false,
            italic_val: false,
            font_name_val: "Calibri".to_string(),
            font_size_val: 11.0,
            font_color_val: None,
            num_format_val: String::new(),
            align_val: None,
            text_wrap_val: false,
            background_color_val: None,
            foreground_color_val: None,
            border_val: None,
            border_color_val: None,
            underline_val: None,
            rotation_val: 0,
            indent_val: 0,
            shrink_val: false,
            unlocked_val: false,
            hidden_val: false,
        }
    }

    #[getter]
    pub fn bold(&self) -> bool {
        self.bold_val
    }

    #[setter]
    pub fn set_bold(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_bold();
        }
        self.bold_val = b;
        Ok(())
    }

    #[getter]
    pub fn italic(&self) -> bool {
        self.italic_val
    }

    #[setter]
    pub fn set_italic(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_italic();
        }
        self.italic_val = b;
        Ok(())
    }

    #[getter]
    pub fn font_name(&self) -> String {
        self.font_name_val.clone()
    }

    #[setter]
    pub fn set_font_name(&mut self, val: Bound<'_, PyAny>) {
        let name = val.to_string();
        let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
        self.inner = old_inner.set_font_name(&name);
        self.font_name_val = name;
    }

    #[getter]
    pub fn font_size(&self) -> f64 {
        self.font_size_val
    }

    #[setter]
    pub fn set_font_size(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let sz = val.extract::<f64>()?;
        let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
        self.inner = old_inner.set_font_size(sz);
        self.font_size_val = sz;
        Ok(())
    }

    #[getter]
    pub fn font_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.font_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_font_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.font_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_font_color(rcolor);
            self.font_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err("Invalid font_color"))
    }

    #[getter]
    pub fn num_format(&self) -> String {
        self.num_format_val.clone()
    }

    #[setter]
    pub fn set_num_format(&mut self, val: Bound<'_, PyAny>) {
        let fmt = val.to_string();
        let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
        self.inner = old_inner.set_num_format(&fmt);
        self.num_format_val = fmt;
    }

    #[getter]
    pub fn align(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.align_val.as_ref().map(|a| a.clone_ref(py))
    }

    #[setter]
    pub fn set_align(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.align_val = None;
            return Ok(());
        }
        if let Ok(a) = val.extract::<FormatAlign>() {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_align(RFormatAlign::from(a));
            self.align_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid FormatAlign",
        ))
    }

    #[getter]
    pub fn text_wrap(&self) -> bool {
        self.text_wrap_val
    }

    #[setter]
    pub fn set_text_wrap(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_text_wrap();
        }
        self.text_wrap_val = b;
        Ok(())
    }

    #[getter]
    pub fn background_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.background_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_background_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.background_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_background_color(rcolor);
            self.background_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid background_color",
        ))
    }

    #[getter]
    pub fn foreground_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.foreground_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_foreground_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.foreground_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_foreground_color(rcolor);
            self.foreground_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid foreground_color",
        ))
    }

    #[getter]
    pub fn border(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.border_val.as_ref().map(|b| b.clone_ref(py))
    }

    #[setter]
    pub fn set_border(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.border_val = None;
            return Ok(());
        }
        if let Ok(b) = val.extract::<FormatBorder>() {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_border(RFormatBorder::from(b));
            self.border_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid FormatBorder",
        ))
    }

    #[getter]
    pub fn border_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.border_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_border_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.border_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_border_color(rcolor);
            self.border_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid border_color",
        ))
    }

    #[getter]
    pub fn underline(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.underline_val.as_ref().map(|u| u.clone_ref(py))
    }

    #[setter]
    pub fn set_underline(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.underline_val = None;
            return Ok(());
        }
        if let Ok(u) = val.extract::<FormatUnderline>() {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_underline(RFormatUnderline::from(u));
            self.underline_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid FormatUnderline",
        ))
    }

    #[getter]
    pub fn rotation(&self) -> i16 {
        self.rotation_val
    }

    #[setter]
    pub fn set_rotation(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let rot = val.extract::<i16>()?;
        let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
        self.inner = old_inner.set_rotation(rot);
        self.rotation_val = rot;
        Ok(())
    }

    #[getter]
    pub fn indent(&self) -> u8 {
        self.indent_val
    }

    #[setter]
    pub fn set_indent(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let ind = val.extract::<u8>()?;
        let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
        self.inner = old_inner.set_indent(ind);
        self.indent_val = ind;
        Ok(())
    }

    #[getter]
    pub fn shrink(&self) -> bool {
        self.shrink_val
    }

    #[setter]
    pub fn set_shrink(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_shrink();
        }
        self.shrink_val = b;
        Ok(())
    }

    #[getter]
    pub fn unlocked(&self) -> bool {
        self.unlocked_val
    }

    #[setter]
    pub fn set_unlocked(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_unlocked();
        }
        self.unlocked_val = b;
        Ok(())
    }

    #[getter]
    pub fn hidden(&self) -> bool {
        self.hidden_val
    }

    #[setter]
    pub fn set_hidden(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        if b {
            let old_inner = std::mem::replace(&mut self.inner, RFormat::new());
            self.inner = old_inner.set_hidden();
        }
        self.hidden_val = b;
        Ok(())
    }
}
