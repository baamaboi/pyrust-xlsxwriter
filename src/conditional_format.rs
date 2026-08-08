use pyo3::prelude::*;
use rust_xlsxwriter::{
    Color as RColor, ConditionalFormat2ColorScale as R2ColorScale,
    ConditionalFormat3ColorScale as R3ColorScale, ConditionalFormatDataBar as RDataBar,
};

use crate::workbook::Color;

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
pub struct ConditionalFormat2ColorScale {
    pub(crate) inner: R2ColorScale,
    min_color_val: Option<Py<PyAny>>,
    max_color_val: Option<Py<PyAny>>,
}

#[pymethods]
impl ConditionalFormat2ColorScale {
    #[new]
    pub fn new() -> Self {
        ConditionalFormat2ColorScale {
            inner: R2ColorScale::new(),
            min_color_val: None,
            max_color_val: None,
        }
    }

    #[getter]
    pub fn min_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.min_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_min_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.min_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, R2ColorScale::new());
            self.inner = old_inner.set_minimum_color(rcolor);
            self.min_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "min_color must be a Color enum or string",
        ))
    }

    #[getter]
    pub fn max_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.max_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_max_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.max_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, R2ColorScale::new());
            self.inner = old_inner.set_maximum_color(rcolor);
            self.max_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "max_color must be a Color enum or string",
        ))
    }
}

#[pyclass]
pub struct ConditionalFormat3ColorScale {
    pub(crate) inner: R3ColorScale,
    min_color_val: Option<Py<PyAny>>,
    mid_color_val: Option<Py<PyAny>>,
    max_color_val: Option<Py<PyAny>>,
}

#[pymethods]
impl ConditionalFormat3ColorScale {
    #[new]
    pub fn new() -> Self {
        ConditionalFormat3ColorScale {
            inner: R3ColorScale::new(),
            min_color_val: None,
            mid_color_val: None,
            max_color_val: None,
        }
    }

    #[getter]
    pub fn min_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.min_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_min_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.min_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, R3ColorScale::new());
            self.inner = old_inner.set_minimum_color(rcolor);
            self.min_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "min_color must be a Color enum or string",
        ))
    }

    #[getter]
    pub fn mid_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.mid_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_mid_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.mid_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, R3ColorScale::new());
            self.inner = old_inner.set_midpoint_color(rcolor);
            self.mid_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "mid_color must be a Color enum or string",
        ))
    }

    #[getter]
    pub fn max_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.max_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_max_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.max_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, R3ColorScale::new());
            self.inner = old_inner.set_maximum_color(rcolor);
            self.max_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "max_color must be a Color enum or string",
        ))
    }
}

#[pyclass]
pub struct ConditionalFormatDataBar {
    pub(crate) inner: RDataBar,
    fill_color_val: Option<Py<PyAny>>,
}

#[pymethods]
impl ConditionalFormatDataBar {
    #[new]
    pub fn new() -> Self {
        ConditionalFormatDataBar {
            inner: RDataBar::new(),
            fill_color_val: None,
        }
    }

    #[getter]
    pub fn fill_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.fill_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_fill_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.fill_color_val = None;
            return Ok(());
        }
        if let Some(rcolor) = extract_color(&val) {
            let old_inner = std::mem::replace(&mut self.inner, RDataBar::new());
            self.inner = old_inner.set_fill_color(rcolor);
            self.fill_color_val = Some(val.unbind());
            return Ok(());
        }
        Err(pyo3::exceptions::PyTypeError::new_err(
            "fill_color must be a Color enum or string",
        ))
    }
}
