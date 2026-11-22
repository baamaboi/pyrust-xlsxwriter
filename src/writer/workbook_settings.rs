//! Workbook-level settings: document properties, defined names, default
//! format and themes, calculation mode, zip options and the temporary
//! directory used by constant/low-memory worksheets.

use std::path::PathBuf;

use chrono::{NaiveDate, NaiveDateTime};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{
    PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyMapping, PyString, PyTzInfoAccess,
};
use pyo3::IntoPyObjectExt;
use rust_xlsxwriter::{
    CalculationMode as RCalculationMode, DocProperties as RDocProperties, FontScheme,
    Format as RFormat,
};

use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::Format;
use crate::writer::workbook::Workbook;
use crate::{py_enum, py_props};

py_enum!(
    /// When Excel recalculates formulas (`Workbook.set_calculation_mode()`).
    CalculationMode => RCalculationMode { Automatic, Manual, Partial }
);

// ---------------------------------------------------------------------------
// Date/time conversion for document properties
// ---------------------------------------------------------------------------

static UTC: PyOnceLock<Py<PyAny>> = PyOnceLock::new();

/// Document property dates are stored in UTC: aware datetimes are converted,
/// naive ones are taken to be UTC already, dates mean midnight.
fn utc_datetime(value: &Bound<'_, PyAny>) -> PyResult<NaiveDateTime> {
    let py = value.py();
    if let Ok(dt) = value.cast::<PyDateTime>() {
        if dt.get_tzinfo().is_some() {
            let utc = UTC.import(py, "datetime", "UTC")?;
            let converted = dt.call_method1("astimezone", (utc,))?;
            let kwargs = PyDict::new(py);
            kwargs.set_item("tzinfo", py.None())?;
            return converted.call_method("replace", (), Some(&kwargs))?.extract();
        }
        return dt.extract();
    }
    if let Ok(d) = value.cast::<PyDate>() {
        let date: NaiveDate = d.extract()?;
        return Ok(date.and_time(chrono::NaiveTime::MIN));
    }
    Err(PyTypeError::new_err(format!(
        "expected a datetime or date, not {}",
        value.get_type().name()?
    )))
}

/// A document property date/time (UTC). Accepts `datetime` or `date`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PropDateTime(pub NaiveDateTime);

impl<'a, 'py> FromPyObject<'a, 'py> for PropDateTime {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        Ok(PropDateTime(utc_datetime(&obj)?))
    }
}

impl<'py> IntoPyObject<'py> for PropDateTime {
    type Target = PyAny;
    type Output = Bound<'py, PyAny>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        self.0.into_bound_py_any(py)
    }
}

// ---------------------------------------------------------------------------
// Custom properties
// ---------------------------------------------------------------------------

/// The value of a custom document property.
#[derive(Clone, Debug, PartialEq)]
pub enum CustomValue {
    Text(String),
    Int(i32),
    Real(f64),
    Bool(bool),
    DateTime(NaiveDateTime),
}

impl<'a, 'py> FromPyObject<'a, 'py> for CustomValue {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(text) = obj.cast::<PyString>() {
            return Ok(CustomValue::Text(text.to_str()?.to_string()));
        }
        if let Ok(b) = obj.cast::<PyBool>() {
            return Ok(CustomValue::Bool(b.is_true()));
        }
        if obj.is_instance_of::<PyInt>() {
            // Excel's integer properties are 32-bit; store larger ints as numbers.
            return Ok(match obj.extract::<i32>() {
                Ok(i) => CustomValue::Int(i),
                Err(_) => CustomValue::Real(obj.extract::<f64>()?),
            });
        }
        if let Ok(f) = obj.cast::<PyFloat>() {
            return Ok(CustomValue::Real(f.value()));
        }
        if obj.cast::<PyDate>().is_ok() {
            return Ok(CustomValue::DateTime(utc_datetime(&obj)?));
        }
        Err(PyTypeError::new_err(format!(
            "custom property values must be str, int, float, bool, datetime or date, not {}",
            obj.get_type().name()?
        )))
    }
}

impl<'py> IntoPyObject<'py> for CustomValue {
    type Target = PyAny;
    type Output = Bound<'py, PyAny>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        match self {
            CustomValue::Text(text) => text.into_bound_py_any(py),
            CustomValue::Int(i) => i.into_bound_py_any(py),
            CustomValue::Real(f) => f.into_bound_py_any(py),
            CustomValue::Bool(b) => b.into_bound_py_any(py),
            CustomValue::DateTime(dt) => dt.into_bound_py_any(py),
        }
    }
}

/// Custom properties in insertion order; a dict on the Python side.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CustomProperties(pub Vec<(String, CustomValue)>);

impl CustomProperties {
    /// Add a property, replacing an existing one with the same name.
    fn set(&mut self, name: String, value: CustomValue) -> PyResult<()> {
        if name.is_empty() {
            return Err(param_err("custom property names can't be empty"));
        }
        match self.0.iter_mut().find(|(existing, _)| *existing == name) {
            Some(slot) => slot.1 = value,
            None => self.0.push((name, value)),
        }
        Ok(())
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for CustomProperties {
    type Error = PyErr;

    /// From a mapping `{name: value}` or an iterable of `(name, value)` pairs.
    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        let pairs = match obj.cast::<PyMapping>() {
            Ok(mapping) => mapping.items()?.into_any(),
            Err(_) => obj.to_owned(),
        };
        let mut out = CustomProperties::default();
        for item in pairs.try_iter()? {
            let (name, value): (String, CustomValue) = item?.extract()?;
            out.set(name, value)?;
        }
        Ok(out)
    }
}

impl<'py> IntoPyObject<'py> for CustomProperties {
    type Target = PyDict;
    type Output = Bound<'py, PyDict>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        let dict = PyDict::new(py);
        for (name, value) in self.0 {
            dict.set_item(name, value)?;
        }
        Ok(dict)
    }
}

// ---------------------------------------------------------------------------
// DocProperties
// ---------------------------------------------------------------------------

/// Plain description of the document properties (File > Info in Excel).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct DocPropertiesSpec {
    pub title: String,
    pub subject: String,
    pub manager: String,
    pub company: String,
    pub category: String,
    pub author: String,
    pub keywords: String,
    pub comment: String,
    pub status: String,
    pub hyperlink_base: String,
    /// None means "the time `set_properties()` is called".
    pub creation_datetime: Option<NaiveDateTime>,
    pub custom_properties: CustomProperties,
}

impl DocPropertiesSpec {
    /// Build the equivalent `rust_xlsxwriter::DocProperties`.
    pub fn to_rust(&self) -> RDocProperties {
        let mut p = RDocProperties::new()
            .set_title(&self.title)
            .set_subject(&self.subject)
            .set_manager(&self.manager)
            .set_company(&self.company)
            .set_category(&self.category)
            .set_author(&self.author)
            .set_keywords(&self.keywords)
            .set_comment(&self.comment)
            .set_status(&self.status)
            .set_hyperlink_base(&self.hyperlink_base);
        if let Some(dt) = &self.creation_datetime {
            p = p.set_creation_datetime(dt);
        }
        for (name, value) in &self.custom_properties.0 {
            p = match value {
                CustomValue::Text(text) => p.set_custom_property(name, text.as_str()),
                CustomValue::Int(i) => p.set_custom_property(name, *i),
                CustomValue::Real(f) => p.set_custom_property(name, *f),
                CustomValue::Bool(b) => p.set_custom_property(name, *b),
                CustomValue::DateTime(dt) => p.set_custom_property(name, dt),
            };
        }
        p
    }
}

/// Document metadata shown in Excel's File > Info and the file's properties.
/// Every property is also a constructor keyword:
///
/// ```python
/// props = DocProperties(title="Q3 report", author="Finance", company="ACME")
/// props.set_custom_property("Checked by", "Ann").set_custom_property("Revision", 3)
/// wb.set_properties(props)
/// ```
///
/// `creation_datetime` is stored in UTC: naive datetimes are taken as UTC,
/// aware ones are converted. It defaults to the time `set_properties()` is
/// called. Custom property values may be str, int, float, bool, datetime or
/// date (ints outside the 32-bit range are stored as numbers).
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct DocProperties {
    pub spec: DocPropertiesSpec,
}

impl DocProperties {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    fn get_creation_datetime(&self) -> Option<PropDateTime> {
        self.spec.creation_datetime.map(PropDateTime)
    }

    fn set_creation_datetime(&mut self, value: Option<PropDateTime>) -> PyResult<()> {
        self.spec.creation_datetime = value.map(|v| v.0);
        Ok(())
    }

    fn get_custom_properties(&self) -> CustomProperties {
        self.spec.custom_properties.clone()
    }

    fn set_custom_properties(&mut self, value: CustomProperties) -> PyResult<()> {
        self.spec.custom_properties = value;
        Ok(())
    }
}

py_props!(DocProperties
    fields {
        title: String,
        subject: String,
        manager: String,
        company: String,
        category: String,
        author: String,
        keywords: String,
        comment: String,
        status: String,
        hyperlink_base: String,
    }
    custom {
        creation_datetime: Option<PropDateTime>,
        custom_properties: CustomProperties,
    }
);

#[pymethods]
impl DocProperties {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut properties = DocProperties { spec: DocPropertiesSpec::default() };
        properties.apply_kwargs(kwargs)?;
        Ok(properties)
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

    /// Add (or replace) a custom property and return self (for chaining).
    /// `value` may be str, int, float, bool, datetime or date.
    fn set_custom_property<'py>(
        mut slf: PyRefMut<'py, Self>,
        name: String,
        value: CustomValue,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.spec.custom_properties.set(name, value)?;
        Ok(slf)
    }

    /// Remove a custom property; KeyError if it doesn't exist.
    fn remove_custom_property(&mut self, name: &str) -> PyResult<()> {
        let props = &mut self.spec.custom_properties.0;
        let index = props.iter().position(|(existing, _)| existing == name).ok_or_else(|| {
            pyo3::exceptions::PyKeyError::new_err(format!("no custom property named {name:?}"))
        })?;
        props.remove(index);
        Ok(())
    }

    /// A copy of these properties.
    fn copy(&self) -> Self {
        self.clone()
    }

    fn __copy__(&self) -> Self {
        self.clone()
    }

    fn __deepcopy__(&self, _memo: &Bound<'_, PyAny>) -> Self {
        self.clone()
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<DocProperties>() {
            Ok(other) => self.spec == other.borrow().spec,
            Err(_) => false,
        }
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let s = &self.spec;
        let mut parts = Vec::new();
        for (name, value) in [
            ("title", &s.title),
            ("subject", &s.subject),
            ("manager", &s.manager),
            ("company", &s.company),
            ("category", &s.category),
            ("author", &s.author),
            ("keywords", &s.keywords),
            ("comment", &s.comment),
            ("status", &s.status),
            ("hyperlink_base", &s.hyperlink_base),
        ] {
            if !value.is_empty() {
                parts.push(format!("{name}={}", PyString::new(py, value).repr()?));
            }
        }
        if let Some(dt) = s.creation_datetime {
            parts.push(format!("creation_datetime={}", dt.into_bound_py_any(py)?.repr()?));
        }
        if !s.custom_properties.0.is_empty() {
            let dict = s.custom_properties.clone().into_pyobject(py)?;
            parts.push(format!("custom_properties={}", dict.repr()?));
        }
        Ok(format!("DocProperties({})", parts.join(", ")))
    }
}

// ---------------------------------------------------------------------------
// Workbook methods
// ---------------------------------------------------------------------------

/// Column widths (pixels) Excel supports for a default format.
const DEFAULT_COL_WIDTHS: [u32; 7] = [56, 64, 72, 80, 96, 104, 120];

impl Workbook {
    fn require_no_worksheets(&self, method: &str) -> PyResult<()> {
        if self.worksheets.is_empty() {
            Ok(())
        } else {
            Err(param_err(format!(
                "{method}() must be called before any worksheet is added to the workbook"
            )))
        }
    }

    fn apply_default_format(&mut self, format: RFormat, row_height: u32, col_width: u32) -> PyResult<()> {
        self.inner.set_default_format(&format, row_height, col_width).py()?;
        self.default_format = Some((format, row_height, col_width));
        Ok(())
    }
}

#[pymethods]
impl Workbook {
    /// Set the document properties (title, author, custom properties, ...)
    /// from a `DocProperties` or a dict of its keywords.
    fn set_properties(&mut self, properties: &Bound<'_, PyAny>) -> PyResult<()> {
        let spec = match properties.cast::<DocProperties>() {
            Ok(p) => p.borrow().spec.clone(),
            Err(_) => {
                let kwargs = properties.cast::<PyDict>().map_err(|_| {
                    PyTypeError::new_err("set_properties() expects a DocProperties or a dict")
                })?;
                DocProperties::py_new(Some(kwargs))?.spec
            }
        };
        self.inner.set_properties(&spec.to_rust());
        Ok(())
    }

    /// Define a name for a range, constant or formula: `define_name("Rate",
    /// "=0.05")`, `define_name("Sales", "=Sheet1!$A$1:$A$10")`. Prefix the
    /// name with a sheet name (`"Sheet1!Local"`) for a worksheet-local name.
    fn define_name(&mut self, name: &str, formula: &str) -> PyResult<()> {
        self.inner.define_name(name, formula).py()?;
        Ok(())
    }

    /// Change the workbook's default cell format (normally Calibri 11).
    /// `row_height` is the default row height and `col_width` the default
    /// column width, both in pixels; `col_width` must be one of 56, 64, 72,
    /// 80, 96, 104 or 120. Must be called before any worksheet is added.
    #[pyo3(signature = (format, row_height=20, col_width=64))]
    fn set_default_format(
        &mut self,
        format: PyRef<'_, Format>,
        row_height: u32,
        col_width: u32,
    ) -> PyResult<()> {
        self.require_no_worksheets("set_default_format")?;
        if !DEFAULT_COL_WIDTHS.contains(&col_width) {
            return Err(param_err(format!(
                "default column width {col_width} is not supported: use one of {DEFAULT_COL_WIDTHS:?} pixels"
            )));
        }
        if row_height == 0 || row_height > 545 {
            return Err(param_err(format!(
                "default row height {row_height} is outside 1..=545 pixels"
            )));
        }
        self.apply_default_format(format.inner.clone(), row_height, col_width)
    }

    /// Use Excel 2023's "Office" theme and its default font (Aptos Narrow 11)
    /// instead of the 2007 theme (Calibri 11). Must be called before any
    /// worksheet is added.
    fn use_excel_2023_theme(&mut self) -> PyResult<()> {
        self.require_no_worksheets("use_excel_2023_theme")?;
        self.inner.use_excel_2023_theme().py()?;
        // Mirror the default format rust_xlsxwriter sets so that worksheets
        // created by this binding use it too.
        let format = RFormat::new()
            .set_font_name("Aptos Narrow")
            .set_font_size(11)
            .set_font_scheme(FontScheme::Body);
        self.default_format = Some((format, 20, 64));
        Ok(())
    }

    /// Use a custom theme from a `.thmx` file, an `.xlsx` file or an
    /// extracted `theme1.xml`. Also call `set_default_format()` with the
    /// theme's body font (`font_scheme=FontScheme.Body`).
    fn use_custom_theme(&mut self, path: PathBuf) -> PyResult<()> {
        self.inner.use_custom_theme(&path).py()?;
        Ok(())
    }

    /// Ask Excel to suggest opening the file read-only.
    fn read_only_recommended(&mut self) {
        self.inner.read_only_recommended();
    }

    /// Use ZIP64 extensions so that worksheets larger than 4 GB uncompressed
    /// can be saved (Excel itself can't open such files reliably).
    #[pyo3(signature = (enable=true))]
    fn use_zip_large_file(&mut self, enable: bool) {
        self.inner.use_zip_large_file(enable);
    }

    /// When Excel recalculates formulas (default: `CalculationMode.Automatic`).
    fn set_calculation_mode(&mut self, mode: CalculationMode) {
        self.inner.set_calculation_mode(mode.into());
    }

    /// Width of the worksheet tab bar as a percentage (0..=100) of the
    /// horizontal scroll bar area (Excel's default is 60).
    fn set_tab_ratio(&mut self, ratio: u8) -> PyResult<()> {
        if ratio > 100 {
            return Err(param_err(format!("tab ratio {ratio} is outside 0..=100")));
        }
        self.inner.set_tab_ratio(ratio);
        Ok(())
    }

    /// ZIP compression level 1 (fastest) to 9 (smallest); the default is 6.
    fn set_zip_compression_level(&mut self, level: u8) -> PyResult<()> {
        if !(1..=9).contains(&level) {
            return Err(param_err(format!("zip compression level {level} is outside 1..=9")));
        }
        self.inner.set_zip_compression_level(level);
        Ok(())
    }

    /// Directory for the temporary files of constant/low-memory worksheets
    /// (default: the system temp directory). It must exist and be writable;
    /// set it before adding those worksheets.
    fn set_tempdir(&mut self, path: PathBuf) -> PyResult<()> {
        self.inner.set_tempdir(&path).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<DocProperties>()?;
    m.add_class::<CalculationMode>()?;
    Ok(())
}
