//! Converting Python values to cell values (and back).
//!
//! `classify()` decides how a Python object is stored in a cell. The order
//! matters: `bool` is checked before `int` (bool is an int subclass), and
//! `datetime` before `date`.
//!
//! | Python                         | Cell                                   |
//! |--------------------------------|----------------------------------------|
//! | `None`, `""`                   | blank                                  |
//! | `bool`                         | boolean                                |
//! | `int`, `float`, numpy scalars  | number                                 |
//! | `decimal.Decimal`              | number                                 |
//! | `str` starting with `=`        | formula (`{=...}` is an array formula) |
//! | other `str`                    | string                                 |
//! | `datetime`, `date`, `time`     | date/time number with a date format    |
//! | `timedelta`                    | number of days, `[h]:mm:ss`            |
//! | `Formula`, `Url`               | formula with cached result, hyperlink  |

use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{
    PyBool, PyDate, PyDateTime, PyDelta, PyDeltaAccess, PyFloat, PyInt, PyString, PyTime, PyType,
    PyTzInfoAccess,
};
use rust_decimal::Decimal;
use rust_xlsxwriter::{Formula as RFormula, Url as RUrl};

use crate::common::errors::param_err;

/// A formula with an optional cached result, e.g. `Formula("=SUM(A1:A3)", result="6")`.
///
/// Excel recalculates formulas when a file opens, but other readers show the
/// cached result, so supplying one is useful for files read by tools.
#[pyclass(module = "pyrust_xlsxwriter", frozen, eq, from_py_object)]
#[derive(Clone, Debug, PartialEq)]
pub struct Formula {
    #[pyo3(get)]
    pub formula: String,
    #[pyo3(get)]
    pub result: Option<String>,
}

impl Formula {
    pub fn to_rust(&self) -> RFormula {
        let formula = RFormula::new(&self.formula);
        match &self.result {
            Some(result) => formula.set_result(result),
            None => formula,
        }
    }
}

#[pymethods]
impl Formula {
    #[new]
    #[pyo3(signature = (formula, result=None))]
    fn new(formula: String, result: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let result = match result {
            None => None,
            Some(value) if value.is_none() => None,
            Some(value) => Some(value.str()?.to_string()),
        };
        Ok(Formula { formula, result })
    }

    fn __repr__(&self) -> String {
        match &self.result {
            Some(result) => format!("Formula({:?}, result={:?})", self.formula, result),
            None => format!("Formula({:?})", self.formula),
        }
    }
}

/// A hyperlink, e.g. `Url("https://example.com", text="Example", tip="Open")`.
///
/// Supported link types: `http://`, `https://`, `ftp://`, `mailto:`,
/// `file://` and internal links such as `internal:Sheet2!A1`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, eq, from_py_object)]
#[derive(Clone, Debug, PartialEq)]
pub struct Url {
    #[pyo3(get)]
    pub link: String,
    #[pyo3(get)]
    pub text: Option<String>,
    #[pyo3(get)]
    pub tip: Option<String>,
}

impl Url {
    pub fn to_rust(&self) -> RUrl {
        let mut url = RUrl::new(&self.link);
        if let Some(text) = &self.text {
            url = url.set_text(text);
        }
        if let Some(tip) = &self.tip {
            url = url.set_tip(tip);
        }
        url
    }
}

#[pymethods]
impl Url {
    #[new]
    #[pyo3(signature = (link, text=None, tip=None))]
    fn new(link: String, text: Option<String>, tip: Option<String>) -> Self {
        Url { link, text, tip }
    }

    fn __repr__(&self) -> String {
        format!("Url({:?}, text={:?}, tip={:?})", self.link, self.text, self.tip)
    }
}

/// A Python value classified for writing to a cell.
pub enum CellValue<'py> {
    Blank,
    Bool(bool),
    Number(f64),
    Decimal(Decimal),
    String(Bound<'py, PyString>),
    Formula(String),
    ArrayFormula(String),
    FormulaObject(Formula),
    Url(Url),
    DateTime(NaiveDateTime),
    Date(NaiveDate),
    Time(NaiveTime),
    /// A duration in days.
    Duration(f64),
}

static DECIMAL_TYPE: PyOnceLock<Py<PyType>> = PyOnceLock::new();

fn decimal_type(py: Python<'_>) -> PyResult<&Bound<'_, PyType>> {
    DECIMAL_TYPE.import(py, "decimal", "Decimal")
}

/// Convert a Python datetime to a naive one, rejecting time zones.
pub fn naive_datetime(dt: &Bound<'_, PyDateTime>) -> PyResult<NaiveDateTime> {
    if dt.get_tzinfo().is_some() {
        return Err(param_err(
            "Excel has no time zones: convert the datetime to naive local time first, \
             e.g. dt.astimezone(tz).replace(tzinfo=None)",
        ));
    }
    dt.extract::<NaiveDateTime>()
}

fn duration_days(delta: &Bound<'_, PyDelta>) -> f64 {
    delta.get_days() as f64
        + (delta.get_seconds() as f64 + delta.get_microseconds() as f64 / 1e6) / 86_400.0
}

/// Classify a Python object for writing to a cell.
pub fn classify<'py>(value: &Bound<'py, PyAny>) -> PyResult<CellValue<'py>> {
    // Fast paths for the common exact types.
    if let Ok(f) = value.cast_exact::<PyFloat>() {
        return Ok(CellValue::Number(f.value()));
    }
    if let Ok(s) = value.cast_exact::<PyString>() {
        return Ok(classify_str(s.clone()));
    }
    if let Ok(b) = value.cast::<PyBool>() {
        return Ok(CellValue::Bool(b.is_true()));
    }
    if value.is_instance_of::<PyInt>() {
        return Ok(match value.extract::<i64>() {
            Ok(i) => CellValue::Number(i as f64),
            Err(_) => CellValue::Number(value.extract::<f64>()?),
        });
    }
    if value.is_none() {
        return Ok(CellValue::Blank);
    }
    // datetime must be checked before date (datetime is a date subclass).
    if let Ok(dt) = value.cast::<PyDateTime>() {
        return Ok(CellValue::DateTime(naive_datetime(dt)?));
    }
    if let Ok(d) = value.cast::<PyDate>() {
        return Ok(CellValue::Date(d.extract::<NaiveDate>()?));
    }
    if let Ok(t) = value.cast::<PyTime>() {
        if t.get_tzinfo().is_some() {
            return Err(param_err("Excel has no time zones: pass a naive time"));
        }
        return Ok(CellValue::Time(t.extract::<NaiveTime>()?));
    }
    if let Ok(delta) = value.cast::<PyDelta>() {
        return Ok(CellValue::Duration(duration_days(delta)));
    }
    if let Ok(s) = value.cast::<PyString>() {
        return Ok(classify_str(s.clone()));
    }
    if let Ok(f) = value.cast::<PyFloat>() {
        return Ok(CellValue::Number(f.value()));
    }
    if let Ok(formula) = value.cast::<Formula>() {
        return Ok(CellValue::FormulaObject(formula.get().clone()));
    }
    if let Ok(url) = value.cast::<Url>() {
        return Ok(CellValue::Url(url.get().clone()));
    }
    let py = value.py();
    if value.is_instance(decimal_type(py)?)? {
        return Ok(match value.extract::<Decimal>() {
            Ok(d) => CellValue::Decimal(d),
            // Outside rust_decimal's range (or NaN/Infinity): use the float value.
            Err(_) => CellValue::Number(value.extract::<f64>()?),
        });
    }
    // numpy and other numeric scalars: bool first, then ints, then floats.
    if value.get_type().module().is_ok_and(|m| m == "numpy")
        && let Ok(b) = value.extract::<bool>() {
            return Ok(CellValue::Bool(b));
        }
    if value.hasattr("__index__")?
        && let Ok(i) = value.call_method0("__index__").and_then(|i| i.extract::<i64>()) {
            return Ok(CellValue::Number(i as f64));
        }
    if value.hasattr("__float__")? {
        return Ok(CellValue::Number(value.call_method0("__float__")?.extract::<f64>()?));
    }
    Err(PyTypeError::new_err(format!(
        "unsupported type for a cell value: {} (convert it to str, int, float, bool, date/time or None)",
        value.get_type().name()?
    )))
}

fn classify_str(s: Bound<'_, PyString>) -> CellValue<'_> {
    let text = match s.to_str() {
        Ok(text) => text,
        Err(_) => return CellValue::String(s),
    };
    if text.is_empty() {
        CellValue::Blank
    } else if text.len() > 1 && text.starts_with('=') {
        CellValue::Formula(text.to_string())
    } else if text.len() > 3 && text.starts_with("{=") && text.ends_with('}') {
        CellValue::ArrayFormula(text[1..text.len() - 1].to_string())
    } else {
        CellValue::String(s)
    }
}

/// Excel serial date (1900 date system) for a datetime.
pub fn excel_serial(dt: &NaiveDateTime) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31)
        .and_then(|d| d.and_hms_opt(0, 0, 0))
        .expect("valid epoch");
    let delta = *dt - epoch;
    let mut days = delta.num_days() as f64;
    let secs = (delta.num_milliseconds() - delta.num_days() * 86_400_000) as f64 / 1000.0;
    // Excel's fictitious 1900-02-29.
    if days >= 60.0 {
        days += 1.0;
    }
    days + secs / 86_400.0
}

pub fn time_fraction(t: &NaiveTime) -> f64 {
    (t.num_seconds_from_midnight() as f64 + t.nanosecond() as f64 / 1e9) / 86_400.0
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Formula>()?;
    m.add_class::<Url>()?;
    Ok(())
}
