//! Worksheet autofilters: `FilterCondition`, `FilterCriteria` and the
//! `Worksheet.autofilter()`, `filter_column()` and `filter_automatic_off()`
//! methods.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyInt, PyString, PyTuple};
use pyo3::IntoPyObjectExt;
use rust_xlsxwriter::{FilterCondition as RFilterCondition, FilterCriteria as RFilterCriteria};

use crate::common::cellref::{split_cols, split_range, Params};
use crate::common::errors::{param_err, IntoPyResult};
use crate::writer::conditional_format::spec_methods;
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// Operators for custom autofilter conditions (`FilterCondition.add_custom_filter`).
    FilterCriteria => RFilterCriteria {
        EqualTo, NotEqualTo, GreaterThan, GreaterThanOrEqualTo, LessThan, LessThanOrEqualTo,
        BeginsWith, DoesNotBeginWith, EndsWith, DoesNotEndWith, Contains, DoesNotContain,
    }
);

/// A filter value: a string or a number.
#[derive(Clone, Debug, PartialEq)]
pub enum FilterValue {
    Text(String),
    Int(i64),
    Number(f64),
}

impl<'a, 'py> FromPyObject<'a, 'py> for FilterValue {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(s) = obj.cast::<PyString>() {
            return Ok(FilterValue::Text(s.to_str()?.to_string()));
        }
        if !obj.is_instance_of::<PyBool>() {
            if obj.is_instance_of::<PyInt>()
                && let Ok(i) = obj.extract::<i64>() {
                    return Ok(FilterValue::Int(i));
                }
            if obj.is_instance_of::<PyInt>() || obj.is_instance_of::<PyFloat>() {
                let value: f64 = obj.extract()?;
                if !value.is_finite() {
                    return Err(param_err(format!("filter value {value} is not a finite number")));
                }
                return Ok(FilterValue::Number(value));
            }
        }
        Err(PyTypeError::new_err(format!(
            "filter values must be str, int or float, not {}",
            obj.get_type().name()?
        )))
    }
}

impl<'py> IntoPyObject<'py> for FilterValue {
    type Target = PyAny;
    type Output = Bound<'py, PyAny>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        match self {
            FilterValue::Text(s) => s.into_bound_py_any(py),
            FilterValue::Int(i) => i.into_bound_py_any(py),
            FilterValue::Number(n) => n.into_bound_py_any(py),
        }
    }
}

/// Plain description of a filter condition.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct FilterConditionSpec {
    pub list_filters: Vec<FilterValue>,
    pub blanks: bool,
    pub custom_filters: Vec<(FilterCriteria, FilterValue)>,
    pub boolean_or: bool,
}

impl FilterConditionSpec {
    fn validate(&self) -> PyResult<()> {
        if self.custom_filters.len() > 2 {
            return Err(param_err(format!(
                "Excel allows at most 2 custom filters per column, got {}",
                self.custom_filters.len()
            )));
        }
        if !self.custom_filters.is_empty() && (!self.list_filters.is_empty() || self.blanks) {
            return Err(param_err(
                "a FilterCondition is either a list filter (list_filters/blanks) or a custom filter \
                 (custom_filters), not both",
            ));
        }
        Ok(())
    }

    /// Build the equivalent `rust_xlsxwriter::FilterCondition`.
    pub fn to_rust(&self) -> PyResult<RFilterCondition> {
        self.validate()?;
        if self.list_filters.is_empty() && self.custom_filters.is_empty() && !self.blanks {
            return Err(param_err(
                "the FilterCondition has no conditions: add list_filters, blanks or custom_filters",
            ));
        }
        let mut fc = RFilterCondition::new();
        for value in &self.list_filters {
            fc = match value {
                FilterValue::Text(s) => fc.add_list_filter(s.as_str()),
                FilterValue::Int(i) => match i32::try_from(*i) {
                    Ok(i) => fc.add_list_filter(i),
                    Err(_) => fc.add_list_filter(*i as f64),
                },
                FilterValue::Number(n) => fc.add_list_filter(*n),
            };
        }
        if self.blanks {
            fc = fc.add_list_blanks_filter();
        }
        for (criteria, value) in &self.custom_filters {
            let criteria: RFilterCriteria = (*criteria).into();
            fc = match value {
                FilterValue::Text(s) => fc.add_custom_filter(criteria, s.as_str()),
                FilterValue::Int(i) => match i32::try_from(*i) {
                    Ok(i) => fc.add_custom_filter(criteria, i),
                    Err(_) => fc.add_custom_filter(criteria, *i as f64),
                },
                FilterValue::Number(n) => fc.add_custom_filter(criteria, *n),
            };
        }
        if self.boolean_or {
            fc = fc.add_custom_boolean_or();
        }
        Ok(fc)
    }
}

/// The filter applied to one column of an autofilter range.
///
/// A list filter shows rows equal to any of `list_filters` (and blanks with
/// `blanks=True`). A custom filter has up to two `(FilterCriteria, value)`
/// conditions in `custom_filters`, combined with "and" (or "or" with
/// `boolean_or=True`). The chainable `add_*()` methods mirror rust_xlsxwriter:
///
/// ```python
/// FilterCondition().add_list_filter("East").add_list_filter("West")
/// FilterCondition(custom_filters=[(FilterCriteria.GreaterThan, 4000)])
/// ```
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct FilterCondition {
    pub spec: FilterConditionSpec,
}

impl FilterCondition {
    fn on_change(&mut self) -> PyResult<()> {
        self.spec.validate()
    }
}

py_props!(FilterCondition
    fields {
        list_filters: Vec<FilterValue>,
        blanks: bool,
        custom_filters: Vec<(FilterCriteria, FilterValue)>,
        boolean_or: bool,
    }
);

#[pymethods]
impl FilterCondition {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut fc = FilterCondition { spec: FilterConditionSpec::default() };
        fc.apply_kwargs(kwargs)?;
        Ok(fc)
    }

    /// Show rows equal to `value` (call repeatedly for "or"); returns self.
    fn add_list_filter(mut slf: PyRefMut<'_, Self>, value: FilterValue) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.list_filters.push(value);
        if let Err(e) = slf.on_change() {
            slf.spec.list_filters.pop();
            return Err(e);
        }
        Ok(slf)
    }

    /// Also show blank cells in a list filter; returns self.
    fn add_list_blanks_filter(mut slf: PyRefMut<'_, Self>) -> PyResult<PyRefMut<'_, Self>> {
        let before = slf.spec.blanks;
        slf.spec.blanks = true;
        if let Err(e) = slf.on_change() {
            slf.spec.blanks = before;
            return Err(e);
        }
        Ok(slf)
    }

    /// Add a custom condition (at most two); returns self.
    fn add_custom_filter(
        mut slf: PyRefMut<'_, Self>,
        criteria: FilterCriteria,
        value: FilterValue,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.custom_filters.push((criteria, value));
        if let Err(e) = slf.on_change() {
            slf.spec.custom_filters.pop();
            return Err(e);
        }
        Ok(slf)
    }

    /// Combine two custom conditions with "or" instead of "and"; returns self.
    fn add_custom_boolean_or(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.spec.boolean_or = true;
        slf
    }
}

spec_methods!(FilterCondition);

#[pymethods]
impl Worksheet {
    /// Add an autofilter to a range: `autofilter("A1:D20")` or
    /// `autofilter(first_row, first_col, last_row, last_col)`.
    #[pyo3(signature = (*args))]
    fn autofilter(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        Params::new("autofilter", rest, None, &[], 0)?;
        self.inner.autofilter(r1, c1, r2, c2).py()?;
        Ok(())
    }

    /// Filter one column of the autofilter range: `filter_column("B", condition)`
    /// or `filter_column(1, condition)`. Rows that don't match are hidden
    /// automatically unless `filter_automatic_off()` is called.
    #[pyo3(signature = (*args, **kwargs))]
    fn filter_column(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        if first != last {
            return Err(param_err("filter_column() takes a single column, not a column span"));
        }
        let p = Params::new("filter_column", rest, kwargs, &["condition"], 1)?;
        let condition: PyRef<'_, FilterCondition> = p.required(0)?;
        let rust = condition.spec.to_rust()?;
        self.inner.filter_column(first, &rust).py()?;
        Ok(())
    }

    /// Don't hide rows that don't match the filters (hide them yourself).
    fn filter_automatic_off(&mut self) {
        self.inner.filter_automatic_off();
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<FilterCondition>()?;
    m.add_class::<FilterCriteria>()?;
    Ok(())
}
