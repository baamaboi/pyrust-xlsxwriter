//! Conditional formats: every rust_xlsxwriter conditional format type, their
//! rule/type enums and `Worksheet.add_conditional_format()`.
//!
//! Each conditional format class keeps its settings in a plain spec struct
//! and builds the `rust_xlsxwriter` object from it when it is added to a
//! worksheet. Rules that carry values (`ConditionalFormatCellRule.between(1, 10)`,
//! `ConditionalFormatTopRule.top(5)`, ...) are small frozen classes with
//! static constructors.
//!
//! This module also holds helpers shared with the other data features:
//! [`RuleValue`] (a number/string/date/formula rule value), the comparison
//! rule classes, [`FormulaText`] and the `spec_methods!` macro.

use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyString, PyTime, PyTuple, PyTzInfoAccess,
};
use pyo3::IntoPyObjectExt;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale as R2ColorScale, ConditionalFormat3ColorScale as R3ColorScale,
    ConditionalFormatAverage as RAverage, ConditionalFormatAverageRule as RAverageRule,
    ConditionalFormatBlank as RBlank, ConditionalFormatCell as RCell,
    ConditionalFormatCellRule as RCellRule, ConditionalFormatCustomIcon as RCustomIcon,
    ConditionalFormatDataBar as RDataBar,
    ConditionalFormatDataBarAxisPosition as RDataBarAxisPosition,
    ConditionalFormatDataBarDirection as RDataBarDirection, ConditionalFormatDate as RDate,
    ConditionalFormatDateRule as RDateRule, ConditionalFormatDuplicate as RDuplicate,
    ConditionalFormatError as RError, ConditionalFormatFormula as RFormulaCf,
    ConditionalFormatIconSet as RIconSet, ConditionalFormatIconType as RIconType,
    ConditionalFormatText as RText, ConditionalFormatTextRule as RTextRule,
    ConditionalFormatTop as RTop, ConditionalFormatTopRule as RTopRule,
    ConditionalFormatType as RType, ConditionalFormatValue as RValue, Formula as RFormula,
    Worksheet as RWorksheet,
};

use crate::common::cellref::{split_range, CellRange, Params};
use crate::common::color::Color;
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::{Format, FormatSpec};
use crate::common::value::{naive_datetime, Formula};
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

// ===========================================================================
// Shared helpers (also used by data_validation.rs, table.rs, ...)
// ===========================================================================

/// A value used in a rule: a number, string, boolean, date/time or formula.
///
/// Python `str` values that start with `=` are formulas; other strings are
/// literal text. `Formula("...")` objects are always formulas.
#[derive(Clone, Debug, PartialEq)]
pub enum RuleValue {
    Int(i64),
    Number(f64),
    Bool(bool),
    Text(String),
    /// A formula; `object` records whether it was given as a `Formula`.
    Formula { text: String, object: bool },
    DateTime(NaiveDateTime),
    Date(NaiveDate),
    Time(NaiveTime),
}

impl RuleValue {
    pub(crate) fn from_py(value: &Bound<'_, PyAny>) -> PyResult<Self> {
        if let Ok(b) = value.cast::<PyBool>() {
            return Ok(RuleValue::Bool(b.is_true()));
        }
        if value.is_instance_of::<PyInt>() {
            return Ok(match value.extract::<i64>() {
                Ok(i) => RuleValue::Int(i),
                Err(_) => RuleValue::Number(finite(value.extract::<f64>()?)?),
            });
        }
        if let Ok(f) = value.cast::<PyFloat>() {
            return Ok(RuleValue::Number(finite(f.value())?));
        }
        if let Ok(s) = value.cast::<PyString>() {
            let s = s.to_str()?;
            if s.len() > 1 && s.starts_with('=') {
                return Ok(RuleValue::Formula { text: s.to_string(), object: false });
            }
            return Ok(RuleValue::Text(s.to_string()));
        }
        if let Ok(f) = value.cast::<Formula>() {
            return Ok(RuleValue::Formula { text: f.get().formula.clone(), object: true });
        }
        if let Ok(dt) = value.cast::<PyDateTime>() {
            return Ok(RuleValue::DateTime(naive_datetime(dt)?));
        }
        if let Ok(d) = value.cast::<PyDate>() {
            return Ok(RuleValue::Date(d.extract::<NaiveDate>()?));
        }
        if let Ok(t) = value.cast::<PyTime>() {
            if t.get_tzinfo().is_some() {
                return Err(param_err("Excel has no time zones: pass a naive time"));
            }
            return Ok(RuleValue::Time(t.extract::<NaiveTime>()?));
        }
        // decimal.Decimal, numpy scalars and other numbers.
        if value.hasattr("__index__")?
            && let Ok(i) = value.call_method0("__index__").and_then(|i| i.extract::<i64>()) {
                return Ok(RuleValue::Int(i));
            }
        if value.hasattr("__float__")? {
            return Ok(RuleValue::Number(finite(value.call_method0("__float__")?.extract::<f64>()?)?));
        }
        Err(PyTypeError::new_err(format!(
            "unsupported rule value type {}: use a number, str, bool, date/time or Formula",
            value.get_type().name()?
        )))
    }

    /// The equivalent `rust_xlsxwriter` conditional format value.
    pub(crate) fn to_cf_value(&self) -> RValue {
        match self {
            RuleValue::Int(v) => match i32::try_from(*v) {
                Ok(i) => i.into(),
                Err(_) => (*v as f64).into(),
            },
            RuleValue::Number(v) => (*v).into(),
            RuleValue::Bool(v) => (*v).into(),
            RuleValue::Text(v) => v.as_str().into(),
            RuleValue::Formula { text, .. } => RFormula::new(text).into(),
            RuleValue::DateTime(v) => (*v).into(),
            RuleValue::Date(v) => (*v).into(),
            RuleValue::Time(v) => (*v).into(),
        }
    }

    pub(crate) fn as_f64(&self) -> Option<f64> {
        match self {
            RuleValue::Int(v) => Some(*v as f64),
            RuleValue::Number(v) => Some(*v),
            _ => None,
        }
    }

    pub(crate) fn is_formula(&self) -> bool {
        matches!(self, RuleValue::Formula { .. })
    }

    /// A short description for error messages.
    pub(crate) fn kind(&self) -> &'static str {
        match self {
            RuleValue::Int(_) => "int",
            RuleValue::Number(_) => "float",
            RuleValue::Bool(_) => "bool",
            RuleValue::Text(_) => "str",
            RuleValue::Formula { .. } => "formula",
            RuleValue::DateTime(_) => "datetime",
            RuleValue::Date(_) => "date",
            RuleValue::Time(_) => "time",
        }
    }
}

fn finite(value: f64) -> PyResult<f64> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(param_err(format!("rule value {value} is not a finite number")))
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for RuleValue {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        RuleValue::from_py(&obj.to_owned())
    }
}

impl<'py> IntoPyObject<'py> for RuleValue {
    type Target = PyAny;
    type Output = Bound<'py, PyAny>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        match self {
            RuleValue::Int(v) => v.into_bound_py_any(py),
            RuleValue::Number(v) => v.into_bound_py_any(py),
            RuleValue::Bool(v) => v.into_bound_py_any(py),
            RuleValue::Text(v) => v.into_bound_py_any(py),
            RuleValue::Formula { text, object: true } => {
                Formula { formula: text, result: None }.into_bound_py_any(py)
            }
            RuleValue::Formula { text, object: false } => text.into_bound_py_any(py),
            RuleValue::DateTime(v) => v.into_bound_py_any(py),
            RuleValue::Date(v) => v.into_bound_py_any(py),
            RuleValue::Time(v) => v.into_bound_py_any(py),
        }
    }
}

/// A formula given as a `str` or a `Formula` object (read back as `str`).
#[derive(Clone, Debug, PartialEq)]
pub struct FormulaText(pub String);

impl<'a, 'py> FromPyObject<'a, 'py> for FormulaText {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(s) = obj.cast::<PyString>() {
            return Ok(FormulaText(s.to_str()?.to_string()));
        }
        if let Ok(f) = obj.cast::<Formula>() {
            return Ok(FormulaText(f.get().formula.clone()));
        }
        Err(PyTypeError::new_err(format!(
            "expected a formula string or Formula, not {}",
            obj.get_type().name()?
        )))
    }
}

impl<'py> IntoPyObject<'py> for FormulaText {
    type Target = PyString;
    type Output = Bound<'py, PyString>;
    type Error = std::convert::Infallible;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(PyString::new(py, &self.0))
    }
}

/// Comparison operators shared by cell and data validation rules.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Criteria {
    EqualTo,
    NotEqualTo,
    GreaterThan,
    GreaterThanOrEqualTo,
    LessThan,
    LessThanOrEqualTo,
    Between,
    NotBetween,
}

impl Criteria {
    pub(crate) fn name(self) -> &'static str {
        match self {
            Criteria::EqualTo => "equal_to",
            Criteria::NotEqualTo => "not_equal_to",
            Criteria::GreaterThan => "greater_than",
            Criteria::GreaterThanOrEqualTo => "greater_than_or_equal_to",
            Criteria::LessThan => "less_than",
            Criteria::LessThanOrEqualTo => "less_than_or_equal_to",
            Criteria::Between => "between",
            Criteria::NotBetween => "not_between",
        }
    }
}

/// A comparison with one value (or two for between/not between).
#[derive(Clone, Debug, PartialEq)]
pub struct ComparisonRule {
    pub(crate) criteria: Criteria,
    pub(crate) values: Vec<RuleValue>,
}

impl ComparisonRule {
    pub(crate) fn first(&self) -> PyResult<&RuleValue> {
        self.values.first().ok_or_else(|| param_err("the rule has no value"))
    }

    pub(crate) fn second(&self) -> PyResult<&RuleValue> {
        self.values.get(1).ok_or_else(|| param_err("the rule needs a second value"))
    }

    pub(crate) fn repr(&self, class: &str, py: Python<'_>) -> PyResult<String> {
        let mut parts = Vec::with_capacity(self.values.len());
        for value in &self.values {
            parts.push(value.clone().into_bound_py_any(py)?.repr()?.to_string());
        }
        Ok(format!("{class}.{}({})", self.criteria.name(), parts.join(", ")))
    }
}

/// Build a rust_xlsxwriter comparison rule enum (`ConditionalFormatCellRule`
/// or `DataValidationRule`) from a [`ComparisonRule`], converting each value
/// with `$conv` (a `Fn(&RuleValue) -> PyResult<T>`).
macro_rules! comparison_to_rust {
    ($enum:ident, $rule:expr, $conv:expr) => {{
        use $crate::writer::conditional_format::Criteria;
        let rule = $rule;
        let conv = $conv;
        let first = conv(rule.first()?)?;
        match rule.criteria {
            Criteria::EqualTo => $enum::EqualTo(first),
            Criteria::NotEqualTo => $enum::NotEqualTo(first),
            Criteria::GreaterThan => $enum::GreaterThan(first),
            Criteria::GreaterThanOrEqualTo => $enum::GreaterThanOrEqualTo(first),
            Criteria::LessThan => $enum::LessThan(first),
            Criteria::LessThanOrEqualTo => $enum::LessThanOrEqualTo(first),
            Criteria::Between => $enum::Between(first, conv(rule.second()?)?),
            Criteria::NotBetween => $enum::NotBetween(first, conv(rule.second()?)?),
        }
    }};
}
pub(crate) use comparison_to_rust;

/// Define a frozen Python rule class with `equal_to(v)`, ..., `between(a, b)`
/// static constructors, wrapping a [`ComparisonRule`].
macro_rules! comparison_rule_class {
    ($(#[$meta:meta])* $name:ident) => {
        $(#[$meta])*
        #[pyo3::pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
        #[derive(Clone, Debug, PartialEq)]
        pub struct $name {
            pub(crate) rule: $crate::writer::conditional_format::ComparisonRule,
        }

        impl $name {
            fn make(
                criteria: $crate::writer::conditional_format::Criteria,
                values: Vec<$crate::writer::conditional_format::RuleValue>,
            ) -> Self {
                $name { rule: $crate::writer::conditional_format::ComparisonRule { criteria, values } }
            }
        }

        #[pyo3::pymethods]
        impl $name {
            /// Values equal to `value`.
            #[staticmethod]
            fn equal_to(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make($crate::writer::conditional_format::Criteria::EqualTo, vec![value])
            }

            /// Values not equal to `value`.
            #[staticmethod]
            fn not_equal_to(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make($crate::writer::conditional_format::Criteria::NotEqualTo, vec![value])
            }

            /// Values greater than `value`.
            #[staticmethod]
            fn greater_than(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make($crate::writer::conditional_format::Criteria::GreaterThan, vec![value])
            }

            /// Values greater than or equal to `value`.
            #[staticmethod]
            fn greater_than_or_equal_to(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make(
                    $crate::writer::conditional_format::Criteria::GreaterThanOrEqualTo,
                    vec![value],
                )
            }

            /// Values less than `value`.
            #[staticmethod]
            fn less_than(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make($crate::writer::conditional_format::Criteria::LessThan, vec![value])
            }

            /// Values less than or equal to `value`.
            #[staticmethod]
            fn less_than_or_equal_to(value: $crate::writer::conditional_format::RuleValue) -> Self {
                Self::make(
                    $crate::writer::conditional_format::Criteria::LessThanOrEqualTo,
                    vec![value],
                )
            }

            /// Values between `minimum` and `maximum` (inclusive).
            #[staticmethod]
            fn between(
                minimum: $crate::writer::conditional_format::RuleValue,
                maximum: $crate::writer::conditional_format::RuleValue,
            ) -> Self {
                Self::make($crate::writer::conditional_format::Criteria::Between, vec![minimum, maximum])
            }

            /// Values outside `minimum` .. `maximum`.
            #[staticmethod]
            fn not_between(
                minimum: $crate::writer::conditional_format::RuleValue,
                maximum: $crate::writer::conditional_format::RuleValue,
            ) -> Self {
                Self::make(
                    $crate::writer::conditional_format::Criteria::NotBetween,
                    vec![minimum, maximum],
                )
            }

            /// The comparison: "equal_to", "between", ... (the constructor name).
            #[getter]
            fn criteria(&self) -> &'static str {
                self.rule.criteria.name()
            }

            /// The rule value(s): one value, or (minimum, maximum) for between rules.
            #[getter]
            fn values<'py>(&self, py: pyo3::Python<'py>) -> pyo3::PyResult<pyo3::Bound<'py, pyo3::types::PyTuple>> {
                pyo3::types::PyTuple::new(py, self.rule.values.clone())
            }

            fn __eq__(&self, other: &pyo3::Bound<'_, pyo3::PyAny>) -> bool {
                match other.cast::<$name>() {
                    Ok(other) => self.rule == other.get().rule,
                    Err(_) => false,
                }
            }

            fn __repr__(&self, py: pyo3::Python<'_>) -> pyo3::PyResult<String> {
                self.rule.repr(stringify!($name), py)
            }
        }
    };
}
pub(crate) use comparison_rule_class;

/// Common methods for spec-first classes: `set(**kwargs)`, `copy()`,
/// `__copy__`, `__deepcopy__` and `__eq__` (comparing specs).
macro_rules! spec_methods {
    ($($ty:ident),* $(,)?) => { $(
        #[pyo3::pymethods]
        impl $ty {
            /// Set several properties at once and return self (for chaining).
            #[pyo3(signature = (**kwargs))]
            fn set<'py>(
                mut slf: pyo3::PyRefMut<'py, Self>,
                kwargs: Option<&pyo3::Bound<'py, pyo3::types::PyDict>>,
            ) -> pyo3::PyResult<pyo3::PyRefMut<'py, Self>> {
                slf.apply_kwargs(kwargs)?;
                Ok(slf)
            }

            /// A copy of this object.
            fn copy(&self) -> Self {
                self.clone()
            }

            fn __copy__(&self) -> Self {
                self.clone()
            }

            fn __deepcopy__(&self, _memo: &pyo3::Bound<'_, pyo3::PyAny>) -> Self {
                self.clone()
            }

            fn __eq__(&self, other: &pyo3::Bound<'_, pyo3::PyAny>) -> bool {
                match other.cast::<$ty>() {
                    Ok(other) => match other.try_borrow() {
                        Ok(other) => self.spec == other.spec,
                        // `other` is `self`, mutably borrowed elsewhere.
                        Err(_) => false,
                    },
                    Err(_) => false,
                }
            }
        }
    )* };
}
pub(crate) use spec_methods;

/// Store a `Format` property as a `FormatSpec` (getter returns a copy).
macro_rules! format_accessors {
    ($ty:ident: $($field:ident),*) => {
        paste::paste! {
            impl $ty {
                $(
                    fn [<get_ $field>](&self) -> Option<Format> {
                        self.spec.$field.clone().map(Format::from_spec)
                    }

                    fn [<set_ $field>](&mut self, value: Option<Format>) -> PyResult<()> {
                        self.spec.$field = value.map(|f| f.spec);
                        Ok(())
                    }
                )*
            }
        }
    };
}
pub(crate) use format_accessors;

// ===========================================================================
// Enums
// ===========================================================================

py_enum!(
    /// Criteria for `ConditionalFormatAverage`.
    ConditionalFormatAverageRule => RAverageRule {
        AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage,
        OneStandardDeviationAbove, OneStandardDeviationBelow,
        TwoStandardDeviationsAbove, TwoStandardDeviationsBelow,
        ThreeStandardDeviationsAbove, ThreeStandardDeviationsBelow,
    }
);

py_enum!(
    /// Criteria for `ConditionalFormatDate` (relative to when the file is opened).
    ConditionalFormatDateRule => RDateRule {
        Yesterday, Today, Tomorrow, Last7Days, LastWeek, ThisWeek, NextWeek,
        LastMonth, ThisMonth, NextMonth,
    }
);

py_enum!(
    /// Threshold types for color scales, data bars and custom icons.
    ConditionalFormatType => RType {
        Automatic, Lowest, Number, Percent, Formula, Percentile, Highest,
    }
);

py_enum!(
    /// Data bar direction.
    ConditionalFormatDataBarDirection => RDataBarDirection { Context, LeftToRight, RightToLeft }
);

py_enum!(
    /// Data bar axis position.
    ConditionalFormatDataBarAxisPosition => RDataBarAxisPosition {
        Automatic, Midpoint, #[pyo3(name = "None_")] None,
    }
);

py_enum!(
    /// Icon set styles for `ConditionalFormatIconSet`.
    ConditionalFormatIconType => RIconType {
        ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights, ThreeTrafficLightsWithRim,
        ThreeSigns, ThreeSymbolsCircled, ThreeSymbols, ThreeStars, ThreeTriangles,
        FourArrows, FourArrowsGray, FourRedToBlack, FourHistograms, FourTrafficLights,
        FiveArrows, FiveArrowsGray, FiveHistograms, FiveQuadrants, FiveBoxes,
    }
);

impl ConditionalFormatIconType {
    /// Number of icons in the set (3, 4 or 5).
    fn icon_count(self) -> usize {
        use ConditionalFormatIconType::*;
        match self {
            ThreeArrows | ThreeArrowsGray | ThreeFlags | ThreeTrafficLights
            | ThreeTrafficLightsWithRim | ThreeSigns | ThreeSymbolsCircled | ThreeSymbols
            | ThreeStars | ThreeTriangles => 3,
            FourArrows | FourArrowsGray | FourRedToBlack | FourHistograms | FourTrafficLights => 4,
            FiveArrows | FiveArrowsGray | FiveHistograms | FiveQuadrants | FiveBoxes => 5,
        }
    }
}

// ===========================================================================
// Rule classes
// ===========================================================================

comparison_rule_class!(
    /// The rule of a `ConditionalFormatCell`: `ConditionalFormatCellRule.greater_than(5)`,
    /// `.between(1, 10)`, ... Values may be numbers, strings (compared as text),
    /// bools, dates/times or formulas (`"=$B$1"` or `Formula(...)`).
    ConditionalFormatCellRule
);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TextCriteria {
    Contains,
    DoesNotContain,
    BeginsWith,
    EndsWith,
}

/// The rule of a `ConditionalFormatText`: `ConditionalFormatTextRule.contains("x")`,
/// `.does_not_contain(...)`, `.begins_with(...)` or `.ends_with(...)`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Debug, PartialEq)]
pub struct ConditionalFormatTextRule {
    criteria: TextCriteria,
    text: String,
}

impl ConditionalFormatTextRule {
    fn to_rust(&self) -> RTextRule {
        let text = self.text.clone();
        match self.criteria {
            TextCriteria::Contains => RTextRule::Contains(text),
            TextCriteria::DoesNotContain => RTextRule::DoesNotContain(text),
            TextCriteria::BeginsWith => RTextRule::BeginsWith(text),
            TextCriteria::EndsWith => RTextRule::EndsWith(text),
        }
    }

    fn criteria_name(&self) -> &'static str {
        match self.criteria {
            TextCriteria::Contains => "contains",
            TextCriteria::DoesNotContain => "does_not_contain",
            TextCriteria::BeginsWith => "begins_with",
            TextCriteria::EndsWith => "ends_with",
        }
    }
}

#[pymethods]
impl ConditionalFormatTextRule {
    /// Text containing `text`.
    #[staticmethod]
    fn contains(text: String) -> Self {
        ConditionalFormatTextRule { criteria: TextCriteria::Contains, text }
    }

    /// Text not containing `text`.
    #[staticmethod]
    fn does_not_contain(text: String) -> Self {
        ConditionalFormatTextRule { criteria: TextCriteria::DoesNotContain, text }
    }

    /// Text beginning with `text`.
    #[staticmethod]
    fn begins_with(text: String) -> Self {
        ConditionalFormatTextRule { criteria: TextCriteria::BeginsWith, text }
    }

    /// Text ending with `text`.
    #[staticmethod]
    fn ends_with(text: String) -> Self {
        ConditionalFormatTextRule { criteria: TextCriteria::EndsWith, text }
    }

    /// "contains", "does_not_contain", "begins_with" or "ends_with".
    #[getter]
    fn criteria(&self) -> &'static str {
        self.criteria_name()
    }

    /// The text to match.
    #[getter]
    fn value(&self) -> String {
        self.text.clone()
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<ConditionalFormatTextRule>() {
            Ok(other) => self == other.get(),
            Err(_) => false,
        }
    }

    fn __repr__(&self) -> String {
        format!("ConditionalFormatTextRule.{}({:?})", self.criteria_name(), self.text)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TopCriteria {
    Top,
    Bottom,
    TopPercent,
    BottomPercent,
}

/// The rule of a `ConditionalFormatTop`: `ConditionalFormatTopRule.top(10)`,
/// `.bottom(n)` (1..=1000) or `.top_percent(p)`, `.bottom_percent(p)` (1..=100).
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Debug, PartialEq)]
pub struct ConditionalFormatTopRule {
    criteria: TopCriteria,
    value: u16,
}

impl ConditionalFormatTopRule {
    fn new(criteria: TopCriteria, value: i64) -> PyResult<Self> {
        let (limit, what) = match criteria {
            TopCriteria::Top | TopCriteria::Bottom => (1000, "rank"),
            TopCriteria::TopPercent | TopCriteria::BottomPercent => (100, "percent"),
        };
        if !(1..=limit).contains(&value) {
            return Err(param_err(format!(
                "top/bottom {what} {value} is outside Excel's range 1..={limit}"
            )));
        }
        Ok(ConditionalFormatTopRule { criteria, value: value as u16 })
    }

    fn to_rust(&self) -> RTopRule {
        match self.criteria {
            TopCriteria::Top => RTopRule::Top(self.value),
            TopCriteria::Bottom => RTopRule::Bottom(self.value),
            TopCriteria::TopPercent => RTopRule::TopPercent(self.value),
            TopCriteria::BottomPercent => RTopRule::BottomPercent(self.value),
        }
    }

    fn criteria_name(&self) -> &'static str {
        match self.criteria {
            TopCriteria::Top => "top",
            TopCriteria::Bottom => "bottom",
            TopCriteria::TopPercent => "top_percent",
            TopCriteria::BottomPercent => "bottom_percent",
        }
    }
}

impl Default for ConditionalFormatTopRule {
    fn default() -> Self {
        ConditionalFormatTopRule { criteria: TopCriteria::Top, value: 10 }
    }
}

#[pymethods]
impl ConditionalFormatTopRule {
    /// The top `rank` values (1..=1000).
    #[staticmethod]
    fn top(rank: i64) -> PyResult<Self> {
        Self::new(TopCriteria::Top, rank)
    }

    /// The bottom `rank` values (1..=1000).
    #[staticmethod]
    fn bottom(rank: i64) -> PyResult<Self> {
        Self::new(TopCriteria::Bottom, rank)
    }

    /// The top `percent`% of values (1..=100).
    #[staticmethod]
    fn top_percent(percent: i64) -> PyResult<Self> {
        Self::new(TopCriteria::TopPercent, percent)
    }

    /// The bottom `percent`% of values (1..=100).
    #[staticmethod]
    fn bottom_percent(percent: i64) -> PyResult<Self> {
        Self::new(TopCriteria::BottomPercent, percent)
    }

    /// "top", "bottom", "top_percent" or "bottom_percent".
    #[getter]
    fn criteria(&self) -> &'static str {
        self.criteria_name()
    }

    /// The rank or percentage.
    #[getter]
    fn value(&self) -> u16 {
        self.value
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<ConditionalFormatTopRule>() {
            Ok(other) => self == other.get(),
            Err(_) => false,
        }
    }

    fn __repr__(&self) -> String {
        format!("ConditionalFormatTopRule.{}({})", self.criteria_name(), self.value)
    }
}

// ===========================================================================
// Rule-based conditional formats (format, multi_range, stop_if_true)
// ===========================================================================

/// Apply `format`, `multi_range` and `stop_if_true` from a spec.
macro_rules! apply_common {
    ($cf:expr, $spec:expr, with_format) => {{
        let mut cf = $cf;
        if let Some(format) = &$spec.format {
            cf = cf.set_format(format.to_rust());
        }
        apply_common!(cf, $spec)
    }};
    ($cf:expr, $spec:expr) => {{
        let mut cf = $cf;
        if let Some(range) = &$spec.multi_range {
            cf = cf.set_multi_range(range.as_str());
        }
        cf.set_stop_if_true($spec.stop_if_true)
    }};
}

fn check_multi_range(range: &Option<String>) -> PyResult<()> {
    if let Some(range) = range
        && range.trim().is_empty() {
            return Err(param_err("multi_range must not be empty (use None to unset it)"));
        }
    Ok(())
}

/// Define a rule-based conditional format class with the common
/// `format`/`multi_range`/`stop_if_true` properties plus its own fields.
macro_rules! rule_cf_class {
    (
        $(#[$meta:meta])*
        $name:ident, $spec:ident { $( $field:ident : $t:ty = $default:expr ),* $(,)? }
    ) => {
        #[derive(Clone, Debug, PartialEq)]
        pub struct $spec {
            $( pub $field: $t, )*
            pub format: Option<FormatSpec>,
            pub multi_range: Option<String>,
            pub stop_if_true: bool,
        }

        impl Default for $spec {
            fn default() -> Self {
                $spec {
                    $( $field: $default, )*
                    format: None,
                    multi_range: None,
                    stop_if_true: false,
                }
            }
        }

        $(#[$meta])*
        #[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
        #[derive(Clone)]
        pub struct $name {
            pub spec: $spec,
        }

        impl $name {
            fn on_change(&mut self) -> PyResult<()> {
                check_multi_range(&self.spec.multi_range)
            }
        }

        format_accessors!($name: format);

        py_props!($name
            fields {
                $( $field: $t, )*
                multi_range: Option<String>,
                stop_if_true: bool,
            }
            custom {
                format: Option<Format>,
            }
        );

        #[pymethods]
        impl $name {
            #[new]
            #[pyo3(signature = (**kwargs))]
            fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
                let mut cf = $name { spec: $spec::default() };
                cf.apply_kwargs(kwargs)?;
                Ok(cf)
            }
        }

        spec_methods!($name);
    };
}

rule_cf_class!(
    /// Highlight cells whose value matches a comparison:
    /// `ConditionalFormatCell(rule=ConditionalFormatCellRule.greater_than(5), format=fmt)`.
    ConditionalFormatCell, CellSpec { rule: Option<ConditionalFormatCellRule> = None }
);

impl ConditionalFormatCell {
    fn to_rust(&self) -> PyResult<RCell> {
        let rule = self.spec.rule.as_ref().ok_or_else(|| {
            param_err("ConditionalFormatCell needs a rule, e.g. rule=ConditionalFormatCellRule.greater_than(5)")
        })?;
        let rule: RCellRule<RValue> =
            comparison_to_rust!(RCellRule, &rule.rule, |v: &RuleValue| -> PyResult<RValue> {
                Ok(v.to_cf_value())
            });
        Ok(apply_common!(RCell::new().set_rule(rule), self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight text cells: `ConditionalFormatText(rule=ConditionalFormatTextRule.contains("x"), format=fmt)`.
    ConditionalFormatText, TextSpec { rule: Option<ConditionalFormatTextRule> = None }
);

impl ConditionalFormatText {
    fn to_rust(&self) -> PyResult<RText> {
        let rule = self.spec.rule.as_ref().ok_or_else(|| {
            param_err("ConditionalFormatText needs a rule, e.g. rule=ConditionalFormatTextRule.contains(\"x\")")
        })?;
        Ok(apply_common!(RText::new().set_rule(rule.to_rust()), self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight dates relative to today: `ConditionalFormatDate(rule=ConditionalFormatDateRule.Today)`.
    ConditionalFormatDate, DateSpec { rule: ConditionalFormatDateRule = ConditionalFormatDateRule::Yesterday }
);

impl ConditionalFormatDate {
    fn to_rust(&self) -> PyResult<RDate> {
        Ok(apply_common!(RDate::new().set_rule(self.spec.rule.into()), self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight values above/below the average (or N standard deviations).
    ConditionalFormatAverage, AverageSpec { rule: ConditionalFormatAverageRule = ConditionalFormatAverageRule::AboveAverage }
);

impl ConditionalFormatAverage {
    fn to_rust(&self) -> PyResult<RAverage> {
        Ok(apply_common!(RAverage::new().set_rule(self.spec.rule.into()), self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight the top/bottom N or N% values (default: top 10).
    ConditionalFormatTop, TopSpec { rule: ConditionalFormatTopRule = ConditionalFormatTopRule::default() }
);

impl ConditionalFormatTop {
    fn to_rust(&self) -> PyResult<RTop> {
        Ok(apply_common!(RTop::new().set_rule(self.spec.rule.to_rust()), self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight duplicate values (`invert=True` highlights unique values).
    ConditionalFormatDuplicate, DuplicateSpec { invert: bool = false }
);

impl ConditionalFormatDuplicate {
    fn to_rust(&self) -> PyResult<RDuplicate> {
        let mut cf = RDuplicate::new();
        if self.spec.invert {
            cf = cf.invert();
        }
        Ok(apply_common!(cf, self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight blank cells (`invert=True` highlights non-blank cells).
    ConditionalFormatBlank, BlankSpec { invert: bool = false }
);

impl ConditionalFormatBlank {
    fn to_rust(&self) -> PyResult<RBlank> {
        let mut cf = RBlank::new();
        if self.spec.invert {
            cf = cf.invert();
        }
        Ok(apply_common!(cf, self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight cells with errors (`invert=True` highlights cells without errors).
    ConditionalFormatError, ErrorSpec { invert: bool = false }
);

impl ConditionalFormatError {
    fn to_rust(&self) -> PyResult<RError> {
        let mut cf = RError::new();
        if self.spec.invert {
            cf = cf.invert();
        }
        Ok(apply_common!(cf, self.spec, with_format))
    }
}

rule_cf_class!(
    /// Highlight cells where a formula is true, relative to the top-left cell:
    /// `ConditionalFormatFormula(rule="=ISODD(A1)", format=fmt)`.
    ConditionalFormatFormula, FormulaSpec { rule: Option<FormulaText> = None }
);

impl ConditionalFormatFormula {
    fn to_rust(&self) -> PyResult<RFormulaCf> {
        let rule = match &self.spec.rule {
            Some(rule) if !rule.0.trim().trim_start_matches('=').is_empty() => rule,
            _ => {
                return Err(param_err(
                    "ConditionalFormatFormula needs a rule formula, e.g. rule=\"=ISODD(A1)\"",
                ))
            }
        };
        Ok(apply_common!(RFormulaCf::new().set_rule(rule.0.as_str()), self.spec, with_format))
    }
}

// ===========================================================================
// Thresholds for color scales and data bars
// ===========================================================================

use ConditionalFormatType as CfType;

const SCALE_MIN: &[CfType] = &[CfType::Lowest, CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile];
const SCALE_MID: &[CfType] = &[CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile];
const SCALE_MAX: &[CfType] = &[CfType::Highest, CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile];
const BAR_MIN: &[CfType] = &[
    CfType::Automatic, CfType::Lowest, CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile,
];
const BAR_MAX: &[CfType] = &[
    CfType::Automatic, CfType::Highest, CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile,
];
const ICON_TYPES: &[CfType] = &[CfType::Number, CfType::Percent, CfType::Formula, CfType::Percentile];

/// Whether a threshold type uses a value.
fn needs_value(kind: CfType) -> bool {
    matches!(kind, CfType::Number | CfType::Percent | CfType::Formula | CfType::Percentile)
}

/// Validate a threshold type/value pair (completeness is checked when built).
fn check_threshold(which: &str, kind: CfType, value: &Option<RuleValue>, allowed: &[CfType]) -> PyResult<()> {
    if !allowed.contains(&kind) {
        let names: Vec<String> = allowed.iter().map(|k| format!("{k:?}")).collect();
        return Err(param_err(format!(
            "{which}_type cannot be ConditionalFormatType.{kind:?} here; use one of {}",
            names.join(", ")
        )));
    }
    if let Some(value) = value {
        if !matches!(
            value,
            RuleValue::Int(_) | RuleValue::Number(_) | RuleValue::Formula { .. }
                | RuleValue::Date(_) | RuleValue::DateTime(_) | RuleValue::Time(_)
        ) {
            return Err(param_err(format!(
                "{which}_value must be a number, date/time or formula, not {}",
                value.kind()
            )));
        }
        if matches!(kind, CfType::Percent | CfType::Percentile)
            && let Some(number) = value.as_f64()
                && !(0.0..=100.0).contains(&number) {
                    return Err(param_err(format!(
                        "{which}_value {number} is outside 0..=100 for {kind:?}"
                    )));
                }
    }
    Ok(())
}

/// The value to pass for a threshold, or None to keep the Rust default.
fn threshold_value(which: &str, kind: CfType, value: &Option<RuleValue>) -> PyResult<Option<RValue>> {
    if !needs_value(kind) {
        return Ok(None);
    }
    match value {
        Some(value) => Ok(Some(value.to_cf_value())),
        None => Err(param_err(format!(
            "{which}_value must be set when {which}_type is ConditionalFormatType.{kind:?}"
        ))),
    }
}

/// Define a colour-scale/data-bar/icon-set class (no `format`).
macro_rules! scale_cf_class {
    (
        $(#[$meta:meta])*
        $name:ident, $spec:ident { $( $field:ident : $t:ty = $default:expr ),* $(,)? }
        custom { $( $cfield:ident : $ct:ty ),* $(,)? }
    ) => {
        #[derive(Clone, Debug, PartialEq)]
        pub struct $spec {
            $( pub $field: $t, )*
            pub multi_range: Option<String>,
            pub stop_if_true: bool,
        }

        impl Default for $spec {
            fn default() -> Self {
                $spec { $( $field: $default, )* multi_range: None, stop_if_true: false }
            }
        }

        $(#[$meta])*
        #[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
        #[derive(Clone)]
        pub struct $name {
            pub spec: $spec,
        }

        py_props!($name
            fields {
                $( $field: $t, )*
                multi_range: Option<String>,
                stop_if_true: bool,
            }
            custom { $( $cfield: $ct, )* }
        );

        #[pymethods]
        impl $name {
            #[new]
            #[pyo3(signature = (**kwargs))]
            fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
                let mut cf = $name { spec: $spec::default() };
                cf.apply_kwargs(kwargs)?;
                Ok(cf)
            }
        }

        spec_methods!($name);
    };
}

// ---------------------------------------------------------------------------
// 2 color scale
// ---------------------------------------------------------------------------

scale_cf_class!(
    /// A 2 color scale from `min_color` (lowest) to `max_color` (highest).
    ///
    /// Thresholds default to the lowest/highest value; change them with
    /// `min_type`/`min_value` and `max_type`/`max_value` (or `set_minimum()`).
    ConditionalFormat2ColorScale, ColorScale2Spec {
        min_type: ConditionalFormatType = CfType::Lowest,
        min_value: Option<RuleValue> = None,
        max_type: ConditionalFormatType = CfType::Highest,
        max_value: Option<RuleValue> = None,
        min_color: Option<Color> = None,
        max_color: Option<Color> = None,
    }
    custom {}
);

impl ConditionalFormat2ColorScale {
    fn on_change(&mut self) -> PyResult<()> {
        let s = &self.spec;
        check_threshold("min", s.min_type, &s.min_value, SCALE_MIN)?;
        check_threshold("max", s.max_type, &s.max_value, SCALE_MAX)?;
        check_multi_range(&s.multi_range)
    }

    fn to_rust(&self) -> PyResult<R2ColorScale> {
        let s = &self.spec;
        let mut cf = R2ColorScale::new();
        if let Some(v) = threshold_value("min", s.min_type, &s.min_value)? {
            cf = cf.set_minimum(s.min_type.into(), v);
        }
        if let Some(v) = threshold_value("max", s.max_type, &s.max_value)? {
            cf = cf.set_maximum(s.max_type.into(), v);
        }
        if let Some(c) = s.min_color {
            cf = cf.set_minimum_color(c);
        }
        if let Some(c) = s.max_color {
            cf = cf.set_maximum_color(c);
        }
        Ok(apply_common!(cf, s))
    }
}

#[pymethods]
impl ConditionalFormat2ColorScale {
    /// Set `min_type` and `min_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_minimum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.min_type = rule_type;
        slf.spec.min_value = value;
        slf.on_change()?;
        Ok(slf)
    }

    /// Set `max_type` and `max_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_maximum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.max_type = rule_type;
        slf.spec.max_value = value;
        slf.on_change()?;
        Ok(slf)
    }
}

// ---------------------------------------------------------------------------
// 3 color scale
// ---------------------------------------------------------------------------

scale_cf_class!(
    /// A 3 color scale: `min_color` (lowest), `mid_color` (50th percentile)
    /// and `max_color` (highest). Thresholds are set with `*_type`/`*_value`.
    ConditionalFormat3ColorScale, ColorScale3Spec {
        min_type: ConditionalFormatType = CfType::Lowest,
        min_value: Option<RuleValue> = None,
        mid_type: ConditionalFormatType = CfType::Percentile,
        mid_value: Option<RuleValue> = Some(RuleValue::Int(50)),
        max_type: ConditionalFormatType = CfType::Highest,
        max_value: Option<RuleValue> = None,
        min_color: Option<Color> = None,
        mid_color: Option<Color> = None,
        max_color: Option<Color> = None,
    }
    custom {}
);

impl ConditionalFormat3ColorScale {
    fn on_change(&mut self) -> PyResult<()> {
        let s = &self.spec;
        check_threshold("min", s.min_type, &s.min_value, SCALE_MIN)?;
        check_threshold("mid", s.mid_type, &s.mid_value, SCALE_MID)?;
        check_threshold("max", s.max_type, &s.max_value, SCALE_MAX)?;
        check_multi_range(&s.multi_range)
    }

    fn to_rust(&self) -> PyResult<R3ColorScale> {
        let s = &self.spec;
        let mut cf = R3ColorScale::new();
        if let Some(v) = threshold_value("min", s.min_type, &s.min_value)? {
            cf = cf.set_minimum(s.min_type.into(), v);
        }
        if let Some(v) = threshold_value("mid", s.mid_type, &s.mid_value)? {
            cf = cf.set_midpoint(s.mid_type.into(), v);
        }
        if let Some(v) = threshold_value("max", s.max_type, &s.max_value)? {
            cf = cf.set_maximum(s.max_type.into(), v);
        }
        if let Some(c) = s.min_color {
            cf = cf.set_minimum_color(c);
        }
        if let Some(c) = s.mid_color {
            cf = cf.set_midpoint_color(c);
        }
        if let Some(c) = s.max_color {
            cf = cf.set_maximum_color(c);
        }
        Ok(apply_common!(cf, s))
    }
}

#[pymethods]
impl ConditionalFormat3ColorScale {
    /// Set `min_type` and `min_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_minimum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.min_type = rule_type;
        slf.spec.min_value = value;
        slf.on_change()?;
        Ok(slf)
    }

    /// Set `mid_type` and `mid_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_midpoint(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.mid_type = rule_type;
        slf.spec.mid_value = value;
        slf.on_change()?;
        Ok(slf)
    }

    /// Set `max_type` and `max_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_maximum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.max_type = rule_type;
        slf.spec.max_value = value;
        slf.on_change()?;
        Ok(slf)
    }
}

// ---------------------------------------------------------------------------
// Data bar
// ---------------------------------------------------------------------------

scale_cf_class!(
    /// An Excel 2010+ style data bar.
    ///
    /// Thresholds default to `ConditionalFormatType.Automatic`. Colours,
    /// `solid_fill`, `border_off`, `direction`, `bar_only`, `axis_position`
    /// and `axis_color` mirror Excel's "Edit Formatting Rule" dialog;
    /// `classic_style=True` writes an Excel 2007 style bar.
    ConditionalFormatDataBar, DataBarSpec {
        min_type: ConditionalFormatType = CfType::Automatic,
        min_value: Option<RuleValue> = None,
        max_type: ConditionalFormatType = CfType::Automatic,
        max_value: Option<RuleValue> = None,
        fill_color: Option<Color> = None,
        border_color: Option<Color> = None,
        negative_fill_color: Option<Color> = None,
        negative_border_color: Option<Color> = None,
        solid_fill: bool = false,
        border_off: bool = false,
        direction: ConditionalFormatDataBarDirection = ConditionalFormatDataBarDirection::Context,
        bar_only: bool = false,
        axis_position: ConditionalFormatDataBarAxisPosition = ConditionalFormatDataBarAxisPosition::Automatic,
        axis_color: Option<Color> = None,
        classic_style: bool = false,
    }
    custom {}
);

impl ConditionalFormatDataBar {
    fn on_change(&mut self) -> PyResult<()> {
        let s = &self.spec;
        check_threshold("min", s.min_type, &s.min_value, BAR_MIN)?;
        check_threshold("max", s.max_type, &s.max_value, BAR_MAX)?;
        check_multi_range(&s.multi_range)
    }

    fn to_rust(&self) -> PyResult<RDataBar> {
        let s = &self.spec;
        let mut cf = RDataBar::new();
        if s.min_type != CfType::Automatic {
            let value = threshold_value("min", s.min_type, &s.min_value)?.unwrap_or_else(|| 0.into());
            cf = cf.set_minimum(s.min_type.into(), value);
        }
        if s.max_type != CfType::Automatic {
            let value = threshold_value("max", s.max_type, &s.max_value)?.unwrap_or_else(|| 0.into());
            cf = cf.set_maximum(s.max_type.into(), value);
        }
        if let Some(c) = s.fill_color {
            cf = cf.set_fill_color(c);
        }
        if let Some(c) = s.border_color {
            cf = cf.set_border_color(c);
        }
        if let Some(c) = s.negative_fill_color {
            cf = cf.set_negative_fill_color(c);
        }
        if let Some(c) = s.negative_border_color {
            cf = cf.set_negative_border_color(c);
        }
        if let Some(c) = s.axis_color {
            cf = cf.set_axis_color(c);
        }
        cf = cf
            .set_solid_fill(s.solid_fill)
            .set_border_off(s.border_off)
            .set_direction(s.direction.into())
            .set_bar_only(s.bar_only)
            .set_axis_position(s.axis_position.into());
        if s.classic_style {
            cf = cf.use_classic_style();
        }
        Ok(apply_common!(cf, s))
    }
}

#[pymethods]
impl ConditionalFormatDataBar {
    /// Set `min_type` and `min_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_minimum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.min_type = rule_type;
        slf.spec.min_value = value;
        slf.on_change()?;
        Ok(slf)
    }

    /// Set `max_type` and `max_value` together; returns self.
    #[pyo3(signature = (rule_type, value=None))]
    fn set_maximum(
        mut slf: PyRefMut<'_, Self>,
        rule_type: ConditionalFormatType,
        value: Option<RuleValue>,
    ) -> PyResult<PyRefMut<'_, Self>> {
        slf.spec.max_type = rule_type;
        slf.spec.max_value = value;
        slf.on_change()?;
        Ok(slf)
    }
}

// ---------------------------------------------------------------------------
// Icon sets
// ---------------------------------------------------------------------------

/// Settings of one icon in a `ConditionalFormatIconSet`.
#[derive(Clone, Debug, PartialEq)]
pub struct CustomIconSpec {
    pub rule_type: ConditionalFormatType,
    pub value: RuleValue,
    pub greater_than: bool,
    pub icon_type: Option<ConditionalFormatIconType>,
    pub icon_index: u8,
    pub no_icon: bool,
}

impl Default for CustomIconSpec {
    fn default() -> Self {
        CustomIconSpec {
            rule_type: CfType::Percent,
            value: RuleValue::Int(0),
            greater_than: false,
            icon_type: None,
            icon_index: 0,
            no_icon: false,
        }
    }
}

impl CustomIconSpec {
    fn validate(&self) -> PyResult<()> {
        check_threshold("rule", self.rule_type, &Some(self.value.clone()), ICON_TYPES)?;
        if let Some(icon_type) = self.icon_type {
            let count = icon_type.icon_count();
            if usize::from(self.icon_index) >= count {
                return Err(param_err(format!(
                    "icon_index {} is outside 0..{count} for {icon_type:?}",
                    self.icon_index
                )));
            }
        }
        Ok(())
    }

    fn to_rust(&self) -> RCustomIcon {
        let mut icon = RCustomIcon::new()
            .set_rule(self.rule_type.into(), self.value.to_cf_value())
            .set_greater_than(self.greater_than);
        if let Some(icon_type) = self.icon_type {
            icon = icon.set_icon_type(icon_type.into(), self.icon_index);
        }
        if self.no_icon {
            icon = icon.set_no_icon(true);
        }
        icon
    }
}

/// The threshold (and optionally a different icon) for one icon of an icon set.
///
/// `ConditionalFormatCustomIcon(rule_type=ConditionalFormatType.Percentile, value=50)`.
/// The rule applies as ">= value" (or "> value" with `greater_than=True`).
/// `icon_type` + `icon_index` pick an icon from another set; `no_icon=True`
/// hides it. Excel always uses ">= 0%" for the first icon.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct ConditionalFormatCustomIcon {
    pub spec: CustomIconSpec,
}

impl ConditionalFormatCustomIcon {
    fn on_change(&mut self) -> PyResult<()> {
        self.spec.validate()
    }
}

py_props!(ConditionalFormatCustomIcon
    fields {
        rule_type: ConditionalFormatType,
        value: RuleValue,
        greater_than: bool,
        icon_type: Option<ConditionalFormatIconType>,
        icon_index: u8,
        no_icon: bool,
    }
);

#[pymethods]
impl ConditionalFormatCustomIcon {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut icon = ConditionalFormatCustomIcon { spec: CustomIconSpec::default() };
        icon.apply_kwargs(kwargs)?;
        Ok(icon)
    }
}

spec_methods!(ConditionalFormatCustomIcon);

impl<'a, 'py> FromPyObject<'a, 'py> for ConditionalFormatCustomIcon {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        Ok(obj.cast::<ConditionalFormatCustomIcon>()?.borrow().clone())
    }
}

/// Settings of a `ConditionalFormatIconSet`.
#[derive(Clone, Debug, PartialEq)]
pub struct IconSetSpec {
    pub icon_type: ConditionalFormatIconType,
    pub reverse_icons: bool,
    pub icons_only: bool,
    pub icons: Vec<CustomIconSpec>,
    pub multi_range: Option<String>,
    pub stop_if_true: bool,
}

impl Default for IconSetSpec {
    fn default() -> Self {
        IconSetSpec {
            icon_type: ConditionalFormatIconType::ThreeTrafficLights,
            reverse_icons: false,
            icons_only: false,
            icons: Vec::new(),
            multi_range: None,
            stop_if_true: false,
        }
    }
}

/// An icon set (traffic lights, arrows, ...).
///
/// `icons` optionally replaces the default thresholds with a list of
/// `ConditionalFormatCustomIcon` (one per icon in the set: 3, 4 or 5).
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct ConditionalFormatIconSet {
    pub spec: IconSetSpec,
}

impl ConditionalFormatIconSet {
    fn get_icons(&self) -> Vec<ConditionalFormatCustomIcon> {
        self.spec.icons.iter().cloned().map(|spec| ConditionalFormatCustomIcon { spec }).collect()
    }

    fn set_icons(&mut self, value: Vec<ConditionalFormatCustomIcon>) -> PyResult<()> {
        for icon in &value {
            icon.spec.validate()?;
        }
        self.spec.icons = value.into_iter().map(|icon| icon.spec).collect();
        Ok(())
    }

    fn on_change(&mut self) -> PyResult<()> {
        check_multi_range(&self.spec.multi_range)
    }

    fn to_rust(&self) -> PyResult<RIconSet> {
        let s = &self.spec;
        let count = s.icon_type.icon_count();
        if !s.icons.is_empty() && s.icons.len() != count {
            return Err(param_err(format!(
                "{:?} has {count} icons but {} custom icons were given",
                s.icon_type,
                s.icons.len()
            )));
        }
        let mut cf = RIconSet::new()
            .set_icon_type(s.icon_type.into())
            .reverse_icons(s.reverse_icons)
            .show_icons_only(s.icons_only);
        if !s.icons.is_empty() {
            let icons: Vec<RCustomIcon> = s.icons.iter().map(CustomIconSpec::to_rust).collect();
            cf = cf.set_icons(&icons);
        }
        Ok(apply_common!(cf, s))
    }
}

py_props!(ConditionalFormatIconSet
    fields {
        icon_type: ConditionalFormatIconType,
        reverse_icons: bool,
        icons_only: bool,
        multi_range: Option<String>,
        stop_if_true: bool,
    }
    custom {
        icons: Vec<ConditionalFormatCustomIcon>,
    }
);

#[pymethods]
impl ConditionalFormatIconSet {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut cf = ConditionalFormatIconSet { spec: IconSetSpec::default() };
        cf.apply_kwargs(kwargs)?;
        Ok(cf)
    }
}

spec_methods!(ConditionalFormatIconSet);

// ===========================================================================
// Worksheet methods
// ===========================================================================

/// Add any conditional format object to a worksheet range.
fn add_any(ws: &mut RWorksheet, range: CellRange, cf: &Bound<'_, PyAny>) -> PyResult<()> {
    let (r1, c1, r2, c2) = range;
    macro_rules! try_add {
        ($($t:ty),*) => {
            $(
                if let Ok(cf) = cf.cast::<$t>() {
                    let rust = cf.borrow().to_rust()?;
                    ws.add_conditional_format(r1, c1, r2, c2, &rust).py()?;
                    return Ok(());
                }
            )*
        };
    }
    try_add!(
        ConditionalFormatCell, ConditionalFormatText, ConditionalFormatDate,
        ConditionalFormatAverage, ConditionalFormatTop, ConditionalFormatDuplicate,
        ConditionalFormatBlank, ConditionalFormatError, ConditionalFormatFormula,
        ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatDataBar,
        ConditionalFormatIconSet
    );
    Err(PyTypeError::new_err(format!(
        "expected a ConditionalFormat* object (ConditionalFormatCell, ConditionalFormatDataBar, ...), not {}",
        cf.get_type().name()?
    )))
}

/// Shared body of the legacy type-specific `add_conditional_format_*` methods.
fn add_legacy<T: pyo3::PyTypeInfo>(
    ws: &mut Worksheet,
    method: &'static str,
    args: &Bound<'_, PyTuple>,
    kwargs: Option<&Bound<'_, PyDict>>,
) -> PyResult<()> {
    let (range, rest) = split_range(args.as_slice())?;
    let p = Params::new(method, rest, kwargs, &["cond_fmt"], 1)?;
    let cf = p.required_any(0)?;
    if !cf.is_instance_of::<T>() {
        return Err(PyTypeError::new_err(format!(
            "{method}() expects a {}, not {}",
            T::type_object(cf.py()).name()?,
            cf.get_type().name()?
        )));
    }
    add_any(&mut ws.inner, range, cf)
}

#[pymethods]
impl Worksheet {
    /// Add a conditional format to a range: `add_conditional_format("A1:A10", cf)`
    /// or `add_conditional_format(first_row, first_col, last_row, last_col, cf)`.
    ///
    /// `cf` is any `ConditionalFormat*` object. Use the object's `multi_range`
    /// to apply it to several ranges at once.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_conditional_format(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let p = Params::new("add_conditional_format", rest, kwargs, &["conditional_format"], 1)?;
        add_any(&mut self.inner, range, p.required_any(0)?)
    }

    /// Add a 2 color scale (kept for compatibility; prefer `add_conditional_format`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_conditional_format_2color_scale(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        add_legacy::<ConditionalFormat2ColorScale>(self, "add_conditional_format_2color_scale", args, kwargs)
    }

    /// Add a 3 color scale (kept for compatibility; prefer `add_conditional_format`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_conditional_format_3color_scale(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        add_legacy::<ConditionalFormat3ColorScale>(self, "add_conditional_format_3color_scale", args, kwargs)
    }

    /// Add a data bar (kept for compatibility; prefer `add_conditional_format`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_conditional_format_data_bar(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        add_legacy::<ConditionalFormatDataBar>(self, "add_conditional_format_data_bar", args, kwargs)
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ConditionalFormatCellRule>()?;
    m.add_class::<ConditionalFormatTextRule>()?;
    m.add_class::<ConditionalFormatTopRule>()?;
    m.add_class::<ConditionalFormatAverageRule>()?;
    m.add_class::<ConditionalFormatDateRule>()?;
    m.add_class::<ConditionalFormatType>()?;
    m.add_class::<ConditionalFormatDataBarDirection>()?;
    m.add_class::<ConditionalFormatDataBarAxisPosition>()?;
    m.add_class::<ConditionalFormatIconType>()?;
    m.add_class::<ConditionalFormatCell>()?;
    m.add_class::<ConditionalFormatText>()?;
    m.add_class::<ConditionalFormatDate>()?;
    m.add_class::<ConditionalFormatAverage>()?;
    m.add_class::<ConditionalFormatTop>()?;
    m.add_class::<ConditionalFormatDuplicate>()?;
    m.add_class::<ConditionalFormatBlank>()?;
    m.add_class::<ConditionalFormatError>()?;
    m.add_class::<ConditionalFormatFormula>()?;
    m.add_class::<ConditionalFormat2ColorScale>()?;
    m.add_class::<ConditionalFormat3ColorScale>()?;
    m.add_class::<ConditionalFormatDataBar>()?;
    m.add_class::<ConditionalFormatIconSet>()?;
    m.add_class::<ConditionalFormatCustomIcon>()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn icon_counts() {
        assert_eq!(ConditionalFormatIconType::ThreeStars.icon_count(), 3);
        assert_eq!(ConditionalFormatIconType::FourRedToBlack.icon_count(), 4);
        assert_eq!(ConditionalFormatIconType::FiveBoxes.icon_count(), 5);
    }

    #[test]
    fn thresholds() {
        assert!(needs_value(CfType::Percentile) && !needs_value(CfType::Lowest));
        assert!(check_threshold("min", CfType::Number, &Some(RuleValue::Int(3)), SCALE_MIN).is_ok());
        assert!(check_threshold("max", CfType::Automatic, &None, BAR_MAX).is_ok());
        assert!(check_threshold("mid", CfType::Percent, &Some(RuleValue::Number(100.0)), SCALE_MID).is_ok());
        // Error cases build Python exceptions, so they are tested from Python.
        assert!(!SCALE_MIN.contains(&CfType::Highest) && !BAR_MAX.contains(&CfType::Lowest));
    }
}
