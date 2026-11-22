//! Data validation: `DataValidation`, `DataValidationRule`,
//! `DataValidationErrorStyle` and `Worksheet.add_data_validation()`.
//!
//! What a validation allows is set with one property per rust_xlsxwriter
//! `allow_*()` method (`whole_number`, `decimal_number`, `list_strings`, ...);
//! setting one replaces the previous kind. The chainable `allow_*()` methods
//! do the same.

use chrono::{NaiveDateTime, NaiveTime};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyTuple};
use rust_xlsxwriter::{
    DataValidation as RDataValidation, DataValidationErrorStyle as RErrorStyle,
    DataValidationRule as RRule, Formula as RFormula,
};

use crate::common::cellref::{split_range, Params};
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::value::{excel_serial, time_fraction};
use crate::writer::conditional_format::{
    comparison_rule_class, comparison_to_rust, spec_methods, ComparisonRule, FormulaText, RuleValue,
};
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// The style of the error dialog shown for invalid input.
    DataValidationErrorStyle => RErrorStyle { Stop, Warning, Information }
);

/// What a `DataValidation` currently allows (see `DataValidation.validation_type`).
#[pyclass(module = "pyrust_xlsxwriter", eq, eq_int, hash, frozen, skip_from_py_object)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum DataValidationType {
    Any,
    WholeNumber,
    DecimalNumber,
    Date,
    Time,
    TextLength,
    List,
    Custom,
}

comparison_rule_class!(
    /// A data validation rule: `DataValidationRule.between(1, 10)`,
    /// `.greater_than(date(2025, 1, 1))`, `.equal_to("=$B$1")`, ...
    ///
    /// The allowed value types depend on the validation: ints (whole numbers,
    /// text length), numbers (decimals), dates/datetimes (dates), times
    /// (times), or formulas (`"=..."` strings or `Formula`) for any of them.
    DataValidationRule
);

/// The kind of input allowed, with its rule.
#[derive(Clone, Debug, PartialEq)]
pub enum Allow {
    Any,
    WholeNumber(ComparisonRule),
    DecimalNumber(ComparisonRule),
    Date(ComparisonRule),
    Time(ComparisonRule),
    TextLength(ComparisonRule),
    ListStrings(Vec<String>),
    ListFormula(String),
    Custom(String),
}

/// Plain description of a data validation.
#[derive(Clone, Debug, PartialEq)]
pub struct DataValidationSpec {
    pub allow: Allow,
    pub ignore_blank: bool,
    pub show_dropdown: bool,
    pub show_input_message: bool,
    pub show_error_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    pub error_style: DataValidationErrorStyle,
    pub multi_range: Option<String>,
}

impl Default for DataValidationSpec {
    fn default() -> Self {
        DataValidationSpec {
            allow: Allow::Any,
            ignore_blank: true,
            show_dropdown: true,
            show_input_message: true,
            show_error_message: true,
            input_title: None,
            input_message: None,
            error_title: None,
            error_message: None,
            error_style: DataValidationErrorStyle::Stop,
            multi_range: None,
        }
    }
}

/// Excel's length limits count UTF-16 code units.
fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

fn check_length(what: &str, text: &Option<String>, limit: usize) -> PyResult<()> {
    if let Some(text) = text {
        let length = utf16_len(text);
        if length > limit {
            return Err(param_err(format!(
                "{what} is {length} characters long; Excel's limit is {limit}"
            )));
        }
    }
    Ok(())
}

/// A formula value as the string written to the file.
fn formula_string(value: &RuleValue) -> String {
    match value {
        RuleValue::Int(v) => v.to_string(),
        RuleValue::Number(v) => v.to_string(),
        RuleValue::Bool(v) => v.to_string().to_uppercase(),
        RuleValue::Text(v) => format!("\"{}\"", v.replace('"', "\"\"")),
        RuleValue::Formula { text, .. } => text.clone(),
        RuleValue::DateTime(v) => excel_serial(v).to_string(),
        RuleValue::Date(v) => excel_serial(&v.and_time(NaiveTime::MIN)).to_string(),
        RuleValue::Time(v) => time_fraction(v).to_string(),
    }
}

/// Check that each rule value is allowed for the validation kind.
fn check_values(kind: &str, rule: &ComparisonRule, ok: impl Fn(&RuleValue) -> bool, expected: &str) -> PyResult<()> {
    for value in &rule.values {
        if !value.is_formula() && !ok(value) {
            return Err(param_err(format!(
                "{kind} rule values must be {expected} or formulas, not {} ({value:?})",
                value.kind()
            )));
        }
    }
    Ok(())
}

fn is_i32(value: &RuleValue) -> bool {
    matches!(value, RuleValue::Int(v) if i32::try_from(*v).is_ok())
}

fn is_u32(value: &RuleValue) -> bool {
    matches!(value, RuleValue::Int(v) if u32::try_from(*v).is_ok())
}

fn is_number(value: &RuleValue) -> bool {
    matches!(value, RuleValue::Int(_) | RuleValue::Number(_))
}

fn is_date(value: &RuleValue) -> bool {
    matches!(value, RuleValue::Date(_) | RuleValue::DateTime(_))
}

fn is_time(value: &RuleValue) -> bool {
    matches!(value, RuleValue::Time(_))
}

impl Allow {
    fn validation_type(&self) -> DataValidationType {
        match self {
            Allow::Any => DataValidationType::Any,
            Allow::WholeNumber(_) => DataValidationType::WholeNumber,
            Allow::DecimalNumber(_) => DataValidationType::DecimalNumber,
            Allow::Date(_) => DataValidationType::Date,
            Allow::Time(_) => DataValidationType::Time,
            Allow::TextLength(_) => DataValidationType::TextLength,
            Allow::ListStrings(_) | Allow::ListFormula(_) => DataValidationType::List,
            Allow::Custom(_) => DataValidationType::Custom,
        }
    }

    fn validate(&self) -> PyResult<()> {
        match self {
            Allow::Any => Ok(()),
            Allow::WholeNumber(rule) => check_values("whole_number", rule, is_i32, "ints (32-bit)"),
            Allow::DecimalNumber(rule) => check_values("decimal_number", rule, is_number, "numbers"),
            Allow::Date(rule) => check_values("date", rule, is_date, "dates/datetimes"),
            Allow::Time(rule) => check_values("time", rule, is_time, "times"),
            Allow::TextLength(rule) => {
                check_values("text_length", rule, is_u32, "non-negative ints")
            }
            Allow::ListStrings(list) => {
                if list.is_empty() {
                    return Err(param_err("list_strings must contain at least one string"));
                }
                let joined: Vec<String> = list.iter().map(|s| s.replace('"', "\"\"")).collect();
                let length = utf16_len(&joined.join(","));
                if length > 255 {
                    return Err(param_err(format!(
                        "list_strings is {length} characters long including commas; Excel's limit is 255 \
                         (put the values in cells and use list_formula instead)"
                    )));
                }
                Ok(())
            }
            Allow::ListFormula(formula) | Allow::Custom(formula) => {
                if formula.trim().trim_start_matches('=').is_empty() {
                    return Err(param_err("the validation formula must not be empty"));
                }
                Ok(())
            }
        }
    }
}

impl DataValidationSpec {
    fn validate(&self) -> PyResult<()> {
        self.allow.validate()?;
        check_length("input_title", &self.input_title, 32)?;
        check_length("error_title", &self.error_title, 32)?;
        check_length("input_message", &self.input_message, 255)?;
        check_length("error_message", &self.error_message, 255)?;
        if let Some(range) = &self.multi_range
            && range.trim().is_empty() {
                return Err(param_err("multi_range must not be empty (use None to unset it)"));
            }
        Ok(())
    }

    /// Build the equivalent `rust_xlsxwriter::DataValidation`.
    pub fn to_rust(&self) -> PyResult<RDataValidation> {
        self.validate()?;
        let formula = |v: &RuleValue| -> PyResult<RFormula> { Ok(RFormula::new(formula_string(v))) };
        let any_formula = |rule: &ComparisonRule| rule.values.iter().any(RuleValue::is_formula);

        let mut dv = RDataValidation::new();
        dv = match &self.allow {
            Allow::Any => dv.allow_any_value(),
            Allow::WholeNumber(rule) if any_formula(rule) => {
                dv.allow_whole_number_formula(comparison_to_rust!(RRule, rule, formula))
            }
            Allow::WholeNumber(rule) => dv.allow_whole_number(comparison_to_rust!(RRule, rule, |v: &RuleValue| {
                match v {
                    RuleValue::Int(i) => i32::try_from(*i)
                        .map_err(|_| param_err(format!("whole_number value {i} is outside the 32-bit range"))),
                    other => Err(param_err(format!("whole_number value {other:?} is not an int"))),
                }
            })),
            Allow::DecimalNumber(rule) if any_formula(rule) => {
                dv.allow_decimal_number_formula(comparison_to_rust!(RRule, rule, formula))
            }
            Allow::DecimalNumber(rule) => dv.allow_decimal_number(comparison_to_rust!(RRule, rule, |v: &RuleValue| {
                v.as_f64().ok_or_else(|| param_err(format!("decimal_number value {v:?} is not a number")))
            })),
            Allow::Date(rule) if any_formula(rule) => {
                dv.allow_date_formula(comparison_to_rust!(RRule, rule, formula))
            }
            Allow::Date(rule) => dv.allow_date(comparison_to_rust!(RRule, rule, |v: &RuleValue| -> PyResult<NaiveDateTime> {
                match v {
                    RuleValue::DateTime(dt) => Ok(*dt),
                    RuleValue::Date(d) => Ok(d.and_time(NaiveTime::MIN)),
                    other => Err(param_err(format!("date value {other:?} is not a date"))),
                }
            })),
            Allow::Time(rule) if any_formula(rule) => {
                dv.allow_time_formula(comparison_to_rust!(RRule, rule, formula))
            }
            Allow::Time(rule) => dv.allow_time(comparison_to_rust!(RRule, rule, |v: &RuleValue| -> PyResult<NaiveTime> {
                match v {
                    RuleValue::Time(t) => Ok(*t),
                    other => Err(param_err(format!("time value {other:?} is not a time"))),
                }
            })),
            Allow::TextLength(rule) if any_formula(rule) => {
                dv.allow_text_length_formula(comparison_to_rust!(RRule, rule, formula))
            }
            Allow::TextLength(rule) => dv.allow_text_length(comparison_to_rust!(RRule, rule, |v: &RuleValue| {
                match v {
                    RuleValue::Int(i) => u32::try_from(*i)
                        .map_err(|_| param_err(format!("text_length value {i} must be a non-negative 32-bit int"))),
                    other => Err(param_err(format!("text_length value {other:?} is not an int"))),
                }
            })),
            Allow::ListStrings(list) => dv.allow_list_strings(list).py()?,
            Allow::ListFormula(text) => dv.allow_list_formula(RFormula::new(text)),
            Allow::Custom(text) => dv.allow_custom(RFormula::new(text)),
        };

        dv = dv
            .ignore_blank(self.ignore_blank)
            .show_dropdown(self.show_dropdown)
            .show_input_message(self.show_input_message)
            .show_error_message(self.show_error_message)
            .set_error_style(self.error_style.into());
        if let Some(text) = &self.input_title {
            dv = dv.set_input_title(text).py()?;
        }
        if let Some(text) = &self.input_message {
            dv = dv.set_input_message(text).py()?;
        }
        if let Some(text) = &self.error_title {
            dv = dv.set_error_title(text).py()?;
        }
        if let Some(text) = &self.error_message {
            dv = dv.set_error_message(text).py()?;
        }
        if let Some(range) = &self.multi_range {
            dv = dv.set_multi_range(range.as_str());
        }
        Ok(dv)
    }
}

/// A data validation for a cell range.
///
/// ```python
/// dv = DataValidation(whole_number=DataValidationRule.between(1, 10),
///                     input_title="Enter a number", input_message="1 to 10")
/// ws.add_data_validation("B2:B10", dv)
/// ```
///
/// The kind of input is set by one of `whole_number`, `decimal_number`,
/// `date`, `time`, `text_length` (each a `DataValidationRule`),
/// `list_strings`, `list_formula`, `custom` (a formula), or the chainable
/// `allow_*()` methods; setting one replaces the previous kind. A new
/// validation allows any value.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct DataValidation {
    pub spec: DataValidationSpec,
}

/// Read a rule-kind property: the rule if `allow` is that kind.
macro_rules! rule_accessors {
    ($($get:ident, $set:ident => $variant:ident);* $(;)?) => {
        impl DataValidation {
            $(
                fn $get(&self) -> Option<DataValidationRule> {
                    match &self.spec.allow {
                        Allow::$variant(rule) => Some(DataValidationRule { rule: rule.clone() }),
                        _ => None,
                    }
                }

                fn $set(&mut self, value: Option<DataValidationRule>) -> PyResult<()> {
                    self.replace_allow(value.map(|r| Allow::$variant(r.rule)), DataValidationType::$variant)
                }
            )*
        }
    };
}

rule_accessors!(
    get_whole_number, set_whole_number => WholeNumber;
    get_decimal_number, set_decimal_number => DecimalNumber;
    get_date, set_date => Date;
    get_time, set_time => Time;
    get_text_length, set_text_length => TextLength;
);

impl DataValidation {
    fn on_change(&mut self) -> PyResult<()> {
        self.spec.validate()
    }

    /// Set `allow`; `None` resets to "any value" if `kind` is the current kind.
    fn replace_allow(&mut self, allow: Option<Allow>, kind: DataValidationType) -> PyResult<()> {
        match allow {
            Some(allow) => {
                allow.validate()?;
                self.spec.allow = allow;
            }
            None => {
                if self.spec.allow.validation_type() == kind {
                    self.spec.allow = Allow::Any;
                }
            }
        }
        Ok(())
    }

    fn get_list_strings(&self) -> Option<Vec<String>> {
        match &self.spec.allow {
            Allow::ListStrings(list) => Some(list.clone()),
            _ => None,
        }
    }

    fn set_list_strings(&mut self, value: Option<Vec<String>>) -> PyResult<()> {
        let is_list_strings = matches!(self.spec.allow, Allow::ListStrings(_));
        match value {
            Some(list) => self.replace_allow(Some(Allow::ListStrings(list)), DataValidationType::List),
            None if is_list_strings => self.replace_allow(None, DataValidationType::List),
            None => Ok(()),
        }
    }

    fn get_list_formula(&self) -> Option<FormulaText> {
        match &self.spec.allow {
            Allow::ListFormula(text) => Some(FormulaText(text.clone())),
            _ => None,
        }
    }

    fn set_list_formula(&mut self, value: Option<FormulaText>) -> PyResult<()> {
        let is_list_formula = matches!(self.spec.allow, Allow::ListFormula(_));
        match value {
            Some(f) => self.replace_allow(Some(Allow::ListFormula(f.0)), DataValidationType::List),
            None if is_list_formula => self.replace_allow(None, DataValidationType::List),
            None => Ok(()),
        }
    }

    fn get_custom(&self) -> Option<FormulaText> {
        match &self.spec.allow {
            Allow::Custom(text) => Some(FormulaText(text.clone())),
            _ => None,
        }
    }

    fn set_custom(&mut self, value: Option<FormulaText>) -> PyResult<()> {
        self.replace_allow(value.map(|f| Allow::Custom(f.0)), DataValidationType::Custom)
    }
}

py_props!(DataValidation
    fields {
        ignore_blank: bool,
        show_dropdown: bool,
        show_input_message: bool,
        show_error_message: bool,
        input_title: Option<String>,
        input_message: Option<String>,
        error_title: Option<String>,
        error_message: Option<String>,
        error_style: DataValidationErrorStyle,
        multi_range: Option<String>,
    }
    custom {
        whole_number: Option<DataValidationRule>,
        decimal_number: Option<DataValidationRule>,
        date: Option<DataValidationRule>,
        time: Option<DataValidationRule>,
        text_length: Option<DataValidationRule>,
        list_strings: Option<Vec<String>>,
        list_formula: Option<FormulaText>,
        custom: Option<FormulaText>,
    }
);

/// Chainable `allow_*()` methods, mirroring rust_xlsxwriter.
macro_rules! allow_methods {
    ($($(#[$meta:meta])* $method:ident => $setter:ident);* $(;)?) => {
        #[pymethods]
        impl DataValidation {
            $(
                $(#[$meta])*
                fn $method(mut slf: PyRefMut<'_, Self>, rule: DataValidationRule) -> PyResult<PyRefMut<'_, Self>> {
                    slf.$setter(Some(rule))?;
                    Ok(slf)
                }
            )*
        }
    };
}

allow_methods!(
    /// Allow whole numbers matching `rule` (int values or formulas).
    allow_whole_number => set_whole_number;
    /// Same as `allow_whole_number()` with formula values: `DataValidationRule.less_than("=B1")`.
    allow_whole_number_formula => set_whole_number;
    /// Allow decimal numbers matching `rule`.
    allow_decimal_number => set_decimal_number;
    /// Same as `allow_decimal_number()` with formula values.
    allow_decimal_number_formula => set_decimal_number;
    /// Allow dates matching `rule` (date/datetime values or formulas).
    allow_date => set_date;
    /// Same as `allow_date()` with formula values.
    allow_date_formula => set_date;
    /// Allow times matching `rule` (time values or formulas).
    allow_time => set_time;
    /// Same as `allow_time()` with formula values.
    allow_time_formula => set_time;
    /// Allow text whose length matches `rule`.
    allow_text_length => set_text_length;
    /// Same as `allow_text_length()` with formula values.
    allow_text_length_formula => set_text_length;
);

#[pymethods]
impl DataValidation {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut dv = DataValidation { spec: DataValidationSpec::default() };
        dv.apply_kwargs(kwargs)?;
        Ok(dv)
    }

    /// Allow one of a list of strings, shown as a dropdown (255 characters max in total).
    fn allow_list_strings(mut slf: PyRefMut<'_, Self>, list: Vec<String>) -> PyResult<PyRefMut<'_, Self>> {
        slf.set_list_strings(Some(list))?;
        Ok(slf)
    }

    /// Allow one of the values in a range/formula, e.g. `"=$E$1:$E$5"`.
    fn allow_list_formula(mut slf: PyRefMut<'_, Self>, formula: FormulaText) -> PyResult<PyRefMut<'_, Self>> {
        slf.set_list_formula(Some(formula))?;
        Ok(slf)
    }

    /// Allow values for which a formula is true, e.g. `"=ISTEXT(A1)"`.
    fn allow_custom(mut slf: PyRefMut<'_, Self>, formula: FormulaText) -> PyResult<PyRefMut<'_, Self>> {
        slf.set_custom(Some(formula))?;
        Ok(slf)
    }

    /// Allow any value (useful with an input message only).
    fn allow_any_value(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.spec.allow = Allow::Any;
        slf
    }

    /// The kind of input currently allowed.
    #[getter]
    fn validation_type(&self) -> DataValidationType {
        self.spec.allow.validation_type()
    }
}

spec_methods!(DataValidation);

#[pymethods]
impl Worksheet {
    /// Add a data validation to a range: `add_data_validation("B2:B10", dv)`
    /// or `add_data_validation(first_row, first_col, last_row, last_col, dv)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_data_validation(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("add_data_validation", rest, kwargs, &["data_validation"], 1)?;
        let dv: PyRef<'_, DataValidation> = p.required(0)?;
        let rust = dv.spec.to_rust()?;
        self.inner.add_data_validation(r1, c1, r2, c2, &rust).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<DataValidation>()?;
    m.add_class::<DataValidationRule>()?;
    m.add_class::<DataValidationErrorStyle>()?;
    m.add_class::<DataValidationType>()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;

    #[test]
    fn formula_strings_for_mixed_rules() {
        assert_eq!(formula_string(&RuleValue::Int(5)), "5");
        assert_eq!(formula_string(&RuleValue::Number(1.5)), "1.5");
        assert_eq!(formula_string(&RuleValue::Bool(true)), "TRUE");
        assert_eq!(formula_string(&RuleValue::Text("a\"b".into())), "\"a\"\"b\"");
        let date = NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
        assert_eq!(formula_string(&RuleValue::Date(date)), "45292");
        assert_eq!(formula_string(&RuleValue::Time(NaiveTime::from_hms_opt(12, 0, 0).unwrap())), "0.5");
    }

    #[test]
    fn utf16_lengths() {
        assert_eq!(utf16_len("abc"), 3);
        assert_eq!(utf16_len("😀"), 2);
    }
}
