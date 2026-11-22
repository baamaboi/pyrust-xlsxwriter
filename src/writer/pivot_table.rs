//! Pivot tables (rust_xlsxwriter 0.103+): `PivotTable`,
//! `PivotTableValueField`, `PivotTableFunction`, `PivotTableLayout`,
//! `PivotTableStyle` and `Worksheet.add_pivot_table()`.
//!
//! rust_xlsxwriter writes the pivot table definition without cached data, so
//! Excel builds the table when the file is opened.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyString, PyTuple};
use rust_xlsxwriter::{
    PivotTable as RPivotTable, PivotTableFunction as RPivotTableFunction,
    PivotTableLayout as RPivotTableLayout, PivotTableStyle as RPivotTableStyle,
    PivotTableValueField as RPivotTableValueField,
};

use crate::common::cellref::{split_cell, Params};
use crate::common::errors::{param_err, IntoPyResult};
use crate::writer::conditional_format::spec_methods;
use crate::writer::sparkline::RangeText;
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// How a pivot table value field summarizes its data (default `Sum`).
    PivotTableFunction => RPivotTableFunction {
        Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP,
    }
);

py_enum!(
    /// The report layout of a pivot table (default `Compact`).
    PivotTableLayout => RPivotTableLayout { Compact, Outline, Tabular }
);

py_enum!(
    /// The built-in Excel pivot table styles (default `Light16`).
    PivotTableStyle => RPivotTableStyle {
        #[pyo3(name = "None_")] None,
        Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10,
        Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20,
        Light21, Light22, Light23, Light24, Light25, Light26, Light27, Light28,
        Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10,
        Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19,
        Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28,
        Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10,
        Dark11, Dark12, Dark13, Dark14, Dark15, Dark16, Dark17, Dark18, Dark19, Dark20,
        Dark21, Dark22, Dark23, Dark24, Dark25, Dark26, Dark27, Dark28,
    }
);

// ---------------------------------------------------------------------------
// PivotTableValueField
// ---------------------------------------------------------------------------

/// Plain description of a pivot table value field.
#[derive(Clone, Debug, PartialEq)]
pub struct PivotTableValueFieldSpec {
    pub field: String,
    pub function: PivotTableFunction,
    pub custom_name: Option<String>,
    pub num_format: Option<String>,
}

impl PivotTableValueFieldSpec {
    fn new(field: String) -> Self {
        PivotTableValueFieldSpec {
            field,
            function: PivotTableFunction::Sum,
            custom_name: None,
            num_format: None,
        }
    }

    fn to_rust(&self) -> RPivotTableValueField {
        let mut value = RPivotTableValueField::new(&self.field).set_function(self.function.into());
        if let Some(name) = &self.custom_name {
            value = value.set_custom_name(name);
        }
        if let Some(format) = &self.num_format {
            value = value.set_num_format(format);
        }
        value
    }
}

/// A source data field summarized in the body of a pivot table:
/// `PivotTableValueField("Volume", function=PivotTableFunction.Average,
/// custom_name="Avg volume", num_format="#,##0")`.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct PivotTableValueField {
    pub spec: PivotTableValueFieldSpec,
}

impl PivotTableValueField {
    fn on_change(&mut self) -> PyResult<()> {
        if self.spec.field.is_empty() {
            return Err(param_err("a pivot table value field needs a source field name"));
        }
        if matches!(&self.spec.custom_name, Some(name) if name.is_empty()) {
            return Err(param_err("custom_name must not be empty (use None for the default)"));
        }
        Ok(())
    }
}

py_props!(PivotTableValueField
    fields {
        field: String,
        function: PivotTableFunction,
        custom_name: Option<String>,
        num_format: Option<String>,
    }
);

#[pymethods]
impl PivotTableValueField {
    #[new]
    #[pyo3(signature = (field, **kwargs))]
    fn py_new(field: String, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut value = PivotTableValueField { spec: PivotTableValueFieldSpec::new(field) };
        value.apply_kwargs(kwargs)?;
        Ok(value)
    }
}

spec_methods!(PivotTableValueField);

/// Value fields given as `PivotTableValueField` objects or field names.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ValueFieldList(pub Vec<PivotTableValueFieldSpec>);

fn value_field_spec(item: &Bound<'_, PyAny>) -> PyResult<PivotTableValueFieldSpec> {
    if let Ok(value) = item.cast::<PivotTableValueField>() {
        return Ok(value.borrow().spec.clone());
    }
    if let Ok(name) = item.cast::<PyString>() {
        let name = name.to_str()?;
        if name.is_empty() {
            return Err(param_err("a pivot table value field needs a source field name"));
        }
        return Ok(PivotTableValueFieldSpec::new(name.to_string()));
    }
    Err(PyTypeError::new_err(format!(
        "value fields must be PivotTableValueField objects or field names, not {}",
        item.get_type().name()?
    )))
}

impl<'a, 'py> FromPyObject<'a, 'py> for ValueFieldList {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if obj.is_instance_of::<PyString>() {
            return Err(PyTypeError::new_err("value_fields must be a list, not a str"));
        }
        let mut fields = Vec::new();
        for item in obj.try_iter()? {
            fields.push(value_field_spec(&item?)?);
        }
        Ok(ValueFieldList(fields))
    }
}

impl<'py> IntoPyObject<'py> for ValueFieldList {
    type Target = PyList;
    type Output = Bound<'py, PyList>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        PyList::new(py, self.0.into_iter().map(|spec| PivotTableValueField { spec }))
    }
}

// ---------------------------------------------------------------------------
// PivotTable
// ---------------------------------------------------------------------------

/// Plain description of a pivot table.
#[derive(Clone, Debug, PartialEq)]
pub struct PivotTableSpec {
    pub data_source: Option<RangeText>,
    pub field_names: Option<Vec<String>>,
    pub row_fields: Vec<String>,
    pub column_fields: Vec<String>,
    pub filter_fields: Vec<String>,
    pub value_fields: ValueFieldList,
    pub style: PivotTableStyle,
    pub layout: PivotTableLayout,
    pub row_grand_totals: bool,
    pub column_grand_totals: bool,
}

impl Default for PivotTableSpec {
    fn default() -> Self {
        PivotTableSpec {
            data_source: None,
            field_names: None,
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            filter_fields: Vec::new(),
            value_fields: ValueFieldList::default(),
            style: PivotTableStyle::Light16,
            layout: PivotTableLayout::Compact,
            row_grand_totals: true,
            column_grand_totals: true,
        }
    }
}

impl PivotTableSpec {
    fn validate(&self) -> PyResult<()> {
        for (kind, names) in [("row", &self.row_fields), ("column", &self.column_fields), ("filter", &self.filter_fields)] {
            if names.iter().any(String::is_empty) {
                return Err(param_err(format!("{kind} field names must not be empty")));
            }
        }
        Ok(())
    }

    /// Build the `rust_xlsxwriter::PivotTable`; a data source without a
    /// sheet name refers to `sheet`.
    pub fn to_rust(&self, sheet: &str) -> PyResult<RPivotTable> {
        self.validate()?;
        let source = self.data_source.as_ref().ok_or_else(|| {
            param_err("the pivot table has no data source: set PivotTable(data_source=\"Data!A1:D9\")")
        })?;
        let mut pivot = RPivotTable::new().set_data_source(source.qualified(sheet).as_str());
        if let Some(names) = &self.field_names {
            pivot = pivot.set_field_names(names);
        }
        for name in &self.row_fields {
            pivot = pivot.add_row_field(name);
        }
        for name in &self.column_fields {
            pivot = pivot.add_column_field(name);
        }
        for name in &self.filter_fields {
            pivot = pivot.add_filter_field(name);
        }
        for value in &self.value_fields.0 {
            pivot = pivot.add_value_field(&value.to_rust());
        }
        Ok(pivot
            .set_style(self.style.into())
            .set_layout(self.layout.into())
            .set_row_grand_totals(self.row_grand_totals)
            .set_column_grand_totals(self.column_grand_totals))
    }
}

/// A pivot table summarizing a range of source data.
///
/// ```python
/// pivot = PivotTable(data_source="Data!A1:D9", row_fields=["Region"],
///                    column_fields=["Item"], value_fields=["Volume"])
/// ws.add_pivot_table("A3", pivot)
/// ```
///
/// The first row of `data_source` holds the field names. Value fields are
/// `PivotTableValueField` objects (or field names, summed). The chainable
/// `add_*_field()` methods mirror rust_xlsxwriter. A `data_source` without a
/// sheet name refers to the sheet the pivot table is added to.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct PivotTable {
    pub spec: PivotTableSpec,
}

impl PivotTable {
    fn on_change(&mut self) -> PyResult<()> {
        self.spec.validate()
    }
}

py_props!(PivotTable
    fields {
        data_source: Option<RangeText>,
        field_names: Option<Vec<String>>,
        row_fields: Vec<String>,
        column_fields: Vec<String>,
        filter_fields: Vec<String>,
        value_fields: ValueFieldList,
        style: PivotTableStyle,
        layout: PivotTableLayout,
        row_grand_totals: bool,
        column_grand_totals: bool,
    }
);

fn check_field_name(name: &str) -> PyResult<()> {
    if name.is_empty() {
        return Err(param_err("a pivot table field name must not be empty"));
    }
    Ok(())
}

#[pymethods]
impl PivotTable {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut pivot = PivotTable { spec: PivotTableSpec::default() };
        pivot.apply_kwargs(kwargs)?;
        Ok(pivot)
    }

    /// Add a row field (a source column whose values become row labels); returns self.
    fn add_row_field(mut slf: PyRefMut<'_, Self>, name: String) -> PyResult<PyRefMut<'_, Self>> {
        check_field_name(&name)?;
        slf.spec.row_fields.push(name);
        Ok(slf)
    }

    /// Add a column field (values become column labels); returns self.
    fn add_column_field(mut slf: PyRefMut<'_, Self>, name: String) -> PyResult<PyRefMut<'_, Self>> {
        check_field_name(&name)?;
        slf.spec.column_fields.push(name);
        Ok(slf)
    }

    /// Add a filter field (shown above the pivot table); returns self.
    fn add_filter_field(mut slf: PyRefMut<'_, Self>, name: String) -> PyResult<PyRefMut<'_, Self>> {
        check_field_name(&name)?;
        slf.spec.filter_fields.push(name);
        Ok(slf)
    }

    /// Add a value field (a `PivotTableValueField` or a field name to sum); returns self.
    fn add_value_field<'py>(
        mut slf: PyRefMut<'py, Self>,
        value_field: &Bound<'py, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let spec = value_field_spec(value_field)?;
        slf.spec.value_fields.0.push(spec);
        Ok(slf)
    }
}

spec_methods!(PivotTable);

#[pymethods]
impl Worksheet {
    /// Add a pivot table with its top-left cell at `cell`:
    /// `add_pivot_table("A3", pivot)` or `add_pivot_table(row, col, pivot)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_pivot_table(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("add_pivot_table", rest, kwargs, &["pivot_table"], 1)?;
        let pivot: PyRef<'_, PivotTable> = p.required(0)?;
        let rust = pivot.spec.to_rust(&self.inner.name())?;
        self.inner.add_pivot_table(row, col, &rust).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PivotTable>()?;
    m.add_class::<PivotTableValueField>()?;
    m.add_class::<PivotTableFunction>()?;
    m.add_class::<PivotTableLayout>()?;
    m.add_class::<PivotTableStyle>()?;
    Ok(())
}
