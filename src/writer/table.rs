//! Worksheet tables: `Table`, `TableColumn`, `TableStyle`, `TableFunction`
//! and `Worksheet.add_table()`.

use std::hash::{Hash, Hasher};

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyString, PyTuple};
use rust_xlsxwriter::{
    Formula as RFormula, Table as RTable, TableColumn as RTableColumn,
    TableFunction as RTableFunction, TableStyle as RTableStyle,
};

use crate::common::cellref::{range_name, split_range, Params};
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::{Format, FormatSpec};
use crate::writer::conditional_format::{format_accessors, spec_methods, FormulaText};
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// The 61 built-in Excel table styles (default `Medium9`).
    TableStyle => RTableStyle {
        #[pyo3(name = "None_")] None,
        Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10,
        Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20,
        Light21,
        Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10,
        Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19,
        Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28,
        Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11,
    }
);

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum FunctionKind {
    None,
    Average,
    Count,
    CountNumbers,
    Max,
    Min,
    Sum,
    StdDev,
    Var,
    Custom,
}

impl FunctionKind {
    fn name(self) -> &'static str {
        match self {
            FunctionKind::None => "None_",
            FunctionKind::Average => "Average",
            FunctionKind::Count => "Count",
            FunctionKind::CountNumbers => "CountNumbers",
            FunctionKind::Max => "Max",
            FunctionKind::Min => "Min",
            FunctionKind::Sum => "Sum",
            FunctionKind::StdDev => "StdDev",
            FunctionKind::Var => "Var",
            FunctionKind::Custom => "Custom",
        }
    }
}

/// A total row function for a `TableColumn`.
///
/// Use a constant (`TableFunction.Sum`, `.Average`, `.Count`, `.CountNumbers`,
/// `.Max`, `.Min`, `.StdDev`, `.Var`, `.None_`) or a custom formula with
/// `TableFunction.custom("=SUM([Sales])")`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableFunction {
    kind: FunctionKind,
    formula: Option<String>,
}

impl TableFunction {
    const fn of(kind: FunctionKind) -> Self {
        TableFunction { kind, formula: None }
    }

    fn to_rust(&self) -> RTableFunction {
        match self.kind {
            FunctionKind::None => RTableFunction::None,
            FunctionKind::Average => RTableFunction::Average,
            FunctionKind::Count => RTableFunction::Count,
            FunctionKind::CountNumbers => RTableFunction::CountNumbers,
            FunctionKind::Max => RTableFunction::Max,
            FunctionKind::Min => RTableFunction::Min,
            FunctionKind::Sum => RTableFunction::Sum,
            FunctionKind::StdDev => RTableFunction::StdDev,
            FunctionKind::Var => RTableFunction::Var,
            FunctionKind::Custom => {
                RTableFunction::Custom(RFormula::new(self.formula.as_deref().unwrap_or_default()))
            }
        }
    }
}

#[pymethods]
impl TableFunction {
    /// No total function (the total row cell is left empty).
    #[classattr]
    #[pyo3(name = "None_")]
    fn none_() -> Self {
        Self::of(FunctionKind::None)
    }

    /// `SUBTOTAL(101, ...)`: the average.
    #[classattr]
    #[pyo3(name = "Average")]
    fn average() -> Self {
        Self::of(FunctionKind::Average)
    }

    /// `SUBTOTAL(103, ...)`: the count of non-empty cells.
    #[classattr]
    #[pyo3(name = "Count")]
    fn count() -> Self {
        Self::of(FunctionKind::Count)
    }

    /// `SUBTOTAL(102, ...)`: the count of numbers.
    #[classattr]
    #[pyo3(name = "CountNumbers")]
    fn count_numbers() -> Self {
        Self::of(FunctionKind::CountNumbers)
    }

    /// `SUBTOTAL(104, ...)`: the maximum.
    #[classattr]
    #[pyo3(name = "Max")]
    fn max() -> Self {
        Self::of(FunctionKind::Max)
    }

    /// `SUBTOTAL(105, ...)`: the minimum.
    #[classattr]
    #[pyo3(name = "Min")]
    fn min() -> Self {
        Self::of(FunctionKind::Min)
    }

    /// `SUBTOTAL(109, ...)`: the sum.
    #[classattr]
    #[pyo3(name = "Sum")]
    fn sum() -> Self {
        Self::of(FunctionKind::Sum)
    }

    /// `SUBTOTAL(107, ...)`: the standard deviation.
    #[classattr]
    #[pyo3(name = "StdDev")]
    fn std_dev() -> Self {
        Self::of(FunctionKind::StdDev)
    }

    /// `SUBTOTAL(110, ...)`: the variance.
    #[classattr]
    #[pyo3(name = "Var")]
    fn var() -> Self {
        Self::of(FunctionKind::Var)
    }

    /// A custom total formula, e.g. `TableFunction.custom("=SUM([Q1])")`.
    #[staticmethod]
    fn custom(formula: FormulaText) -> PyResult<Self> {
        if formula.0.trim().trim_start_matches('=').is_empty() {
            return Err(param_err("TableFunction.custom() needs a non-empty formula"));
        }
        Ok(TableFunction { kind: FunctionKind::Custom, formula: Some(formula.0) })
    }

    /// The function name: "Sum", "Average", ..., "None_" or "Custom".
    #[getter]
    fn name(&self) -> &'static str {
        self.kind.name()
    }

    /// The formula of a custom function, otherwise None.
    #[getter]
    fn formula(&self) -> Option<String> {
        self.formula.clone()
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<TableFunction>() {
            Ok(other) => self == other.get(),
            Err(_) => false,
        }
    }

    fn __hash__(&self) -> u64 {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        self.kind.hash(&mut hasher);
        self.formula.hash(&mut hasher);
        hasher.finish()
    }

    fn __repr__(&self) -> String {
        match &self.formula {
            Some(formula) => format!("TableFunction.custom({formula:?})"),
            None => format!("TableFunction.{}", self.kind.name()),
        }
    }
}

// ---------------------------------------------------------------------------
// TableColumn
// ---------------------------------------------------------------------------

/// Plain description of a table column.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct TableColumnSpec {
    pub header: Option<String>,
    pub total_function: Option<TableFunction>,
    pub total_label: Option<String>,
    pub formula: Option<FormulaText>,
    pub format: Option<FormatSpec>,
    pub header_format: Option<FormatSpec>,
}

impl TableColumnSpec {
    fn to_rust(&self) -> RTableColumn {
        let mut column = RTableColumn::new();
        if let Some(header) = &self.header {
            column = column.set_header(header);
        }
        if let Some(function) = &self.total_function {
            column = column.set_total_function(function.to_rust());
        }
        if let Some(label) = &self.total_label {
            column = column.set_total_label(label);
        }
        if let Some(formula) = &self.formula {
            column = column.set_formula(RFormula::new(&formula.0));
        }
        if let Some(format) = &self.format {
            column = column.set_format(format.to_rust());
        }
        if let Some(format) = &self.header_format {
            column = column.set_header_format(format.to_rust());
        }
        column
    }
}

/// Settings for one column of a `Table`.
///
/// `TableColumn("Sales", total_function=TableFunction.Sum, format=money)`.
/// `formula` fills the column's data rows, e.g. `"=[@Q1]+[@Q2]"`; `format`
/// applies to the data (and total) cells; `header_format` to the header.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct TableColumn {
    pub spec: TableColumnSpec,
}

impl TableColumn {
    fn on_change(&mut self) -> PyResult<()> {
        if let Some(header) = &self.spec.header
            && header.encode_utf16().count() > 255 {
                return Err(param_err("a table column header is limited to 255 characters"));
            }
        Ok(())
    }
}

format_accessors!(TableColumn: format, header_format);

py_props!(TableColumn
    fields {
        header: Option<String>,
        total_function: Option<TableFunction>,
        total_label: Option<String>,
        formula: Option<FormulaText>,
    }
    custom {
        format: Option<Format>,
        header_format: Option<Format>,
    }
);

#[pymethods]
impl TableColumn {
    #[new]
    #[pyo3(signature = (header=None, **kwargs))]
    fn py_new(header: Option<String>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut column = TableColumn {
            spec: TableColumnSpec { header, ..TableColumnSpec::default() },
        };
        column.apply_kwargs(kwargs)?;
        Ok(column)
    }
}

spec_methods!(TableColumn);

/// A list of columns given as `TableColumn` objects or header strings.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ColumnList(pub Vec<TableColumnSpec>);

impl<'a, 'py> FromPyObject<'a, 'py> for ColumnList {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if obj.is_instance_of::<PyString>() {
            return Err(PyTypeError::new_err("columns must be a list of TableColumn or str, not a str"));
        }
        let mut columns = Vec::new();
        for item in obj.try_iter()? {
            let item = item?;
            if let Ok(column) = item.cast::<TableColumn>() {
                columns.push(column.borrow().spec.clone());
            } else if let Ok(header) = item.cast::<PyString>() {
                columns.push(TableColumnSpec {
                    header: Some(header.to_str()?.to_string()),
                    ..TableColumnSpec::default()
                });
            } else {
                return Err(PyTypeError::new_err(format!(
                    "table columns must be TableColumn objects or header strings, not {}",
                    item.get_type().name()?
                )));
            }
        }
        Ok(ColumnList(columns))
    }
}

impl<'py> IntoPyObject<'py> for ColumnList {
    type Target = PyList;
    type Output = Bound<'py, PyList>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        PyList::new(py, self.0.into_iter().map(|spec| TableColumn { spec }))
    }
}

// ---------------------------------------------------------------------------
// Table
// ---------------------------------------------------------------------------

/// Plain description of a worksheet table.
#[derive(Clone, Debug, PartialEq)]
pub struct TableSpec {
    pub header_row: bool,
    pub total_row: bool,
    pub banded_rows: bool,
    pub banded_columns: bool,
    pub first_column: bool,
    pub last_column: bool,
    pub autofilter: bool,
    pub columns: ColumnList,
    pub name: Option<String>,
    pub style: TableStyle,
    pub alt_text: Option<String>,
    pub alt_text_title: Option<String>,
}

impl Default for TableSpec {
    fn default() -> Self {
        TableSpec {
            header_row: true,
            total_row: false,
            banded_rows: true,
            banded_columns: false,
            first_column: false,
            last_column: false,
            autofilter: true,
            columns: ColumnList::default(),
            name: None,
            style: TableStyle::Medium9,
            alt_text: None,
            alt_text_title: None,
        }
    }
}

impl TableSpec {
    /// Build the equivalent `rust_xlsxwriter::Table`.
    pub fn to_rust(&self) -> RTable {
        let mut table = RTable::new()
            .set_header_row(self.header_row)
            .set_total_row(self.total_row)
            .set_banded_rows(self.banded_rows)
            .set_banded_columns(self.banded_columns)
            .set_first_column(self.first_column)
            .set_last_column(self.last_column)
            .set_autofilter(self.autofilter)
            .set_style(self.style.into());
        if !self.columns.0.is_empty() {
            let columns: Vec<RTableColumn> = self.columns.0.iter().map(TableColumnSpec::to_rust).collect();
            table = table.set_columns(&columns);
        }
        if let Some(name) = &self.name {
            table = table.set_name(name);
        }
        if let Some(text) = &self.alt_text {
            table = table.set_alt_text(text);
        }
        if let Some(text) = &self.alt_text_title {
            table = table.set_alt_text_title(text);
        }
        table
    }
}

/// A worksheet table (Excel's "Format as Table").
///
/// ```python
/// table = Table(columns=["Product", "Q1", TableColumn("Total", formula="=[@Q1]*2")],
///               style=TableStyle.Medium2, total_row=True)
/// ws.add_table("B2:D6", table)
/// ```
///
/// Write the data with the normal `write*()` methods; the header row is
/// written by the table (from `columns`, or "Column1", ... by default).
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct Table {
    pub spec: TableSpec,
}

impl Table {
    fn on_change(&mut self) -> PyResult<()> {
        if let Some(name) = &self.spec.name
            && name.is_empty() {
                return Err(param_err("a table name must not be empty (use None for the default)"));
            }
        Ok(())
    }
}

py_props!(Table
    fields {
        header_row: bool,
        total_row: bool,
        banded_rows: bool,
        banded_columns: bool,
        first_column: bool,
        last_column: bool,
        autofilter: bool,
        columns: ColumnList,
        name: Option<String>,
        style: TableStyle,
        alt_text: Option<String>,
        alt_text_title: Option<String>,
    }
);

#[pymethods]
impl Table {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut table = Table { spec: TableSpec::default() };
        table.apply_kwargs(kwargs)?;
        Ok(table)
    }
}

spec_methods!(Table);

#[pymethods]
impl Worksheet {
    /// Add a table over a range: `add_table("B2:D6", table)` or
    /// `add_table(first_row, first_col, last_row, last_col, table)`.
    /// Without `table` a default table is added.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_table(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let (r1, c1, r2, c2) = range;
        let p = Params::new("add_table", rest, kwargs, &["table"], 0)?;
        let table: Option<PyRef<'_, Table>> = p.extract(0)?;
        let spec = table.map(|t| t.spec.clone()).unwrap_or_default();
        let width = usize::from(c2 - c1) + 1;
        if spec.columns.0.len() > width {
            return Err(param_err(format!(
                "the table has {} columns but the range {} is only {width} columns wide",
                spec.columns.0.len(),
                range_name(range, false)
            )));
        }
        self.inner.add_table(r1, c1, r2, c2, &spec.to_rust()).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Table>()?;
    m.add_class::<TableColumn>()?;
    m.add_class::<TableStyle>()?;
    m.add_class::<TableFunction>()?;
    Ok(())
}
