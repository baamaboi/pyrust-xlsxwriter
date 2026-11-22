//! `Worksheet` for new files: the struct, its properties, cell writing and
//! row/column settings.
//!
//! Other feature areas (tables, images, charts, page setup, ...) add their
//! methods in their own files with additional `#[pymethods] impl Worksheet`
//! blocks (enabled by PyO3's `multiple-pymethods` feature).

use std::sync::LazyLock;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyString, PyTuple};
use rust_xlsxwriter::{Format as RFormat, Worksheet as RWorksheet};

use crate::common::cellref::{split_cell, split_cols, split_range, split_row, Params};
use crate::common::color::Color;
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::Format;
use crate::common::value::{classify, CellValue, Formula, Url};
use crate::py_props;

pub(crate) static DATETIME_FORMAT: LazyLock<RFormat> =
    LazyLock::new(|| RFormat::new().set_num_format("yyyy-mm-dd hh:mm:ss"));
pub(crate) static DATE_FORMAT: LazyLock<RFormat> =
    LazyLock::new(|| RFormat::new().set_num_format("yyyy-mm-dd"));
pub(crate) static TIME_FORMAT: LazyLock<RFormat> =
    LazyLock::new(|| RFormat::new().set_num_format("hh:mm:ss"));
pub(crate) static DURATION_FORMAT: LazyLock<RFormat> =
    LazyLock::new(|| RFormat::new().set_num_format("[h]:mm:ss"));

/// Standard Excel paper sizes (pass the enum or the raw int).
#[pyclass(module = "pyrust_xlsxwriter", eq, eq_int, hash, frozen, from_py_object)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum PaperSize {
    PrinterDefault = 0,
    Letter = 1,
    LetterSmall = 2,
    Tabloid = 3,
    Ledger = 4,
    Legal = 5,
    Statement = 6,
    Executive = 7,
    A3 = 8,
    A4 = 9,
    A4Small = 10,
    A5 = 11,
    B4 = 12,
    B5 = 13,
    Folio = 14,
    Quarto = 15,
    Note = 18,
    Envelope9 = 19,
    Envelope10 = 20,
    Envelope11 = 21,
    Envelope12 = 22,
    Envelope14 = 23,
    EnvelopeDL = 27,
    EnvelopeC3 = 28,
    EnvelopeC4 = 29,
    EnvelopeC5 = 30,
    EnvelopeC6 = 31,
}

impl PaperSize {
    pub fn from_u8(value: u8) -> Option<Self> {
        use PaperSize::*;
        [
            PrinterDefault, Letter, LetterSmall, Tabloid, Ledger, Legal, Statement, Executive, A3,
            A4, A4Small, A5, B4, B5, Folio, Quarto, Note, Envelope9, Envelope10, Envelope11,
            Envelope12, Envelope14, EnvelopeDL, EnvelopeC3, EnvelopeC4, EnvelopeC5, EnvelopeC6,
        ]
        .into_iter()
        .find(|p| *p as u8 == value)
    }
}

/// Settings exposed as worksheet properties.
#[derive(Clone, Debug, PartialEq)]
pub struct WorksheetSpec {
    pub active: bool,
    pub selected: bool,
    pub hidden: bool,
    pub tab_color: Option<Color>,
    pub default_row_height: Option<f64>,
    pub paper_size: Option<u8>,
    pub header: String,
    pub footer: String,
    pub print_center_horizontally: bool,
    pub print_center_vertically: bool,
    pub screen_gridlines: bool,
    pub print_gridlines: bool,
    pub print_headings: bool,
    pub zoom: u16,
    pub right_to_left: bool,
}

impl Default for WorksheetSpec {
    fn default() -> Self {
        WorksheetSpec {
            active: false,
            selected: false,
            hidden: false,
            tab_color: None,
            default_row_height: None,
            paper_size: None,
            header: String::new(),
            footer: String::new(),
            print_center_horizontally: false,
            print_center_vertically: false,
            screen_gridlines: true,
            print_gridlines: false,
            print_headings: false,
            zoom: 100,
            right_to_left: false,
        }
    }
}

/// A worksheet in a new workbook.
///
/// Create one with `Workbook.add_worksheet()` (or `Worksheet()` and
/// `Workbook.push_worksheet()`). Rows and columns are zero-based; every
/// method that takes a cell or range also accepts "A1" / "A1:B2" strings.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct Worksheet {
    pub inner: RWorksheet,
    pub spec: WorksheetSpec,
    /// The settings already pushed to `inner` (only changes are applied, so
    /// untouched settings never add XML and methods such as
    /// `set_very_hidden()` aren't undone by later property changes).
    applied: WorksheetSpec,
    /// Set for constant/low-memory worksheets, which must be written in row order.
    pub row_order: Option<RowOrder>,
}

/// Row-order bookkeeping for constant/low-memory worksheets.
#[derive(Clone, Copy, Debug, Default)]
pub struct RowOrder {
    /// The highest row written so far (earlier rows are already flushed).
    pub current: Option<u32>,
    /// Rows up to and including this one were completed by a save.
    pub saved_through: Option<u32>,
}

impl Worksheet {
    pub fn create(name: Option<&str>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut worksheet = Worksheet {
            inner: RWorksheet::new(),
            spec: WorksheetSpec::default(),
            applied: WorksheetSpec::default(),
            row_order: None,
        };
        if let Some(name) = name {
            worksheet.inner.set_name(name).py()?;
        }
        worksheet.apply_kwargs(kwargs)?;
        Ok(worksheet)
    }

    /// Wrap an existing Rust worksheet (e.g. a chartsheet) with default settings.
    pub fn from_inner(inner: RWorksheet) -> Self {
        Worksheet { inner, spec: WorksheetSpec::default(), applied: WorksheetSpec::default(), row_order: None }
    }

    /// Swap in a different Rust worksheet (e.g. a constant-memory one) and
    /// push all current settings to it.
    pub fn replace_inner(&mut self, inner: RWorksheet, row_ordered: bool) -> PyResult<()> {
        self.inner = inner;
        self.applied = WorksheetSpec::default();
        self.row_order = row_ordered.then(RowOrder::default);
        self.on_change()
    }

    /// Record that `inner` already reflects the current spec (for methods
    /// that change `inner` directly, such as `set_very_hidden()`).
    pub fn sync_applied(&mut self) {
        self.applied = self.spec.clone();
    }

    /// Called after a save: in row-ordered mode the current row is complete.
    pub fn mark_saved(&mut self) {
        if let Some(order) = self.row_order.as_mut() {
            order.saved_through = order.current;
        }
    }

    /// Reject writes that a constant/low-memory worksheet would silently ignore.
    pub fn check_row_order(&mut self, row: u32) -> PyResult<()> {
        let Some(order) = self.row_order.as_mut() else { return Ok(()) };
        if let Some(saved) = order.saved_through
            && row <= saved {
                return Err(param_err(format!(
                    "row {row} can't be written: this constant/low-memory worksheet was saved after row {saved} was written, so continue on a later row"
                )));
            }
        if let Some(current) = order.current
            && row < current {
                return Err(param_err(format!(
                    "row {row} can't be written: this constant/low-memory worksheet must be written in row order and row {current} has already been written"
                )));
            }
        order.current = Some(order.current.map_or(row, |c| c.max(row)));
        Ok(())
    }

    /// Push changed properties to the wrapped worksheet.
    fn on_change(&mut self) -> PyResult<()> {
        if !(10..=400).contains(&self.spec.zoom) {
            return Err(param_err(format!("zoom {} is outside 10..=400", self.spec.zoom)));
        }
        let s = self.spec.clone();
        let a = &self.applied;
        let ws = &mut self.inner;
        if s.active != a.active {
            ws.set_active(s.active);
        }
        if s.selected != a.selected {
            ws.set_selected(s.selected);
        }
        if s.hidden != a.hidden {
            ws.set_hidden(s.hidden);
        }
        if s.tab_color != a.tab_color {
            ws.set_tab_color(s.tab_color.map(|c| c.0).unwrap_or_default());
        }
        if s.default_row_height != a.default_row_height
            && let Some(height) = s.default_row_height {
                ws.set_default_row_height(height);
            }
        if s.paper_size != a.paper_size {
            ws.set_paper_size(s.paper_size.unwrap_or(0));
        }
        if s.header != a.header {
            ws.set_header(&s.header);
        }
        if s.footer != a.footer {
            ws.set_footer(&s.footer);
        }
        if s.print_center_horizontally != a.print_center_horizontally {
            ws.set_print_center_horizontally(s.print_center_horizontally);
        }
        if s.print_center_vertically != a.print_center_vertically {
            ws.set_print_center_vertically(s.print_center_vertically);
        }
        if s.screen_gridlines != a.screen_gridlines {
            ws.set_screen_gridlines(s.screen_gridlines);
        }
        if s.print_gridlines != a.print_gridlines {
            ws.set_print_gridlines(s.print_gridlines);
        }
        if s.print_headings != a.print_headings {
            ws.set_print_headings(s.print_headings);
        }
        if s.zoom != a.zoom {
            ws.set_zoom(s.zoom);
        }
        if s.right_to_left != a.right_to_left {
            ws.set_right_to_left(s.right_to_left);
        }
        self.applied = s;
        Ok(())
    }

    fn get_name(&self) -> String {
        self.inner.name()
    }

    fn set_name(&mut self, name: String) -> PyResult<()> {
        self.inner.set_name(name).py()?;
        Ok(())
    }

    fn get_paper_size(&self) -> Option<Py<PyAny>> {
        Python::attach(|py| {
            self.spec.paper_size.map(|size| match PaperSize::from_u8(size) {
                Some(p) => Py::new(py, p).map(|p| p.into_any()).ok(),
                None => size.into_pyobject(py).ok().map(|v| v.into_any().unbind()),
            })?
        })
    }

    fn set_paper_size(&mut self, value: Option<Py<PyAny>>) -> PyResult<()> {
        self.spec.paper_size = Python::attach(|py| -> PyResult<Option<u8>> {
            Ok(match value {
                None => None,
                Some(v) => {
                    let v = v.bind(py);
                    if v.is_none() {
                        None
                    } else if let Ok(p) = v.extract::<PaperSize>() {
                        Some(p as u8)
                    } else {
                        Some(v.extract::<u8>()?)
                    }
                }
            })
        })?;
        Ok(())
    }

    /// Write any supported Python value to a cell.
    pub fn write_value(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'_, PyAny>,
        format: Option<&RFormat>,
    ) -> PyResult<()> {
        let value = classify(value)?;
        self.write_classified(row, col, value, format)
    }

    pub fn write_classified(
        &mut self,
        row: u32,
        col: u16,
        value: CellValue<'_>,
        format: Option<&RFormat>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let ws = &mut self.inner;
        match value {
            CellValue::Blank => match format {
                Some(f) => {
                    ws.write_blank(row, col, f).py()?;
                }
                None => {
                    ws.clear_cell(row, col);
                }
            },
            CellValue::Bool(b) => {
                match format {
                    Some(f) => ws.write_boolean_with_format(row, col, b, f),
                    None => ws.write_boolean(row, col, b),
                }
                .py()?;
            }
            CellValue::Number(n) => {
                match format {
                    Some(f) => ws.write_number_with_format(row, col, n, f),
                    None => ws.write_number(row, col, n),
                }
                .py()?;
            }
            CellValue::Decimal(d) => {
                match format {
                    Some(f) => ws.write_with_format(row, col, d, f),
                    None => ws.write(row, col, d),
                }
                .py()?;
            }
            CellValue::String(s) => {
                let s = s.to_str()?;
                match format {
                    Some(f) => ws.write_string_with_format(row, col, s, f),
                    None => ws.write_string(row, col, s),
                }
                .py()?;
            }
            CellValue::Formula(text) => {
                match format {
                    Some(f) => ws.write_formula_with_format(row, col, text.as_str(), f),
                    None => ws.write_formula(row, col, text.as_str()),
                }
                .py()?;
            }
            CellValue::ArrayFormula(text) => {
                match format {
                    Some(f) => ws.write_array_formula_with_format(row, col, row, col, text.as_str(), f),
                    None => ws.write_array_formula(row, col, row, col, text.as_str()),
                }
                .py()?;
            }
            CellValue::FormulaObject(formula) => {
                match format {
                    Some(f) => ws.write_formula_with_format(row, col, formula.to_rust(), f),
                    None => ws.write_formula(row, col, formula.to_rust()),
                }
                .py()?;
            }
            CellValue::Url(url) => {
                match format {
                    Some(f) => ws.write_url_with_format(row, col, url.to_rust(), f),
                    None => ws.write_url(row, col, url.to_rust()),
                }
                .py()?;
            }
            CellValue::DateTime(dt) => {
                ws.write_datetime_with_format(row, col, dt, format.unwrap_or(&DATETIME_FORMAT))
                    .py()?;
            }
            CellValue::Date(d) => {
                ws.write_datetime_with_format(row, col, d, format.unwrap_or(&DATE_FORMAT)).py()?;
            }
            CellValue::Time(t) => {
                ws.write_datetime_with_format(row, col, t, format.unwrap_or(&TIME_FORMAT)).py()?;
            }
            CellValue::Duration(days) => {
                ws.write_number_with_format(row, col, days, format.unwrap_or(&DURATION_FORMAT))
                    .py()?;
            }
        }
        Ok(())
    }

    fn write_formula_impl(
        &mut self,
        row: u32,
        col: u16,
        formula: &Bound<'_, PyAny>,
        format: Option<&RFormat>,
        value: Option<&Bound<'_, PyAny>>,
        kind: FormulaKind,
        range: Option<(u32, u16, u32, u16)>,
    ) -> PyResult<()> {
        let mut formula: Formula = match formula.cast::<Formula>() {
            Ok(f) => f.get().clone(),
            Err(_) => Formula { formula: formula.extract::<String>()?, result: None },
        };
        if let Some(value) = value {
            formula.result = Some(value.str()?.to_string());
        }
        let f = formula.to_rust();
        let (r1, c1, r2, c2) = range.unwrap_or((row, col, row, col));
        self.check_row_order(r1)?;
        let ws = &mut self.inner;
        match (kind, format) {
            (FormulaKind::Normal, Some(fmt)) => ws.write_formula_with_format(row, col, f, fmt),
            (FormulaKind::Normal, None) => ws.write_formula(row, col, f),
            (FormulaKind::Array, Some(fmt)) => ws.write_array_formula_with_format(r1, c1, r2, c2, f, fmt),
            (FormulaKind::Array, None) => ws.write_array_formula(r1, c1, r2, c2, f),
            (FormulaKind::DynamicArray, Some(fmt)) => {
                ws.write_dynamic_array_formula_with_format(r1, c1, r2, c2, f, fmt)
            }
            (FormulaKind::DynamicArray, None) => ws.write_dynamic_array_formula(r1, c1, r2, c2, f),
            (FormulaKind::Dynamic, Some(fmt)) => ws.write_dynamic_formula_with_format(row, col, f, fmt),
            (FormulaKind::Dynamic, None) => ws.write_dynamic_formula(row, col, f),
        }
        .py()?;
        Ok(())
    }

    fn write_sequence(
        &mut self,
        row: u32,
        col: u16,
        data: &Bound<'_, PyAny>,
        format: Option<&RFormat>,
        by_row: bool,
    ) -> PyResult<()> {
        for (i, item) in data.try_iter()?.enumerate() {
            let item = item?;
            let (r, c) = offset(row, col, i, by_row)?;
            self.write_value(r, c, &item, format)?;
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum FormulaKind {
    Normal,
    Array,
    DynamicArray,
    Dynamic,
}

fn offset(row: u32, col: u16, i: usize, by_row: bool) -> PyResult<(u32, u16)> {
    let out_of_range = || crate::common::errors::range_err("data extends past Excel's row/column limits");
    if by_row {
        let c = u16::try_from(col as usize + i).map_err(|_| out_of_range())?;
        Ok((row, c))
    } else {
        let r = u32::try_from(row as usize + i).map_err(|_| out_of_range())?;
        Ok((r, col))
    }
}

fn format_ref<'a>(format: &'a Option<PyRef<'_, Format>>) -> Option<&'a RFormat> {
    format.as_ref().map(|f| &f.inner)
}

py_props!(Worksheet
    fields {
        active: bool,
        selected: bool,
        hidden: bool,
        tab_color: Option<Color>,
        default_row_height: Option<f64>,
        header: String,
        footer: String,
        print_center_horizontally: bool,
        print_center_vertically: bool,
        screen_gridlines: bool,
        print_gridlines: bool,
        print_headings: bool,
        zoom: u16,
        right_to_left: bool,
    }
    custom {
        name: String,
        paper_size: Option<Py<PyAny>>,
    }
);

#[pymethods]
impl Worksheet {
    #[new]
    #[pyo3(signature = (name=None, **kwargs))]
    fn py_new(name: Option<&str>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Worksheet::create(name, kwargs)
    }

    fn __repr__(&self) -> String {
        format!("<Worksheet {:?}>", self.inner.name())
    }

    // ------------------------------------------------------------------
    // Writing cells
    // ------------------------------------------------------------------

    /// Write any value: `write(row, col, value, format=None)` or `write("A1", value, format=None)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn write(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write", rest, kwargs, &["value", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_value(row, col, p.required_any(0)?, format_ref(&format))
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_number(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_number", rest, kwargs, &["number", "format"], 1)?;
        let number: f64 = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_classified(row, col, CellValue::Number(number), format_ref(&format))
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_string(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_string", rest, kwargs, &["string", "format"], 1)?;
        let string = p.required_any(0)?.cast::<PyString>()?.clone();
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_classified(row, col, CellValue::String(string), format_ref(&format))
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_boolean(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_boolean", rest, kwargs, &["value", "format"], 1)?;
        let value: bool = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_classified(row, col, CellValue::Bool(value), format_ref(&format))
    }

    /// Write a blank cell. Without a format this clears the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_blank(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_blank", rest, kwargs, &["format"], 0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(0)?;
        self.write_classified(row, col, CellValue::Blank, format_ref(&format))
    }

    /// Write a `datetime`, `date`, `time` or `timedelta`. Without a format a
    /// default date/time format is applied so the cell shows a date.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_datetime(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_datetime", rest, kwargs, &["datetime", "format"], 1)?;
        let value = classify(p.required_any(0)?)?;
        if !matches!(
            value,
            CellValue::DateTime(_) | CellValue::Date(_) | CellValue::Time(_) | CellValue::Duration(_)
        ) {
            return Err(pyo3::exceptions::PyTypeError::new_err(
                "write_datetime() expects a datetime, date, time or timedelta",
            ));
        }
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_classified(row, col, value, format_ref(&format))
    }

    /// Write a formula, optionally with a cached result `value`.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_formula(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_formula_impl(row, col, p.required_any(0)?, format_ref(&format), p.raw(2), FormulaKind::Normal, None)
    }

    /// Write a legacy (Ctrl-Shift-Enter) array formula over a range.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_array_formula(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let p = Params::new("write_array_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_formula_impl(range.0, range.1, p.required_any(0)?, format_ref(&format), p.raw(2), FormulaKind::Array, Some(range))
    }

    /// Write a dynamic array formula (spills in Excel 365) over a range.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_dynamic_array_formula(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let p = Params::new("write_dynamic_array_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_formula_impl(range.0, range.1, p.required_any(0)?, format_ref(&format), p.raw(2), FormulaKind::DynamicArray, Some(range))
    }

    /// Write a single-cell dynamic formula (for functions such as `UNIQUE` or `FILTER`).
    #[pyo3(signature = (*args, **kwargs))]
    fn write_dynamic_formula(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_dynamic_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_formula_impl(row, col, p.required_any(0)?, format_ref(&format), p.raw(2), FormulaKind::Dynamic, None)
    }

    /// Write a hyperlink: `write_url("A1", "https://...", text=None, tip=None, format=None)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_url(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_url", rest, kwargs, &["url", "text", "tip", "format"], 1)?;
        let url_arg = p.required_any(0)?;
        let mut url: Url = match url_arg.cast::<Url>() {
            Ok(u) => u.get().clone(),
            Err(_) => Url { link: url_arg.extract()?, text: None, tip: None },
        };
        if let Some(text) = p.extract::<String>(1)? {
            url.text = Some(text);
        }
        if let Some(tip) = p.extract::<String>(2)? {
            url.tip = Some(tip);
        }
        let format: Option<PyRef<'_, Format>> = p.extract(3)?;
        self.write_classified(row, col, CellValue::Url(url), format_ref(&format))
    }

    /// Write a string made of differently formatted segments:
    /// `write_rich_string("A1", ["plain ", (bold, "bold"), " plain"], format=None)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_rich_string(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_rich_string", rest, kwargs, &["segments", "format"], 1)?;
        let default = RFormat::new();
        let mut owned: Vec<(Option<RFormat>, String)> = Vec::new();
        for item in p.required_any(0)?.try_iter()? {
            let item = item?;
            if let Ok(text) = item.cast::<PyString>() {
                owned.push((None, text.to_str()?.to_string()));
            } else {
                let (fmt, text): (PyRef<'_, Format>, String) = item.extract().map_err(|_| {
                    pyo3::exceptions::PyTypeError::new_err(
                        "rich string segments must be str or (Format, str) pairs",
                    )
                })?;
                owned.push((Some(fmt.inner.clone()), text));
            }
        }
        if owned.is_empty() {
            return Err(param_err("write_rich_string() needs at least one segment"));
        }
        self.check_row_order(row)?;
        let segments: Vec<(&RFormat, &str)> = owned
            .iter()
            .map(|(f, t)| (f.as_ref().unwrap_or(&default), t.as_str()))
            .collect();
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        match format_ref(&format) {
            Some(f) => self.inner.write_rich_string_with_format(row, col, &segments, f),
            None => self.inner.write_rich_string(row, col, &segments),
        }
        .py()?;
        Ok(())
    }

    /// Write an iterable across a row starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_row(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_row", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_sequence(row, col, p.required_any(0)?, format_ref(&format), true)
    }

    /// Write an iterable down a column starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_column(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_column", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_sequence(row, col, p.required_any(0)?, format_ref(&format), false)
    }

    /// Write a 2-D iterable (an iterable of rows) starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_rows(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_rows", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        for (i, line) in p.required_any(0)?.try_iter()?.enumerate() {
            let (r, _) = offset(row, col, i, false)?;
            self.write_sequence(r, col, &line?, format_ref(&format), true)?;
        }
        Ok(())
    }

    /// Write a 2-D iterable of columns starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_columns(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_columns", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        for (i, column) in p.required_any(0)?.try_iter()?.enumerate() {
            let (_, c) = offset(row, col, i, true)?;
            self.write_sequence(row, c, &column?, format_ref(&format), false)?;
        }
        Ok(())
    }

    /// Merge a range and write a value into it: `merge_range("A1:D1", "Title", format)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn merge_range(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("merge_range", rest, kwargs, &["value", "format"], 0)?;
        self.check_row_order(r1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let default = RFormat::new();
        let fmt = format_ref(&format).unwrap_or(&default);
        match p.raw(0) {
            Some(value) => match value.cast::<PyString>() {
                Ok(s) if !s.to_str()?.starts_with('=') => {
                    self.inner.merge_range(r1, c1, r2, c2, s.to_str()?, fmt).py()?;
                }
                _ => {
                    // Non-string values: merge, then overwrite the first cell.
                    self.inner.merge_range(r1, c1, r2, c2, "", fmt).py()?;
                    self.write_value(r1, c1, value, Some(fmt))?;
                }
            },
            None => {
                self.inner.merge_range(r1, c1, r2, c2, "", fmt).py()?;
            }
        }
        Ok(())
    }

    /// Set the cached result shown for a formula cell by non-Excel readers.
    #[pyo3(signature = (*args))]
    fn set_formula_result(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("set_formula_result", rest, None, &["result"], 1)?;
        let result = p.required_any(0)?.str()?.to_string();
        self.inner.set_formula_result(row, col, result);
        Ok(())
    }

    /// Set the default cached result for formulas (Excel shows 0 otherwise in some viewers).
    fn set_formula_result_default(&mut self, result: String) {
        self.inner.set_formula_result_default(result);
    }

    /// Apply a format to a cell, keeping its value.
    #[pyo3(signature = (*args))]
    fn set_cell_format(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("set_cell_format", rest, None, &["format"], 1)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        self.inner.set_cell_format(row, col, &format.inner).py()?;
        Ok(())
    }

    /// Apply a format to every cell in a range, keeping values.
    #[pyo3(signature = (*args))]
    fn set_range_format(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("set_range_format", rest, None, &["format"], 1)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        self.inner.set_range_format(r1, c1, r2, c2, &format.inner).py()?;
        Ok(())
    }

    /// Apply a format to a range plus a border format around its outside.
    #[pyo3(signature = (*args))]
    fn set_range_format_with_border(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("set_range_format_with_border", rest, None, &["format", "border_format"], 2)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        let border: PyRef<'_, Format> = p.required(1)?;
        self.inner
            .set_range_format_with_border(r1, c1, r2, c2, &format.inner, &border.inner)
            .py()?;
        Ok(())
    }

    /// Remove a cell's value and format.
    #[pyo3(signature = (*args))]
    fn clear_cell(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.inner.clear_cell(row, col);
        Ok(())
    }

    /// Remove a cell's format, keeping its value.
    #[pyo3(signature = (*args))]
    fn clear_cell_format(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.inner.clear_cell_format(row, col);
        Ok(())
    }

    /// Text written instead of NaN (Excel has no NaN). Default "#NUM!".
    fn set_nan_value(&mut self, value: String) {
        self.inner.set_nan_value(value);
    }

    /// Text written instead of +inf. Default "#DIV/0!".
    fn set_infinity_value(&mut self, value: String) {
        self.inner.set_infinity_value(value);
    }

    /// Text written instead of -inf. Default "#DIV/0!".
    fn set_neg_infinity_value(&mut self, value: String) {
        self.inner.set_neg_infinity_value(value);
    }

    // ------------------------------------------------------------------
    // Columns
    // ------------------------------------------------------------------

    /// Set column width (in character units) for a column or span: `set_column_width("B:D", 20)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn set_column_width(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        let p = Params::new("set_column_width", rest, kwargs, &["width", "format"], 1)?;
        let width: f64 = p.required(0)?;
        self.inner.set_column_range_width(first, last, width).py()?;
        if let Some(format) = p.extract::<PyRef<'_, Format>>(1)? {
            self.inner.set_column_range_format(first, last, &format.inner).py()?;
        }
        Ok(())
    }

    /// Set the width of columns `first..=last` (kept for backwards compatibility).
    #[pyo3(signature = (first_col, last_col, width, format=None))]
    fn set_column_range_width(
        &mut self,
        first_col: u16,
        last_col: u16,
        width: f64,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        self.inner.set_column_range_width(first_col, last_col, width).py()?;
        if let Some(format) = format {
            self.inner.set_column_range_format(first_col, last_col, &format.inner).py()?;
        }
        Ok(())
    }

    /// Set column width in pixels.
    #[pyo3(signature = (*args, **kwargs))]
    fn set_column_width_pixels(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        let p = Params::new("set_column_width_pixels", rest, kwargs, &["width", "format"], 1)?;
        let width: u32 = p.required(0)?;
        self.inner.set_column_range_width_pixels(first, last, width).py()?;
        if let Some(format) = p.extract::<PyRef<'_, Format>>(1)? {
            self.inner.set_column_range_format(first, last, &format.inner).py()?;
        }
        Ok(())
    }

    /// Set the default format of a column or span.
    #[pyo3(signature = (*args))]
    fn set_column_format(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        let p = Params::new("set_column_format", rest, None, &["format"], 1)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        self.inner.set_column_range_format(first, last, &format.inner).py()?;
        Ok(())
    }

    /// Hide a column or span.
    #[pyo3(signature = (*args))]
    fn set_column_hidden(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((first, last), _) = split_cols(args.as_slice())?;
        self.inner.set_column_range_hidden(first, last).py()?;
        Ok(())
    }

    /// Set a column's width in pixels as if it had been autofitted.
    #[pyo3(signature = (*args))]
    fn set_column_autofit_width(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        let p = Params::new("set_column_autofit_width", rest, None, &["width"], 1)?;
        let width: u32 = p.required(0)?;
        for col in first..=last {
            self.inner.set_column_autofit_width(col, width).py()?;
        }
        Ok(())
    }

    /// Group columns into an outline: `group_columns("B:D", collapsed=False)`.
    #[pyo3(signature = (*args, collapsed=false))]
    fn group_columns(&mut self, args: &Bound<'_, PyTuple>, collapsed: bool) -> PyResult<()> {
        let (first, last) = cols_or_pair(args)?;
        if collapsed {
            self.inner.group_columns_collapsed(first, last).py()?;
        } else {
            self.inner.group_columns(first, last).py()?;
        }
        Ok(())
    }

    /// Show outline expand/collapse buttons to the left of groups (default: right).
    #[pyo3(signature = (enable=true))]
    fn group_symbols_to_left(&mut self, enable: bool) {
        self.inner.group_symbols_to_left(enable);
    }

    // ------------------------------------------------------------------
    // Rows
    // ------------------------------------------------------------------

    /// Set row height in points, optionally with a row format.
    #[pyo3(signature = (*args, **kwargs))]
    fn set_row_height(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (row, rest) = split_row(args.as_slice())?;
        let p = Params::new("set_row_height", rest, kwargs, &["height", "format"], 1)?;
        let height: f64 = p.required(0)?;
        self.inner.set_row_height(row, height).py()?;
        if let Some(format) = p.extract::<PyRef<'_, Format>>(1)? {
            self.inner.set_row_format(row, &format.inner).py()?;
        }
        Ok(())
    }

    /// Set row height in pixels, optionally with a row format.
    #[pyo3(signature = (*args, **kwargs))]
    fn set_row_height_pixels(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (row, rest) = split_row(args.as_slice())?;
        let p = Params::new("set_row_height_pixels", rest, kwargs, &["height", "format"], 1)?;
        let height: u32 = p.required(0)?;
        self.inner.set_row_height_pixels(row, height).py()?;
        if let Some(format) = p.extract::<PyRef<'_, Format>>(1)? {
            self.inner.set_row_format(row, &format.inner).py()?;
        }
        Ok(())
    }

    /// Set the default format of a row.
    fn set_row_format(&mut self, row: u32, format: PyRef<'_, Format>) -> PyResult<()> {
        self.inner.set_row_format(row, &format.inner).py()?;
        Ok(())
    }

    /// Hide a row.
    fn set_row_hidden(&mut self, row: u32) -> PyResult<()> {
        self.inner.set_row_hidden(row).py()?;
        Ok(())
    }

    /// Unhide a row hidden with `set_row_hidden()`.
    fn set_row_unhidden(&mut self, row: u32) -> PyResult<()> {
        self.inner.set_row_unhidden(row).py()?;
        Ok(())
    }

    /// Set the default row height in pixels.
    fn set_default_row_height_pixels(&mut self, height: u32) {
        self.inner.set_default_row_height_pixels(height);
    }

    /// Hide all rows without data (useful with a small used area).
    #[pyo3(signature = (enable=true))]
    fn hide_unused_rows(&mut self, enable: bool) {
        self.inner.hide_unused_rows(enable);
    }

    /// Group rows `first..=last` into an outline.
    #[pyo3(signature = (first_row, last_row, collapsed=false))]
    fn group_rows(&mut self, first_row: u32, last_row: u32, collapsed: bool) -> PyResult<()> {
        if collapsed {
            self.inner.group_rows_collapsed(first_row, last_row).py()?;
        } else {
            self.inner.group_rows(first_row, last_row).py()?;
        }
        Ok(())
    }

    /// Show outline expand/collapse buttons above groups (default: below).
    #[pyo3(signature = (enable=true))]
    fn group_symbols_above(&mut self, enable: bool) {
        self.inner.group_symbols_above(enable);
    }

    // ------------------------------------------------------------------
    // Autofit, panes, misc
    // ------------------------------------------------------------------

    /// Adjust column widths to fit their contents (an approximation, as in
    /// XlsxWriter). `max_width` caps the width in pixels.
    #[pyo3(signature = (max_width=None))]
    fn autofit(&mut self, max_width: Option<u32>) {
        if let Some(width) = max_width {
            self.inner.set_autofit_max_width(width);
        }
        self.inner.autofit();
    }

    /// Only consider rows up to `max_row` when autofitting (faster on large sheets).
    fn set_autofit_max_row(&mut self, max_row: u32) {
        self.inner.set_autofit_max_row(max_row);
    }

    /// Freeze the rows above and columns left of a cell: `freeze_panes(1, 0)` or `freeze_panes("A2")`.
    #[pyo3(signature = (*args))]
    fn freeze_panes(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.inner.set_freeze_panes(row, col).py()?;
        Ok(())
    }

    /// Set the worksheet's VBA code name (for macros referring to it).
    fn set_vba_name(&mut self, name: &str) -> PyResult<()> {
        self.inner.set_vba_name(name).py()?;
        Ok(())
    }

    fn set_landscape(&mut self) {
        self.inner.set_landscape();
    }

    fn set_portrait(&mut self) {
        self.inner.set_portrait();
    }

    /// Fit the printed area to `width` x `height` pages (0 = unlimited).
    fn set_fit_to_pages(&mut self, width: u16, height: u16) {
        self.inner.set_print_fit_to_pages(width, height);
    }

    // Area markers: feature modules add more `#[pymethods] impl Worksheet`
    // blocks in their own files.
}

/// Columns as `"B:D"`, `"B"`, or `first, last` ints.
fn cols_or_pair(args: &Bound<'_, PyTuple>) -> PyResult<(u16, u16)> {
    let items = args.as_slice();
    if let Some(first) = items.first()
        && first.cast::<PyString>().is_ok() {
            return Ok(split_cols(items)?.0);
        }
    if items.len() != 2 {
        return Err(pyo3::exceptions::PyTypeError::new_err(
            "expected a column span like \"B:D\" or first_col, last_col",
        ));
    }
    Ok((items[0].extract()?, items[1].extract()?))
}

/// Helper for feature modules: extract a list of Python objects.
#[allow(dead_code)]
pub(crate) fn as_list<'py>(value: &Bound<'py, PyAny>) -> PyResult<Bound<'py, PyList>> {
    let items: Vec<Bound<'py, PyAny>> = value.try_iter()?.collect::<PyResult<_>>()?;
    PyList::new(value.py(), items)
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Worksheet>()?;
    m.add_class::<PaperSize>()?;
    Ok(())
}
