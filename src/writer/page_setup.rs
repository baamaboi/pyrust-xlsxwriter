//! Worksheet page setup, views, printing and protection.
//!
//! * views: selection, top-left cell, page-layout / page-break views, zoom to fit;
//! * printing: margins, print area, repeated rows/columns, page breaks, scale,
//!   page order, black-and-white / draft, header/footer scaling;
//! * protection: `protect()` with an optional password and `ProtectionOptions`,
//!   unprotected ranges;
//! * misc: very hidden, first tab, hidden zeroes, ignored errors.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use rust_xlsxwriter::{IgnoreError as RIgnoreError, ProtectionOptions as RProtectionOptions};

use crate::common::cellref::{
    parse_range, split_cell, split_cols, split_range, CellRange, Params, MAX_COLS, MAX_ROWS,
};
use crate::common::errors::{param_err, range_err, IntoPyResult};
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// Excel's cell error/warning types (the green triangle) that
    /// `Worksheet.ignore_error()` can turn off.
    IgnoreError => RIgnoreError {
        NumberStoredAsText,
        FormulaError,
        FormulaDiffers,
        FormulaRefersToEmptyCells,
        FormulaOmitsCells,
        DataValidationError,
        TwoDigitTextYear,
        UnlockedCellsWithFormula,
        InconsistentColumnFormula,
    }
);

// ---------------------------------------------------------------------------
// ProtectionOptions
// ---------------------------------------------------------------------------

/// Plain description of the worksheet protection options. `true` means the
/// action is still allowed while the worksheet is protected. The defaults
/// match Excel's.
#[derive(Clone, Debug, PartialEq)]
pub struct ProtectionSpec {
    pub select_locked_cells: bool,
    pub select_unlocked_cells: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_links: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub sort: bool,
    pub use_autofilter: bool,
    pub use_pivot_tables: bool,
    pub edit_scenarios: bool,
    pub edit_objects: bool,
    pub contents: bool,
}

impl Default for ProtectionSpec {
    fn default() -> Self {
        let d = RProtectionOptions::new();
        ProtectionSpec {
            select_locked_cells: d.select_locked_cells,
            select_unlocked_cells: d.select_unlocked_cells,
            format_cells: d.format_cells,
            format_columns: d.format_columns,
            format_rows: d.format_rows,
            insert_columns: d.insert_columns,
            insert_rows: d.insert_rows,
            insert_links: d.insert_links,
            delete_columns: d.delete_columns,
            delete_rows: d.delete_rows,
            sort: d.sort,
            use_autofilter: d.use_autofilter,
            use_pivot_tables: d.use_pivot_tables,
            edit_scenarios: d.edit_scenarios,
            edit_objects: d.edit_objects,
            contents: d.contents,
        }
    }
}

macro_rules! protection_fields {
    ($m:ident) => {
        $m!(
            select_locked_cells, select_unlocked_cells, format_cells, format_columns, format_rows,
            insert_columns, insert_rows, insert_links, delete_columns, delete_rows, sort,
            use_autofilter, use_pivot_tables, edit_scenarios, edit_objects, contents
        )
    };
}

impl ProtectionSpec {
    /// Build the equivalent `rust_xlsxwriter::ProtectionOptions`.
    pub fn to_rust(&self) -> RProtectionOptions {
        let mut options = RProtectionOptions::new();
        macro_rules! copy {
            ($($field:ident),*) => { $( options.$field = self.$field; )* };
        }
        protection_fields!(copy);
        options
    }

    /// `(name, value)` for every option, in declaration order.
    fn items(&self) -> Vec<(&'static str, bool)> {
        macro_rules! list {
            ($($field:ident),*) => { vec![ $( (stringify!($field), self.$field) ),* ] };
        }
        protection_fields!(list)
    }
}

/// What users may still do on a protected worksheet. Every option is a
/// property and a constructor keyword; `True` means "allowed":
///
/// ```python
/// options = ProtectionOptions(format_cells=True, insert_rows=True)
/// ws.protect(password="secret", options=options)
/// ```
///
/// Defaults (as in Excel): `select_locked_cells`, `select_unlocked_cells` and
/// `contents` are True, everything else is False.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct ProtectionOptions {
    pub spec: ProtectionSpec,
}

impl ProtectionOptions {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }
}

py_props!(ProtectionOptions
    fields {
        select_locked_cells: bool,
        select_unlocked_cells: bool,
        format_cells: bool,
        format_columns: bool,
        format_rows: bool,
        insert_columns: bool,
        insert_rows: bool,
        insert_links: bool,
        delete_columns: bool,
        delete_rows: bool,
        sort: bool,
        use_autofilter: bool,
        use_pivot_tables: bool,
        edit_scenarios: bool,
        edit_objects: bool,
        contents: bool,
    }
);

#[pymethods]
impl ProtectionOptions {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut options = ProtectionOptions { spec: ProtectionSpec::default() };
        options.apply_kwargs(kwargs)?;
        Ok(options)
    }

    /// Set several options at once and return self (for chaining).
    #[pyo3(signature = (**kwargs))]
    fn set<'py>(
        mut slf: PyRefMut<'py, Self>,
        kwargs: Option<&Bound<'py, PyDict>>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.apply_kwargs(kwargs)?;
        Ok(slf)
    }

    /// Every option as a dict.
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        for (name, value) in self.spec.items() {
            dict.set_item(name, value)?;
        }
        Ok(dict)
    }

    /// A copy of these options.
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
        match other.cast::<ProtectionOptions>() {
            Ok(other) => self.spec == other.borrow().spec,
            Err(_) => false,
        }
    }

    fn __repr__(&self) -> String {
        let defaults = ProtectionSpec::default().items();
        let parts: Vec<String> = self
            .spec
            .items()
            .into_iter()
            .zip(defaults)
            .filter(|((_, value), (_, default))| value != default)
            .map(|((name, value), _)| format!("{name}={}", if value { "True" } else { "False" }))
            .collect();
        format!("ProtectionOptions({})", parts.join(", "))
    }
}

// ---------------------------------------------------------------------------
// Argument helpers
// ---------------------------------------------------------------------------

/// A cell or range in either order: `"B2"`, `"A1:C3"`, `"C3:A1"`, `(row, col)`
/// or `(first_row, first_col, last_row, last_col)`. Used where Excel accepts a
/// reversed range (a selection keeps its first cell as the active cell).
fn split_any_range<'a, 'py>(
    args: &'a [Bound<'py, PyAny>],
) -> PyResult<(CellRange, &'a [Bound<'py, PyAny>])> {
    let first = args.first().ok_or_else(|| {
        PyTypeError::new_err("missing range: pass an \"A1:B2\" string or row/col ints")
    })?;
    if let Ok(text) = first.cast::<PyString>() {
        let text = text.to_str()?;
        let range = parse_range(text)
            .ok_or_else(|| param_err(format!("invalid range reference {text:?}")))?;
        return Ok((range, &args[1..]));
    }
    let ((r1, c1), rest) = split_cell(args)?;
    if rest.is_empty() {
        return Ok(((r1, c1, r1, c1), rest));
    }
    let ((r2, c2), rest) = split_cell(rest)?;
    Ok(((r1, c1, r2, c2), rest))
}

/// Reject positional arguments left over after a cell/range.
fn no_more_args(method: &'static str, rest: &[Bound<'_, PyAny>]) -> PyResult<()> {
    Params::new(method, rest, None, &[], 0).map(|_| ())
}

fn checked_row(value: i64) -> PyResult<u32> {
    if (0..MAX_ROWS as i64).contains(&value) {
        Ok(value as u32)
    } else {
        Err(range_err(format!("row {value} is outside 0..{MAX_ROWS}")))
    }
}

/// Validate and sort page breaks. Break 0 (before the first row/column) is
/// meaningless and dropped, duplicates are removed.
fn page_breaks(breaks: Vec<i64>, limit: u32, what: &str) -> PyResult<Vec<u32>> {
    let mut out = Vec::with_capacity(breaks.len());
    for value in breaks {
        if !(0..limit as i64).contains(&value) {
            return Err(range_err(format!("page break {what} {value} is outside 0..{limit}")));
        }
        if value > 0 {
            out.push(value as u32);
        }
    }
    out.sort_unstable();
    out.dedup();
    if out.len() > 1023 {
        return Err(param_err(format!(
            "Excel allows at most 1023 {what} page breaks, got {}",
            out.len()
        )));
    }
    Ok(out)
}

fn check_margin(name: &str, value: f64) -> PyResult<f64> {
    if value.is_finite() && value >= 0.0 {
        Ok(value)
    } else {
        Err(param_err(format!("margin {name}={value} must be a finite number >= 0 (inches)")))
    }
}

// ---------------------------------------------------------------------------
// Worksheet methods
// ---------------------------------------------------------------------------

#[pymethods]
impl Worksheet {
    // ---- views -----------------------------------------------------------

    /// Select a cell or range when the file opens: `set_selection("B3")`,
    /// `set_selection("B3:D6")` or `set_selection(2, 1, 5, 3)`. The first
    /// cell is the active cell, so a reversed range such as `"D6:B3"` makes
    /// D6 active.
    #[pyo3(signature = (*args))]
    fn set_selection(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_any_range(args.as_slice())?;
        no_more_args("set_selection", rest)?;
        self.inner.set_selection(r1, c1, r2, c2).py()?;
        Ok(())
    }

    /// Scroll the view so that this cell is the top-left visible cell:
    /// `set_top_left_cell("E20")` or `set_top_left_cell(19, 4)`.
    #[pyo3(signature = (*args))]
    fn set_top_left_cell(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        no_more_args("set_top_left_cell", rest)?;
        self.inner.set_top_left_cell(row, col).py()?;
        Ok(())
    }

    /// The top-left cell of the scrolling pane below/right of frozen panes
    /// (call after `freeze_panes()`): `set_freeze_panes_top_cell("A20")`.
    #[pyo3(signature = (*args))]
    fn set_freeze_panes_top_cell(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        no_more_args("set_freeze_panes_top_cell", rest)?;
        self.inner.set_freeze_panes_top_cell(row, col).py()?;
        Ok(())
    }

    /// Show the worksheet in Excel's "Normal" view (the default).
    fn set_view_normal(&mut self) {
        self.inner.set_view_normal();
    }

    /// Show the worksheet in Excel's "Page Layout" view.
    fn set_view_page_layout(&mut self) {
        self.inner.set_view_page_layout();
    }

    /// Show the worksheet in Excel's "Page Break Preview" view.
    fn set_view_page_break_preview(&mut self) {
        self.inner.set_view_page_break_preview();
    }

    /// Zoom a chartsheet to fit the window (used by chartsheets only).
    #[pyo3(signature = (enable=true))]
    fn set_zoom_to_fit(&mut self, enable: bool) {
        self.inner.set_zoom_to_fit(enable);
    }

    /// Hide the worksheet so that it can only be unhidden with VBA (not from
    /// Excel's "Unhide" dialog). `hidden` reads True afterwards.
    ///
    /// Note: assigning a worksheet property afterwards (e.g. `ws.zoom = 120`)
    /// re-applies the plain `hidden` state, so call this last.
    #[pyo3(signature = (enable=true))]
    fn set_very_hidden(&mut self, enable: bool) {
        self.spec.hidden = enable;
        if enable {
            self.spec.active = false;
            self.spec.selected = false;
        }
        self.inner.set_very_hidden(enable);
        self.sync_applied();
    }

    /// Make this the first visible tab in the tab bar (when there are many
    /// worksheets and the first ones are scrolled out of view).
    #[pyo3(signature = (enable=true))]
    fn set_first_tab(&mut self, enable: bool) {
        self.inner.set_first_tab(enable);
    }

    /// Display zero values as blank cells (Excel's "Show a zero in cells that
    /// have zero value" option turned off).
    #[pyo3(signature = (enable=true))]
    fn set_hidden_zeroes(&mut self, enable: bool) {
        self.inner.set_hidden_zeroes(enable);
    }

    /// Turn off one kind of Excel warning (green triangle) for a cell:
    /// `ignore_error("A1", IgnoreError.NumberStoredAsText)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn ignore_error(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("ignore_error", rest, kwargs, &["error_type"], 1)?;
        let error_type: IgnoreError = p.required(0)?;
        self.inner.ignore_error(row, col, error_type.into()).py()?;
        Ok(())
    }

    /// Turn off one kind of Excel warning for a range:
    /// `ignore_error_range("A1:C10", IgnoreError.FormulaDiffers)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn ignore_error_range(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("ignore_error_range", rest, kwargs, &["error_type"], 1)?;
        let error_type: IgnoreError = p.required(0)?;
        self.inner.ignore_error_range(r1, c1, r2, c2, error_type.into()).py()?;
        Ok(())
    }

    // ---- printing ----------------------------------------------------------

    /// Set horizontal page breaks: a page break is inserted *before* each
    /// zero-based row, e.g. `set_page_breaks([20, 40])`. Replaces earlier
    /// breaks; row 0 is ignored, as is an empty list.
    fn set_page_breaks(&mut self, rows: Vec<i64>) -> PyResult<()> {
        let breaks = page_breaks(rows, MAX_ROWS, "row")?;
        if !breaks.is_empty() {
            self.inner.set_page_breaks(&breaks).py()?;
        }
        Ok(())
    }

    /// Set vertical page breaks: a page break is inserted *before* each
    /// zero-based column. Replaces earlier breaks; column 0 is ignored, as is
    /// an empty list.
    fn set_vertical_page_breaks(&mut self, cols: Vec<i64>) -> PyResult<()> {
        let breaks = page_breaks(cols, MAX_COLS as u32, "column")?;
        if !breaks.is_empty() {
            self.inner.set_vertical_page_breaks(&breaks).py()?;
        }
        Ok(())
    }

    /// Set the printed page margins, in inches (Excel's defaults are shown).
    #[pyo3(signature = (left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3))]
    fn set_margins(
        &mut self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> PyResult<()> {
        self.inner.set_margins(
            check_margin("left", left)?,
            check_margin("right", right)?,
            check_margin("top", top)?,
            check_margin("bottom", bottom)?,
            check_margin("header", header)?,
            check_margin("footer", footer)?,
        );
        Ok(())
    }

    /// Print only this range: `set_print_area("A1:H20")` or
    /// `set_print_area(0, 0, 19, 7)`. Stored as the `_xlnm.Print_Area`
    /// defined name.
    #[pyo3(signature = (*args))]
    fn set_print_area(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        no_more_args("set_print_area", rest)?;
        self.inner.set_print_area(r1, c1, r2, c2).py()?;
        Ok(())
    }

    /// Repeat zero-based rows `first_row..=last_row` at the top of every
    /// printed page (`last_row` defaults to `first_row`).
    #[pyo3(signature = (first_row, last_row=None))]
    fn set_repeat_rows(&mut self, first_row: i64, last_row: Option<i64>) -> PyResult<()> {
        let first = checked_row(first_row)?;
        let last = checked_row(last_row.unwrap_or(first_row))?;
        self.inner.set_repeat_rows(first, last).py()?;
        Ok(())
    }

    /// Repeat columns at the left of every printed page:
    /// `set_repeat_columns("A:B")`, `set_repeat_columns("A")`,
    /// `set_repeat_columns(0)` or `set_repeat_columns(0, 1)`.
    #[pyo3(signature = (*args))]
    fn set_repeat_columns(&mut self, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let items = args.as_slice();
        let ((first, mut last), mut rest) = split_cols(items)?;
        let first_is_text = items.first().is_some_and(|v| v.cast::<PyString>().is_ok());
        if !first_is_text && !rest.is_empty() {
            let ((col, _), more) = split_cols(rest)?;
            last = col;
            rest = more;
        }
        no_more_args("set_repeat_columns", rest)?;
        self.inner.set_repeat_columns(first, last).py()?;
        Ok(())
    }

    /// Print at `scale` percent (10..=400). Overrides `set_fit_to_pages()`.
    fn set_print_scale(&mut self, scale: u16) -> PyResult<()> {
        if !(10..=400).contains(&scale) {
            return Err(param_err(format!("print scale {scale} is outside 10..=400")));
        }
        self.inner.set_print_scale(scale);
        Ok(())
    }

    /// Number the first printed page `page_number`.
    ///
    /// Only 1 is accepted for now: rust_xlsxwriter 0.103 writes other values
    /// as `useFirstPageNumber="N"`, which is not valid OOXML (Excel expects
    /// `firstPageNumber="N" useFirstPageNumber="1"`), so a ParameterError is
    /// raised instead of producing a file Excel would repair.
    fn set_print_first_page_number(&mut self, page_number: u16) -> PyResult<()> {
        if page_number != 1 {
            return Err(param_err(format!(
                "set_print_first_page_number({page_number}) is not supported yet: \
                 rust_xlsxwriter 0.103 writes page numbers other than 1 as invalid XML"
            )));
        }
        self.inner.set_print_first_page_number(page_number);
        Ok(())
    }

    /// Printed page order: `True` (the default) prints down, then over;
    /// `False` prints over, then down.
    fn set_page_order(&mut self, down_then_over: bool) {
        self.inner.set_page_order(down_then_over);
    }

    /// Print in black and white.
    #[pyo3(signature = (enable=true))]
    fn set_print_black_and_white(&mut self, enable: bool) {
        self.inner.set_print_black_and_white(enable);
    }

    /// Print in draft quality.
    #[pyo3(signature = (enable=true))]
    fn set_print_draft(&mut self, enable: bool) {
        self.inner.set_print_draft(enable);
    }

    /// Scale the header and footer with the document (Excel's default: True).
    fn set_header_footer_scale_with_doc(&mut self, enable: bool) {
        self.inner.set_header_footer_scale_with_doc(enable);
    }

    /// Align the header and footer with the page margins (Excel's default: True).
    fn set_header_footer_align_with_page(&mut self, enable: bool) {
        self.inner.set_header_footer_align_with_page(enable);
    }

    // ---- protection --------------------------------------------------------

    /// Protect the worksheet so locked cells can't be edited. Optionally with
    /// a `password` (Excel's weak legacy hash: it deters, it doesn't secure)
    /// and `ProtectionOptions` listing what users may still do.
    #[pyo3(signature = (password=None, options=None))]
    fn protect(
        &mut self,
        password: Option<&str>,
        options: Option<PyRef<'_, ProtectionOptions>>,
    ) {
        match password {
            Some(password) => self.inner.protect_with_password(password),
            None => self.inner.protect(),
        };
        if let Some(options) = options {
            self.inner.protect_with_options(&options.spec.to_rust());
        }
    }

    /// Protect the worksheet with a password (same as `protect(password=...)`).
    fn protect_with_password(&mut self, password: &str) {
        self.inner.protect_with_password(password);
    }

    /// Protect the worksheet with options (same as `protect(options=...)`).
    fn protect_with_options(&mut self, options: PyRef<'_, ProtectionOptions>) {
        self.inner.protect_with_options(&options.spec.to_rust());
    }

    /// Let users edit a range of a protected worksheet:
    /// `unprotect_range("A1:B5", title="Inputs", password=None)`. Without a
    /// title Excel's "Range1", "Range2", ... is used.
    #[pyo3(signature = (*args, **kwargs))]
    fn unprotect_range(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("unprotect_range", rest, kwargs, &["title", "password"], 0)?;
        let title: String = p.get_or(0, String::new())?;
        let password: String = p.get_or(1, String::new())?;
        self.inner.unprotect_range_with_options(r1, c1, r2, c2, &title, &password).py()?;
        Ok(())
    }

    /// Same as `unprotect_range(range, title, password)` with both required.
    #[pyo3(signature = (*args, **kwargs))]
    fn unprotect_range_with_options(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("unprotect_range_with_options", rest, kwargs, &["title", "password"], 2)?;
        let title: String = p.required(0)?;
        let password: String = p.required(1)?;
        self.inner.unprotect_range_with_options(r1, c1, r2, c2, &title, &password).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ProtectionOptions>()?;
    m.add_class::<IgnoreError>()?;
    Ok(())
}
