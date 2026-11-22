//! `ReadWorksheet`: values, formulas, styles, merged ranges and hyperlinks
//! of one sheet of a `ReadWorkbook`.
//!
//! Coordinates are absolute and zero-based: row 0 is Excel row 1 whatever
//! the first used row is, and lists are padded with None from row/column 0
//! up to the last used cell.

use std::collections::HashMap;
use std::sync::Arc;

use calamine::{Data, Range};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyString, PyTuple};

use crate::common::cellref::{
    Cell, MAX_COLS, MAX_ROWS, cell_name, col_to_name, range_name, split_cell,
};
use crate::common::errors::range_err;
use crate::common::format::Format;
use crate::reader::convert::data_to_py;
use crate::reader::shared::{Shared, SheetData};
use crate::reader::style::style_to_spec;

fn none(py: Python<'_>) -> Bound<'_, PyAny> {
    py.None().into_bound(py)
}

/// The cells of absolute row `row` and the absolute column of the first one.
fn row_slice<T: calamine::CellType>(range: &Range<T>, row: u32) -> Option<(&[T], u32)> {
    let (first_row, first_col) = range.start()?;
    let (last_row, _) = range.end()?;
    if row < first_row || row > last_row {
        return None;
    }
    Some((&range[(row - first_row) as usize], first_col))
}

/// The values of absolute row `row`, columns `c1..=c2` (padded with None),
/// converted by `convert`. Lists and tuples are built straight from this.
fn row_values<'a, 'py>(
    py: Python<'py>,
    range: &'a Range<Data>,
    row: u32,
    c1: u32,
    c2: u32,
    mut convert: impl FnMut(&'a Data) -> Bound<'py, PyAny>,
) -> impl ExactSizeIterator<Item = Bound<'py, PyAny>> {
    let slice = row_slice(range, row);
    (c1..c2 + 1).map(move |c| match slice {
        Some((cells, first)) if c >= first => match cells.get((c - first) as usize) {
            Some(value) => convert(value),
            None => none(py),
        },
        _ => none(py),
    })
}

/// One row as a list, columns `c1..=c2` (padded with None).
fn row_list<'py>(
    py: Python<'py>,
    range: &Range<Data>,
    row: u32,
    c1: u32,
    c2: u32,
    is_1904: bool,
) -> PyResult<Bound<'py, PyList>> {
    PyList::new(py, row_values(py, range, row, c1, c2, |v| data_to_py(py, v, is_1904)))
}

/// Converts cells for one call that returns many rows: equal texts become
/// one shared Python `str` (as openpyxl does with the shared string table)
/// instead of one object per cell.
struct ValueCache<'a, 'py> {
    py: Python<'py>,
    is_1904: bool,
    strings: HashMap<&'a str, Bound<'py, PyString>>,
}

impl<'a, 'py> ValueCache<'a, 'py> {
    fn new(py: Python<'py>, is_1904: bool) -> Self {
        ValueCache { py, is_1904, strings: HashMap::new() }
    }

    fn get(&mut self, value: &'a Data) -> Bound<'py, PyAny> {
        match value {
            Data::String(s) => {
                let py = self.py;
                self.strings.entry(s.as_str()).or_insert_with(|| PyString::new(py, s)).clone().into_any()
            }
            other => data_to_py(self.py, other, self.is_1904),
        }
    }
}

/// Pauses Python's cyclic garbage collector while a large list of lists is
/// built (fresh lists of scalars cannot form cycles, but each new list counts
/// towards a collection). The previous state is restored on drop.
struct GcPause {
    was_enabled: bool,
}

impl GcPause {
    /// Only worth it for big results.
    fn for_rows(_py: Python<'_>, rows: u32) -> Option<Self> {
        if rows < 1000 {
            return None;
        }
        // SAFETY: the `Python` token proves we are attached to the interpreter.
        let was_enabled = unsafe { pyo3::ffi::PyGC_Disable() } != 0;
        Some(GcPause { was_enabled })
    }
}

impl Drop for GcPause {
    fn drop(&mut self) {
        if self.was_enabled {
            // SAFETY: created while attached and dropped in the same scope.
            unsafe {
                pyo3::ffi::PyGC_Enable();
            }
        }
    }
}

/// Rows `r1..=r2`, columns `c1..=c2`, as a list of lists.
pub fn rows_to_list<'py>(
    py: Python<'py>,
    range: &Range<Data>,
    r1: u32,
    r2: u32,
    c1: u32,
    c2: u32,
    is_1904: bool,
) -> PyResult<Bound<'py, PyList>> {
    let _pause = GcPause::for_rows(py, r2 - r1 + 1);
    let mut cache = ValueCache::new(py, is_1904);
    let mut rows = Vec::with_capacity((r2 - r1) as usize + 1);
    for row in r1..=r2 {
        rows.push(PyList::new(py, row_values(py, range, row, c1, c2, |v| cache.get(v)))?);
    }
    PyList::new(py, rows)
}

/// Rows `r1..=r2`, columns `c1..=c2`, as dicts keyed by `headers`, built
/// straight from the cells (no list of rows in between). With `skip_empty`,
/// rows without any value are left out.
pub fn records_from_range<'py>(
    py: Python<'py>,
    range: &Range<Data>,
    r1: u32,
    r2: u32,
    c1: u32,
    c2: u32,
    is_1904: bool,
    headers: &[Bound<'py, PyAny>],
    skip_empty: bool,
) -> PyResult<Bound<'py, PyList>> {
    let _pause = GcPause::for_rows(py, r2 - r1 + 1);
    let mut cache = ValueCache::new(py, is_1904);
    let out = PyList::empty(py);
    for row in r1..=r2 {
        let record = PyDict::new(py);
        let mut empty = true;
        for (key, value) in headers.iter().zip(row_values(py, range, row, c1, c2, |v| cache.get(v))) {
            empty &= value.is_none();
            record.set_item(key, value)?;
        }
        if !(skip_empty && empty) {
            out.append(record)?;
        }
    }
    Ok(out)
}

fn formula_text<'py>(py: Python<'py>, formula: &str) -> Bound<'py, PyAny> {
    // ods formulas carry an "of:" namespace prefix: "of:=[.A1]*2".
    let formula = formula.strip_prefix("of:").unwrap_or(formula);
    if formula.starts_with('=') {
        PyString::new(py, formula).into_any()
    } else {
        PyString::new(py, &format!("={formula}")).into_any()
    }
}

fn formula_at<'py>(
    py: Python<'py>,
    formulas: &Range<String>,
    row: u32,
    col: u32,
) -> Option<Bound<'py, PyAny>> {
    let (cells, first) = row_slice(formulas, row)?;
    let text = cells.get(col.checked_sub(first)? as usize)?;
    (!text.is_empty()).then(|| formula_text(py, text))
}

/// Parse `(row, col)` or `"A1"` and reject extra arguments.
fn one_cell(method: &str, args: &Bound<'_, PyTuple>) -> PyResult<Cell> {
    let items: Vec<Bound<'_, PyAny>> = args.iter().collect();
    let (cell, rest) = split_cell(&items)?;
    if !rest.is_empty() {
        return Err(PyTypeError::new_err(format!(
            "{method}() takes a cell: (row, col) or \"A1\", got {} extra argument(s)",
            rest.len()
        )));
    }
    Ok(cell)
}

fn bound_arg(name: &str, value: Option<i64>, limit: u32) -> PyResult<Option<u32>> {
    match value {
        None => Ok(None),
        Some(v) if (0..limit as i64).contains(&v) => Ok(Some(v as u32)),
        Some(v) => Err(range_err(format!("{name}={v} is outside 0..{limit}"))),
    }
}

/// A cell yielded by `iter_rows(values_only=False)`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, skip_from_py_object)]
pub struct ReadCell {
    /// Zero-based row.
    #[pyo3(get)]
    row: u32,
    /// Zero-based column.
    #[pyo3(get)]
    col: u32,
    /// The cell value.
    #[pyo3(get)]
    value: Py<PyAny>,
}

#[pymethods]
impl ReadCell {
    /// The cell reference, e.g. "B3".
    #[getter]
    fn coordinate(&self) -> String {
        cell_name(self.row, self.col as u16, false, false)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!(
            "<ReadCell {} value={}>",
            self.coordinate(),
            self.value.bind(py).repr()?
        ))
    }
}

/// A hyperlink in a worksheet (xlsx only).
#[pyclass(module = "pyrust_xlsxwriter", frozen, skip_from_py_object)]
pub struct HyperlinkInfo {
    /// `(first_row, first_col, last_row, last_col)` of the linked cells.
    #[pyo3(get)]
    range: (u32, u32, u32, u32),
    /// The external target (URL, mailto:, file path), or None for internal links.
    #[pyo3(get)]
    target: Option<String>,
    /// The location inside the target or workbook, e.g. "Sheet2!A1".
    #[pyo3(get)]
    location: Option<String>,
    /// The display text stored with the link, if any.
    #[pyo3(get)]
    display: Option<String>,
    /// The tooltip, if any.
    #[pyo3(get)]
    tooltip: Option<String>,
}

#[pymethods]
impl HyperlinkInfo {
    /// `(row, col)` of the top-left linked cell.
    #[getter]
    fn cell(&self) -> (u32, u32) {
        (self.range.0, self.range.1)
    }

    /// The linked range as "A1" / "A1:B2".
    #[getter]
    fn reference(&self) -> String {
        let (r1, c1, r2, c2) = self.range;
        range_name((r1, c1 as u16, r2, c2 as u16), false)
    }

    fn __repr__(&self) -> String {
        let show = |v: &Option<String>| v.as_ref().map_or("None".to_string(), |s| format!("{s:?}"));
        format!(
            "<HyperlinkInfo {} target={} location={}>",
            self.reference(),
            show(&self.target),
            show(&self.location)
        )
    }
}

/// Iterator returned by `ReadWorksheet.iter_rows()`.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct RowIterator {
    data: Arc<SheetData>,
    is_1904: bool,
    next: u32,
    last: u32,
    done: bool,
    min_col: u32,
    max_col: u32,
    values_only: bool,
}

#[pymethods]
impl RowIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.done || self.next > self.last {
            self.done = true;
            return Ok(None);
        }
        let row = self.next;
        if row == self.last {
            self.done = true;
        } else {
            self.next += 1;
        }
        // One row at a time, so no shared-string cache: it would keep every
        // distinct text alive for the whole iteration.
        let is_1904 = self.is_1904;
        let values = row_values(py, &self.data.values, row, self.min_col, self.max_col, |v| {
            data_to_py(py, v, is_1904)
        });
        if self.values_only {
            return Ok(Some(PyTuple::new(py, values)?));
        }
        let cells = values
            .zip(self.min_col..)
            .map(|(value, col)| {
                Bound::new(
                    py,
                    ReadCell {
                        row,
                        col,
                        value: value.unbind(),
                    },
                )
            })
            .collect::<PyResult<Vec<_>>>()?;
        Ok(Some(PyTuple::new(py, cells)?))
    }
}

/// A sheet of a `ReadWorkbook` (from `book.worksheet(...)` or `book[...]`).
///
/// All rows and columns are zero-based and absolute: `values()[0]` is Excel
/// row 1 even if the data starts lower down.
#[pyclass(module = "pyrust_xlsxwriter", frozen)]
pub struct ReadWorksheet {
    shared: Arc<Shared>,
    index: usize,
    data: Arc<SheetData>,
}

impl ReadWorksheet {
    pub fn new(shared: Arc<Shared>, index: usize, data: Arc<SheetData>) -> Self {
        ReadWorksheet {
            shared,
            index,
            data,
        }
    }

    fn end(&self) -> Option<(u32, u32)> {
        self.data.values.end()
    }
}

#[pymethods]
impl ReadWorksheet {
    /// The sheet name.
    #[getter]
    fn name(&self) -> &str {
        &self.shared.sheets[self.index].name
    }

    /// The zero-based position of the sheet in the workbook.
    #[getter]
    fn index(&self) -> usize {
        self.index
    }

    /// "worksheet", "chartsheet", "dialogsheet", "macrosheet" or "vba".
    #[getter]
    fn kind(&self) -> &'static str {
        self.shared.sheets[self.index].kind
    }

    /// "visible", "hidden" or "very_hidden".
    #[getter]
    fn visibility(&self) -> &'static str {
        self.shared.sheets[self.index].visibility
    }

    /// `(first_row, first_col, last_row, last_col)` of the cells with a value,
    /// or None for an empty sheet.
    #[getter]
    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        let (r1, c1) = self.data.values.start()?;
        let (r2, c2) = self.end()?;
        Some((r1, c1, r2, c2))
    }

    /// Number of rows up to the last used one (last row + 1; 0 if empty).
    #[getter]
    fn max_row(&self) -> u32 {
        self.end().map_or(0, |(r, _)| r + 1)
    }

    /// Number of columns up to the last used one (last column + 1; 0 if empty).
    #[getter]
    fn max_column(&self) -> u32 {
        self.end().map_or(0, |(_, c)| c + 1)
    }

    /// All values as a list of rows, from A1 to the last used cell, padded with None.
    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        match self.end() {
            None => Ok(PyList::empty(py)),
            Some((last_row, last_col)) => rows_to_list(
                py,
                &self.data.values,
                0,
                last_row,
                0,
                last_col,
                self.shared.is_1904,
            ),
        }
    }

    /// Iterate over rows as tuples, like openpyxl's `iter_rows()` but
    /// zero-based. Bounds are inclusive; `max_row` / `max_col` default to the
    /// last used row / column (an empty sheet then yields nothing). With
    /// `values_only=False` the tuples hold `ReadCell` objects.
    #[pyo3(signature = (min_row=0, max_row=None, min_col=0, max_col=None, values_only=true))]
    fn iter_rows(
        &self,
        min_row: i64,
        max_row: Option<i64>,
        min_col: i64,
        max_col: Option<i64>,
        values_only: bool,
    ) -> PyResult<RowIterator> {
        let min_row = bound_arg("min_row", Some(min_row), MAX_ROWS)?.unwrap_or(0);
        let min_col = bound_arg("min_col", Some(min_col), MAX_COLS as u32)?.unwrap_or(0);
        let max_row = bound_arg("max_row", max_row, MAX_ROWS)?;
        let max_col = bound_arg("max_col", max_col, MAX_COLS as u32)?;
        let end = self.end();
        let (last, last_col, empty) =
            match (max_row.or(end.map(|e| e.0)), max_col.or(end.map(|e| e.1))) {
                (Some(r), Some(c)) => (r, c, r < min_row || c < min_col),
                _ => (0, 0, true),
            };
        Ok(RowIterator {
            data: self.data.clone(),
            is_1904: self.shared.is_1904,
            next: min_row,
            last,
            done: empty,
            min_col,
            max_col: last_col,
            values_only,
        })
    }

    /// The value of one cell: `cell(row, col)` or `cell("A1")` (None if empty).
    #[pyo3(signature = (*args))]
    fn cell<'py>(
        &self,
        py: Python<'py>,
        args: &Bound<'py, PyTuple>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let (row, col) = one_cell("cell", args)?;
        Ok(match self.data.values.get_value((row, col as u32)) {
            Some(value) => data_to_py(py, value, self.shared.is_1904),
            None => none(py),
        })
    }

    /// The formula of one cell with a leading "=" (as openpyxl shows it), or None.
    #[pyo3(signature = (*args))]
    fn formula<'py>(
        &self,
        py: Python<'py>,
        args: &Bound<'py, PyTuple>,
    ) -> PyResult<Option<Bound<'py, PyAny>>> {
        let (row, col) = one_cell("formula", args)?;
        let formulas = self.shared.formulas(py, self.index, &self.data)?;
        Ok(formula_at(py, formulas, row, col as u32))
    }

    /// Every formula ("=..." text, None for cells without one) in the same
    /// absolute layout as `values()`, extended if a formula lies outside it.
    fn formulas<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let formulas = self.shared.formulas(py, self.index, &self.data)?;
        let (last_row, last_col) = match (self.end(), formulas.end()) {
            (None, None) => return Ok(PyList::empty(py)),
            (Some(a), None) | (None, Some(a)) => a,
            (Some(a), Some(b)) => (a.0.max(b.0), a.1.max(b.1)),
        };
        let _pause = GcPause::for_rows(py, last_row + 1);
        let mut rows = Vec::with_capacity(last_row as usize + 1);
        for row in 0..=last_row {
            let items = (0..last_col + 1)
                .map(|col| formula_at(py, formulas, row, col).unwrap_or_else(|| none(py)));
            rows.push(PyList::new(py, items)?);
        }
        PyList::new(py, rows)
    }

    /// The format of one cell as a `Format`, or None for the default format.
    /// Only xlsx/xlsm files carry styles; other formats always give None.
    #[pyo3(signature = (*args))]
    fn style(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<Option<Format>> {
        let (row, col) = one_cell("style", args)?;
        let styles = self.shared.styles(py, self.index, &self.data)?;
        Ok(styles
            .get(row, col as u32)
            .map(|style| Format::from_spec(style_to_spec(style))))
    }

    /// The merged ranges as `(first_row, first_col, last_row, last_col)`
    /// (xlsx and xls; empty for other formats).
    #[getter]
    fn merged_ranges(&self, py: Python<'_>) -> PyResult<Vec<(u32, u32, u32, u32)>> {
        Ok(self.shared.merged(py, self.index, &self.data)?.to_vec())
    }

    /// The hyperlinks (xlsx only; empty for other formats).
    #[getter]
    fn hyperlinks(&self, py: Python<'_>) -> PyResult<Vec<HyperlinkInfo>> {
        Ok(self
            .shared
            .hyperlinks(py, self.index, &self.data)?
            .iter()
            .map(|h| HyperlinkInfo {
                range: (
                    h.range.start.0,
                    h.range.start.1,
                    h.range.end.0,
                    h.range.end.1,
                ),
                target: h.target.clone(),
                location: h.location.clone(),
                display: h.displayed_text.clone(),
                tooltip: h.tooltip.clone(),
            })
            .collect())
    }

    /// The rows below `header_row` as dicts keyed by the header cells.
    ///
    /// Columns span the used range; an empty header cell is keyed by its
    /// column letter, a repeated header gets a ".1", ".2", ... suffix, and
    /// entirely empty rows are skipped.
    #[pyo3(signature = (header_row=0))]
    fn records<'py>(&self, py: Python<'py>, header_row: i64) -> PyResult<Bound<'py, PyList>> {
        let header_row = bound_arg("header_row", Some(header_row), MAX_ROWS)?.unwrap_or(0);
        let (Some((_, first_col)), Some((last_row, last_col))) =
            (self.data.values.start(), self.end())
        else {
            return Ok(PyList::empty(py));
        };
        if header_row >= last_row {
            return Ok(PyList::empty(py));
        }
        let is_1904 = self.shared.is_1904;
        let header = row_list(
            py,
            &self.data.values,
            header_row,
            first_col,
            last_col,
            is_1904,
        )?;
        let mut keys: Vec<Bound<'py, PyAny>> = Vec::with_capacity(header.len());
        let mut seen: Vec<String> = Vec::with_capacity(header.len());
        for (value, col) in header.iter().zip(first_col..) {
            let base = if value.is_none() {
                col_to_name(col as u16)
            } else {
                value.str()?.to_string()
            };
            let mut key = base.clone();
            let mut n = 0;
            while seen.contains(&key) {
                n += 1;
                key = format!("{base}.{n}");
            }
            keys.push(PyString::new(py, &key).into_any());
            seen.push(key);
        }
        records_from_range(
            py,
            &self.data.values,
            header_row + 1,
            last_row,
            first_col,
            last_col,
            is_1904,
            &keys,
            true,
        )
    }

    fn __repr__(&self) -> String {
        let dims = match self.dimensions() {
            Some((r1, c1, r2, c2)) => range_name((r1, c1 as u16, r2, c2 as u16), false),
            None => "empty".to_string(),
        };
        format!("<ReadWorksheet {:?} {dims}>", self.name())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ReadWorksheet>()?;
    m.add_class::<ReadCell>()?;
    m.add_class::<HyperlinkInfo>()?;
    Ok(())
}
