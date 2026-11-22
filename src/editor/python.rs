//! Python classes for editing existing workbooks: `load_workbook()`,
//! `EditWorkbook` and `EditWorksheet`.
//!
//! Writing uses the same calling conventions as the writer's `Worksheet`
//! ("A1" or (row, col), `format=`), with one difference that suits
//! templates: writing without a format keeps the cell's existing style.

use std::fs::File;
use std::io::BufWriter;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, MutexGuard};

use calamine::{Data, DataType};
use chrono::NaiveDate;
use pyo3::exceptions::{PyIndexError, PyKeyError, PyTypeError};
use pyo3::prelude::*;
use pyo3::sync::MutexExt;
use pyo3::types::{PyBytes, PyDelta, PyDict, PyList, PyString, PyTuple};
use pyo3::IntoPyObjectExt;

use crate::common::bytes::{write_to_file, SharedBytes};
use crate::common::cellref::{
    check_cell, range_name, split_cell, split_cols, split_range, split_row, CellRange, Params,
};
use crate::common::color::Color;
use crate::common::errors::{edit_err, file_err, new_err, param_err, read_err, ErrorKind};
use crate::common::format::{Format, FormatSpec};
use crate::common::value::{classify, excel_serial, time_fraction, CellValue, Url};
use crate::editor::book::{BookState, Visibility};
use crate::editor::package::Package;
use crate::editor::sheet::{CellEdit, ColEdit, EditValue, RowEdit, StyleOp};
use crate::writer::workbook::Workbook;

type Shared = Arc<Mutex<BookState>>;

fn lock<'a>(py: Python<'_>, book: &'a Shared) -> PyResult<MutexGuard<'a, BookState>> {
    book.lock_py_attached(py).map_err(|_| edit_err("the workbook is in an inconsistent state after an earlier panic"))
}

/// The bytes of a path, bytes object or binary file object, without copying
/// them: Python `bytes` are shared in place, a file's contents are moved in.
pub fn read_source(py: Python<'_>, source: &Bound<'_, PyAny>) -> PyResult<(SharedBytes, Option<PathBuf>)> {
    if source.is_instance_of::<PyBytes>() {
        return Ok((SharedBytes::from_py(source)?, None));
    }
    if source.hasattr("read")? && !source.is_instance_of::<PyString>() {
        let data = source.call_method0("read")?;
        if !data.is_instance_of::<PyBytes>() {
            return Err(PyTypeError::new_err("file object must be opened in binary mode"));
        }
        return Ok((SharedBytes::from_py(&data)?, None));
    }
    let path: PathBuf = source.extract().map_err(|_| {
        PyTypeError::new_err("source must be a path, bytes, or a binary file object")
    })?;
    let data = py
        .detach(|| std::fs::read(&path))
        .map_err(|e| file_err(format!("cannot read {}: {e}", path.display())))?;
    Ok((SharedBytes::from_vec(data), Some(path)))
}

/// Convert a calamine value to Python.
pub fn data_to_py(py: Python<'_>, value: &Data) -> PyResult<Py<PyAny>> {
    Ok(match value {
        Data::Empty => py.None(),
        Data::Int(i) => i.into_py_any(py)?,
        Data::Float(f) => {
            if f.fract() == 0.0 && f.abs() < 9.007_199_254_740_992e15 {
                (*f as i64).into_py_any(py)?
            } else {
                f.into_py_any(py)?
            }
        }
        Data::String(s) => s.into_py_any(py)?,
        Data::Bool(b) => b.into_py_any(py)?,
        Data::Error(e) => e.to_string().into_py_any(py)?,
        Data::DateTimeIso(s) | Data::DurationIso(s) => s.into_py_any(py)?,
        Data::DateTime(dt) => {
            if dt.is_duration() {
                match dt.as_duration() {
                    Some(d) => d.into_py_any(py)?,
                    None => dt.as_f64().into_py_any(py)?,
                }
            } else if dt.as_f64() < 1.0 && dt.as_f64() >= 0.0 {
                match value.as_datetime() {
                    Some(d) => d.time().into_py_any(py)?,
                    None => dt.as_f64().into_py_any(py)?,
                }
            } else {
                match dt.as_datetime() {
                    Some(d) => d.into_py_any(py)?,
                    None => dt.as_f64().into_py_any(py)?,
                }
            }
        }
    })
}

fn serial_to_py(py: Python<'_>, serial: f64, code: &str, is_1904: bool) -> PyResult<Py<PyAny>> {
    if code == "[h]:mm:ss" {
        let seconds = (serial * 86_400.0).round() as i64;
        return PyDelta::new(py, (seconds / 86_400) as i32, (seconds % 86_400) as i32, 0, true)?.into_py_any(py);
    }
    let base = if is_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1)
    } else if serial < 60.0 {
        NaiveDate::from_ymd_opt(1899, 12, 31)
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)
    }
    .and_then(|d| d.and_hms_opt(0, 0, 0))
    .ok_or_else(|| param_err("invalid date"))?;
    let millis = (serial * 86_400_000.0).round() as i64;
    let dt = base + chrono::Duration::milliseconds(millis);
    if code == "hh:mm:ss" {
        return dt.time().into_py_any(py);
    }
    dt.into_py_any(py)
}

/// What a Python value becomes in the editor.
struct Prepared {
    value: EditValue,
    style: StyleOp,
    link: Option<Url>,
}

/// Turn a Python value into a cell edit. `format` is shared (not copied) by
/// every cell it is applied to.
fn prepare(value: &Bound<'_, PyAny>, format: Option<&Arc<FormatSpec>>, is_1904: bool) -> PyResult<Prepared> {
    let explicit = format.map(|f| StyleOp::Replace(f.clone()));
    let keep = || explicit.clone().unwrap_or(StyleOp::Keep);
    let date_style = |code: &'static str| explicit.clone().unwrap_or(StyleOp::DefaultNumFormat(code));
    let serial = |dt: &chrono::NaiveDateTime| -> f64 {
        let s = excel_serial(dt);
        if is_1904 { s - 1462.0 } else { s }
    };
    let mut link = None;
    let (value, style) = match classify(value)? {
        CellValue::Blank => (EditValue::Blank, keep()),
        CellValue::Bool(b) => (EditValue::Bool(b), keep()),
        CellValue::Number(n) => (EditValue::Number(n), keep()),
        CellValue::Decimal(d) => (EditValue::Number(d.to_string().parse::<f64>().unwrap_or(f64::NAN)), keep()),
        CellValue::String(s) => (EditValue::Str(s.to_str()?.to_string()), keep()),
        CellValue::Formula(text) => (EditValue::formula(strip_equals(text), None, None), keep()),
        CellValue::ArrayFormula(text) => (EditValue::formula(strip_equals(text), None, Some(String::new())), keep()),
        CellValue::FormulaObject(f) => (EditValue::formula(strip_equals(f.formula), f.result, None), keep()),
        CellValue::Url(url) => {
            let text = url.text.clone().unwrap_or_else(|| url.link.trim_start_matches("internal:").to_string());
            link = Some(url);
            let style = explicit.clone().unwrap_or_else(|| StyleOp::Merge(Arc::new(FormatSpec { hyperlink: true, ..Default::default() })));
            (EditValue::Str(text), style)
        }
        CellValue::DateTime(dt) => (EditValue::Number(serial(&dt)), date_style("yyyy-mm-dd hh:mm:ss")),
        CellValue::Date(d) => {
            let dt = d.and_hms_opt(0, 0, 0).ok_or_else(|| param_err("invalid date"))?;
            (EditValue::Number(serial(&dt)), date_style("yyyy-mm-dd"))
        }
        CellValue::Time(t) => (EditValue::Number(time_fraction(&t)), date_style("hh:mm:ss")),
        CellValue::Duration(days) => (EditValue::Number(days), date_style("[h]:mm:ss")),
    };
    Ok(Prepared { value, style, link })
}

/// Open an existing .xlsx/.xlsm file for editing.
///
/// `source` may be a path, `bytes`, or a binary file object. Everything the
/// editor doesn't change (charts, pivot tables, images, VBA, ...) is kept
/// exactly as it was.
#[pyfunction]
pub fn load_workbook(py: Python<'_>, source: &Bound<'_, PyAny>) -> PyResult<EditWorkbook> {
    let (data, path) = read_source(py, source)?;
    let state = py.detach(|| BookState::open(data)).map_err(read_err)?;
    Ok(EditWorkbook { book: Arc::new(Mutex::new(state)), path })
}

/// An existing workbook opened for editing (see `load_workbook()`).
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct EditWorkbook {
    book: Shared,
    path: Option<PathBuf>,
}

impl EditWorkbook {
    fn sheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<EditWorksheet> {
        let state = lock(py, &self.book)?;
        let index = if let Ok(name) = key.cast::<PyString>() {
            let name = name.to_str()?;
            state.index_by_name(name).ok_or_else(|| PyKeyError::new_err(format!("no worksheet named {name:?}")))?
        } else {
            let index: isize = key.extract().map_err(|_| PyTypeError::new_err("worksheet key must be a name or an index"))?;
            let len = state.sheets.len() as isize;
            let index = if index < 0 { index + len } else { index };
            if index < 0 || index >= len {
                return Err(PyIndexError::new_err("worksheet index out of range"));
            }
            index as usize
        };
        Ok(EditWorksheet { book: self.book.clone(), id: state.sheets[index].id })
    }

    /// Build the edited package (without the GIL); write it with `Package::save`.
    fn build(&self, py: Python<'_>) -> PyResult<Package> {
        let book = self.book.clone();
        py.detach(move || {
            let mut state = book.lock().map_err(|_| "the workbook is in an inconsistent state".to_string())?;
            state.build()
        })
        .map_err(edit_err)
    }

    /// The edited file as bytes in memory.
    fn build_bytes(&self, py: Python<'_>) -> PyResult<Vec<u8>> {
        let mut package = self.build(py)?;
        py.detach(|| package.save_to_vec()).map_err(edit_err)
    }
}

#[pymethods]
impl EditWorkbook {
    /// Worksheet names in tab order.
    #[getter]
    fn sheet_names(&self, py: Python<'_>) -> PyResult<Vec<String>> {
        Ok(lock(py, &self.book)?.sheets.iter().map(|s| s.name.clone()).collect())
    }

    /// All sheets in tab order.
    #[getter]
    fn worksheets(&self, py: Python<'_>) -> PyResult<Vec<EditWorksheet>> {
        Ok(lock(py, &self.book)?
            .sheets
            .iter()
            .map(|s| EditWorksheet { book: self.book.clone(), id: s.id })
            .collect())
    }

    /// Look up a worksheet by name or index.
    fn worksheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<EditWorksheet> {
        self.sheet(py, key)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<EditWorksheet> {
        self.sheet(py, key)
    }

    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(lock(py, &self.book)?.sheets.len())
    }

    fn __contains__(&self, py: Python<'_>, name: &str) -> PyResult<bool> {
        Ok(lock(py, &self.book)?.index_by_name(name).is_some())
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let sheets = self.worksheets(py)?;
        PyList::new(py, sheets)?.call_method0("__iter__").map(|i| i.unbind())
    }

    /// Whether the workbook uses the 1904 date system.
    #[getter]
    fn is_1904(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(lock(py, &self.book)?.is_1904)
    }

    /// Index of the active (selected) worksheet.
    #[getter]
    fn active(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(lock(py, &self.book)?.active_index())
    }

    #[setter]
    fn set_active(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let index = match value.cast::<EditWorksheet>() {
            Ok(ws) => {
                let id = ws.borrow().id;
                lock(py, &self.book)?.index_of(id).map_err(edit_err)?
            }
            Err(_) => value.extract::<usize>()?,
        };
        lock(py, &self.book)?.set_active_index(index).map_err(edit_err)
    }

    /// Add a new, empty worksheet (at the end, or at `index`).
    #[pyo3(signature = (name=None, index=None))]
    fn add_worksheet(&self, py: Python<'_>, name: Option<&str>, index: Option<usize>) -> PyResult<EditWorksheet> {
        let id = lock(py, &self.book)?
            .add_sheet(name, index)
            .map_err(|e| new_err(ErrorKind::SheetName, e))?;
        Ok(EditWorksheet { book: self.book.clone(), id })
    }

    /// Remove a worksheet (by name, index or object).
    fn remove_worksheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<()> {
        let id = match key.cast::<EditWorksheet>() {
            Ok(ws) => ws.borrow().id,
            Err(_) => self.sheet(py, key)?.id,
        };
        let mut state = lock(py, &self.book)?;
        let index = state.index_of(id).map_err(edit_err)?;
        state.remove_sheet(index).map_err(edit_err)
    }

    /// Move a worksheet to a new tab position.
    fn move_worksheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>, index: usize) -> PyResult<()> {
        let id = match key.cast::<EditWorksheet>() {
            Ok(ws) => ws.borrow().id,
            Err(_) => self.sheet(py, key)?.id,
        };
        let mut state = lock(py, &self.book)?;
        let from = state.index_of(id).map_err(edit_err)?;
        state.move_sheet(from, index).map_err(edit_err)
    }

    /// Defined names as `(name, formula, sheet_name_or_None)`.
    #[getter]
    fn defined_names(&self, py: Python<'_>) -> PyResult<Vec<(String, String, Option<String>)>> {
        let state = lock(py, &self.book)?;
        Ok(state
            .defined_names()
            .into_iter()
            .map(|(name, formula, local)| (name, formula, local.and_then(|i| state.sheets.get(i)).map(|s| s.name.clone())))
            .collect())
    }

    /// Define (or replace) a name, e.g. `define_name("Sales", "=Data!$B$2:$B$100")`.
    /// `sheet` makes it local to one worksheet.
    #[pyo3(signature = (name, formula, sheet=None))]
    fn define_name(&self, py: Python<'_>, name: &str, formula: &str, sheet: Option<&str>) -> PyResult<()> {
        let mut state = lock(py, &self.book)?;
        let local = match sheet {
            Some(s) => Some(state.index_by_name(s).ok_or_else(|| PyKeyError::new_err(format!("no worksheet named {s:?}")))?),
            None => None,
        };
        state.define_name(name, formula, local).map_err(param_err)
    }

    /// Remove a defined name; returns whether it existed.
    #[pyo3(signature = (name, sheet=None))]
    fn remove_name(&self, py: Python<'_>, name: &str, sheet: Option<&str>) -> PyResult<bool> {
        let mut state = lock(py, &self.book)?;
        let local = match sheet {
            Some(s) => Some(state.index_by_name(s).ok_or_else(|| PyKeyError::new_err(format!("no worksheet named {s:?}")))?),
            None => None,
        };
        state.remove_name(name, local).map_err(edit_err)
    }

    /// Document properties (title, subject, author, keywords, comments, ...).
    #[getter]
    fn properties<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let props = lock(py, &self.book)?.properties().map_err(edit_err)?;
        let dict = PyDict::new(py);
        for (k, v) in props {
            dict.set_item(k, v)?;
        }
        Ok(dict)
    }

    /// Set document properties: title, subject, author, keywords, comments,
    /// category, status, last_modified_by, created, modified (datetimes or ISO strings).
    #[pyo3(signature = (**kwargs))]
    fn set_properties(&self, py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        const KEYS: &[&str] = &["title", "subject", "author", "keywords", "comments", "category", "status", "last_modified_by", "created", "modified"];
        let mut updates = Vec::new();
        if let Some(kwargs) = kwargs {
            for (k, v) in kwargs.iter() {
                let key: String = k.extract()?;
                if !KEYS.contains(&key.as_str()) {
                    return Err(PyTypeError::new_err(format!("set_properties() got an unexpected keyword argument '{key}'")));
                }
                let text = if key == "created" || key == "modified" {
                    match v.extract::<chrono::NaiveDateTime>() {
                        Ok(dt) => dt.format("%Y-%m-%dT%H:%M:%SZ").to_string(),
                        Err(_) => v.extract::<String>()?,
                    }
                } else {
                    v.extract::<String>()?
                };
                updates.push((key, text));
            }
        }
        let mut state = lock(py, &self.book)?;
        for (k, v) in updates {
            state.set_property(&k, &v);
        }
        Ok(())
    }

    /// Save to `path` (default: the path the workbook was loaded from).
    #[pyo3(signature = (path=None))]
    fn save(&self, py: Python<'_>, path: Option<PathBuf>) -> PyResult<()> {
        let path = path
            .or_else(|| self.path.clone())
            .ok_or_else(|| PyTypeError::new_err("save() needs a path (the workbook was not loaded from a file)"))?;
        // Streamed straight into the file: the result is never held in memory whole.
        let mut package = self.build(py)?;
        py.detach(|| write_atomically(&path, |file| package.save(file).map_err(std::io::Error::other)))
            .map_err(|e| file_err(format!("cannot write {}: {e}", path.display())))
    }

    /// The edited file as `bytes`.
    fn save_to_buffer<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let data = self.build_bytes(py)?;
        Ok(PyBytes::new(py, &data))
    }

    /// Write the edited file to a binary file object.
    fn save_to_file(&self, py: Python<'_>, file: &Bound<'_, PyAny>) -> PyResult<()> {
        write_to_file(file, &self.build_bytes(py)?)
    }

    /// Save back to the file the workbook was loaded from.
    fn close(&self, py: Python<'_>) -> PyResult<()> {
        self.save(py, None)
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    /// Save on a clean exit (when loaded from a path); do nothing if the block raised.
    fn __exit__(
        &self,
        py: Python<'_>,
        exc_type: Option<&Bound<'_, PyAny>>,
        _exc: Option<&Bound<'_, PyAny>>,
        _tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        if exc_type.is_none_or(|t| t.is_none()) && self.path.is_some() {
            self.close(py)?;
        }
        Ok(false)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!("<EditWorkbook sheets={:?}>", self.sheet_names(py)?))
    }
}

/// Write via a temporary file and rename, so a failed save never leaves a
/// truncated file in place of the original. `write` streams the content into
/// the (buffered) temporary file; on any failure the temporary file is removed.
fn write_atomically(path: &Path, write: impl FnOnce(BufWriter<File>) -> std::io::Result<BufWriter<File>>) -> std::io::Result<()> {
    let dir = path.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or(Path::new("."));
    let name = path.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    let tmp = dir.join(format!(".{name}.{}.tmp", std::process::id()));
    let result = File::create(&tmp)
        .and_then(|file| write(BufWriter::new(file)))
        .and_then(|out| out.into_inner().map_err(|e| e.into_error()))
        .and_then(|file| {
            drop(file);
            std::fs::rename(&tmp, path)
        });
    if result.is_err() {
        let _ = std::fs::remove_file(&tmp);
    }
    result
}

/// A worksheet of an `EditWorkbook`. Rows and columns are zero-based; every
/// method taking a cell or range also accepts "A1" / "A1:B2".
#[pyclass(module = "pyrust_xlsxwriter", frozen)]
pub struct EditWorksheet {
    book: Shared,
    id: u64,
}

impl EditWorksheet {
    fn with<R>(&self, py: Python<'_>, f: impl FnOnce(&mut BookState, usize) -> PyResult<R>) -> PyResult<R> {
        let mut state = lock(py, &self.book)?;
        let index = state.index_of(self.id).map_err(edit_err)?;
        if !state.sheets[index].is_worksheet {
            return Err(edit_err(format!("{:?} is a chartsheet, not a worksheet", state.sheets[index].name)));
        }
        f(&mut state, index)
    }

    fn is_1904(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(lock(py, &self.book)?.is_1904)
    }

    /// Prepare (outside the lock) and apply one write.
    fn write_one(&self, py: Python<'_>, row: u32, col: u16, value: &Bound<'_, PyAny>, format: Option<&Arc<FormatSpec>>) -> PyResult<()> {
        let prepared = prepare(value, format, self.is_1904(py)?)?;
        self.apply(py, vec![(row, col, prepared)])
    }

    fn apply(&self, py: Python<'_>, items: Vec<(u32, u16, Prepared)>) -> PyResult<()> {
        self.with(py, |state, index| {
            // Values in a merged range's hidden cells would be invisible (Excel
            // and openpyxl refuse such writes too).
            let mut merges = state.current_merges(index).map_err(read_err)?;
            {
                let edits = &state.sheets[index].edits;
                merges.retain(|m| !edits.merges_remove.contains(m));
                merges.extend(edits.merges_add.iter().copied());
            }
            for (row, col, _) in &items {
                if let Some(m) = merges.iter().find(|m| *row >= m.0 && *row <= m.2 && *col >= m.1 && *col <= m.3 && (*row, *col) != (m.0, m.1)) {
                    return Err(param_err(format!(
                        "cell {} is inside the merged range {}; write to its top-left cell {} or unmerge it first",
                        range_name((*row, *col, *row, *col), false),
                        range_name(*m, false),
                        range_name((m.0, m.1, m.0, m.1), false)
                    )));
                }
            }
            let edits = &mut state.sheets[index].edits;
            for (row, col, mut prepared) in items {
                check_cell(row, col)?;
                if let EditValue::Formula(f) = &mut prepared.value
                    && let Some(r) = &mut f.array_ref
                    && r.is_empty() {
                        *r = range_name((row, col, row, col), false);
                    }
                if let Some(link) = prepared.link.take() {
                    edits.hyperlinks.retain(|((r, c), _)| (*r, *c) != (row, col));
                    edits.hyperlinks.push(((row, col), link));
                }
                edits.set_cell(row, col, CellEdit::Set { value: prepared.value, style: prepared.style });
            }
            Ok(())
        })
    }

    fn write_grid(&self, py: Python<'_>, row: u32, col: u16, data: &Bound<'_, PyAny>, format: Option<&Arc<FormatSpec>>, rows_outer: bool) -> PyResult<()> {
        let is_1904 = self.is_1904(py)?;
        let mut items = Vec::new();
        for (i, line) in data.try_iter()?.enumerate() {
            for (j, value) in line?.try_iter()?.enumerate() {
                let (r, c) = if rows_outer { (i, j) } else { (j, i) };
                let r = u32::try_from(row as usize + r).map_err(|_| crate::common::errors::range_err("data extends past the last row"))?;
                let c = u16::try_from(col as usize + c).map_err(|_| crate::common::errors::range_err("data extends past the last column"))?;
                items.push((r, c, prepare(&value?, format, is_1904)?));
            }
        }
        self.apply(py, items)
    }

    fn write_line(&self, py: Python<'_>, row: u32, col: u16, data: &Bound<'_, PyAny>, format: Option<&Arc<FormatSpec>>, across: bool) -> PyResult<()> {
        let is_1904 = self.is_1904(py)?;
        let mut items = Vec::new();
        for (i, value) in data.try_iter()?.enumerate() {
            let (r, c) = if across { (row as usize, col as usize + i) } else { (row as usize + i, col as usize) };
            let r = u32::try_from(r).map_err(|_| crate::common::errors::range_err("data extends past the last row"))?;
            let c = u16::try_from(c).map_err(|_| crate::common::errors::range_err("data extends past the last column"))?;
            items.push((r, c, prepare(&value?, format, is_1904)?));
        }
        self.apply(py, items)
    }

    /// The current value of a cell (edits first, then the original file).
    /// The edit and the original value are borrowed, not copied.
    fn value_at(&self, py: Python<'_>, state: &mut BookState, index: usize, row: u32, col: u16) -> PyResult<Py<PyAny>> {
        // Load the original values so they can be borrowed alongside the edits;
        // if they can't be read the original cell reads as empty.
        let _ = state.values(index);
        let state = &*state;
        let original = || match state.original_value(index, row, col) {
            Some(d) => data_to_py(py, d),
            None => Ok(py.None()),
        };
        match state.sheets[index].edits.cells.get(&(row, col)) {
            Some(edit) => edit_to_py(py, edit, state.is_1904, original),
            None => original(),
        }
    }

    fn bounds(state: &mut BookState, index: usize) -> PyResult<Option<CellRange>> {
        let values = state.values(index).map_err(read_err)?;
        let original = match (values.start(), values.end()) {
            (Some((r1, c1)), Some((r2, c2))) if !values.is_empty() => Some((r1, c1 as u16, r2, c2 as u16)),
            _ => None,
        };
        let mut bounds = original.and_then(|range| state.current_range(index, range));
        for ((r, c), edit) in &state.sheets[index].edits.cells {
            if matches!(edit, CellEdit::Clear) {
                continue;
            }
            bounds = Some(match bounds {
                None => (*r, *c, *r, *c),
                Some((r1, c1, r2, c2)) => (r1.min(*r), c1.min(*c), r2.max(*r), c2.max(*c)),
            });
        }
        Ok(bounds)
    }

    /// Render objects with the writer (same arguments as the writer's method)
    /// into a scratch workbook, then record them to be grafted onto this sheet.
    fn graft(&self, py: Python<'_>, method: &str, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (target, others) = {
            let state = lock(py, &self.book)?;
            let index = state.index_of(self.id).map_err(edit_err)?;
            if !state.sheets[index].is_worksheet {
                return Err(edit_err("objects can only be added to worksheets"));
            }
            let target = state.sheets[index].name.clone();
            let others: Vec<String> = state.sheets.iter().filter(|s| s.id != self.id).map(|s| s.name.clone()).collect();
            (target, others)
        };
        let module = py.import("pyrust_xlsxwriter._pyrust_xlsxwriter")?;
        let scratch = module.getattr("Workbook")?.call0()?;
        let sheet = scratch.call_method1("add_worksheet", (target,))?;
        for name in others {
            scratch.call_method1("add_worksheet", (name,))?;
        }
        sheet.call_method(method, args, kwargs)?;
        // Render straight into a Rust buffer (no round trip through Python bytes).
        let data = scratch.cast::<Workbook>()?.borrow_mut().with_sheets(py, |wb| wb.save_to_buffer())?;
        self.with(py, |state, index| state.record_graft(index, data).map_err(edit_err))
    }

    fn structural(&self, py: Python<'_>, rows: bool, insert: bool, at: i64, count: i64) -> PyResult<()> {
        use crate::common::cellref::{MAX_COLS, MAX_ROWS};
        use crate::editor::formula::StructuralEdit;
        let limit = if rows { MAX_ROWS as i64 } else { MAX_COLS as i64 };
        if count < 1 {
            return Err(param_err("count must be at least 1"));
        }
        if at >= limit {
            return Err(crate::common::errors::range_err(format!("{} {at} is outside the sheet", if rows { "row" } else { "column" })));
        }
        self.with(py, |state, index| {
            if insert
                && let Some((_, _, last_row, last_col)) = Self::bounds(state, index)? {
                    let last = if rows { last_row as i64 } else { last_col as i64 };
                    if last >= at && last + count >= limit {
                        return Err(edit_err("inserting would push cells off the end of the sheet"));
                    }
                }
            let sheet = state.sheets[index].name.clone();
            let count_rows = count.min(limit) as u32;
            let edit = match (rows, insert) {
                (true, true) => StructuralEdit::InsertRows { sheet, at: at as u32, count: count_rows },
                (true, false) => StructuralEdit::DeleteRows { sheet, at: at as u32, count: count_rows },
                (false, true) => StructuralEdit::InsertCols { sheet, at: at as u16, count: count.min(limit) as u16 },
                (false, false) => StructuralEdit::DeleteCols { sheet, at: at as u16, count: count.min(limit) as u16 },
            };
            state.structural_edit(index, edit).map_err(edit_err)
        })
    }

    /// Style a range; the format is stored once and shared by its cells.
    fn set_style(&self, py: Python<'_>, range: CellRange, format: &FormatSpec, merge: bool) -> PyResult<()> {
        let format = Arc::new(format.clone());
        let op = if merge { StyleOp::Merge(format) } else { StyleOp::Replace(format) };
        self.with(py, |state, index| {
            let edits = &mut state.sheets[index].edits;
            for row in range.0..=range.2 {
                for col in range.1..=range.3 {
                    edits.set_cell(row, col, CellEdit::Style(op.clone()));
                }
            }
            Ok(())
        })
    }
}

fn edit_to_py(
    py: Python<'_>,
    edit: &CellEdit,
    is_1904: bool,
    original: impl FnOnce() -> PyResult<Py<PyAny>>,
) -> PyResult<Py<PyAny>> {
    match edit {
        CellEdit::Clear => Ok(py.None()),
        CellEdit::Style(_) => original(),
        CellEdit::Set { value, style } => Ok(match value {
            EditValue::Blank => py.None(),
            EditValue::Number(n) => match style {
                StyleOp::DefaultNumFormat(code) => serial_to_py(py, *n, code, is_1904)?,
                _ if n.fract() == 0.0 && n.abs() < 9.007_199_254_740_992e15 => (*n as i64).into_py_any(py)?,
                _ => n.into_py_any(py)?,
            },
            EditValue::Str(s) => s.into_py_any(py)?,
            EditValue::Bool(b) => b.into_py_any(py)?,
            EditValue::Error(e) => e.into_py_any(py)?,
            EditValue::Formula(f) => match &f.result {
                Some(r) => match r.parse::<f64>() {
                    Ok(n) => n.into_py_any(py)?,
                    Err(_) => r.into_py_any(py)?,
                },
                None => py.None(),
            },
        }),
    }
}

/// The format argument, copied once into an `Arc` that all its cells share.
fn fmt_spec(format: &Option<PyRef<'_, Format>>) -> Option<Arc<FormatSpec>> {
    format.as_ref().map(|f| Arc::new(f.spec.clone()))
}

#[pymethods]
impl EditWorksheet {
    /// The worksheet name; assigning renames it (defined names are updated).
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<String> {
        let state = lock(py, &self.book)?;
        let index = state.index_of(self.id).map_err(edit_err)?;
        Ok(state.sheets[index].name.clone())
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        let mut state = lock(py, &self.book)?;
        let index = state.index_of(self.id).map_err(edit_err)?;
        state.rename(index, name).map_err(|e| new_err(ErrorKind::SheetName, e))
    }

    /// Position in the tab order.
    #[getter]
    fn index(&self, py: Python<'_>) -> PyResult<usize> {
        lock(py, &self.book)?.index_of(self.id).map_err(edit_err)
    }

    /// "visible", "hidden" or "very_hidden".
    #[getter]
    fn state(&self, py: Python<'_>) -> PyResult<&'static str> {
        let state = lock(py, &self.book)?;
        let index = state.index_of(self.id).map_err(edit_err)?;
        Ok(state.visibility(index).as_str())
    }

    #[setter]
    fn set_state(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let visibility = match value {
            "visible" => Visibility::Visible,
            "hidden" => Visibility::Hidden,
            "very_hidden" | "veryHidden" => Visibility::VeryHidden,
            other => return Err(param_err(format!("state must be 'visible', 'hidden' or 'very_hidden', not {other:?}"))),
        };
        let mut state = lock(py, &self.book)?;
        let index = state.index_of(self.id).map_err(edit_err)?;
        state.set_visibility(index, visibility).map_err(edit_err)
    }

    #[getter]
    fn hidden(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.state(py)? != "visible")
    }

    #[setter]
    fn set_hidden(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_state(py, if value { "hidden" } else { "visible" })
    }

    /// Set the tab colour (None removes it).
    #[setter]
    fn set_tab_color(&self, py: Python<'_>, color: Option<Color>) -> PyResult<()> {
        self.with(py, |state, index| {
            state.sheets[index].edits.tab_color = Some(color);
            Ok(())
        })
    }

    /// Set the zoom percentage (10-400).
    #[setter]
    fn set_zoom(&self, py: Python<'_>, zoom: u16) -> PyResult<()> {
        if !(10..=400).contains(&zoom) {
            return Err(param_err(format!("zoom {zoom} is outside 10..=400")));
        }
        self.with(py, |state, index| {
            state.sheets[index].edits.zoom = Some(zoom);
            Ok(())
        })
    }

    // ------------------------------------------------------------------
    // Reading
    // ------------------------------------------------------------------

    /// Used range (first_row, first_col, last_row, last_col), or None if empty.
    #[getter]
    fn dimensions(&self, py: Python<'_>) -> PyResult<Option<CellRange>> {
        self.with(py, Self::bounds)
    }

    /// Number of rows up to the last used one.
    #[getter]
    fn max_row(&self, py: Python<'_>) -> PyResult<u32> {
        Ok(self.dimensions(py)?.map_or(0, |(_, _, r, _)| r + 1))
    }

    /// Number of columns up to the last used one.
    #[getter]
    fn max_column(&self, py: Python<'_>) -> PyResult<u16> {
        Ok(self.dimensions(py)?.map_or(0, |(_, _, _, c)| c + 1))
    }

    /// The value of a cell (formulas give their cached result).
    #[pyo3(signature = (*args))]
    fn cell(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<Py<PyAny>> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.with(py, |state, index| self.value_at(py, state, index, row, col))
    }

    /// All values from A1 to the last used cell, as a list of rows.
    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        self.with(py, |state, index| {
            let rows = PyList::empty(py);
            let Some((_, _, last_row, last_col)) = Self::bounds(state, index)? else { return Ok(rows) };
            for r in 0..=last_row {
                let row = PyList::empty(py);
                for c in 0..=last_col {
                    row.append(self.value_at(py, state, index, r, c)?)?;
                }
                rows.append(row)?;
            }
            Ok(rows)
        })
    }

    /// Rows of values as tuples, optionally limited to a block (zero-based,
    /// inclusive). Rows are read one at a time as the iterator advances, so
    /// edits made while iterating show up in later rows.
    #[pyo3(signature = (min_row=0, max_row=None, min_col=0, max_col=None))]
    fn iter_rows(&self, py: Python<'_>, min_row: u32, max_row: Option<u32>, min_col: u16, max_col: Option<u16>) -> PyResult<EditRowIterator> {
        let bounds = self.with(py, Self::bounds)?;
        let last = match (max_row.or(bounds.map(|b| b.2)), max_col.or(bounds.map(|b| b.3))) {
            (Some(row), Some(col)) if row >= min_row => Some((row, col)),
            _ => None,
        };
        Ok(EditRowIterator { sheet: EditWorksheet { book: self.book.clone(), id: self.id }, next: min_row, last, min_col })
    }

    /// The formula of a cell ("=..."), or None.
    #[pyo3(signature = (*args))]
    fn formula(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<Option<String>> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.with(py, |state, index| {
            match state.sheets[index].edits.cells.get(&(row, col)) {
                Some(CellEdit::Set { value: EditValue::Formula(f), .. }) => return Ok(Some(format!("={}", f.text))),
                Some(CellEdit::Set { .. }) | Some(CellEdit::Clear) => return Ok(None),
                _ => {}
            }
            let Some(original) = state.to_original(index, row, col) else { return Ok(None) };
            let text = state.formulas(index).map_err(read_err)?.get(&original).cloned();
            Ok(text.map(|f| format!("={}", state.replay_formula(index, &f))))
        })
    }

    /// The cell's format (after pending edits), or None for the default style.
    #[pyo3(signature = (*args))]
    fn style(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<Option<Format>> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.with(py, |state, index| {
            let original = match state.to_original(index, row, col) {
                Some(pos) => state.cell_styles(index).map_err(read_err)?.get(&pos).copied(),
                None => None,
            };
            let base = original.map(|xf| state.styles.spec(xf)).unwrap_or_default();
            let op = match state.sheets[index].edits.cells.get(&(row, col)) {
                Some(CellEdit::Set { style, .. }) | Some(CellEdit::Style(style)) => Some(style.clone()),
                Some(CellEdit::Clear) => return Ok(None),
                None => None,
            };
            let spec = match op {
                None | Some(StyleOp::Keep) => base,
                Some(StyleOp::Replace(spec)) => Arc::unwrap_or_clone(spec),
                Some(StyleOp::Merge(spec)) => base.overlay(&spec),
                Some(StyleOp::DefaultNumFormat(code)) => {
                    if original.is_some_and(|xf| state.styles.is_date_xf(xf)) {
                        base
                    } else {
                        base.overlay(&FormatSpec { num_format: Some(code.to_string()), ..Default::default() })
                    }
                }
            };
            Ok(if spec == FormatSpec::default() { None } else { Some(Format::from_spec(spec)) })
        })
    }

    /// Merged ranges after pending edits.
    #[getter]
    fn merged_ranges(&self, py: Python<'_>) -> PyResult<Vec<CellRange>> {
        self.with(py, |state, index| {
            let mut merges = state.current_merges(index).map_err(read_err)?;
            let edits = &state.sheets[index].edits;
            merges.retain(|m| !edits.merges_remove.contains(m));
            for m in &edits.merges_add {
                if !merges.contains(m) {
                    merges.push(*m);
                }
            }
            Ok(merges)
        })
    }

    // ------------------------------------------------------------------
    // Writing
    // ------------------------------------------------------------------

    /// Write any value. Without `format` the cell keeps its existing style
    /// (dates get a date format only if the cell doesn't already have one).
    #[pyo3(signature = (*args, **kwargs))]
    fn write(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write", rest, kwargs, &["value", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_one(py, row, col, p.required_any(0)?, fmt_spec(&format).as_ref())
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_number(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_number", rest, kwargs, &["number", "format"], 1)?;
        let number: f64 = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        self.apply(py, vec![(row, col, Prepared { value: EditValue::Number(number), style, link: None })])
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_string(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_string", rest, kwargs, &["string", "format"], 1)?;
        let string: String = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        self.apply(py, vec![(row, col, Prepared { value: EditValue::Str(string), style, link: None })])
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_boolean(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_boolean", rest, kwargs, &["value", "format"], 1)?;
        let value: bool = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        self.apply(py, vec![(row, col, Prepared { value: EditValue::Bool(value), style, link: None })])
    }

    /// Clear a cell's value; with `format`, also set its style.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_blank(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_blank", rest, kwargs, &["format"], 0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(0)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        self.apply(py, vec![(row, col, Prepared { value: EditValue::Blank, style, link: None })])
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_datetime(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_datetime", rest, kwargs, &["datetime", "format"], 1)?;
        let value = p.required_any(0)?;
        if !matches!(classify(value)?, CellValue::DateTime(_) | CellValue::Date(_) | CellValue::Time(_) | CellValue::Duration(_)) {
            return Err(PyTypeError::new_err("write_datetime() expects a datetime, date, time or timedelta"));
        }
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_one(py, row, col, value, fmt_spec(&format).as_ref())
    }

    /// Write a formula; `value` is an optional cached result.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_formula(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let formula = formula_text(p.required_any(0)?)?;
        let result = match p.raw(2) {
            Some(v) => Some(v.str()?.to_string()),
            None => formula.1,
        };
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        self.apply(py, vec![(row, col, Prepared { value: EditValue::formula(formula.0, result, None), style, link: None })])
    }

    /// Write an array formula over a range (stored in its top-left cell).
    #[pyo3(signature = (*args, **kwargs))]
    fn write_array_formula(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let p = Params::new("write_array_formula", rest, kwargs, &["formula", "format", "value"], 1)?;
        let formula = formula_text(p.required_any(0)?)?;
        let text = formula.0.trim_start_matches('{').trim_end_matches('}').trim_start_matches('=').to_string();
        let result = match p.raw(2) {
            Some(v) => Some(v.str()?.to_string()),
            None => formula.1,
        };
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace).unwrap_or(StyleOp::Keep);
        let mut items = vec![(range.0, range.1, Prepared {
            value: EditValue::formula(text, result, Some(range_name(range, false))),
            style: style.clone(),
            link: None,
        })];
        for r in range.0..=range.2 {
            for c in range.1..=range.3 {
                if (r, c) != (range.0, range.1) {
                    items.push((r, c, Prepared { value: EditValue::Blank, style: style.clone(), link: None }));
                }
            }
        }
        self.apply(py, items)
    }

    /// Write a hyperlink: `write_url("A1", "https://...", text=None, tip=None, format=None)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_url(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_url", rest, kwargs, &["url", "text", "tip", "format"], 1)?;
        let arg = p.required_any(0)?;
        let mut url: Url = match arg.cast::<Url>() {
            Ok(u) => u.get().clone(),
            Err(_) => Url { link: arg.extract()?, text: None, tip: None },
        };
        if let Some(text) = p.extract::<String>(1)? {
            url.text = Some(text);
        }
        if let Some(tip) = p.extract::<String>(2)? {
            url.tip = Some(tip);
        }
        let format: Option<PyRef<'_, Format>> = p.extract(3)?;
        let value = Bound::new(py, url)?.into_any();
        self.write_one(py, row, col, &value, fmt_spec(&format).as_ref())
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_row(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_row", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_line(py, row, col, p.required_any(0)?, fmt_spec(&format).as_ref(), true)
    }

    #[pyo3(signature = (*args, **kwargs))]
    fn write_column(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_column", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_line(py, row, col, p.required_any(0)?, fmt_spec(&format).as_ref(), false)
    }

    /// Write an iterable of rows starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_rows(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_rows", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_grid(py, row, col, p.required_any(0)?, fmt_spec(&format).as_ref(), true)
    }

    /// Write an iterable of columns starting at the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn write_columns(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("write_columns", rest, kwargs, &["data", "format"], 1)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        self.write_grid(py, row, col, p.required_any(0)?, fmt_spec(&format).as_ref(), false)
    }

    /// Set a cell's format, keeping its value. `merge=True` lays the format's
    /// set properties over the existing style instead of replacing it.
    #[pyo3(signature = (*args, merge=false))]
    fn set_cell_format(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, merge: bool) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("set_cell_format", rest, None, &["format"], 1)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        self.set_style(py, (row, col, row, col), &format.spec, merge)
    }

    /// Set the format of every cell in a range (see `set_cell_format`).
    #[pyo3(signature = (*args, merge=false))]
    fn set_range_format(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, merge: bool) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        let p = Params::new("set_range_format", rest, None, &["format"], 1)?;
        let format: PyRef<'_, Format> = p.required(0)?;
        self.set_style(py, range, &format.spec, merge)
    }

    /// Remove a cell entirely (value and format).
    #[pyo3(signature = (*args))]
    fn clear_cell(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.with(py, |state, index| {
            state.sheets[index].edits.set_cell(row, col, CellEdit::Clear);
            Ok(())
        })
    }

    /// Clear the values in a range, keeping formats (like Excel's Delete key).
    #[pyo3(signature = (*args))]
    fn clear_contents(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let (range, _) = split_range(args.as_slice())?;
        let cells = (range.2 - range.0 + 1) as u64 * (range.3 - range.1 + 1) as u64;
        if cells > 5_000_000 {
            return Err(param_err("clear_contents() range is too large (over 5,000,000 cells)"));
        }
        self.with(py, |state, index| {
            let edits = &mut state.sheets[index].edits;
            for r in range.0..=range.2 {
                for c in range.1..=range.3 {
                    edits.set_cell(r, c, CellEdit::Set { value: EditValue::Blank, style: StyleOp::Keep });
                }
            }
            Ok(())
        })
    }

    /// Merge a range, optionally writing a value (and format) into it.
    #[pyo3(signature = (*args, **kwargs))]
    fn merge_range(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let (range, rest) = split_range(args.as_slice())?;
        if range.0 == range.2 && range.1 == range.3 {
            return Err(param_err("a merged range must span more than one cell"));
        }
        let p = Params::new("merge_range", rest, kwargs, &["value", "format"], 0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let spec = fmt_spec(&format);
        let first = match p.raw(0) {
            Some(value) => Some(prepare(value, spec.as_ref(), self.is_1904(py)?)?),
            None => None,
        };
        self.with(py, |state, index| {
            let mut existing = state.current_merges(index).map_err(read_err)?;
            let edits = &state.sheets[index].edits;
            existing.retain(|m| !edits.merges_remove.contains(m));
            existing.extend(edits.merges_add.iter().copied());
            if let Some(other) = existing.iter().find(|m| overlaps(m, &range)) {
                return Err(new_err(
                    ErrorKind::Overlap,
                    format!("merged range {} overlaps existing merged range {}", range_name(range, false), range_name(*other, false)),
                ));
            }
            let edits = &mut state.sheets[index].edits;
            edits.merges_remove.retain(|m| *m != range);
            edits.merges_add.push(range);
            if let Some(first) = first {
                edits.set_cell(range.0, range.1, CellEdit::Set { value: first.value, style: first.style.clone() });
                if let Some(spec) = &spec {
                    for r in range.0..=range.2 {
                        for c in range.1..=range.3 {
                            if (r, c) != (range.0, range.1) {
                                edits.set_cell(r, c, CellEdit::Style(StyleOp::Replace(spec.clone())));
                            }
                        }
                    }
                }
            }
            Ok(())
        })
    }

    /// Remove a merged range (values are kept).
    #[pyo3(signature = (*args))]
    fn unmerge_range(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let (range, _) = split_range(args.as_slice())?;
        self.with(py, |state, index| {
            let original = state.current_merges(index).map_err(read_err)?;
            let edits = &mut state.sheets[index].edits;
            let had_new = edits.merges_add.contains(&range);
            edits.merges_add.retain(|m| *m != range);
            if original.contains(&range) {
                edits.merges_remove.push(range);
            } else if !had_new {
                return Err(param_err(format!("{} is not a merged range", range_name(range, false))));
            }
            Ok(())
        })
    }

    /// Set the width of a column or span ("B:D"), optionally with a column format.
    #[pyo3(signature = (*args, **kwargs))]
    fn set_column_width(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((first, last), rest) = split_cols(args.as_slice())?;
        let p = Params::new("set_column_width", rest, kwargs, &["width", "format"], 1)?;
        let width: f64 = p.required(0)?;
        if !(0.0..=255.0).contains(&width) {
            return Err(param_err(format!("column width {width} is outside 0..=255")));
        }
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        let style = fmt_spec(&format).map(StyleOp::Replace);
        self.with(py, |state, index| {
            state.sheets[index].edits.cols.push(ColEdit { first, last, width: Some(width), hidden: None, style });
            Ok(())
        })
    }

    /// Hide (or unhide) a column or span.
    #[pyo3(signature = (*args, hidden=true))]
    fn set_column_hidden(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, hidden: bool) -> PyResult<()> {
        let ((first, last), _) = split_cols(args.as_slice())?;
        self.with(py, |state, index| {
            state.sheets[index].edits.cols.push(ColEdit { first, last, width: None, hidden: Some(hidden), style: None });
            Ok(())
        })
    }

    /// Set a row's height in points (None restores the default height).
    #[pyo3(signature = (*args))]
    fn set_row_height(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let (row, rest) = split_row(args.as_slice())?;
        let p = Params::new("set_row_height", rest, None, &["height"], 1)?;
        let height: Option<f64> = match p.raw_any(0) {
            Some(v) if v.is_none() => None,
            _ => Some(p.required(0)?),
        };
        if height.is_some_and(|h| !(0.0..=409.0).contains(&h)) {
            return Err(param_err("row height must be between 0 and 409 points"));
        }
        self.with(py, |state, index| {
            let edit = state.sheets[index].edits.rows.entry(row).or_insert_with(RowEdit::default);
            edit.height = Some(height);
            Ok(())
        })
    }

    /// Hide (or unhide) a row.
    #[pyo3(signature = (row, hidden=true))]
    fn set_row_hidden(&self, py: Python<'_>, row: u32, hidden: bool) -> PyResult<()> {
        check_cell(row, 0)?;
        self.with(py, |state, index| {
            let edit = state.sheets[index].edits.rows.entry(row).or_insert_with(RowEdit::default);
            edit.hidden = Some(hidden);
            Ok(())
        })
    }

    /// Insert `count` empty rows before `row`. Formulas, merged ranges,
    /// conditional formats, validations, tables, charts, images and defined
    /// names that refer to moved cells are updated.
    #[pyo3(signature = (row, count=1))]
    fn insert_rows(&self, py: Python<'_>, row: u32, count: u32) -> PyResult<()> {
        self.structural(py, true, true, row as i64, count as i64)
    }

    /// Delete `count` rows starting at `row`. References to deleted cells become `#REF!`.
    #[pyo3(signature = (row, count=1))]
    fn delete_rows(&self, py: Python<'_>, row: u32, count: u32) -> PyResult<()> {
        self.structural(py, true, false, row as i64, count as i64)
    }

    /// Insert `count` empty columns before `col` (an index or letters such as "C").
    #[pyo3(signature = (col, count=1))]
    fn insert_cols(&self, py: Python<'_>, col: &Bound<'_, PyAny>, count: u16) -> PyResult<()> {
        let ((col, _), _) = split_cols(std::slice::from_ref(col))?;
        self.structural(py, false, true, col as i64, count as i64)
    }

    /// Delete `count` columns starting at `col` (an index or letters such as "C").
    #[pyo3(signature = (col, count=1))]
    fn delete_cols(&self, py: Python<'_>, col: &Bound<'_, PyAny>, count: u16) -> PyResult<()> {
        let ((col, _), _) = split_cols(std::slice::from_ref(col))?;
        self.structural(py, false, false, col as i64, count as i64)
    }

    /// Add a conditional format (same arguments as `Worksheet.add_conditional_format`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_conditional_format(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "add_conditional_format", args, kwargs)
    }

    /// Add data validation (same arguments as `Worksheet.add_data_validation`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_data_validation(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "add_data_validation", args, kwargs)
    }

    /// Add a table (same arguments as `Worksheet.add_table`); header cells are written too.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_table(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "add_table", args, kwargs)
    }

    /// Add a sparkline (same arguments as `Worksheet.add_sparkline`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_sparkline(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "add_sparkline", args, kwargs)
    }

    /// Add a sparkline group (same arguments as `Worksheet.add_sparkline_group`).
    #[pyo3(signature = (*args, **kwargs))]
    fn add_sparkline_group(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "add_sparkline_group", args, kwargs)
    }

    /// Insert an image (same arguments as `Worksheet.insert_image`).
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_image(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "insert_image", args, kwargs)
    }

    /// Insert a chart (same arguments as `Worksheet.insert_chart`).
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_chart(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "insert_chart", args, kwargs)
    }

    /// Insert a shape or text box (same arguments as `Worksheet.insert_shape`).
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_shape(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "insert_shape", args, kwargs)
    }

    /// Add a note (same arguments as `Worksheet.insert_note`). Sheets that already have notes aren't supported yet.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_note(&self, py: Python<'_>, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        self.graft(py, "insert_note", args, kwargs)
    }

    /// Freeze the rows above and columns left of a cell.
    #[pyo3(signature = (*args))]
    fn freeze_panes(&self, py: Python<'_>, args: &Bound<'_, PyTuple>) -> PyResult<()> {
        let ((row, col), _) = split_cell(args.as_slice())?;
        self.with(py, |state, index| {
            state.sheets[index].edits.freeze = Some(Some((row, col)));
            Ok(())
        })
    }

    /// Remove frozen panes.
    fn unfreeze_panes(&self, py: Python<'_>) -> PyResult<()> {
        self.with(py, |state, index| {
            state.sheets[index].edits.freeze = Some(None);
            Ok(())
        })
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!("<EditWorksheet {:?}>", self.name(py)?))
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other.cast::<EditWorksheet>().is_ok_and(|o| Arc::ptr_eq(&o.get().book, &self.book) && o.get().id == self.id)
    }

    fn __hash__(&self) -> u64 {
        self.id ^ (Arc::as_ptr(&self.book) as u64)
    }
}

/// Iterator returned by `EditWorksheet.iter_rows()`: reads the current values
/// of one row per step, so the rows are never all in memory at once.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct EditRowIterator {
    sheet: EditWorksheet,
    next: u32,
    /// The last row and column to read; None once the iterator is exhausted.
    last: Option<(u32, u16)>,
    min_col: u16,
}

#[pymethods]
impl EditRowIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        let Some((last_row, last_col)) = self.last else { return Ok(None) };
        let row = self.next;
        if row == last_row {
            self.last = None;
        } else {
            self.next += 1;
        }
        let (sheet, min_col) = (&self.sheet, self.min_col);
        let values = sheet.with(py, |state, index| {
            (min_col..=last_col).map(|col| sheet.value_at(py, state, index, row, col)).collect::<PyResult<Vec<_>>>()
        })?;
        Ok(Some(PyTuple::new(py, values)?))
    }
}

fn overlaps(a: &CellRange, b: &CellRange) -> bool {
    !(a.2 < b.0 || a.0 > b.2 || a.3 < b.1 || a.1 > b.3)
}

/// The text (without '=') and cached result of a formula argument.
fn formula_text(value: &Bound<'_, PyAny>) -> PyResult<(String, Option<String>)> {
    if let Ok(f) = value.cast::<crate::common::value::Formula>() {
        let f = f.get();
        return Ok((f.formula.trim_start_matches('=').to_string(), f.result.clone()));
    }
    Ok((strip_equals(value.extract()?), None))
}

/// Drop a formula's leading '=' in place (no new string).
fn strip_equals(mut text: String) -> String {
    let n = text.len() - text.trim_start_matches('=').len();
    text.drain(..n);
    text
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<EditWorkbook>()?;
    m.add_class::<EditWorksheet>()?;
    Ok(())
}
