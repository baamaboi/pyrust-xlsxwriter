//! `open_workbook()`, `ReadWorkbook`, `SheetInfo` and `TableInfo`.

use std::path::PathBuf;
use std::sync::{Arc, Mutex, PoisonError};

use calamine::WorkbookProperties;
use pyo3::exceptions::{PyIndexError, PyKeyError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyByteArray, PyBytes, PyDict, PyList, PyMemoryView, PyString};

use crate::reader::convert::{IsoValue, parse_iso_datetime};
use crate::reader::shared::{Shared, TableMeta};
use crate::common::bytes::SharedBytes;
use crate::reader::source::Origin;
use crate::reader::worksheet::{ReadWorksheet, records_from_range, rows_to_list};

/// In-memory workbook bytes: `bytes` are read in place, `bytearray` is copied once.
fn bytes_origin(data: &Bound<'_, PyAny>) -> PyResult<Origin> {
    Ok(Origin::Bytes(SharedBytes::from_py(data)?))
}

/// Turn the `source` argument into a path or in-memory bytes.
fn origin_from_py(source: &Bound<'_, PyAny>) -> PyResult<Origin> {
    if let Ok(path) = source.cast::<PyString>() {
        return Ok(Origin::Path(PathBuf::from(path.to_str()?)));
    }
    if source.is_instance_of::<PyBytes>() || source.is_instance_of::<PyByteArray>() {
        return bytes_origin(source);
    }
    if source.is_instance_of::<PyMemoryView>() {
        return bytes_origin(&source.call_method0("tobytes")?);
    }
    if source.hasattr("__fspath__")? {
        return Ok(Origin::Path(source.extract::<PathBuf>()?));
    }
    if source.hasattr("read")? {
        let data = source.call_method0("read")?;
        if data.is_instance_of::<PyBytes>() || data.is_instance_of::<PyByteArray>() {
            return bytes_origin(&data);
        }
        return Err(PyTypeError::new_err(format!(
            "the file object's read() returned {}, not bytes: open the file in binary mode ('rb')",
            data.get_type().name()?
        )));
    }
    Err(PyTypeError::new_err(format!(
        "open_workbook() expects a path (str or os.PathLike), bytes, or a binary file object, not {}",
        source.get_type().name()?
    )))
}

/// Open a spreadsheet for reading.
///
/// `source` is a path (`str` or `os.PathLike`), the file's contents as
/// `bytes`, or a binary file object (anything with `.read()`). The format
/// (xlsx, xlsm, xlsb, xls or ods) is detected from the content. Raises
/// `FileError` if the file cannot be opened and `ReadError` if it is not a
/// readable spreadsheet.
#[pyfunction]
pub fn open_workbook(py: Python<'_>, source: &Bound<'_, PyAny>) -> PyResult<ReadWorkbook> {
    let origin = origin_from_py(source)?;
    let shared = py.detach(move || Shared::open(origin))?;
    Ok(ReadWorkbook::new(shared))
}

/// Name, type and visibility of a sheet in a `ReadWorkbook`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, skip_from_py_object)]
#[derive(Clone)]
pub struct SheetInfo {
    /// The sheet name.
    #[pyo3(get)]
    pub name: String,
    /// The zero-based position in the workbook.
    #[pyo3(get)]
    pub index: usize,
    /// "worksheet", "chartsheet", "dialogsheet", "macrosheet" or "vba".
    #[pyo3(get)]
    pub kind: &'static str,
    /// "visible", "hidden" or "very_hidden".
    #[pyo3(get)]
    pub visibility: &'static str,
}

#[pymethods]
impl SheetInfo {
    fn __repr__(&self) -> String {
        format!(
            "SheetInfo(name={:?}, index={}, kind={:?}, visibility={:?})",
            self.name, self.index, self.kind, self.visibility
        )
    }
}

/// A table (ListObject) in an xlsx workbook.
#[pyclass(module = "pyrust_xlsxwriter", frozen, skip_from_py_object)]
pub struct TableInfo {
    shared: Arc<Shared>,
    meta: Arc<TableMeta>,
}

impl TableInfo {
    fn data_rows(&self) -> (u32, u32) {
        let (r1, _, r2, _) = self.meta.range;
        let first = r1.saturating_add(self.meta.header_rows);
        let last = r2.saturating_sub(self.meta.totals_rows);
        (first, last)
    }
}

#[pymethods]
impl TableInfo {
    /// The table name.
    #[getter]
    fn name(&self) -> &str {
        &self.meta.name
    }

    /// The name of the worksheet holding the table.
    #[getter]
    fn sheet_name(&self) -> &str {
        &self.shared.sheets[self.meta.sheet_index].name
    }

    /// The column names.
    #[getter]
    fn columns(&self) -> Vec<String> {
        self.meta.columns.clone()
    }

    /// `(first_row, first_col, last_row, last_col)` of the whole table,
    /// including the header row and any totals row.
    #[getter]
    fn range(&self) -> (u32, u32, u32, u32) {
        self.meta.range
    }

    /// True if the table has a header row.
    #[getter]
    fn has_header(&self) -> bool {
        self.meta.header_rows > 0
    }

    /// True if the table has a totals row.
    #[getter]
    fn has_totals(&self) -> bool {
        self.meta.totals_rows > 0
    }

    /// The data rows (without the header and totals rows) as lists of values.
    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let data = self.shared.sheet(py, self.meta.sheet_index)?;
        let (first, last) = self.data_rows();
        let (_, c1, _, c2) = self.meta.range;
        if first > last {
            return Ok(PyList::empty(py));
        }
        rows_to_list(py, &data.values, first, last, c1, c2, self.shared.is_1904)
    }

    /// The data rows as dicts keyed by column name.
    fn records<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let data = self.shared.sheet(py, self.meta.sheet_index)?;
        let (first, last) = self.data_rows();
        let (_, c1, _, c2) = self.meta.range;
        if first > last {
            return Ok(PyList::empty(py));
        }
        let headers: Vec<Bound<'py, PyAny>> = self
            .meta
            .columns
            .iter()
            .map(|c| PyString::new(py, c).into_any())
            .collect();
        records_from_range(py, &data.values, first, last, c1, c2, self.shared.is_1904, &headers, false)
    }

    fn __repr__(&self) -> String {
        let (r1, c1, r2, c2) = self.meta.range;
        format!(
            "<TableInfo {:?} sheet={:?} range={}>",
            self.meta.name,
            self.sheet_name(),
            crate::common::cellref::range_name((r1, c1 as u16, r2, c2 as u16), false)
        )
    }
}

/// A spreadsheet opened for reading with `open_workbook()`.
///
/// Sheets are parsed on first access and cached. Use it as a context manager
/// (or call `close()`) to release the file; sheets that were already loaded
/// stay readable after closing.
///
/// ```python
/// with open_workbook("data.xlsx") as book:
///     rows = book.worksheet(0).values()
/// ```
#[pyclass(module = "pyrust_xlsxwriter", frozen)]
pub struct ReadWorkbook {
    shared: Arc<Shared>,
    /// One worksheet object per sheet, so `book[0] is book[0]` while open;
    /// emptied by `close()` so their parsed data can be freed.
    worksheets: Vec<Mutex<Option<Py<ReadWorksheet>>>>,
}

impl ReadWorkbook {
    fn new(shared: Shared) -> Self {
        let worksheets = shared.sheets.iter().map(|_| Mutex::new(None)).collect();
        ReadWorkbook {
            shared: Arc::new(shared),
            worksheets,
        }
    }

    fn find_index(&self, key: &Bound<'_, PyAny>) -> PyResult<usize> {
        if let Ok(name) = key.cast::<PyString>() {
            let name = name.to_str()?;
            let sheets = &self.shared.sheets;
            return sheets
                .iter()
                .position(|s| s.name == name)
                .or_else(|| {
                    sheets
                        .iter()
                        .position(|s| s.name.to_lowercase() == name.to_lowercase())
                })
                .ok_or_else(|| {
                    PyKeyError::new_err(format!(
                        "no sheet named {name:?}; the sheets are {:?}",
                        sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>()
                    ))
                });
        }
        if key.is_instance_of::<pyo3::types::PyBool>() {
            return Err(PyTypeError::new_err(
                "sheet key must be a name (str) or an index (int)",
            ));
        }
        let index: isize = key.extract().map_err(|_| {
            PyTypeError::new_err("sheet key must be a name (str) or an index (int)")
        })?;
        let len = self.shared.sheets.len() as isize;
        let resolved = if index < 0 { index + len } else { index };
        if resolved < 0 || resolved >= len {
            return Err(PyIndexError::new_err(format!(
                "sheet index {index} out of range (the workbook has {len} sheet(s))"
            )));
        }
        Ok(resolved as usize)
    }

    pub(crate) fn sheet_at(&self, py: Python<'_>, index: usize) -> PyResult<Py<ReadWorksheet>> {
        let slot = || self.worksheets[index].lock().unwrap_or_else(PoisonError::into_inner);
        if let Some(ws) = slot().as_ref() {
            return Ok(ws.clone_ref(py));
        }
        // Loaded without holding the lock: loading releases the GIL.
        let data = self.shared.sheet(py, index)?;
        let ws = Py::new(py, ReadWorksheet::new(self.shared.clone(), index, data))?;
        if self.shared.is_closed() {
            return Ok(ws);
        }
        Ok(slot().get_or_insert(ws).clone_ref(py))
    }
}

fn put_text(dict: &Bound<'_, PyDict>, key: &str, value: &Option<String>) -> PyResult<()> {
    if let Some(value) = value.as_deref().filter(|v| !v.is_empty()) {
        dict.set_item(key, value)?;
    }
    Ok(())
}

fn put_date(dict: &Bound<'_, PyDict>, key: &str, value: &Option<String>) -> PyResult<()> {
    let Some(text) = value.as_deref().filter(|v| !v.is_empty()) else {
        return Ok(());
    };
    match parse_iso_datetime(text) {
        Some(IsoValue::DateTime(dt)) => dict.set_item(key, dt),
        _ => dict.set_item(key, text),
    }
}

fn properties_dict<'py>(py: Python<'py>, p: &WorkbookProperties) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);
    put_text(&dict, "title", &p.title)?;
    put_text(&dict, "subject", &p.subject)?;
    put_text(&dict, "author", &p.creator)?;
    put_text(&dict, "manager", &p.manager)?;
    put_text(&dict, "company", &p.company)?;
    put_text(&dict, "category", &p.category)?;
    put_text(&dict, "keywords", &p.keywords)?;
    put_text(&dict, "comment", &p.description)?;
    put_text(&dict, "status", &p.content_status)?;
    put_date(&dict, "created", &p.created)?;
    put_date(&dict, "modified", &p.modified)?;
    put_text(&dict, "last_modified_by", &p.last_modified_by)?;
    put_text(&dict, "revision", &p.revision)?;
    put_text(&dict, "version", &p.version)?;
    put_text(&dict, "application", &p.application)?;
    put_text(&dict, "app_version", &p.app_version)?;
    put_text(&dict, "template", &p.template)?;
    if !p.custom_properties.is_empty() {
        let custom = PyDict::new(py);
        let mut entries: Vec<_> = p.custom_properties.iter().collect();
        entries.sort();
        for (key, value) in entries {
            custom.set_item(key, value)?;
        }
        dict.set_item("custom", custom)?;
    }
    Ok(dict)
}

/// Iterator over the sheets of a `ReadWorkbook`.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct SheetIterator {
    book: Py<ReadWorkbook>,
    next: usize,
}

#[pymethods]
impl SheetIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<ReadWorksheet>>> {
        let book = self.book.get();
        if self.next >= book.worksheets.len() {
            return Ok(None);
        }
        let sheet = book.sheet_at(py, self.next)?;
        self.next += 1;
        Ok(Some(sheet))
    }
}

#[pymethods]
impl ReadWorkbook {
    /// Same as `open_workbook(source)`.
    #[new]
    fn py_new(py: Python<'_>, source: &Bound<'_, PyAny>) -> PyResult<Self> {
        open_workbook(py, source)
    }

    /// "xlsx" (also for xlsm), "xlsb", "xls" or "ods".
    #[getter]
    fn format(&self) -> &'static str {
        self.shared.kind.name()
    }

    /// The sheet names in workbook order (chartsheets included).
    #[getter]
    fn sheet_names(&self) -> Vec<String> {
        self.shared.sheets.iter().map(|s| s.name.clone()).collect()
    }

    /// Name, kind and visibility of every sheet.
    #[getter]
    fn sheets(&self) -> Vec<SheetInfo> {
        self.shared
            .sheets
            .iter()
            .enumerate()
            .map(|(index, s)| SheetInfo {
                name: s.name.clone(),
                index,
                kind: s.kind,
                visibility: s.visibility,
            })
            .collect()
    }

    /// `(name, formula)` pairs of the workbook's defined names.
    #[getter]
    fn defined_names(&self) -> Vec<(String, String)> {
        self.shared.defined_names.clone()
    }

    /// Document properties that are set: title, subject, author, manager,
    /// company, category, keywords, comment, status, created, modified
    /// (datetimes), last_modified_by, revision, version, application,
    /// app_version, template and custom (a dict of strings).
    #[getter]
    fn properties<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let props = self.shared.properties(py)?;
        properties_dict(py, props)
    }

    /// True if the workbook uses the 1904 date system.
    #[getter]
    fn is_1904(&self) -> bool {
        self.shared.is_1904
    }

    /// True once `close()` has been called.
    #[getter]
    fn closed(&self) -> bool {
        self.shared.is_closed()
    }

    /// Look up a sheet by name or zero-based index (negative counts from the end).
    /// An exact name match is preferred; otherwise names match case-insensitively
    /// as in Excel. Raises KeyError / IndexError if there is no such sheet.
    fn worksheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<ReadWorksheet>> {
        let index = self.find_index(key)?;
        self.sheet_at(py, index)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<ReadWorksheet>> {
        self.worksheet(py, key)
    }

    fn __len__(&self) -> usize {
        self.shared.sheets.len()
    }

    fn __contains__(&self, name: &Bound<'_, PyAny>) -> PyResult<bool> {
        match name.cast::<PyString>() {
            Ok(name) => {
                let name = name.to_str()?;
                Ok(self.shared.sheets.iter().any(|s| s.name == name))
            }
            Err(_) => Ok(false),
        }
    }

    fn __iter__(slf: Bound<'_, Self>) -> SheetIterator {
        SheetIterator {
            book: slf.unbind(),
            next: 0,
        }
    }

    /// The tables defined in the workbook (xlsx only; empty for other formats).
    #[getter]
    fn tables(&self, py: Python<'_>) -> PyResult<Vec<TableInfo>> {
        Ok(self
            .shared
            .tables(py)?
            .iter()
            .map(|meta| TableInfo {
                shared: self.shared.clone(),
                meta: meta.clone(),
            })
            .collect())
    }

    /// Look up a table by name (case-insensitive, as in Excel). Raises KeyError.
    fn table(&self, py: Python<'_>, name: &str) -> PyResult<TableInfo> {
        let tables = self.shared.tables(py)?;
        let found = tables.iter().find(|t| t.name == name).or_else(|| {
            tables
                .iter()
                .find(|t| t.name.to_lowercase() == name.to_lowercase())
        });
        match found {
            Some(meta) => Ok(TableInfo {
                shared: self.shared.clone(),
                meta: meta.clone(),
            }),
            None => Err(PyKeyError::new_err(format!("no table named {name:?}"))),
        }
    }

    /// The embedded images as `(extension, data)` pairs.
    #[getter]
    fn pictures<'py>(&self, py: Python<'py>) -> PyResult<Vec<(String, Bound<'py, PyBytes>)>> {
        Ok(self
            .shared
            .pictures(py)?
            .into_iter()
            .map(|(ext, data)| (ext, PyBytes::new(py, &data)))
            .collect())
    }

    /// The VBA modules as `{name: source code}`, or None without a VBA project.
    #[getter]
    fn vba_modules<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let Some(modules) = self.shared.vba_modules(py)? else {
            return Ok(None);
        };
        let dict = PyDict::new(py);
        for (name, code) in modules {
            dict.set_item(name, code)?;
        }
        Ok(Some(dict))
    }

    /// Release the file, and free parsed sheets that nothing uses any more.
    /// Sheets still held (e.g. in a variable) stay readable; other reads
    /// raise ReadError.
    fn close(&self, py: Python<'_>) {
        self.shared.close(py);
        for slot in &self.worksheets {
            let worksheet = slot.lock().unwrap_or_else(PoisonError::into_inner).take();
            drop(worksheet);
        }
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    /// Close the workbook; exceptions are not suppressed.
    #[pyo3(signature = (exc_type, exc, tb))]
    #[allow(unused_variables)]
    fn __exit__(
        &self,
        py: Python<'_>,
        exc_type: Option<&Bound<'_, PyAny>>,
        exc: Option<&Bound<'_, PyAny>>,
        tb: Option<&Bound<'_, PyAny>>,
    ) -> bool {
        self.close(py);
        false
    }

    fn __repr__(&self) -> String {
        format!(
            "<ReadWorkbook format={:?} sheets={:?}{}>",
            self.shared.kind.name(),
            self.shared
                .sheets
                .iter()
                .map(|s| s.name.as_str())
                .collect::<Vec<_>>(),
            if self.shared.is_closed() {
                " closed"
            } else {
                ""
            }
        )
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(open_workbook, m)?)?;
    m.add_class::<ReadWorkbook>()?;
    m.add_class::<SheetInfo>()?;
    m.add_class::<TableInfo>()?;
    Ok(())
}
