//! `Workbook` for creating new files.
//!
//! The Python object owns one persistent `rust_xlsxwriter::Workbook`, so
//! workbook-level settings (VBA, properties, defined names, default format,
//! constant-memory worksheets, ...) are applied to it directly. Worksheets are
//! owned by their Python objects so they stay editable; for a save they are
//! moved into the workbook and moved back afterwards.
//!
//! Constant/low-memory worksheets (`add_worksheet(memory="constant")`) share
//! the format table (and for "low" the string table) of the Rust workbook that
//! created them, which is why that workbook is kept for the object's lifetime;
//! they can't be moved to another workbook.

use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{LazyLock, Mutex};

use pyo3::exceptions::{PyIndexError, PyKeyError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyString};
use rust_xlsxwriter::{Format as RFormat, Workbook as RWorkbook, Worksheet as RWorksheet, XlsxError};

use crate::common::bytes::write_to_file;
use crate::common::errors::{new_err, param_err, ErrorKind, IntoPyResult};
use crate::writer::worksheet::Worksheet;

/// Source of unique workbook ids.
static NEXT_WORKBOOK_ID: AtomicU64 = AtomicU64::new(1);

/// Constant/low-memory worksheets (by object address) -> id of the workbook
/// that created them. The workbook holds a reference to each of its
/// worksheets, so an address can't be reused while it is registered. When the
/// workbook is dropped, entries for worksheets that die with it are removed;
/// a worksheet still referenced elsewhere stays registered (and is kept in
/// `ORPHANED_SHEETS` so its address is never reused), because it can't be
/// saved correctly in any other workbook.
static MEMORY_SHEETS: LazyLock<Mutex<HashMap<usize, u64>>> = LazyLock::new(Default::default);

static ORPHANED_SHEETS: Mutex<Vec<Py<Worksheet>>> = Mutex::new(Vec::new());

fn memory_sheets() -> std::sync::MutexGuard<'static, HashMap<usize, u64>> {
    MEMORY_SHEETS.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Worksheet storage mode for `Workbook.add_worksheet(memory=...)`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MemoryMode {
    Standard,
    Constant,
    Low,
}

impl MemoryMode {
    fn parse(value: Option<&str>) -> PyResult<Self> {
        match value {
            None => Ok(MemoryMode::Standard),
            Some("constant") => Ok(MemoryMode::Constant),
            Some("low") => Ok(MemoryMode::Low),
            Some(other) => Err(param_err(format!(
                "memory must be None, \"constant\" or \"low\", not {other:?}"
            ))),
        }
    }
}

/// A new Excel workbook.
///
/// ```python
/// with Workbook("report.xlsx") as wb:      # saved on exit
///     ws = wb.add_worksheet("Data")
///     ws.write("A1", "Hello")
/// ```
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct Workbook {
    pub inner: RWorkbook,
    pub worksheets: Vec<Py<Worksheet>>,
    pub path: Option<PathBuf>,
    /// Default format, row height and column width (pixels) set with
    /// `set_default_format()` / `use_excel_2023_theme()`; applied to every
    /// worksheet added afterwards.
    pub default_format: Option<(RFormat, u32, u32)>,
    /// Unique id, used to tie constant/low-memory worksheets to this workbook.
    pub id: u64,
}

impl Drop for Workbook {
    fn drop(&mut self) {
        // Lock order everywhere: Python thread state first, then the registry.
        Python::attach(|py| {
            let mut registry = memory_sheets();
            if !registry.values().any(|owner| *owner == self.id) {
                return;
            }
            for sheet in &self.worksheets {
                let key = sheet.as_ptr() as usize;
                if registry.get(&key) != Some(&self.id) {
                    continue;
                }
                if sheet.get_refcnt(py) > 1 {
                    let mut orphans = ORPHANED_SHEETS.lock().unwrap_or_else(|p| p.into_inner());
                    orphans.push(sheet.clone_ref(py));
                } else {
                    registry.remove(&key);
                }
            }
        });
    }
}

impl Workbook {
    fn sheet_names(&self, py: Python<'_>) -> Vec<String> {
        self.worksheets.iter().map(|ws| ws.borrow(py).inner.name()).collect()
    }

    fn next_sheet_name(&self, py: Python<'_>) -> String {
        let names = self.sheet_names(py);
        let mut n = self.worksheets.len() + 1;
        loop {
            let candidate = format!("Sheet{n}");
            if !names.iter().any(|existing| existing.eq_ignore_ascii_case(&candidate)) {
                return candidate;
            }
            n += 1;
        }
    }

    fn check_unique(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        if self.sheet_names(py).iter().any(|existing| existing.eq_ignore_ascii_case(name)) {
            return Err(new_err(
                ErrorKind::SheetName,
                format!("worksheet name {name:?} is already used in this workbook"),
            ));
        }
        Ok(())
    }

    /// Move every worksheet into the Rust workbook, run `op` without holding
    /// the GIL, then move the worksheets back to their Python objects.
    pub fn with_sheets<R: Send>(
        &mut self,
        py: Python<'_>,
        op: impl FnOnce(&mut RWorkbook) -> Result<R, XlsxError> + Send,
    ) -> PyResult<R> {
        let mut guards = Vec::with_capacity(self.worksheets.len());
        for sheet in &self.worksheets {
            guards.push(sheet.bind(py).try_borrow_mut().map_err(|_| {
                param_err("a worksheet is in use elsewhere and cannot be saved right now")
            })?);
        }
        for guard in guards.iter_mut() {
            let sheet = std::mem::replace(&mut guard.inner, RWorksheet::new());
            self.inner.push_worksheet(sheet);
        }

        let inner = &mut self.inner;
        let result = py.detach(move || op(inner));

        let sheets = std::mem::take(self.inner.worksheets_mut());
        for (guard, sheet) in guards.iter_mut().zip(sheets) {
            guard.inner = sheet;
            guard.mark_saved();
        }
        result.py()
    }

    /// Apply the workbook's custom default format (if any) to a worksheet.
    fn init_default_format(&self, sheet: &mut RWorksheet) -> PyResult<()> {
        if let Some((format, row_height, col_width)) = &self.default_format {
            sheet.set_default_format(format, *row_height, *col_width).map_err(|e| {
                param_err(format!(
                    "can't apply the workbook's default format to this worksheet ({e}); \
                     add worksheets with add_worksheet(), or push them before writing formatted cells"
                ))
            })?;
        }
        Ok(())
    }

    /// Create the Rust worksheet for `add_worksheet()`.
    fn new_rust_worksheet(&mut self, mode: MemoryMode) -> PyResult<RWorksheet> {
        let mut sheet = match mode {
            MemoryMode::Standard => RWorksheet::new(),
            MemoryMode::Constant => self.inner.new_worksheet_with_constant_memory().py()?,
            MemoryMode::Low => self.inner.new_worksheet_with_low_memory().py()?,
        };
        self.init_default_format(&mut sheet)?;
        Ok(sheet)
    }

    pub fn find_index(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<usize> {
        if let Ok(name) = key.cast::<PyString>() {
            let name = name.to_str()?;
            return self
                .sheet_names(py)
                .iter()
                .position(|n| n == name)
                .ok_or_else(|| PyKeyError::new_err(format!("no worksheet named {name:?}")));
        }
        let index: isize = key.extract().map_err(|_| {
            PyTypeError::new_err("worksheet key must be a name (str) or an index (int)")
        })?;
        let len = self.worksheets.len() as isize;
        let index = if index < 0 { index + len } else { index };
        if index < 0 || index >= len {
            return Err(PyIndexError::new_err("worksheet index out of range"));
        }
        Ok(index as usize)
    }
}

#[pymethods]
impl Workbook {
    /// `path` is optional; it is used by `close()` and the `with` statement.
    #[new]
    #[pyo3(signature = (path=None))]
    fn py_new(path: Option<PathBuf>) -> Self {
        Workbook {
            inner: RWorkbook::new(),
            worksheets: Vec::new(),
            path,
            default_format: None,
            id: NEXT_WORKBOOK_ID.fetch_add(1, Ordering::Relaxed),
        }
    }

    /// Add a worksheet. Keyword arguments set worksheet properties, e.g.
    /// `add_worksheet("Data", tab_color="#4472C4", zoom=120)`.
    ///
    /// `memory` selects how cell data is stored:
    ///
    /// * `None` (default): in memory, so cells can be written in any order.
    /// * `"constant"`: each row is flushed to a temporary file as soon as a
    ///   later row is written, keeping memory use flat for very large sheets.
    ///   Strings are stored inline (larger files).
    /// * `"low"`: like `"constant"` but strings go to the shared string table
    ///   (smaller files, memory grows with the number of unique strings).
    ///
    /// In the constant/low modes data must be written in row order: once a
    /// cell in row `n` is written, writes to rows `< n` are ignored (with a
    /// warning on stderr), and formats set separately from data (e.g.
    /// `set_cell_format()`) only work on the current row. Saving is allowed
    /// any number of times, but it completes the current row, so continue on
    /// a later row after a save. Such worksheets are tied to this workbook
    /// (they share its format and string tables) and can't be pushed to
    /// another one; call `set_tempdir()` before adding them to choose where
    /// the temporary files go.
    #[pyo3(signature = (name=None, memory=None, **kwargs))]
    fn add_worksheet<'py>(
        &mut self,
        py: Python<'py>,
        name: Option<&str>,
        memory: Option<&str>,
        kwargs: Option<&Bound<'py, PyDict>>,
    ) -> PyResult<Bound<'py, Worksheet>> {
        let mode = MemoryMode::parse(memory)?;
        let name = match name {
            Some(name) => name.to_string(),
            None => self.next_sheet_name(py),
        };
        self.check_unique(py, &name)?;
        // Validate the name and properties first, then swap in the Rust
        // worksheet for the memory mode / default format and re-apply them.
        let mut worksheet = Worksheet::create(Some(&name), kwargs)?;
        if mode != MemoryMode::Standard || self.default_format.is_some() {
            let mut inner = self.new_rust_worksheet(mode)?;
            inner.set_name(&name).py()?;
            worksheet.replace_inner(inner, mode != MemoryMode::Standard)?;
        }
        let sheet = Bound::new(py, worksheet)?;
        if mode != MemoryMode::Standard {
            memory_sheets().insert(sheet.as_ptr() as usize, self.id);
        }
        self.worksheets.push(sheet.clone().unbind());
        Ok(sheet)
    }

    /// Add a worksheet created separately with `Worksheet()`.
    fn push_worksheet(&mut self, py: Python<'_>, worksheet: Bound<'_, Worksheet>) -> PyResult<()> {
        if self.worksheets.iter().any(|ws| ws.bind(py).is(&worksheet)) {
            return Err(param_err("this worksheet is already in the workbook"));
        }
        if memory_sheets().contains_key(&(worksheet.as_ptr() as usize)) {
            return Err(param_err(
                "a constant/low-memory worksheet belongs to the workbook that created it \
                 and can't be added to another workbook",
            ));
        }
        let name = worksheet.borrow().inner.name();
        if name.is_empty() {
            let name = self.next_sheet_name(py);
            worksheet.borrow_mut().inner.set_name(name).py()?;
        } else {
            self.check_unique(py, &name)?;
        }
        if self.default_format.is_some() {
            let mut sheet = worksheet.borrow_mut();
            self.init_default_format(&mut sheet.inner)?;
            // The default format resets the default row height: re-apply the
            // worksheet's own properties.
            sheet.apply_kwargs(None)?;
        }
        self.worksheets.push(worksheet.unbind());
        Ok(())
    }

    /// The worksheets, in tab order.
    #[getter]
    fn worksheets(&self, py: Python<'_>) -> Vec<Py<Worksheet>> {
        self.worksheets.iter().map(|ws| ws.clone_ref(py)).collect()
    }

    /// The worksheet names, in tab order.
    #[getter(sheet_names)]
    fn py_sheet_names(&self, py: Python<'_>) -> Vec<String> {
        self.sheet_names(py)
    }

    /// Look up a worksheet by name or index.
    fn worksheet(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<Worksheet>> {
        let index = self.find_index(py, key)?;
        Ok(self.worksheets[index].clone_ref(py))
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<Worksheet>> {
        self.worksheet(py, key)
    }

    fn __len__(&self) -> usize {
        self.worksheets.len()
    }

    fn __contains__(&self, py: Python<'_>, name: &str) -> bool {
        self.sheet_names(py).iter().any(|n| n == name)
    }

    /// Save to `path` (or the path given to the constructor).
    #[pyo3(signature = (path=None))]
    fn save(&mut self, py: Python<'_>, path: Option<PathBuf>) -> PyResult<()> {
        let path = path
            .or_else(|| self.path.clone())
            .ok_or_else(|| PyTypeError::new_err("save() needs a path"))?;
        self.with_sheets(py, move |wb| wb.save(&path))
    }

    /// Return the file as `bytes` instead of writing it to disk.
    fn save_to_buffer<'py>(&mut self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let data = self.with_sheets(py, |wb| wb.save_to_buffer())?;
        Ok(PyBytes::new(py, &data))
    }

    /// Write the file to a binary file object (anything with `.write(bytes)`).
    fn save_to_file(&mut self, py: Python<'_>, file: &Bound<'_, PyAny>) -> PyResult<()> {
        let data = self.with_sheets(py, |wb| wb.save_to_buffer())?;
        write_to_file(file, &data)
    }

    /// Save to the path given to the constructor.
    fn close(&mut self, py: Python<'_>) -> PyResult<()> {
        if self.path.is_none() {
            return Err(PyTypeError::new_err(
                "close() needs the workbook to be created with a path: Workbook(path)",
            ));
        }
        self.save(py, None)
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    /// Save on a clean exit; do nothing if the block raised.
    fn __exit__(
        &mut self,
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

    /// Add a VBA project (`vbaProject.bin` extracted from an .xlsm) and,
    /// optionally, its signature. Save the file with an `.xlsm` extension.
    #[pyo3(signature = (path, signature=None))]
    fn add_vba_project(&mut self, path: PathBuf, signature: Option<PathBuf>) -> PyResult<()> {
        match signature {
            Some(signature) => self.inner.add_vba_project_with_signature(&path, &signature).py()?,
            None => self.inner.add_vba_project(&path).py()?,
        };
        Ok(())
    }

    /// Same as `add_vba_project(path, signature)`.
    fn add_vba_project_with_signature(
        &mut self,
        project_path: PathBuf,
        signature_path: PathBuf,
    ) -> PyResult<()> {
        self.inner.add_vba_project_with_signature(&project_path, &signature_path).py()?;
        Ok(())
    }

    /// Set the workbook's VBA code name (default "ThisWorkbook").
    fn set_vba_name(&mut self, name: &str) -> PyResult<()> {
        self.inner.set_vba_name(name).py()?;
        Ok(())
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        format!("<Workbook sheets={:?}>", self.sheet_names(py))
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()
}
