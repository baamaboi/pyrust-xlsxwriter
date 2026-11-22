use pyo3::prelude::*;
use rust_xlsxwriter::Format as ExcelFormat;

#[pyclass]
struct Format {
    inner: ExcelFormat,
}

#[pymethods]
impl Format {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: ExcelFormat::new(),
        }
    }

    pub fn set_bold(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        let inner = slf.inner.clone();
        slf.inner = inner.set_bold();
        slf
    }
}
