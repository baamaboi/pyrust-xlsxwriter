//! pyrust_xlsxwriter: create, read and edit Excel files from Python.
//!
//! * `writer`  - new files, backed by rust_xlsxwriter.
//! * `reader`  - read-only access to xlsx/xlsm/xlsb/xls/ods, backed by calamine.
//! * `editor`  - in-place editing of existing xlsx/xlsm files.
//! * `common`  - values, colours, cell references, formats and errors shared by all three.

pub mod common;
pub mod editor;
pub mod reader;
pub mod writer;

use pyo3::prelude::*;

#[pymodule]
fn _pyrust_xlsxwriter(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    common::register(m)?;
    writer::register(m)?;
    reader::register(m)?;
    editor::register(m)?;
    Ok(())
}
