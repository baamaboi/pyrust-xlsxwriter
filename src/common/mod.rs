//! Building blocks shared by the writer, reader and editor.

#[macro_use]
pub mod macros;
pub mod bytes;
pub mod cellref;
pub mod color;
pub mod errors;
pub mod format;
pub mod value;

use pyo3::prelude::*;

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    errors::register(m)?;
    color::register(m)?;
    format::register(m)?;
    value::register(m)?;
    cellref::register(m)?;
    Ok(())
}
