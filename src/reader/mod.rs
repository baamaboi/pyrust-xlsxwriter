//! Read-only access to spreadsheets, backed by calamine.
//!
//! * `source`    - path / bytes sources and format detection.
//! * `shared`    - the open calamine reader and lazily parsed sheet data.
//! * `convert`   - calamine cell values to Python objects.
//! * `style`     - calamine cell styles to `FormatSpec`.
//! * `workbook`  - `open_workbook()`, `ReadWorkbook`, `SheetInfo`, `TableInfo`.
//! * `worksheet` - `ReadWorksheet`, `ReadCell`, `HyperlinkInfo`.

pub mod convert;
pub mod shared;
pub mod source;
pub mod style;
pub mod workbook;
pub mod worksheet;

use pyo3::prelude::*;

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    workbook::register(m)?;
    worksheet::register(m)?;
    Ok(())
}
