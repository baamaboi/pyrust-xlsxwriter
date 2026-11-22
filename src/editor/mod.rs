//! In-place editing of existing xlsx/xlsm files.
//!
//! * `xml`     - a span-preserving XML scanner and a small DOM;
//! * `package` - the zip package: on-demand reads, raw copy of untouched parts;
//! * `styles`  - patching `styles.xml` (fonts/fills/borders/xfs, de-duplicated);
//! * `sst`     - the append-only shared string table;
//! * `sheet`   - recorded worksheet edits and the streaming sheet rewrite;
//! * `book`    - workbook structure (sheets, names, properties) and saving;
//! * `formula` - formula tokenizer and reference shifting;
//! * `structure` - inserting/deleting rows and columns, sheet renames;
//! * `graft`   - adding charts, tables, images, CF/DV, ... to existing sheets;
//! * `python`  - `load_workbook()`, `EditWorkbook`, `EditWorksheet`.

pub mod book;
pub mod formula;
pub mod graft;
pub mod package;
pub mod python;
pub mod sheet;
pub mod sst;
pub mod structure;
pub mod styles;
pub mod xml;

use pyo3::prelude::*;

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    python::register(m)
}
