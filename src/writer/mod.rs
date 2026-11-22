//! New workbooks, backed by rust_xlsxwriter.
//!
//! `worksheet.rs` defines the `Worksheet` struct and core cell writing. Each
//! feature area lives in its own file and adds methods to `Worksheet` (or
//! `Workbook`) with a separate `#[pymethods] impl` block, then registers its
//! classes in its own `register()`.

pub mod autofilter;
pub mod button;
pub mod chart;
pub mod conditional_format;
pub mod data_validation;
pub mod image;
pub mod note;
pub mod object_movement;
pub mod page_setup;
pub mod pivot_table;
pub mod records;
pub mod shape;
pub mod sparkline;
pub mod table;
pub mod workbook;
pub mod workbook_settings;
pub mod worksheet;

use pyo3::prelude::*;

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    workbook::register(m)?;
    worksheet::register(m)?;
    object_movement::register(m)?;
    autofilter::register(m)?;
    button::register(m)?;
    chart::register(m)?;
    conditional_format::register(m)?;
    data_validation::register(m)?;
    image::register(m)?;
    note::register(m)?;
    page_setup::register(m)?;
    pivot_table::register(m)?;
    records::register(m)?;
    shape::register(m)?;
    sparkline::register(m)?;
    table::register(m)?;
    workbook_settings::register(m)?;
    Ok(())
}
