use pyo3::PyErr;
use pyo3::exceptions::{PyIOError, PyRuntimeError};
use rust_xlsxwriter::XlsxError as RustXlsxwriterCrateError;

#[derive(Debug)]
pub enum Error {
    RustXlsxError(RustXlsxwriterCrateError),
}

impl From<Error> for PyErr {
    fn from(val: Error) -> Self {
        match val {
            Error::RustXlsxError(xlsx_error) => match xlsx_error {
                RustXlsxwriterCrateError::IoError(io_err) => PyIOError::new_err(io_err.to_string()),
                _ => PyRuntimeError::new_err(xlsx_error.to_string()),
            },
        }
    }
}
