//! Exception hierarchy shared by the writer, reader and editor engines.
//!
//! ```text
//! XlsxError(Exception)
//! ├── ParameterError(XlsxError, ValueError)
//! │   ├── SheetNameError
//! │   ├── OverlapError        merged ranges / tables / autofilters overlap
//! │   └── ImageError
//! ├── RangeError(XlsxError, IndexError)   row/column outside Excel's limits
//! ├── FileError(XlsxError, OSError)
//! ├── ReadError(XlsxError)                calamine could not read the file
//! └── EditError(XlsxError)                the editor could not apply a change
//! ```
//!
//! The classes are built at import time with `type()` so they can have two
//! bases (`except ValueError` keeps working for parameter problems).

use pyo3::exceptions::{PyException, PyIndexError, PyOSError, PyValueError};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyDict, PyTuple, PyType};
use rust_xlsxwriter::XlsxError;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ErrorKind {
    Base = 0,
    Parameter,
    SheetName,
    Overlap,
    Image,
    Range,
    File,
    Read,
    Edit,
}

const KIND_COUNT: usize = 9;

static TYPES: [PyOnceLock<Py<PyType>>; KIND_COUNT] = [const { PyOnceLock::new() }; KIND_COUNT];

const DEFINITIONS: [(ErrorKind, &str, &str); KIND_COUNT] = [
    (ErrorKind::Base, "XlsxError", "Base class for all pyrust_xlsxwriter errors."),
    (
        ErrorKind::Parameter,
        "ParameterError",
        "An argument value is invalid for Excel (also a ValueError).",
    ),
    (ErrorKind::SheetName, "SheetNameError", "Invalid, duplicate or unknown worksheet name."),
    (
        ErrorKind::Overlap,
        "OverlapError",
        "A merged range, table or autofilter overlaps an existing one.",
    ),
    (ErrorKind::Image, "ImageError", "The image type or dimensions are not supported."),
    (
        ErrorKind::Range,
        "RangeError",
        "A row/column is outside Excel's limits or a range is reversed (also an IndexError).",
    ),
    (ErrorKind::File, "FileError", "Reading or writing a file failed (also an OSError)."),
    (ErrorKind::Read, "ReadError", "A spreadsheet could not be read."),
    (ErrorKind::Edit, "EditError", "A change could not be applied to an existing workbook."),
];

/// Create the exception classes and add them to the module.
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = m.py();
    let type_fn = py.import("builtins")?.getattr("type")?;

    for (kind, name, doc) in DEFINITIONS {
        let bases: Vec<Bound<'_, PyAny>> = match kind {
            ErrorKind::Base => vec![py.get_type::<PyException>().into_any()],
            ErrorKind::Parameter => vec![
                exception_type(py, ErrorKind::Base).into_any(),
                py.get_type::<PyValueError>().into_any(),
            ],
            ErrorKind::SheetName | ErrorKind::Overlap | ErrorKind::Image => {
                vec![exception_type(py, ErrorKind::Parameter).into_any()]
            }
            ErrorKind::Range => vec![
                exception_type(py, ErrorKind::Base).into_any(),
                py.get_type::<PyIndexError>().into_any(),
            ],
            ErrorKind::File => vec![
                exception_type(py, ErrorKind::Base).into_any(),
                py.get_type::<PyOSError>().into_any(),
            ],
            ErrorKind::Read | ErrorKind::Edit => {
                vec![exception_type(py, ErrorKind::Base).into_any()]
            }
        };

        let namespace = PyDict::new(py);
        namespace.set_item("__module__", "pyrust_xlsxwriter")?;
        namespace.set_item("__doc__", doc)?;
        let cls = type_fn
            .call1((name, PyTuple::new(py, bases)?, namespace))?
            .cast_into::<PyType>()?;

        TYPES[kind as usize]
            .set(py, cls.clone().unbind())
            .map_err(|_| PyValueError::new_err("exception types initialised twice"))?;
        m.add(name, cls)?;
    }
    Ok(())
}

fn exception_type(py: Python<'_>, kind: ErrorKind) -> Bound<'_, PyType> {
    match TYPES[kind as usize].get(py) {
        Some(ty) => ty.bind(py).clone(),
        // Only reachable if an error is raised before module init finished.
        None => py.get_type::<PyException>(),
    }
}

/// Build a Python exception of the given kind.
pub fn new_err(kind: ErrorKind, message: impl Into<String>) -> PyErr {
    let message = message.into();
    Python::attach(|py| PyErr::from_type(exception_type(py, kind), (message,)))
}

pub fn param_err(message: impl Into<String>) -> PyErr {
    new_err(ErrorKind::Parameter, message)
}

pub fn range_err(message: impl Into<String>) -> PyErr {
    new_err(ErrorKind::Range, message)
}

pub fn edit_err(message: impl Into<String>) -> PyErr {
    new_err(ErrorKind::Edit, message)
}

pub fn read_err(message: impl Into<String>) -> PyErr {
    new_err(ErrorKind::Read, message)
}

pub fn file_err(message: impl Into<String>) -> PyErr {
    new_err(ErrorKind::File, message)
}

/// Translate a `rust_xlsxwriter` error into the matching Python exception.
pub fn xlsx_err(error: XlsxError) -> PyErr {
    let kind = match &error {
        XlsxError::RowColumnLimitError | XlsxError::RowColumnOrderError => ErrorKind::Range,
        XlsxError::SheetnameCannotBeBlank(..)
        | XlsxError::SheetnameLengthExceeded(..)
        | XlsxError::SheetnameReused(..)
        | XlsxError::SheetnameContainsInvalidCharacter(..)
        | XlsxError::SheetnameStartsOrEndsWithApostrophe(..)
        | XlsxError::UnknownWorksheetNameOrIndex(..) => ErrorKind::SheetName,
        XlsxError::MergeRangeOverlaps(..)
        | XlsxError::TableRangeOverlaps(..)
        | XlsxError::AutofilterRangeOverlaps(..) => ErrorKind::Overlap,
        XlsxError::UnknownImageType | XlsxError::ImageDimensionError => ErrorKind::Image,
        XlsxError::IoError(..) | XlsxError::ZipError(..) | XlsxError::TempFileError(..) => {
            ErrorKind::File
        }
        _ => ErrorKind::Parameter,
    };
    new_err(kind, error.to_string())
}

/// `result.py()?` converts any `rust_xlsxwriter` result into a `PyResult`.
pub trait IntoPyResult<T> {
    fn py(self) -> PyResult<T>;
}

impl<T> IntoPyResult<T> for Result<T, XlsxError> {
    fn py(self) -> PyResult<T> {
        self.map_err(xlsx_err)
    }
}
