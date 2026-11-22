//! `Worksheet.write_records()`: write a header row and one row per record.
//!
//! Records may be dicts (or any mapping), dataclass instances, NamedTuples or
//! plain objects. The record kind is detected once per record type, and
//! values are read with plain item/attribute access, so the per-cell cost is
//! one dict lookup or `getattr` plus the cell write.

use pyo3::exceptions::{PyKeyError, PyTypeError};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyBool, PyDict, PyFloat, PyInt, PyList, PyMapping, PyString, PyTuple, PyType};
use rust_xlsxwriter::Format as RFormat;

use crate::common::cellref::{split_cell, Params, MAX_COLS};
use crate::common::errors::{param_err, range_err};
use crate::common::format::Format;
use crate::common::value::CellValue;
use crate::writer::worksheet::Worksheet;

static DATACLASS_FIELDS: PyOnceLock<Py<PyAny>> = PyOnceLock::new();

/// How values are read from records of one type.
enum Access {
    /// `dict`: `record.get(column)`, missing keys are blank.
    Dict,
    /// Other mappings: `record[column]`, missing keys are blank.
    Mapping,
    /// NamedTuple: position of each column in the tuple (None: use getattr).
    Tuple(Vec<Option<usize>>),
    /// Dataclasses and other objects: `getattr(record, column)`.
    Attr,
}

/// Per-type cache of the access method.
struct Reader<'py> {
    columns: Vec<Bound<'py, PyAny>>,
    last_type: Option<Bound<'py, PyType>>,
    access: Access,
}

fn type_name(value: &Bound<'_, PyAny>) -> String {
    value.get_type().name().map(|n| n.to_string()).unwrap_or_else(|_| "?".into())
}

fn unsupported(record: &Bound<'_, PyAny>) -> PyErr {
    PyTypeError::new_err(format!(
        "write_records() records must be dicts, dataclass instances, NamedTuples or objects \
         with attributes, not {}",
        type_name(record)
    ))
}

/// The NamedTuple `_fields` of a record, if it is one.
fn namedtuple_fields<'py>(record: &Bound<'py, PyAny>) -> PyResult<Option<Vec<String>>> {
    if record.cast::<PyTuple>().is_err() {
        return Ok(None);
    }
    match record.get_type().getattr("_fields") {
        Ok(fields) => Ok(Some(fields.extract()?)),
        Err(_) => Ok(None),
    }
}

/// Plain values (and sequences) are not records.
fn is_scalar_or_sequence(record: &Bound<'_, PyAny>) -> bool {
    record.is_instance_of::<PyString>()
        || record.is_instance_of::<PyInt>()
        || record.is_instance_of::<PyFloat>()
        || record.is_instance_of::<PyBool>()
        || record.is_instance_of::<PyList>()
        || record.is_instance_of::<PyTuple>()
        || record.is_none()
}

/// Column names derived from the first record.
fn default_columns<'py>(record: &Bound<'py, PyAny>) -> PyResult<Vec<Bound<'py, PyAny>>> {
    let py = record.py();
    if let Ok(dict) = record.cast::<PyDict>() {
        return Ok(dict.keys().iter().collect());
    }
    if let Some(fields) = namedtuple_fields(record)? {
        return fields.into_iter().map(|f| Ok(PyString::new(py, &f).into_any())).collect();
    }
    if let Ok(mapping) = record.cast::<PyMapping>() {
        return mapping.keys()?.iter().map(Ok).collect();
    }
    if is_scalar_or_sequence(record) {
        return Err(unsupported(record));
    }
    if record.get_type().hasattr("__dataclass_fields__")? {
        let fields = DATACLASS_FIELDS.import(py, "dataclasses", "fields")?.call1((record,))?;
        return fields.try_iter()?.map(|f| f?.getattr("name")).collect();
    }
    if let Ok(attrs) = record.getattr("__dict__")
        && let Ok(attrs) = attrs.cast::<PyDict>()
    {
        let mut out = Vec::new();
        for key in attrs.keys().iter() {
            let public = key.cast::<PyString>().is_ok_and(|k| k.to_str().is_ok_and(|k| !k.starts_with('_')));
            if public {
                out.push(key);
            }
        }
        return Ok(out);
    }
    Err(PyTypeError::new_err(format!(
        "write_records() can't find the fields of {} records: pass columns=[...]",
        type_name(record)
    )))
}

impl<'py> Reader<'py> {
    fn new(columns: Vec<Bound<'py, PyAny>>) -> Self {
        Reader { columns, last_type: None, access: Access::Attr }
    }

    /// Work out how to read records of this record's type.
    fn classify(&mut self, record: &Bound<'py, PyAny>) -> PyResult<()> {
        let record_type = record.get_type();
        if self.last_type.as_ref().is_some_and(|t| t.is(&record_type)) {
            return Ok(());
        }
        self.access = if record.cast::<PyDict>().is_ok() {
            Access::Dict
        } else if let Some(fields) = namedtuple_fields(record)? {
            let positions = self
                .columns
                .iter()
                .map(|c| {
                    let name = c.cast::<PyString>().ok().and_then(|s| s.to_str().ok());
                    name.and_then(|name| fields.iter().position(|f| f == name))
                })
                .collect();
            Access::Tuple(positions)
        } else if record.cast::<PyMapping>().is_ok() {
            Access::Mapping
        } else if is_scalar_or_sequence(record) {
            return Err(unsupported(record));
        } else {
            for column in &self.columns {
                if column.cast::<PyString>().is_err() {
                    return Err(PyTypeError::new_err(format!(
                        "write_records() columns must be attribute names (str) for {} records, not {}",
                        type_name(record),
                        type_name(column)
                    )));
                }
            }
            Access::Attr
        };
        self.last_type = Some(record_type);
        Ok(())
    }

    /// The value of column `i` of a record (None for a blank cell).
    fn value(&self, record: &Bound<'py, PyAny>, i: usize) -> PyResult<Option<Bound<'py, PyAny>>> {
        let column = &self.columns[i];
        match &self.access {
            Access::Dict => record.cast::<PyDict>()?.get_item(column),
            Access::Mapping => match record.get_item(column) {
                Ok(value) => Ok(Some(value)),
                Err(e) if e.is_instance_of::<PyKeyError>(record.py()) => Ok(None),
                Err(e) => Err(e),
            },
            Access::Tuple(positions) => match positions[i] {
                Some(index) => record.cast::<PyTuple>()?.get_item(index).map(Some),
                None => record.getattr(column.cast::<PyString>()?).map(Some),
            },
            Access::Attr => record.getattr(column.cast::<PyString>()?).map(Some),
        }
    }
}

/// The data format: one for every column, or one per column.
enum DataFormat {
    All(Option<RFormat>),
    PerColumn(Vec<Option<RFormat>>),
}

impl DataFormat {
    fn new(value: Option<&Bound<'_, PyAny>>, columns: &[Bound<'_, PyAny>]) -> PyResult<Self> {
        let Some(value) = value else { return Ok(DataFormat::All(None)) };
        if let Ok(format) = value.cast::<Format>() {
            return Ok(DataFormat::All(Some(format.borrow().inner.clone())));
        }
        let dict = value.cast::<PyDict>().map_err(|_| {
            PyTypeError::new_err(format!(
                "write_records() argument 'format' must be a Format or a dict of column -> Format, not {}",
                type_name(value)
            ))
        })?;
        for key in dict.keys().iter() {
            let mut found = false;
            for column in columns {
                if column.eq(&key)? {
                    found = true;
                    break;
                }
            }
            if !found {
                return Err(param_err(format!(
                    "write_records() format given for unknown column {}",
                    key.repr()?
                )));
            }
        }
        let mut formats = Vec::with_capacity(columns.len());
        for column in columns {
            formats.push(match dict.get_item(column)? {
                Some(f) if !f.is_none() => Some(f.extract::<PyRef<'_, Format>>()?.inner.clone()),
                _ => None,
            });
        }
        Ok(DataFormat::PerColumn(formats))
    }

    fn get(&self, i: usize) -> Option<&RFormat> {
        match self {
            DataFormat::All(format) => format.as_ref(),
            DataFormat::PerColumn(formats) => formats[i].as_ref(),
        }
    }
}

#[pymethods]
impl Worksheet {
    /// Write records as a table: a header row, then one row per record.
    ///
    /// ```python
    /// rows = [{"name": "Ann", "age": 31}, {"name": "Bob", "age": 27}]
    /// first_row, first_col, last_row, last_col = ws.write_records("A1", rows)
    /// ws.add_table(first_row, first_col, last_row, last_col)   # e.g.
    /// ```
    ///
    /// `records` is an iterable of dicts (or mappings), dataclass instances,
    /// NamedTuples or objects with attributes. Without `columns`, the columns
    /// are the first record's keys / fields / public attributes (later dict
    /// keys not in the first record are ignored); `columns` selects and orders
    /// them. Missing dict keys give blank cells; missing attributes raise
    /// AttributeError. Values are written as by `write()`.
    ///
    /// `header` writes the column names (as text) in the first row, with
    /// `header_format`. `format` is a Format for every data cell or a dict
    /// `{column: Format}`.
    ///
    /// Returns the written range `(first_row, first_col, last_row, last_col)`,
    /// including the header row. Raises ParameterError if there is nothing to
    /// write (no records and no header).
    #[pyo3(signature = (*args, **kwargs))]
    fn write_records(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<(u32, u16, u32, u16)> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        if rest.len() > 1 {
            return Err(PyTypeError::new_err(
                "write_records() takes the cell and the records positionally; pass columns, \
                 header, header_format and format as keywords",
            ));
        }
        let p = Params::new(
            "write_records",
            rest,
            kwargs,
            &["records", "columns", "header", "header_format", "format"],
            1,
        )?;
        let header: bool = p.get_or(2, true)?;
        let header_format: Option<PyRef<'_, Format>> = p.extract(3)?;

        let mut records = p.required_any(0)?.try_iter().map_err(|_| {
            PyTypeError::new_err("write_records() argument 'records' must be an iterable of records")
        })?;
        let first = records.next().transpose()?;

        let columns: Vec<Bound<'_, PyAny>> = match p.raw(1) {
            Some(columns) => {
                if columns.cast::<PyString>().is_ok() {
                    return Err(PyTypeError::new_err(
                        "write_records() argument 'columns' must be a list of column names, not a str",
                    ));
                }
                columns.try_iter()?.collect::<PyResult<_>>()?
            }
            None => match &first {
                Some(record) => default_columns(record)?,
                None => {
                    return Err(param_err(
                        "write_records() got no records: pass columns=[...] to write just the header",
                    ))
                }
            },
        };
        if columns.is_empty() {
            return Err(param_err("write_records() has no columns to write"));
        }
        if first.is_none() && !header {
            return Err(param_err("write_records() got no records and header=False: nothing to write"));
        }
        let last_col = col as usize + columns.len() - 1;
        if last_col >= MAX_COLS as usize {
            return Err(range_err(format!(
                "write_records() needs {} columns from column {col}, past Excel's limit of {MAX_COLS}",
                columns.len()
            )));
        }
        let last_col = last_col as u16;
        let data_format = DataFormat::new(p.raw(4), &columns)?;

        let mut r = row;
        if header {
            let hfmt = header_format.as_ref().map(|f| &f.inner);
            for (i, column) in columns.iter().enumerate() {
                let text = match column.cast::<PyString>() {
                    Ok(s) => s.clone(),
                    Err(_) => column.str()?,
                };
                self.write_classified(r, col + i as u16, CellValue::String(text), hfmt)?;
            }
        }

        let mut reader = Reader::new(columns);
        let ncols = reader.columns.len();
        let mut wrote_rows = header;
        let mut next = first;
        while let Some(record) = next {
            if wrote_rows {
                r = r.checked_add(1).ok_or_else(|| range_err("write_records() row overflow"))?;
            }
            wrote_rows = true;
            reader.classify(&record)?;
            for i in 0..ncols {
                let c = col + i as u16;
                match reader.value(&record, i)? {
                    Some(value) => self.write_value(r, c, &value, data_format.get(i))?,
                    None => self.write_classified(r, c, CellValue::Blank, data_format.get(i))?,
                }
            }
            next = records.next().transpose()?;
        }
        Ok((row, col, r, last_col))
    }
}

pub fn register(_m: &Bound<'_, PyModule>) -> PyResult<()> {
    Ok(())
}
