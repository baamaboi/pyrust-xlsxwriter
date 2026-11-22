//! Cell references: `(row, col)` ints and Excel "A1" strings.
//!
//! Every worksheet method that takes a cell or range accepts either form,
//! XlsxWriter style:
//!
//! ```python
//! ws.write(0, 0, "x")            ws.write("A1", "x")
//! ws.merge_range(0, 0, 1, 3, "x") ws.merge_range("A1:D2", "x")
//! ```
//!
//! Methods use `*args` and call [`split_cell`] / [`split_range`] /
//! [`split_cols`]; the remaining positional/keyword arguments are read with
//! [`Params`]. All indexes are zero-based.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyInt, PyString};

use crate::common::errors::{param_err, range_err};

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u16 = 16_384;

pub type Cell = (u32, u16);
pub type CellRange = (u32, u16, u32, u16);

/// Zero-based column index to letters: 0 -> "A", 27 -> "AB".
pub fn col_to_name(col: u16) -> String {
    let mut col = col as u32 + 1;
    let mut name = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        name.push(b'A' + rem as u8);
        col = (col - 1) / 26;
    }
    name.reverse();
    String::from_utf8(name).unwrap_or_default()
}

/// Column letters to zero-based index: "A" -> 0, "$AB" -> 27.
pub fn name_to_col(name: &str) -> Option<u16> {
    let name = name.trim_start_matches('$');
    if name.is_empty() || name.len() > 3 {
        return None;
    }
    let mut col: u32 = 0;
    for c in name.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    let col = col.checked_sub(1)?;
    if col < MAX_COLS as u32 { Some(col as u16) } else { None }
}

/// Parse "A1" / "$A$1" to a zero-based (row, col).
pub fn parse_cell(text: &str) -> Option<Cell> {
    let text = text.trim();
    let split = text.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = text.split_at(split);
    let letters = letters.trim_end_matches('$');
    let col = name_to_col(letters)?;
    if digits.is_empty() || !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > MAX_ROWS {
        return None;
    }
    Some((row - 1, col))
}

/// Parse "A1:B2" (or a single "A1") to a zero-based range.
/// A sheet prefix such as "Sheet1!A1:B2" is ignored.
pub fn parse_range(text: &str) -> Option<CellRange> {
    let text = text.rsplit('!').next()?.trim();
    match text.split_once(':') {
        Some((a, b)) => {
            let (r1, c1) = parse_cell(a)?;
            let (r2, c2) = parse_cell(b)?;
            Some((r1, c1, r2, c2))
        }
        None => {
            let (r, c) = parse_cell(text)?;
            Some((r, c, r, c))
        }
    }
}

/// Parse "A" or "A:C" (also "A1:C1", taking only the columns).
pub fn parse_cols(text: &str) -> Option<(u16, u16)> {
    let text = text.trim();
    let one = |part: &str| -> Option<u16> {
        let letters: String = part.chars().take_while(|c| !c.is_ascii_digit()).collect();
        name_to_col(letters.trim_end_matches('$'))
    };
    match text.split_once(':') {
        Some((a, b)) => Some((one(a)?, one(b)?)),
        None => {
            let c = one(text)?;
            Some((c, c))
        }
    }
}

/// Zero-based (row, col) to "A1", optionally with absolute `$` markers.
pub fn cell_name(row: u32, col: u16, row_abs: bool, col_abs: bool) -> String {
    format!(
        "{}{}{}{}",
        if col_abs { "$" } else { "" },
        col_to_name(col),
        if row_abs { "$" } else { "" },
        row + 1
    )
}

pub fn range_name(range: CellRange, absolute: bool) -> String {
    let (r1, c1, r2, c2) = range;
    if r1 == r2 && c1 == c2 {
        cell_name(r1, c1, absolute, absolute)
    } else {
        format!(
            "{}:{}",
            cell_name(r1, c1, absolute, absolute),
            cell_name(r2, c2, absolute, absolute)
        )
    }
}

pub fn check_cell(row: u32, col: u16) -> PyResult<()> {
    if row >= MAX_ROWS || col >= MAX_COLS {
        Err(range_err(format!(
            "cell ({row}, {col}) is outside Excel's limits of {MAX_ROWS} rows and {MAX_COLS} columns"
        )))
    } else {
        Ok(())
    }
}

fn is_int(value: &Bound<'_, PyAny>) -> bool {
    value.is_instance_of::<PyInt>() && !value.is_instance_of::<PyBool>()
}

fn extract_row(value: &Bound<'_, PyAny>) -> PyResult<u32> {
    if !is_int(value) {
        return Err(PyTypeError::new_err(format!(
            "row must be an int, not {}",
            value.get_type().name()?
        )));
    }
    let row: i64 = value.extract()?;
    if !(0..MAX_ROWS as i64).contains(&row) {
        return Err(range_err(format!("row {row} is outside 0..{MAX_ROWS}")));
    }
    Ok(row as u32)
}

fn extract_col(value: &Bound<'_, PyAny>) -> PyResult<u16> {
    if let Ok(text) = value.cast::<PyString>() {
        let text = text.to_str()?;
        return name_to_col(text)
            .ok_or_else(|| param_err(format!("invalid column name {text:?}")));
    }
    if !is_int(value) {
        return Err(PyTypeError::new_err(format!(
            "column must be an int, not {}",
            value.get_type().name()?
        )));
    }
    let col: i64 = value.extract()?;
    if !(0..MAX_COLS as i64).contains(&col) {
        return Err(range_err(format!("column {col} is outside 0..{MAX_COLS}")));
    }
    Ok(col as u16)
}

fn missing(what: &str) -> PyErr {
    PyTypeError::new_err(format!("missing {what}: pass (row, col) ints or an \"A1\" string"))
}

/// Take a cell from the start of `args`: `"A1", ...` or `row, col, ...`.
pub fn split_cell<'a, 'py>(
    args: &'a [Bound<'py, PyAny>],
) -> PyResult<(Cell, &'a [Bound<'py, PyAny>])> {
    let first = args.first().ok_or_else(|| missing("cell"))?;
    if let Ok(text) = first.cast::<PyString>() {
        let text = text.to_str()?;
        let cell = parse_cell(text)
            .ok_or_else(|| param_err(format!("invalid cell reference {text:?}")))?;
        return Ok((cell, &args[1..]));
    }
    let col = args.get(1).ok_or_else(|| missing("column"))?;
    Ok(((extract_row(first)?, extract_col(col)?), &args[2..]))
}

/// Take a range from the start of `args`: `"A1:B2", ...` or
/// `first_row, first_col, last_row, last_col, ...`.
pub fn split_range<'a, 'py>(
    args: &'a [Bound<'py, PyAny>],
) -> PyResult<(CellRange, &'a [Bound<'py, PyAny>])> {
    let first = args.first().ok_or_else(|| missing("range"))?;
    let (range, rest) = if let Ok(text) = first.cast::<PyString>() {
        let text = text.to_str()?;
        let range = parse_range(text)
            .ok_or_else(|| param_err(format!("invalid range reference {text:?}")))?;
        (range, &args[1..])
    } else {
        if args.len() < 4 {
            return Err(missing("range (first_row, first_col, last_row, last_col)"));
        }
        (
            (
                extract_row(&args[0])?,
                extract_col(&args[1])?,
                extract_row(&args[2])?,
                extract_col(&args[3])?,
            ),
            &args[4..],
        )
    };
    let (r1, c1, r2, c2) = range;
    if r1 > r2 || c1 > c2 {
        return Err(range_err(format!(
            "range {} is reversed: the first cell must be above/left of the last",
            range_name(range, false)
        )));
    }
    Ok((range, rest))
}

/// Take one column or a column span: `"B"`, `"B:D"`, or an int.
pub fn split_cols<'a, 'py>(
    args: &'a [Bound<'py, PyAny>],
) -> PyResult<((u16, u16), &'a [Bound<'py, PyAny>])> {
    let first = args.first().ok_or_else(|| missing("column"))?;
    if let Ok(text) = first.cast::<PyString>() {
        let text = text.to_str()?;
        let cols = parse_cols(text)
            .ok_or_else(|| param_err(format!("invalid column reference {text:?}")))?;
        return Ok((cols, &args[1..]));
    }
    let col = extract_col(first)?;
    Ok(((col, col), &args[1..]))
}

/// Take a row number from the start of `args`.
pub fn split_row<'a, 'py>(
    args: &'a [Bound<'py, PyAny>],
) -> PyResult<(u32, &'a [Bound<'py, PyAny>])> {
    let first = args.first().ok_or_else(|| missing("row"))?;
    Ok((extract_row(first)?, &args[1..]))
}

/// The positional/keyword parameters that follow a cell/range in `*args`.
///
/// ```ignore
/// let ((row, col), rest) = split_cell(args.as_slice())?;
/// let p = Params::new("write", rest, kwargs, &["value", "format"], 1)?;
/// let value = p.required(0)?;
/// let format: Option<PyRef<Format>> = p.extract(1)?;
/// ```
pub struct Params<'a, 'py> {
    values: Vec<Option<Bound<'py, PyAny>>>,
    method: &'static str,
    names: &'static [&'static str],
    _marker: std::marker::PhantomData<&'a ()>,
}

impl<'a, 'py> Params<'a, 'py> {
    pub fn new(
        method: &'static str,
        positional: &'a [Bound<'py, PyAny>],
        kwargs: Option<&Bound<'py, PyDict>>,
        names: &'static [&'static str],
        required: usize,
    ) -> PyResult<Self> {
        if positional.len() > names.len() {
            return Err(PyTypeError::new_err(format!(
                "{method}() takes at most {} argument(s) after the cell/range, got {}",
                names.len(),
                positional.len()
            )));
        }
        let mut values: Vec<Option<Bound<'py, PyAny>>> = vec![None; names.len()];
        for (i, value) in positional.iter().enumerate() {
            values[i] = Some(value.clone());
        }
        if let Some(kwargs) = kwargs {
            for (key, value) in kwargs.iter() {
                let key: String = key.extract()?;
                let index = names.iter().position(|n| *n == key).ok_or_else(|| {
                    PyTypeError::new_err(format!(
                        "{method}() got an unexpected keyword argument '{key}'"
                    ))
                })?;
                if values[index].is_some() {
                    return Err(PyTypeError::new_err(format!(
                        "{method}() got multiple values for argument '{key}'"
                    )));
                }
                values[index] = Some(value);
            }
        }
        for (i, name) in names.iter().enumerate().take(required) {
            if values[i].is_none() {
                return Err(PyTypeError::new_err(format!(
                    "{method}() missing required argument: '{name}'"
                )));
            }
        }
        Ok(Params { values, method, names, _marker: std::marker::PhantomData })
    }

    /// The raw value (None if not given or given as Python None).
    pub fn raw(&self, index: usize) -> Option<&Bound<'py, PyAny>> {
        self.values[index].as_ref().filter(|v| !v.is_none())
    }

    /// The raw value even if it is Python None.
    pub fn raw_any(&self, index: usize) -> Option<&Bound<'py, PyAny>> {
        self.values[index].as_ref()
    }

    pub fn required_any(&self, index: usize) -> PyResult<&Bound<'py, PyAny>> {
        self.values[index].as_ref().ok_or_else(|| {
            PyTypeError::new_err(format!(
                "{}() missing required argument: '{}'",
                self.method, self.names[index]
            ))
        })
    }

    pub fn extract<T>(&self, index: usize) -> PyResult<Option<T>>
    where
        T: for<'b> FromPyObject<'b, 'py>,
    {
        match self.raw(index) {
            None => Ok(None),
            Some(value) => value.extract::<T>().map(Some).map_err(|e| {
                let e: PyErr = e.into();
                PyTypeError::new_err(format!(
                    "{}() argument '{}': {}",
                    self.method,
                    self.names[index],
                    e
                ))
            }),
        }
    }

    pub fn required<T>(&self, index: usize) -> PyResult<T>
    where
        T: for<'b> FromPyObject<'b, 'py>,
    {
        self.extract(index)?.ok_or_else(|| {
            PyTypeError::new_err(format!(
                "{}() missing required argument: '{}'",
                self.method, self.names[index]
            ))
        })
    }

    pub fn get_or<T>(&self, index: usize, default: T) -> PyResult<T>
    where
        T: for<'b> FromPyObject<'b, 'py>,
    {
        Ok(self.extract(index)?.unwrap_or(default))
    }
}

// ---------------------------------------------------------------------------
// Utility functions exposed to Python.
// ---------------------------------------------------------------------------

/// Convert a zero-based column index to letters: 0 -> "A".
#[pyfunction]
pub fn column_number_to_name(col: u16) -> PyResult<String> {
    if col >= MAX_COLS {
        return Err(range_err(format!("column {col} is outside 0..{MAX_COLS}")));
    }
    Ok(col_to_name(col))
}

/// Convert column letters to a zero-based index: "A" -> 0.
#[pyfunction]
pub fn column_name_to_number(name: &str) -> PyResult<u16> {
    name_to_col(name).ok_or_else(|| param_err(format!("invalid column name {name:?}")))
}

/// Zero-based (row, col) to "A1".
#[pyfunction]
#[pyo3(signature = (row, col, row_absolute=false, col_absolute=false))]
pub fn row_col_to_cell(row: u32, col: u16, row_absolute: bool, col_absolute: bool) -> PyResult<String> {
    check_cell(row, col)?;
    Ok(cell_name(row, col, row_absolute, col_absolute))
}

/// "A1" to zero-based (row, col).
#[pyfunction]
pub fn cell_to_row_col(cell: &str) -> PyResult<Cell> {
    parse_cell(cell).ok_or_else(|| param_err(format!("invalid cell reference {cell:?}")))
}

/// Zero-based range to "A1:B2" (a single cell gives "A1").
#[pyfunction]
#[pyo3(signature = (first_row, first_col, last_row, last_col, absolute=false))]
pub fn cell_range(
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    absolute: bool,
) -> PyResult<String> {
    check_cell(first_row, first_col)?;
    check_cell(last_row, last_col)?;
    Ok(range_name((first_row, first_col, last_row, last_col), absolute))
}

/// "A1:B2" to zero-based (first_row, first_col, last_row, last_col).
#[pyfunction]
pub fn range_to_row_col(range: &str) -> PyResult<CellRange> {
    parse_range(range).ok_or_else(|| param_err(format!("invalid range reference {range:?}")))
}

/// Quote a worksheet name for use in a formula if needed: "My Sheet" -> "'My Sheet'".
#[pyfunction]
pub fn quote_sheet_name(name: &str) -> String {
    rust_xlsxwriter::utility::quote_sheet_name(name)
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(column_number_to_name, m)?)?;
    m.add_function(wrap_pyfunction!(column_name_to_number, m)?)?;
    m.add_function(wrap_pyfunction!(row_col_to_cell, m)?)?;
    m.add_function(wrap_pyfunction!(cell_to_row_col, m)?)?;
    m.add_function(wrap_pyfunction!(cell_range, m)?)?;
    m.add_function(wrap_pyfunction!(range_to_row_col, m)?)?;
    m.add_function(wrap_pyfunction!(quote_sheet_name, m)?)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn columns_round_trip() {
        for (col, name) in [(0, "A"), (25, "Z"), (26, "AA"), (701, "ZZ"), (702, "AAA"), (16383, "XFD")] {
            assert_eq!(col_to_name(col), name);
            assert_eq!(name_to_col(name), Some(col));
        }
        assert_eq!(name_to_col("XFE"), None);
        assert_eq!(name_to_col(""), None);
    }

    #[test]
    fn cells_and_ranges() {
        assert_eq!(parse_cell("A1"), Some((0, 0)));
        assert_eq!(parse_cell("$B$3"), Some((2, 1)));
        assert_eq!(parse_cell("b3"), Some((2, 1)));
        assert_eq!(parse_cell("A0"), None);
        assert_eq!(parse_cell("A1048577"), None);
        assert_eq!(parse_range("A1:C4"), Some((0, 0, 3, 2)));
        assert_eq!(parse_range("Sheet1!$A$1:$C$4"), Some((0, 0, 3, 2)));
        assert_eq!(parse_range("B2"), Some((1, 1, 1, 1)));
        assert_eq!(parse_cols("B:D"), Some((1, 3)));
        assert_eq!(parse_cols("C"), Some((2, 2)));
        assert_eq!(range_name((0, 0, 3, 2), true), "$A$1:$C$4");
    }
}
