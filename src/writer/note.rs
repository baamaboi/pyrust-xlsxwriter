//! `Note`: cell notes (the yellow pop-up "comments" of older Excel versions).

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use rust_xlsxwriter::Note as RNote;

use crate::common::cellref::{split_cell, Params};
use crate::common::color::Color;
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::Format;
use crate::py_props;
use crate::writer::image::{check_alt_text, positive_size, read_child, repr_str, utf16_len};
use crate::writer::object_movement::ObjectMovement;
use crate::writer::worksheet::Worksheet;

/// Excel's limit for note author names.
const MAX_AUTHOR_LEN: usize = 52;

fn check_author(name: &str) -> PyResult<()> {
    if utf16_len(name) > MAX_AUTHOR_LEN {
        return Err(param_err(format!(
            "author {name:?} is {} characters long; Excel's limit is {MAX_AUTHOR_LEN}",
            utf16_len(name)
        )));
    }
    Ok(())
}

/// Settings of a note. `None` means "Excel's default" (Tahoma 8pt on a
/// light-yellow background, visible only on hover).
pub struct NoteSpec {
    pub text: String,
    pub author: Option<String>,
    pub add_author_prefix: bool,
    pub width: u32,
    pub height: u32,
    pub visible: Option<bool>,
    pub background_color: Option<Color>,
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_family: Option<u8>,
    pub format: Option<Py<Format>>,
    pub alt_text: String,
    pub object_movement: ObjectMovement,
}

impl NoteSpec {
    fn new(text: String) -> Self {
        NoteSpec {
            text,
            author: None,
            add_author_prefix: true,
            width: 128,
            height: 74,
            visible: None,
            background_color: None,
            font_name: None,
            font_size: None,
            font_family: None,
            format: None,
            alt_text: String::new(),
            object_movement: ObjectMovement::DontMoveOrSizeWithCells,
        }
    }

    /// A copy; with `deep` the format is copied too instead of shared.
    fn duplicate(&self, py: Python<'_>, deep: bool) -> PyResult<Self> {
        let format = match &self.format {
            None => None,
            Some(f) if deep => Some(Py::new(py, read_child(py, f)?.clone())?),
            Some(f) => Some(f.clone_ref(py)),
        };
        Ok(NoteSpec {
            text: self.text.clone(),
            author: self.author.clone(),
            add_author_prefix: self.add_author_prefix,
            width: self.width,
            height: self.height,
            visible: self.visible,
            background_color: self.background_color,
            font_name: self.font_name.clone(),
            font_size: self.font_size,
            font_family: self.font_family,
            format,
            alt_text: self.alt_text.clone(),
            object_movement: self.object_movement,
        })
    }

    fn equals(&self, py: Python<'_>, other: &NoteSpec) -> PyResult<bool> {
        let formats_equal = match (&self.format, &other.format) {
            (None, None) => true,
            (Some(a), Some(b)) => a.is(b) || read_child(py, a)?.spec == read_child(py, b)?.spec,
            _ => false,
        };
        Ok(formats_equal
            && self.text == other.text
            && self.author == other.author
            && self.add_author_prefix == other.add_author_prefix
            && self.width == other.width
            && self.height == other.height
            && self.visible == other.visible
            && self.background_color == other.background_color
            && self.font_name == other.font_name
            && self.font_size == other.font_size
            && self.font_family == other.font_family
            && self.alt_text == other.alt_text
            && self.object_movement == other.object_movement)
    }
}

/// A cell note (a pop-up comment attached to a cell).
///
/// ```python
/// ws.insert_note("B2", Note("Check this value", author="Ana", visible=True))
/// ws.insert_note("B3", "A plain-text note")
/// ```
///
/// Every property can be passed to the constructor as a keyword argument.
/// `format` replaces the default note format (Tahoma 8pt, light yellow);
/// `font_name`, `font_size`, `font_family` and `background_color` are then
/// applied on top of it.
#[pyclass(module = "pyrust_xlsxwriter")]
pub struct Note {
    pub spec: NoteSpec,
}

impl Note {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    /// Build the `rust_xlsxwriter::Note` for insertion.
    pub fn to_rust(&self, py: Python<'_>) -> PyResult<RNote> {
        let s = &self.spec;
        let mut note = RNote::new(s.text.as_str());
        if let Some(format) = &s.format {
            note = note.set_format(read_child(py, format)?.inner.clone());
        }
        if let Some(author) = &s.author {
            note = note.set_author(author.as_str());
        }
        note = note
            .add_author_prefix(s.add_author_prefix)
            .set_width(s.width)
            .set_height(s.height)
            .set_object_movement(s.object_movement.into());
        if let Some(visible) = s.visible {
            note = note.set_visible(visible);
        }
        if let Some(color) = s.background_color {
            note = note.set_background_color(color.0);
        }
        if let Some(name) = &s.font_name {
            note = note.set_font_name(name.as_str());
        }
        if let Some(size) = s.font_size {
            note = note.set_font_size(size);
        }
        if let Some(family) = s.font_family {
            note = note.set_font_family(family);
        }
        if !s.alt_text.is_empty() {
            note = note.set_alt_text(s.alt_text.as_str());
        }
        Ok(note)
    }

    fn get_author(&self) -> Option<String> {
        self.spec.author.clone()
    }

    fn set_author(&mut self, value: Option<String>) -> PyResult<()> {
        if let Some(name) = &value {
            check_author(name)?;
        }
        self.spec.author = value;
        Ok(())
    }

    fn get_width(&self) -> i64 {
        self.spec.width as i64
    }

    fn set_width(&mut self, value: i64) -> PyResult<()> {
        self.spec.width = positive_size("width", value, 1)?;
        Ok(())
    }

    fn get_height(&self) -> i64 {
        self.spec.height as i64
    }

    fn set_height(&mut self, value: i64) -> PyResult<()> {
        self.spec.height = positive_size("height", value, 1)?;
        Ok(())
    }

    fn get_font_size(&self) -> Option<f64> {
        self.spec.font_size
    }

    fn set_font_size(&mut self, value: Option<f64>) -> PyResult<()> {
        if let Some(size) = value
            && !(size.is_finite() && size > 0.0)
        {
            return Err(param_err(format!("font_size must be a positive number, not {size}")));
        }
        self.spec.font_size = value;
        Ok(())
    }

    fn get_format(&self) -> Option<Py<Format>> {
        Python::attach(|py| self.spec.format.as_ref().map(|f| f.clone_ref(py)))
    }

    fn set_format(&mut self, value: Option<Py<Format>>) -> PyResult<()> {
        self.spec.format = value;
        Ok(())
    }

    fn get_alt_text(&self) -> String {
        self.spec.alt_text.clone()
    }

    fn set_alt_text(&mut self, value: String) -> PyResult<()> {
        check_alt_text(&value)?;
        self.spec.alt_text = value;
        Ok(())
    }
}

py_props!(Note
    fields {
        text: String,
        add_author_prefix: bool,
        visible: Option<bool>,
        background_color: Option<Color>,
        font_name: Option<String>,
        font_family: Option<u8>,
        object_movement: ObjectMovement,
    }
    custom {
        author: Option<String>,
        width: i64,
        height: i64,
        font_size: Option<f64>,
        format: Option<Py<Format>>,
        alt_text: String,
    }
);

#[pymethods]
impl Note {
    /// `Note(text, **properties)`, e.g. `Note("Hi", author="Ana", width=200)`.
    #[new]
    #[pyo3(signature = (text=String::new(), **kwargs))]
    fn py_new(text: String, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut note = Note { spec: NoteSpec::new(text) };
        note.apply_kwargs(kwargs)?;
        Ok(note)
    }

    /// Replace the note text (same as setting `text`).
    fn reset_text(&mut self, text: String) {
        self.spec.text = text;
    }

    /// A copy of this note (with its own copy of `format`).
    fn copy(&self, py: Python<'_>) -> PyResult<Self> {
        Ok(Note { spec: self.spec.duplicate(py, true)? })
    }

    fn __copy__(&self, py: Python<'_>) -> PyResult<Self> {
        Ok(Note { spec: self.spec.duplicate(py, false)? })
    }

    fn __deepcopy__(&self, py: Python<'_>, _memo: &Bound<'_, PyAny>) -> PyResult<Self> {
        self.copy(py)
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> PyResult<bool> {
        match other.cast::<Note>() {
            Ok(other) => {
                let other = other.try_borrow().map_err(|_| param_err("the Note is in use elsewhere"))?;
                self.spec.equals(other.py(), &other.spec)
            }
            Err(_) => Ok(false),
        }
    }

    fn __repr__(&self) -> String {
        let mut parts = vec![repr_str(&self.spec.text)];
        if let Some(author) = &self.spec.author {
            parts.push(format!("author={}", repr_str(author)));
        }
        if let Some(visible) = self.spec.visible {
            parts.push(format!("visible={}", if visible { "True" } else { "False" }));
        }
        format!("Note({})", parts.join(", "))
    }
}

#[pymethods]
impl Worksheet {
    // ---- area: notes ----

    /// Add a note to a cell: `insert_note("B2", Note("text", author="Ana"))`
    /// or simply `insert_note("B2", "text")`.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_note(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_note", rest, kwargs, &["note"], 1)?;
        let value = p.required_any(0)?;
        let note = if let Ok(note) = value.cast::<Note>() {
            let note = note.try_borrow().map_err(|_| param_err("the Note is in use elsewhere"))?;
            note.to_rust(value.py())?
        } else if let Ok(text) = value.cast::<PyString>() {
            RNote::new(text.to_str()?)
        } else {
            return Err(PyTypeError::new_err(format!(
                "insert_note() expects a Note or a str, not {}",
                value.get_type().name()?
            )));
        };
        self.inner.insert_note(row, col, &note).py()?;
        Ok(())
    }

    /// Show every note on the worksheet, not only on hover.
    #[pyo3(signature = (enable=true))]
    fn show_all_notes(&mut self, enable: bool) {
        self.inner.show_all_notes(enable);
    }

    /// Set the author used for notes that don't set their own.
    fn set_default_note_author(&mut self, name: &str) -> PyResult<()> {
        check_author(name)?;
        self.inner.set_default_note_author(name);
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Note>()
}
