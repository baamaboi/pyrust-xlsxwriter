//! `Button` (form-control buttons that run a VBA macro) and checkboxes.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyTuple};
use rust_xlsxwriter::Button as RButton;

use crate::common::cellref::{split_cell, Params};
use crate::common::errors::{param_err, IntoPyResult};
use crate::common::format::Format;
use crate::py_props;
use crate::writer::image::{check_alt_text, offset_arg, positive_size, repr_str, utf16_len};
use crate::writer::object_movement::ObjectMovement;
use crate::writer::worksheet::Worksheet;

/// Settings of a button.
#[derive(Clone, Debug, PartialEq)]
pub struct ButtonSpec {
    pub caption: Option<String>,
    pub macro_name: Option<String>,
    pub width: u32,
    pub height: u32,
    pub alt_text: String,
    pub object_movement: ObjectMovement,
}

impl Default for ButtonSpec {
    fn default() -> Self {
        ButtonSpec {
            caption: None,
            macro_name: None,
            width: 64,
            height: 20,
            alt_text: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
        }
    }
}

/// A form-control button, usually tied to a VBA macro.
///
/// ```python
/// wb.add_vba_project("vbaProject.bin")
/// ws.insert_button("B2", Button(caption="Run", macro_name="say_hello"))
/// wb.save("macros.xlsm")
/// ```
///
/// Every property can be passed to the constructor as a keyword argument.
/// Without a caption Excel shows "Button 1", "Button 2", ...
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct Button {
    pub spec: ButtonSpec,
}

impl Button {
    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    /// Build the `rust_xlsxwriter::Button` for insertion.
    pub fn to_rust(&self) -> RButton {
        let s = &self.spec;
        let mut button = RButton::new()
            .set_width(s.width)
            .set_height(s.height)
            .set_object_movement(s.object_movement.into());
        if let Some(caption) = &s.caption {
            button = button.set_caption(caption.as_str());
        }
        if let Some(name) = &s.macro_name {
            button = button.set_macro(name.as_str());
        }
        if !s.alt_text.is_empty() {
            button = button.set_alt_text(s.alt_text.as_str());
        }
        button
    }

    fn get_caption(&self) -> Option<String> {
        self.spec.caption.clone()
    }

    fn set_caption(&mut self, value: Option<String>) -> PyResult<()> {
        if let Some(caption) = &value
            && utf16_len(caption) > 255
        {
            return Err(param_err(format!(
                "caption is {} characters long; Excel's limit is 255",
                utf16_len(caption)
            )));
        }
        self.spec.caption = value;
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

    fn get_alt_text(&self) -> String {
        self.spec.alt_text.clone()
    }

    fn set_alt_text(&mut self, value: String) -> PyResult<()> {
        check_alt_text(&value)?;
        self.spec.alt_text = value;
        Ok(())
    }
}

py_props!(Button
    fields {
        macro_name: Option<String>,
        object_movement: ObjectMovement,
    }
    custom {
        caption: Option<String>,
        width: i64,
        height: i64,
        alt_text: String,
    }
);

#[pymethods]
impl Button {
    /// `Button(**properties)`, e.g. `Button(caption="Run", macro_name="say_hello")`.
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut button = Button { spec: ButtonSpec::default() };
        button.apply_kwargs(kwargs)?;
        Ok(button)
    }

    /// A copy of this button.
    fn copy(&self) -> Self {
        self.clone()
    }

    fn __copy__(&self) -> Self {
        self.clone()
    }

    fn __deepcopy__(&self, _memo: &Bound<'_, PyAny>) -> Self {
        self.clone()
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        match other.cast::<Button>() {
            Ok(other) => other.try_borrow().is_ok_and(|other| self.spec == other.spec),
            Err(_) => false,
        }
    }

    fn __repr__(&self) -> String {
        let mut parts = Vec::new();
        if let Some(caption) = &self.spec.caption {
            parts.push(format!("caption={}", repr_str(caption)));
        }
        if let Some(name) = &self.spec.macro_name {
            parts.push(format!("macro_name={}", repr_str(name)));
        }
        format!("Button({})", parts.join(", "))
    }
}

#[pymethods]
impl Worksheet {
    // ---- area: buttons and checkboxes ----

    /// Insert a button: `insert_button("B2", Button(caption="Run", macro_name="run"))`.
    /// Without `button` a default button is inserted. Save as `.xlsm` when
    /// the workbook has a VBA project.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_button(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_button", rest, kwargs, &["button", "x_offset", "y_offset"], 0)?;
        let button = match p.extract::<PyRef<'_, Button>>(0)? {
            Some(button) => button.to_rust(),
            None => RButton::new(),
        };
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner.insert_button_with_offset(row, col, &button, x_offset, y_offset).py()?;
        Ok(())
    }

    /// Insert a button with pixel offsets from the top-left corner of the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_button_with_offset(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_button_with_offset", rest, kwargs, &["button", "x_offset", "y_offset"], 3)?;
        let button: PyRef<'_, Button> = p.required(0)?;
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner
            .insert_button_with_offset(row, col, &button.to_rust(), x_offset, y_offset)
            .py()?;
        Ok(())
    }

    /// Insert a checkbox (Excel 365) showing a boolean cell value:
    /// `insert_checkbox("B2", True)`. An optional `format` is applied too.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_checkbox(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_checkbox", rest, kwargs, &["checked", "format"], 1)?;
        let checked: bool = p.required(0)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        match format {
            Some(format) => self.inner.insert_checkbox_with_format(row, col, checked, &format.inner),
            None => self.inner.insert_checkbox(row, col, checked),
        }
        .py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Button>()
}
