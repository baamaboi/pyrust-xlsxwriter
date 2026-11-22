//! `Image`: pictures placed over cells, embedded in cells, used as the
//! worksheet background or in headers/footers.
//!
//! The image bytes are kept by the Python object; the `rust_xlsxwriter::Image`
//! is built from them (plus the settings in `ImageSpec`) when the image is
//! inserted, so one `Image` can be inserted many times with different settings.
//!
//! This module also holds small helpers shared by the other embedded-object
//! modules (notes, shapes, buttons): URL arguments, alt-text checks and
//! object offsets.

use std::path::PathBuf;

use pyo3::buffer::PyBuffer;
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyByteArray, PyBytes, PyDict, PyString, PyTuple};
use pyo3::PyClass;
use rust_xlsxwriter::{
    HeaderImagePosition as RHeaderImagePosition, Image as RImage, Shape as RShape, XlsxError,
};

use crate::common::bytes::SharedBytes;
use crate::common::cellref::{split_cell, Params};
use crate::common::errors::{file_err, new_err, param_err, ErrorKind, IntoPyResult};
use crate::common::format::Format;
use crate::common::value::Url;
use crate::py_props;
use crate::writer::object_movement::ObjectMovement;
use crate::writer::worksheet::Worksheet;

crate::py_enum!(
    /// Where a header/footer image goes. The header/footer text must contain
    /// a matching `&[Picture]` (or `&G`) placeholder in that section, e.g.
    /// `ws.header = "&C&[Picture]"` for `HeaderImagePosition.Center`.
    HeaderImagePosition => RHeaderImagePosition { Left, Center, Right }
);

// ---------------------------------------------------------------------------
// Helpers shared by images, notes, shapes and buttons.
// ---------------------------------------------------------------------------

/// Length of a string in UTF-16 code units (Excel's unit for length limits).
pub(crate) fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

/// A Python-style repr of a string (`'text'`).
pub(crate) fn repr_str(text: &str) -> String {
    Python::attach(|py| {
        PyString::new(py, text)
            .repr()
            .map(|r| r.to_string())
            .unwrap_or_else(|_| format!("{text:?}"))
    })
}

/// Excel limits alternative text to 255 characters.
pub(crate) fn check_alt_text(text: &str) -> PyResult<()> {
    if utf16_len(text) > 255 {
        return Err(param_err(format!(
            "alt_text is {} characters long; Excel's limit is 255",
            utf16_len(text)
        )));
    }
    Ok(())
}

/// Borrow a child object without panicking if it is already borrowed mutably.
pub(crate) fn read_child<'py, T: PyClass>(py: Python<'py>, obj: &'py Py<T>) -> PyResult<PyRef<'py, T>> {
    obj.try_borrow(py)
        .map_err(|_| param_err(format!("a {} object is in use elsewhere", <T as PyClass>::NAME)))
}

/// A hyperlink argument: a `Url` or a plain link string.
#[derive(Clone, Debug, PartialEq)]
pub struct UrlArg(pub Url);

impl<'a, 'py> FromPyObject<'a, 'py> for UrlArg {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(url) = obj.cast::<Url>() {
            return Ok(UrlArg(url.get().clone()));
        }
        if let Ok(text) = obj.cast::<PyString>() {
            return Ok(UrlArg(Url { link: text.to_str()?.to_string(), text: None, tip: None }));
        }
        Err(PyTypeError::new_err(format!(
            "expected a Url or a link string, not {}",
            obj.get_type().name()?
        )))
    }
}

impl<'py> IntoPyObject<'py> for UrlArg {
    type Target = Url;
    type Output = Bound<'py, Url>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Bound::new(py, self.0)
    }
}

/// Check a link for an image or shape (same rules as `rust_xlsxwriter`).
pub(crate) fn check_object_url(url: &Url) -> PyResult<()> {
    RShape::textbox().set_url(url.to_rust()).py()?;
    Ok(())
}

/// Read an `x_offset`/`y_offset` argument (pixels, >= 0).
pub(crate) fn offset_arg(p: &Params<'_, '_>, index: usize, name: &str) -> PyResult<u32> {
    let value: i64 = p.get_or(index, 0)?;
    if !(0..=MAX_OBJECT_PIXELS as i64).contains(&value) {
        return Err(param_err(format!(
            "{name} must be between 0 and {MAX_OBJECT_PIXELS} pixels, not {value}"
        )));
    }
    Ok(value as u32)
}

/// Largest object offset or size accepted, in pixels. rust_xlsxwriter walks
/// the columns under an object with a 16-bit counter, so far larger values
/// would overflow it when the file is saved.
pub(crate) const MAX_OBJECT_PIXELS: u32 = 1_000_000;

/// Check a pixel size (width/height of an object).
pub(crate) fn positive_size(name: &str, value: i64, minimum: i64) -> PyResult<u32> {
    if value < minimum || value > MAX_OBJECT_PIXELS as i64 {
        return Err(param_err(format!(
            "{name} must be between {minimum} and {MAX_OBJECT_PIXELS} pixels, not {value}"
        )));
    }
    Ok(value as u32)
}

// ---------------------------------------------------------------------------
// Image
// ---------------------------------------------------------------------------

/// Settings of an image (everything except the picture itself).
#[derive(Clone, Debug, PartialEq)]
pub struct ImageSpec {
    pub scale_width: f64,
    pub scale_height: f64,
    pub alt_text: String,
    pub decorative: bool,
    pub object_movement: ObjectMovement,
    pub url: Option<Url>,
}

impl Default for ImageSpec {
    fn default() -> Self {
        ImageSpec {
            scale_width: 1.0,
            scale_height: 1.0,
            alt_text: String::new(),
            decorative: false,
            object_movement: ObjectMovement::MoveButDontSizeWithCells,
            url: None,
        }
    }
}

/// An image (PNG, JPEG, GIF or BMP) to insert into a worksheet.
///
/// ```python
/// logo = Image("logo.png", scale_width=0.5, scale_height=0.5, alt_text="Logo")
/// ws.insert_image("B2", logo)
/// ws.embed_image("D2", Image.from_bytes(png_bytes))
/// ```
///
/// Every property can be passed to the constructor as a keyword argument.
/// `width`/`height` and `scale_width`/`scale_height` are two views of the
/// same setting: the displayed size is the pixel size times the scale.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct Image {
    pub spec: ImageSpec,
    /// The image file's bytes, shared by copies of this Image (and, for
    /// images made from Python `bytes`, with that `bytes` object).
    data: SharedBytes,
    /// Name used for the image in header/footer VML (the file stem).
    name: String,
    pixel_width: f64,
    pixel_height: f64,
    width_dpi: f64,
    height_dpi: f64,
}

fn image_error(error: XlsxError, source: &str) -> PyErr {
    match error {
        XlsxError::UnknownImageType => new_err(
            ErrorKind::Image,
            format!("{source}: unknown image type; supported formats are PNG, JPEG, GIF and BMP"),
        ),
        XlsxError::ImageDimensionError => new_err(
            ErrorKind::Image,
            format!("{source}: the image dimensions are zero or could not be read"),
        ),
        XlsxError::IoError(e) => file_err(format!("{source}: {e}")),
        other => crate::common::errors::xlsx_err(other),
    }
}

/// Bytes from `bytes` (shared in place), `bytearray`, any byte buffer
/// (`memoryview`, ...) or a binary file object (`getvalue()` / `read()`).
fn read_bytes(data: &Bound<'_, PyAny>) -> PyResult<SharedBytes> {
    if data.is_instance_of::<PyBytes>() || data.is_instance_of::<PyByteArray>() {
        return SharedBytes::from_py(data);
    }
    if data.is_instance_of::<PyString>() {
        return Err(PyTypeError::new_err(
            "expected image data (bytes); use Image(path) to load an image file",
        ));
    }
    if let Ok(buffer) = PyBuffer::<u8>::get(data) {
        return Ok(SharedBytes::from_vec(buffer.to_vec(data.py())?));
    }
    for method in ["getvalue", "read"] {
        if data.hasattr(method)? {
            let content = data.call_method0(method)?;
            if content.is_instance_of::<PyBytes>() {
                return SharedBytes::from_py(&content);
            }
            return Err(PyTypeError::new_err(format!(
                "{method}() returned {}, expected bytes (open the file in binary mode)",
                content.get_type().name()?
            )));
        }
    }
    Err(PyTypeError::new_err(format!(
        "expected bytes, a byte buffer or a binary file object, not {}",
        data.get_type().name()?
    )))
}

impl Image {
    /// Read and check an image from raw bytes.
    pub fn from_data(data: SharedBytes, name: String, source: &str) -> PyResult<Self> {
        let probe = RImage::new_from_buffer(&data).map_err(|e| image_error(e, source))?;
        Ok(Image {
            spec: ImageSpec::default(),
            pixel_width: probe.width(),
            pixel_height: probe.height(),
            width_dpi: probe.width_dpi(),
            height_dpi: probe.height_dpi(),
            data,
            name,
        })
    }

    /// Load an image file.
    pub fn from_path(path: &PathBuf) -> PyResult<Self> {
        let source = path.display().to_string();
        let data = std::fs::read(path)
            .map_err(|e| file_err(format!("cannot read image {source:?}: {e}")))?;
        let name = path
            .file_stem()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_else(|| "image".to_string());
        Image::from_data(SharedBytes::from_vec(data), name, &source)
    }

    /// Build the `rust_xlsxwriter::Image` for insertion.
    pub fn to_rust(&self) -> PyResult<RImage> {
        let s = &self.spec;
        self.check_display_size(s.scale_width, s.scale_height)?;
        let mut image = RImage::new_from_buffer(&self.data)
            .map_err(|e| image_error(e, "image"))?
            .set_vml_name(self.name.as_str())
            .set_scale_width(s.scale_width)
            .set_scale_height(s.scale_height)
            .set_decorative(s.decorative)
            .set_object_movement(s.object_movement.into());
        if !s.alt_text.is_empty() {
            image = image.set_alt_text(s.alt_text.as_str());
        }
        if let Some(url) = &s.url {
            image = image.set_url(url.to_rust()).py()?;
        }
        Ok(image)
    }

    fn on_change(&mut self) -> PyResult<()> {
        Ok(())
    }

    fn check_scale(name: &str, value: f64) -> PyResult<f64> {
        if !(value.is_finite() && value > 0.0) {
            return Err(param_err(format!("{name} must be a positive number, not {value}")));
        }
        Ok(value)
    }

    /// Check that the image, displayed with these scales, isn't absurdly large.
    fn check_display_size(&self, scale_width: f64, scale_height: f64) -> PyResult<()> {
        let width = self.pixel_width * scale_width * 96.0 / self.width_dpi;
        let height = self.pixel_height * scale_height * 96.0 / self.height_dpi;
        let max = MAX_OBJECT_PIXELS as f64;
        if width > max || height > max {
            return Err(param_err(format!(
                "the image would be displayed at {width:.0}x{height:.0} pixels; the limit is {MAX_OBJECT_PIXELS}"
            )));
        }
        Ok(())
    }

    fn get_width(&self) -> f64 {
        self.pixel_width * self.spec.scale_width
    }

    fn set_width(&mut self, value: f64) -> PyResult<()> {
        let scale = Self::check_scale("width", value)? / self.pixel_width;
        self.check_display_size(scale, self.spec.scale_height)?;
        self.spec.scale_width = scale;
        Ok(())
    }

    fn get_height(&self) -> f64 {
        self.pixel_height * self.spec.scale_height
    }

    fn set_height(&mut self, value: f64) -> PyResult<()> {
        let scale = Self::check_scale("height", value)? / self.pixel_height;
        self.check_display_size(self.spec.scale_width, scale)?;
        self.spec.scale_height = scale;
        Ok(())
    }

    fn get_scale_width(&self) -> f64 {
        self.spec.scale_width
    }

    fn set_scale_width(&mut self, value: f64) -> PyResult<()> {
        let scale = Self::check_scale("scale_width", value)?;
        self.check_display_size(scale, self.spec.scale_height)?;
        self.spec.scale_width = scale;
        Ok(())
    }

    fn get_scale_height(&self) -> f64 {
        self.spec.scale_height
    }

    fn set_scale_height(&mut self, value: f64) -> PyResult<()> {
        let scale = Self::check_scale("scale_height", value)?;
        self.check_display_size(self.spec.scale_width, scale)?;
        self.spec.scale_height = scale;
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

    fn get_url(&self) -> Option<UrlArg> {
        self.spec.url.clone().map(UrlArg)
    }

    fn set_url(&mut self, value: Option<UrlArg>) -> PyResult<()> {
        if let Some(url) = &value {
            check_object_url(&url.0)?;
        }
        self.spec.url = value.map(|u| u.0);
        Ok(())
    }

    fn kind(&self) -> &'static str {
        let d = &self.data;
        if d.len() >= 4 && &d[1..4] == b"PNG" {
            "png"
        } else if d.starts_with(&[0xFF, 0xD8]) {
            "jpeg"
        } else if d.starts_with(b"BM") {
            "bmp"
        } else if d.starts_with(b"GIF8") {
            "gif"
        } else {
            "unknown"
        }
    }
}

py_props!(Image
    fields {
        decorative: bool,
        object_movement: ObjectMovement,
    }
    custom {
        width: f64,
        height: f64,
        scale_width: f64,
        scale_height: f64,
        alt_text: String,
        url: Option<UrlArg>,
    }
);

#[pymethods]
impl Image {
    /// Load an image file (PNG, JPEG, GIF or BMP). Keyword arguments set
    /// properties, e.g. `Image("logo.png", alt_text="Logo")`.
    #[new]
    #[pyo3(signature = (path, **kwargs))]
    fn py_new(path: PathBuf, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut image = Image::from_path(&path)?;
        image.apply_kwargs(kwargs)?;
        Ok(image)
    }

    /// Create an image from its bytes (or a byte buffer / binary file object).
    #[staticmethod]
    #[pyo3(signature = (data, **kwargs))]
    fn from_bytes(data: &Bound<'_, PyAny>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let mut image = Image::from_data(read_bytes(data)?, "image".to_string(), "image data")?;
        image.apply_kwargs(kwargs)?;
        Ok(image)
    }

    /// Same as `from_bytes()` (the name used by rust_xlsxwriter).
    #[staticmethod]
    #[pyo3(signature = (data, **kwargs))]
    fn from_buffer(data: &Bound<'_, PyAny>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Image::from_bytes(data, kwargs)
    }

    /// Scale the image to fit inside `width` x `height` pixels (keeping the
    /// aspect ratio if requested). Sets `scale_width` and `scale_height`.
    #[pyo3(signature = (width, height, keep_aspect_ratio=false))]
    fn scale_to_size(&mut self, width: f64, height: f64, keep_aspect_ratio: bool) -> PyResult<()> {
        Self::check_scale("width", width)?;
        Self::check_scale("height", height)?;
        let mut scale_width = (width / self.pixel_width) * (self.width_dpi / 96.0);
        let mut scale_height = (height / self.pixel_height) * (self.height_dpi / 96.0);
        if keep_aspect_ratio {
            let scale = scale_width.min(scale_height);
            scale_width = scale;
            scale_height = scale;
        }
        self.check_display_size(scale_width, scale_height)?;
        self.spec.scale_width = scale_width;
        self.spec.scale_height = scale_height;
        Ok(())
    }

    /// The image width in pixels (as stored in the file).
    #[getter]
    fn pixel_width(&self) -> f64 {
        self.pixel_width
    }

    /// The image height in pixels (as stored in the file).
    #[getter]
    fn pixel_height(&self) -> f64 {
        self.pixel_height
    }

    /// Horizontal resolution in dots per inch (96 if the file doesn't say).
    #[getter]
    fn width_dpi(&self) -> f64 {
        self.width_dpi
    }

    /// Vertical resolution in dots per inch (96 if the file doesn't say).
    #[getter]
    fn height_dpi(&self) -> f64 {
        self.height_dpi
    }

    /// The image format: "png", "jpeg", "gif" or "bmp".
    #[getter]
    fn image_type(&self) -> &'static str {
        self.kind()
    }

    /// The raw image bytes.
    #[getter]
    fn data<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.data)
    }

    /// A copy of this image (the bytes are shared).
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
        match other.cast::<Image>() {
            Ok(other) => other.try_borrow().is_ok_and(|other| {
                self.spec == other.spec
                    && self.name == other.name
                    && (self.data.same_buffer(&other.data) || self.data[..] == other.data[..])
            }),
            Err(_) => false,
        }
    }

    fn __repr__(&self) -> String {
        format!(
            "<Image {} {}x{} scale=({}, {})>",
            self.kind(),
            self.pixel_width,
            self.pixel_height,
            self.spec.scale_width,
            self.spec.scale_height
        )
    }
}

/// An image argument: an `Image`, or a path (str / os.PathLike) to load.
fn image_arg(value: &Bound<'_, PyAny>) -> PyResult<RImage> {
    if let Ok(image) = value.cast::<Image>() {
        let image = image
            .try_borrow()
            .map_err(|_| param_err("the Image is in use elsewhere"))?;
        return image.to_rust();
    }
    if value.is_instance_of::<PyBytes>() || value.is_instance_of::<PyByteArray>() {
        return Err(PyTypeError::new_err(
            "pass image bytes as Image.from_bytes(data), not as bytes",
        ));
    }
    match value.extract::<PathBuf>() {
        Ok(path) => Image::from_path(&path)?.to_rust(),
        Err(_) => Err(PyTypeError::new_err(format!(
            "expected an Image or a path to an image file, not {}",
            value.get_type().name()?
        ))),
    }
}

/// Run `f`, turning a panic inside rust_xlsxwriter into `None` (without
/// printing the default panic message).
///
/// `Worksheet::insert_image_fit_to_cell_centered()` subtracts the image size
/// from the cell size with unsigned arithmetic, which overflows for hidden or
/// zero-size cells. The cell sizes aren't readable from the Rust worksheet,
/// so the overflow is caught instead. The panic hook is process-wide; it is
/// swapped only for this short call made while holding the GIL.
fn without_panic<R>(f: impl FnOnce() -> R) -> Option<R> {
    let hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(f));
    std::panic::set_hook(hook);
    result.ok()
}

fn header_image_error(error: XlsxError, what: &str, position: HeaderImagePosition) -> PyErr {
    match error {
        XlsxError::ParameterError(message) => {
            let code = match position {
                HeaderImagePosition::Left => "&L",
                HeaderImagePosition::Center => "&C",
                HeaderImagePosition::Right => "&R",
            };
            param_err(format!(
                "{message}. Set the {what} text first, e.g. ws.{what} = \"{code}&[Picture]\""
            ))
        }
        other => crate::common::errors::xlsx_err(other),
    }
}

#[pymethods]
impl Worksheet {
    // ---- area: images ----

    /// Insert an image over the cells, with its top-left corner in the cell
    /// (plus optional pixel offsets): `insert_image("B2", image)` or
    /// `insert_image(1, 1, image, x_offset=5, y_offset=5)`. `image` may also
    /// be a path to an image file.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_image(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_image", rest, kwargs, &["image", "x_offset", "y_offset"], 1)?;
        let image = image_arg(p.required_any(0)?)?;
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner.insert_image_with_offset(row, col, &image, x_offset, y_offset).py()?;
        Ok(())
    }

    /// Insert an image with pixel offsets from the top-left corner of the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_image_with_offset(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("insert_image_with_offset", rest, kwargs, &["image", "x_offset", "y_offset"], 3)?;
        let image = image_arg(p.required_any(0)?)?;
        let x_offset = offset_arg(&p, 1, "x_offset")?;
        let y_offset = offset_arg(&p, 2, "y_offset")?;
        self.inner.insert_image_with_offset(row, col, &image, x_offset, y_offset).py()?;
        Ok(())
    }

    /// Embed an image inside a cell (Excel 365's "Place in Cell"). The image
    /// scales with the cell. An optional `format` applies to the cell.
    #[pyo3(signature = (*args, **kwargs))]
    fn embed_image(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("embed_image", rest, kwargs, &["image", "format"], 1)?;
        let image = image_arg(p.required_any(0)?)?;
        let format: Option<PyRef<'_, Format>> = p.extract(1)?;
        match format {
            Some(format) => self.inner.embed_image_with_format(row, col, &image, &format.inner),
            None => self.inner.embed_image(row, col, &image),
        }
        .py()?;
        Ok(())
    }

    /// Insert an image scaled to fit the cell. With `keep_aspect_ratio`
    /// (default) the image keeps its proportions; with `centered` it is also
    /// centered in the cell. The cell must not be hidden (centering an image
    /// in a hidden row/column raises ParameterError).
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_image_fit_to_cell(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new(
            "insert_image_fit_to_cell",
            rest,
            kwargs,
            &["image", "centered", "keep_aspect_ratio"],
            1,
        )?;
        let image = image_arg(p.required_any(0)?)?;
        let centered: bool = p.get_or(1, false)?;
        let keep_aspect_ratio: bool = p.get_or(2, true)?;
        if centered && keep_aspect_ratio {
            let ws = &mut self.inner;
            match without_panic(|| ws.insert_image_fit_to_cell_centered(row, col, &image).map(|_| ())) {
                Some(result) => result.py()?,
                None => {
                    return Err(param_err(format!(
                        "cannot center an image in cell ({row}, {col}): the row or column is hidden or has zero size"
                    )))
                }
            }
        } else {
            self.inner.insert_image_fit_to_cell(row, col, &image, keep_aspect_ratio).py()?;
        }
        Ok(())
    }

    /// Set a background image for the worksheet (tiled by Excel, screen only).
    fn insert_background_image(&mut self, image: &Bound<'_, PyAny>) -> PyResult<()> {
        let image = image_arg(image)?;
        self.inner.insert_background_image(&image);
        Ok(())
    }

    /// Add an image to the page header. The `header` text must contain a
    /// `&[Picture]` (or `&G`) placeholder in the matching section, e.g.
    /// `ws.header = "&C&[Picture]"` then
    /// `ws.set_header_image(image, HeaderImagePosition.Center)`.
    fn set_header_image(&mut self, image: &Bound<'_, PyAny>, position: HeaderImagePosition) -> PyResult<()> {
        let image = image_arg(image)?;
        self.inner
            .set_header_image(&image, position.into())
            .map_err(|e| header_image_error(e, "header", position))?;
        Ok(())
    }

    /// Add an image to the page footer. The `footer` text must contain a
    /// `&[Picture]` (or `&G`) placeholder in the matching section, e.g.
    /// `ws.footer = "&L&[Picture]"` then
    /// `ws.set_footer_image(image, HeaderImagePosition.Left)`.
    fn set_footer_image(&mut self, image: &Bound<'_, PyAny>, position: HeaderImagePosition) -> PyResult<()> {
        let image = image_arg(image)?;
        self.inner
            .set_footer_image(&image, position.into())
            .map_err(|e| header_image_error(e, "footer", position))?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Image>()?;
    m.add_class::<HeaderImagePosition>()?;
    Ok(())
}
