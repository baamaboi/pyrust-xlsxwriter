//! Sparklines: `Sparkline`, `SparklineType`, `SparklineEmptyCells` and the
//! `Worksheet.add_sparkline()` / `add_sparkline_group()` methods.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use rust_xlsxwriter::{
    ChartEmptyCells as RChartEmptyCells, Sparkline as RSparkline, SparklineType as RSparklineType,
};

use crate::common::cellref::{check_cell, range_name, split_cell, split_range, Params};
use crate::common::color::Color;
use crate::common::errors::{param_err, IntoPyResult};
use crate::writer::conditional_format::spec_methods;
use crate::writer::worksheet::Worksheet;
use crate::{py_enum, py_props};

py_enum!(
    /// Sparkline types.
    SparklineType => RSparklineType { Line, Column, WinLoss => WinLose }
);

py_enum!(
    /// How a sparkline shows empty cells.
    SparklineEmptyCells => RChartEmptyCells { Gaps, Zero, Connected }
);

/// A worksheet range given as `"Sheet1!A1:E1"` (the sheet is optional) or
/// a `(sheet_name, first_row, first_col, last_row, last_col)` tuple. Read
/// back as a string.
#[derive(Clone, Debug, PartialEq)]
pub struct RangeText(pub String);

impl RangeText {
    /// The range with a sheet name, using `sheet` when it has none.
    pub fn qualified(&self, sheet: &str) -> String {
        if self.0.contains('!') {
            self.0.clone()
        } else {
            format!("{}!{}", rust_xlsxwriter::utility::quote_sheet_name(sheet), self.0)
        }
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for RangeText {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> Result<Self, Self::Error> {
        if let Ok(s) = obj.cast::<PyString>() {
            let text = s.to_str()?.trim().to_string();
            if text.is_empty() {
                return Err(param_err("a range must not be empty"));
            }
            return Ok(RangeText(text));
        }
        if let Ok(t) = obj.cast::<PyTuple>()
            && t.len() == 5 {
                let sheet: String = t.get_item(0)?.extract()?;
                let r1: u32 = t.get_item(1)?.extract()?;
                let c1: u16 = t.get_item(2)?.extract()?;
                let r2: u32 = t.get_item(3)?.extract()?;
                let c2: u16 = t.get_item(4)?.extract()?;
                check_cell(r1, c1)?;
                check_cell(r2, c2)?;
                if r1 > r2 || c1 > c2 {
                    return Err(crate::common::errors::range_err(
                        "range is reversed: the first cell must be above/left of the last",
                    ));
                }
                let range = format!(
                    "{}!{}",
                    rust_xlsxwriter::utility::quote_sheet_name(&sheet),
                    range_name((r1, c1, r2, c2), true)
                );
                return Ok(RangeText(range));
            }
        Err(PyTypeError::new_err(format!(
            "expected a range string like \"Sheet1!A1:E1\" or a (sheet_name, first_row, first_col, last_row, last_col) tuple, not {}",
            obj.get_type().name()?
        )))
    }
}

impl<'py> IntoPyObject<'py> for RangeText {
    type Target = PyString;
    type Output = Bound<'py, PyString>;
    type Error = std::convert::Infallible;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(PyString::new(py, &self.0))
    }
}

/// Plain description of a sparkline.
#[derive(Clone, Debug, PartialEq)]
pub struct SparklineSpec {
    pub range: Option<RangeText>,
    pub sparkline_type: SparklineType,
    pub style: u8,
    pub high_point: bool,
    pub low_point: bool,
    pub first_point: bool,
    pub last_point: bool,
    pub negative_points: bool,
    pub markers: bool,
    pub axis: bool,
    pub hidden_data: bool,
    pub empty_cells: SparklineEmptyCells,
    pub right_to_left: bool,
    pub sparkline_color: Option<Color>,
    pub high_point_color: Option<Color>,
    pub low_point_color: Option<Color>,
    pub first_point_color: Option<Color>,
    pub last_point_color: Option<Color>,
    pub negative_points_color: Option<Color>,
    pub markers_color: Option<Color>,
    pub line_weight: Option<f64>,
    pub custom_min: Option<f64>,
    pub custom_max: Option<f64>,
    pub group_min: bool,
    pub group_max: bool,
    pub date_range: Option<RangeText>,
    pub column_order: bool,
}

impl Default for SparklineSpec {
    fn default() -> Self {
        SparklineSpec {
            range: None,
            sparkline_type: SparklineType::Line,
            style: 1,
            high_point: false,
            low_point: false,
            first_point: false,
            last_point: false,
            negative_points: false,
            markers: false,
            axis: false,
            hidden_data: false,
            empty_cells: SparklineEmptyCells::Gaps,
            right_to_left: false,
            sparkline_color: None,
            high_point_color: None,
            low_point_color: None,
            first_point_color: None,
            last_point_color: None,
            negative_points_color: None,
            markers_color: None,
            line_weight: None,
            custom_min: None,
            custom_max: None,
            group_min: false,
            group_max: false,
            date_range: None,
            column_order: false,
        }
    }
}

impl SparklineSpec {
    fn validate(&self) -> PyResult<()> {
        if !(1..=36).contains(&self.style) {
            return Err(param_err(format!("sparkline style {} is outside Excel's range 1..=36", self.style)));
        }
        if let Some(weight) = self.line_weight
            && !(weight.is_finite() && weight > 0.0) {
                return Err(param_err(format!("line_weight {weight} must be a positive number")));
            }
        for (name, value) in [("custom_min", self.custom_min), ("custom_max", self.custom_max)] {
            if let Some(v) = value
                && !v.is_finite() {
                    return Err(param_err(format!("{name} {v} must be a finite number")));
                }
        }
        Ok(())
    }

    /// Build the `rust_xlsxwriter::Sparkline`; ranges without a sheet name
    /// refer to `sheet`.
    pub fn to_rust(&self, sheet: &str) -> PyResult<RSparkline> {
        self.validate()?;
        let range = self
            .range
            .as_ref()
            .ok_or_else(|| param_err("the sparkline has no data range: set Sparkline(range=\"A1:E1\")"))?;
        // The style sets every colour, so it goes first; the colour setters
        // also turn on their point, so the show flags go after them.
        let mut s = RSparkline::new()
            .set_style(self.style)
            .set_range(range.qualified(sheet).as_str())
            .set_type(self.sparkline_type.into());
        if let Some(c) = self.sparkline_color {
            s = s.set_sparkline_color(c);
        }
        if let Some(c) = self.high_point_color {
            s = s.set_high_point_color(c);
        }
        if let Some(c) = self.low_point_color {
            s = s.set_low_point_color(c);
        }
        if let Some(c) = self.first_point_color {
            s = s.set_first_point_color(c);
        }
        if let Some(c) = self.last_point_color {
            s = s.set_last_point_color(c);
        }
        if let Some(c) = self.negative_points_color {
            s = s.set_negative_points_color(c);
        }
        if let Some(c) = self.markers_color {
            s = s.set_markers_color(c);
        }
        s = s
            .show_high_point(self.high_point)
            .show_low_point(self.low_point)
            .show_first_point(self.first_point)
            .show_last_point(self.last_point)
            .show_negative_points(self.negative_points)
            .show_markers(self.markers)
            .show_axis(self.axis)
            .show_hidden_data(self.hidden_data)
            .show_empty_cells_as(self.empty_cells.into())
            .set_right_to_left(self.right_to_left)
            .set_column_order(self.column_order);
        if let Some(weight) = self.line_weight {
            s = s.set_line_weight(weight);
        }
        if let Some(v) = self.custom_min {
            s = s.set_custom_min(v);
        }
        if let Some(v) = self.custom_max {
            s = s.set_custom_max(v);
        }
        if self.group_min {
            s = s.set_group_min(true);
        }
        if self.group_max {
            s = s.set_group_max(true);
        }
        if let Some(dates) = &self.date_range {
            s = s.set_date_range(dates.qualified(sheet).as_str());
        }
        Ok(s)
    }
}

/// A sparkline: a small chart in a cell.
///
/// ```python
/// ws.add_sparkline("F1", Sparkline("A1:E1", markers=True))
/// ws.add_sparkline_group("F1:F3", Sparkline("A1:E3", sparkline_type=SparklineType.Column))
/// ```
///
/// `range` (and `date_range`) may omit the sheet name, meaning the sheet the
/// sparkline is added to. Setting a point colour (`high_point_color`, ...)
/// also turns that point on. `custom_min`/`custom_max` and
/// `group_min`/`group_max` are alternatives: setting one clears the other.
#[pyclass(module = "pyrust_xlsxwriter", skip_from_py_object)]
#[derive(Clone)]
pub struct Sparkline {
    pub spec: SparklineSpec,
}

/// Colour properties that also turn their point on (as rust_xlsxwriter does).
macro_rules! point_colors {
    ($($get:ident, $set:ident => $color:ident, $flag:ident);* $(;)?) => {
        impl Sparkline {
            $(
                fn $get(&self) -> Option<Color> {
                    self.spec.$color
                }

                fn $set(&mut self, value: Option<Color>) -> PyResult<()> {
                    if value.is_some() {
                        self.spec.$flag = true;
                    }
                    self.spec.$color = value;
                    Ok(())
                }
            )*
        }
    };
}

point_colors!(
    get_high_point_color, set_high_point_color => high_point_color, high_point;
    get_low_point_color, set_low_point_color => low_point_color, low_point;
    get_first_point_color, set_first_point_color => first_point_color, first_point;
    get_last_point_color, set_last_point_color => last_point_color, last_point;
    get_negative_points_color, set_negative_points_color => negative_points_color, negative_points;
    get_markers_color, set_markers_color => markers_color, markers;
);

impl Sparkline {
    fn on_change(&mut self) -> PyResult<()> {
        self.spec.validate()
    }

    fn get_custom_min(&self) -> Option<f64> {
        self.spec.custom_min
    }

    fn set_custom_min(&mut self, value: Option<f64>) -> PyResult<()> {
        if value.is_some() {
            self.spec.group_min = false;
        }
        self.spec.custom_min = value;
        Ok(())
    }

    fn get_custom_max(&self) -> Option<f64> {
        self.spec.custom_max
    }

    fn set_custom_max(&mut self, value: Option<f64>) -> PyResult<()> {
        if value.is_some() {
            self.spec.group_max = false;
        }
        self.spec.custom_max = value;
        Ok(())
    }

    fn get_group_min(&self) -> bool {
        self.spec.group_min
    }

    fn set_group_min(&mut self, value: bool) -> PyResult<()> {
        if value {
            self.spec.custom_min = None;
        }
        self.spec.group_min = value;
        Ok(())
    }

    fn get_group_max(&self) -> bool {
        self.spec.group_max
    }

    fn set_group_max(&mut self, value: bool) -> PyResult<()> {
        if value {
            self.spec.custom_max = None;
        }
        self.spec.group_max = value;
        Ok(())
    }
}

py_props!(Sparkline
    fields {
        range: Option<RangeText>,
        sparkline_type: SparklineType,
        style: u8,
        high_point: bool,
        low_point: bool,
        first_point: bool,
        last_point: bool,
        negative_points: bool,
        markers: bool,
        axis: bool,
        hidden_data: bool,
        empty_cells: SparklineEmptyCells,
        right_to_left: bool,
        sparkline_color: Option<Color>,
        line_weight: Option<f64>,
        date_range: Option<RangeText>,
        column_order: bool,
    }
    custom {
        high_point_color: Option<Color>,
        low_point_color: Option<Color>,
        first_point_color: Option<Color>,
        last_point_color: Option<Color>,
        negative_points_color: Option<Color>,
        markers_color: Option<Color>,
        custom_min: Option<f64>,
        custom_max: Option<f64>,
        group_min: bool,
        group_max: bool,
    }
);

#[pymethods]
impl Sparkline {
    #[new]
    #[pyo3(signature = (range=None, sparkline_type=None, **kwargs))]
    fn py_new(
        range: Option<&Bound<'_, PyAny>>,
        sparkline_type: Option<SparklineType>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let mut spec = SparklineSpec::default();
        if let Some(range) = range {
            // `""` (the old default) means "not set yet".
            let empty = range.cast::<PyString>().map(|s| s.to_str().map(str::is_empty).unwrap_or(false));
            if !matches!(empty, Ok(true)) && !range.is_none() {
                spec.range = Some(range.extract::<RangeText>()?);
            }
        }
        if let Some(kind) = sparkline_type {
            spec.sparkline_type = kind;
        }
        let mut sparkline = Sparkline { spec };
        sparkline.apply_kwargs(kwargs)?;
        Ok(sparkline)
    }
}

spec_methods!(Sparkline);

#[pymethods]
impl Worksheet {
    /// Add a sparkline to a cell: `add_sparkline("F1", sparkline)` or
    /// `add_sparkline(row, col, sparkline)`. The data range must be 1D.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_sparkline(&mut self, args: &Bound<'_, PyTuple>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let ((row, col), rest) = split_cell(args.as_slice())?;
        let p = Params::new("add_sparkline", rest, kwargs, &["sparkline"], 1)?;
        let sparkline: PyRef<'_, Sparkline> = p.required(0)?;
        let rust = sparkline.spec.to_rust(&self.inner.name())?;
        self.inner.add_sparkline(row, col, &rust).py()?;
        Ok(())
    }

    /// Add a group of sparklines, one per row or column of a 2D data range:
    /// `add_sparkline_group("F1:F3", Sparkline("A1:E3"))`.
    #[pyo3(signature = (*args, **kwargs))]
    fn add_sparkline_group(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let ((r1, c1, r2, c2), rest) = split_range(args.as_slice())?;
        let p = Params::new("add_sparkline_group", rest, kwargs, &["sparkline"], 1)?;
        let sparkline: PyRef<'_, Sparkline> = p.required(0)?;
        let rust = sparkline.spec.to_rust(&self.inner.name())?;
        self.inner.add_sparkline_group(r1, c1, r2, c2, &rust).py()?;
        Ok(())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Sparkline>()?;
    m.add_class::<SparklineType>()?;
    m.add_class::<SparklineEmptyCells>()?;
    Ok(())
}
