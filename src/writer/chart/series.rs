//! Series and their decorations: `ChartSeries`, `ChartMarker`,
//! `ChartDataLabel`, `ChartPoint`, `ChartTrendline`, `ChartErrorBars`.

use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict};
use rust_xlsxwriter::{
    ChartDataLabel as RChartDataLabel, ChartErrorBars as RChartErrorBars,
    ChartMarker as RChartMarker, ChartPoint as RChartPoint, ChartSeries as RChartSeries,
    ChartTrendline as RChartTrendline,
};

use super::enums::{
    ChartDataLabelPosition, ChartErrorBarsDirection, ChartErrorBarsType, ChartMarkerType,
    ChartTrendlineType,
};
use super::format::{ChartFont, ChartFormat, ChartSolidFill};
use super::node::{
    Build, Node, Props, RangeSpec, Slot, check_data_range, check_finite, check_gap,
    check_marker_size, check_non_negative, check_overlap, check_separator, coerce_default,
    coerce_list, copy_list, new_with, repr_of,
};
use crate::common::color::Color;

// ---------------------------------------------------------------------------
// ChartMarker
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartMarkerSpec {
    pub automatic: bool,
    pub none: bool,
    pub type_: Option<ChartMarkerType>,
    pub size: Option<u8>,
}

/// Series marker for line, scatter and radar charts:
/// `series.marker = ChartMarker(type=ChartMarkerType.Circle, size=7)`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartMarker {
    pub spec: ChartMarkerSpec,
    pub format: Slot<ChartFormat>,
}

chart_props!(ChartMarker
    fields {
        /// Use Excel's automatic marker.
        automatic: bool,
        /// Turn markers off.
        none: bool,
        /// Marker symbol.
        type_ as r#type: Option<ChartMarkerType>,
        /// Marker size (2 to 72).
        size: Option<u8> => check_marker_size,
    }
    children {
        /// Border and fill of the marker.
        format: ChartFormat,
    }
    custom {}
);

impl Node for ChartMarker {}

impl Build for ChartMarker {
    type Rust = RChartMarker;

    fn build(&self, py: Python<'_>) -> PyResult<RChartMarker> {
        let s = &self.spec;
        let mut marker = RChartMarker::new();
        if s.automatic {
            marker.set_automatic();
        }
        if s.none {
            marker.set_none();
        }
        if let Some(kind) = s.type_ {
            marker.set_type(kind.into());
        }
        if let Some(size) = s.size {
            marker.set_size(size);
        }
        if let Some(mut format) = self.format.build(py)? {
            marker.set_format(&mut format);
        }
        Ok(marker)
    }
}

#[pymethods]
impl ChartMarker {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartMarker", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartDataLabel
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartDataLabelSpec {
    pub show_value: bool,
    pub show_category_name: bool,
    pub show_series_name: bool,
    pub show_leader_lines: bool,
    pub show_legend_key: bool,
    pub show_percentage: bool,
    pub position: Option<ChartDataLabelPosition>,
    pub num_format: Option<String>,
    pub separator: Option<char>,
    pub value: Option<RangeSpec>,
    pub hidden: bool,
}

/// Data labels of a series (`series.data_label.show_value = True`) or one
/// custom label in `series.custom_data_labels`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartDataLabel {
    pub spec: ChartDataLabelSpec,
    pub format: Slot<ChartFormat>,
    pub font: Slot<ChartFont>,
}

chart_props!(ChartDataLabel
    fields {
        /// Show the value (Y value for scatter charts).
        show_value: bool,
        /// Show the category name (X value for scatter charts).
        show_category_name: bool,
        /// Show the series name.
        show_series_name: bool,
        /// Show leader lines (pie and doughnut charts).
        show_leader_lines: bool,
        /// Show the legend key.
        show_legend_key: bool,
        /// Show the percentage (pie and doughnut charts).
        show_percentage: bool,
        /// Label position.
        position: Option<ChartDataLabelPosition>,
        /// Number format, e.g. "0.00%".
        num_format: Option<String>,
        /// Separator between multiple items: ',', ';', '.', ' ' or '\n'.
        separator: Option<char> => check_separator,
        /// Custom label text, a formula such as "=Sheet1!$A$1", or a (sheet, row, col) tuple.
        value: Option<RangeSpec>,
        /// Hide this label (for custom labels).
        hidden: bool,
    }
    children {
        /// Border and fill of the label.
        format: ChartFormat,
        /// Label font.
        font: ChartFont,
    }
    custom {
        /// Alias of `show_value` (scatter charts).
        show_y_value,
        /// Alias of `show_category_name` (scatter charts).
        show_x_value,
    }
);

impl ChartDataLabel {
    fn get_show_y_value(&self, _py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.show_value)
    }

    fn set_show_y_value(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.show_value = value;
        Ok(())
    }

    fn get_show_x_value(&self, _py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.show_category_name)
    }

    fn set_show_x_value(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.show_category_name = value;
        Ok(())
    }
}

impl Node for ChartDataLabel {}

impl Build for ChartDataLabel {
    type Rust = RChartDataLabel;

    fn build(&self, py: Python<'_>) -> PyResult<RChartDataLabel> {
        let s = &self.spec;
        let mut label = RChartDataLabel::new();
        // set_value() must come before set_format(), as in rust_xlsxwriter's examples.
        if let Some(value) = &s.value {
            label.set_value(&value.to_rust());
        }
        if s.show_value {
            label.show_value();
        }
        if s.show_category_name {
            label.show_category_name();
        }
        if s.show_series_name {
            label.show_series_name();
        }
        if s.show_leader_lines {
            label.show_leader_lines();
        }
        if s.show_legend_key {
            label.show_legend_key();
        }
        if s.show_percentage {
            label.show_percentage();
        }
        if let Some(position) = s.position {
            label.set_position(position.into());
        }
        if let Some(mut format) = self.format.build(py)? {
            label.set_format(&mut format);
        }
        if let Some(font) = self.font.build(py)? {
            label.set_font(&font);
        }
        if let Some(num_format) = &s.num_format {
            label.set_num_format(num_format);
        }
        if let Some(separator) = s.separator {
            label.set_separator(separator);
        }
        if s.hidden {
            label.set_hidden();
        }
        Ok(label)
    }
}

#[pymethods]
impl ChartDataLabel {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    /// A copy of this label, for use in `custom_data_labels`.
    fn to_custom(&self, py: Python<'_>) -> PyResult<Self> {
        self.props_copy(py)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartDataLabel", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartPoint
// ---------------------------------------------------------------------------

/// Formatting of one data point: `ChartPoint(format=ChartSolidFill(color="red"))`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartPoint {
    pub format: Slot<ChartFormat>,
}

chart_props!(ChartPoint
    fields {}
    children {
        /// Border and fill of the point.
        format: ChartFormat,
    }
    custom {}
);

impl Node for ChartPoint {}

impl Build for ChartPoint {
    type Rust = RChartPoint;

    fn build(&self, py: Python<'_>) -> PyResult<RChartPoint> {
        Ok(match self.format.build(py)? {
            Some(mut format) => RChartPoint::new().set_format(&mut format),
            None => RChartPoint::new(),
        })
    }
}

impl ChartPoint {
    /// A point with a solid fill colour.
    fn with_color(py: Python<'_>, color: Color) -> PyResult<Py<ChartPoint>> {
        let mut fill = ChartSolidFill::default();
        fill.spec.color = Some(color);
        let mut format = ChartFormat::default();
        format.solid_fill.put(py, Some(Py::new(py, fill)?));
        let mut point = ChartPoint::default();
        point.format.put(py, Some(Py::new(py, format)?));
        Py::new(py, point)
    }

    /// The solid fill colour of the point, if any.
    fn solid_color(&self, py: Python<'_>) -> PyResult<Option<Color>> {
        let color = self.format.with(py, |format| {
            Ok(format
                .solid_fill
                .with(py, |fill| Ok(fill.spec.color))?
                .flatten())
        })?;
        Ok(color.flatten())
    }
}

#[pymethods]
impl ChartPoint {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartPoint", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartTrendline
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartTrendlineSpec {
    pub type_: ChartTrendlineType,
    pub name: Option<String>,
    pub forward_period: Option<f64>,
    pub backward_period: Option<f64>,
    pub display_equation: bool,
    pub display_r_squared: bool,
    pub intercept: Option<f64>,
    pub delete_from_legend: bool,
}

/// A trendline: `series.trendline.type = ChartTrendlineType.Polynomial(3)`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartTrendline {
    pub spec: ChartTrendlineSpec,
    pub format: Slot<ChartFormat>,
    pub label_format: Slot<ChartFormat>,
    pub label_font: Slot<ChartFont>,
}

chart_props!(ChartTrendline
    fields {
        /// Trendline type (default None_, no trendline).
        type_ as r#type: ChartTrendlineType,
        /// Name shown in the legend.
        name: Option<String>,
        /// Extend the trendline forward by this many periods.
        forward_period: Option<f64> => check_non_negative,
        /// Extend the trendline backward by this many periods.
        backward_period: Option<f64> => check_non_negative,
        /// Show the equation on the chart.
        display_equation: bool,
        /// Show the R-squared value on the chart.
        display_r_squared: bool,
        /// Y-axis intercept (linear, polynomial and exponential trendlines).
        intercept: Option<f64> => check_finite,
        /// Remove the trendline from the legend.
        delete_from_legend: bool,
    }
    children {
        /// Line format of the trendline.
        format: ChartFormat,
        /// Border and fill of the equation/R-squared label.
        label_format: ChartFormat,
        /// Font of the equation/R-squared label.
        label_font: ChartFont,
    }
    custom {}
);

impl Node for ChartTrendline {}

impl Build for ChartTrendline {
    type Rust = RChartTrendline;

    fn build(&self, py: Python<'_>) -> PyResult<RChartTrendline> {
        let s = &self.spec;
        let mut trendline = RChartTrendline::new();
        trendline.set_type(s.type_.0);
        if let Some(name) = &s.name {
            trendline.set_name(name);
        }
        if let Some(period) = s.forward_period {
            trendline.set_forward_period(period);
        }
        if let Some(period) = s.backward_period {
            trendline.set_backward_period(period);
        }
        if s.display_equation {
            trendline.display_equation(true);
        }
        if s.display_r_squared {
            trendline.display_r_squared(true);
        }
        if let Some(intercept) = s.intercept {
            trendline.set_intercept(intercept);
        }
        if s.delete_from_legend {
            trendline.delete_from_legend(true);
        }
        if let Some(mut format) = self.format.build(py)? {
            trendline.set_format(&mut format);
        }
        if let Some(mut format) = self.label_format.build(py)? {
            trendline.set_label_format(&mut format);
        }
        if let Some(font) = self.label_font.build(py)? {
            trendline.set_label_font(&font);
        }
        Ok(trendline)
    }
}

#[pymethods]
impl ChartTrendline {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartTrendline", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartErrorBars
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ChartErrorBarsSpec {
    pub type_: ChartErrorBarsType,
    pub direction: Option<ChartErrorBarsDirection>,
    pub end_cap: bool,
}

impl Default for ChartErrorBarsSpec {
    fn default() -> Self {
        ChartErrorBarsSpec {
            type_: ChartErrorBarsType::default(),
            direction: None,
            end_cap: true,
        }
    }
}

/// Error bars. Assign them to show them:
/// `series.y_error_bars = ChartErrorBars(type=ChartErrorBarsType.Percentage(5))`
/// (or `= True` for standard-error bars).
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartErrorBars {
    pub spec: ChartErrorBarsSpec,
    pub format: Slot<ChartFormat>,
}

chart_props!(ChartErrorBars
    fields {
        /// How the error amount is calculated (default StandardError).
        type_ as r#type: ChartErrorBarsType,
        /// Plus, minus or both (default Both).
        direction: Option<ChartErrorBarsDirection>,
        /// Draw end caps (default True).
        end_cap: bool,
    }
    children {
        /// Line format of the error bars.
        format: ChartFormat,
    }
    custom {}
);

impl Node for ChartErrorBars {
    /// Also accepts True (standard-error bars) and False (no error bars).
    fn coerce(value: &Bound<'_, PyAny>) -> PyResult<Option<Py<Self>>> {
        if let Ok(flag) = value.cast::<PyBool>() {
            return Ok(if flag.is_true() {
                Some(Py::new(value.py(), ChartErrorBars::default())?)
            } else {
                None
            });
        }
        coerce_default::<Self>(value)
    }
}

impl Build for ChartErrorBars {
    type Rust = RChartErrorBars;

    fn build(&self, py: Python<'_>) -> PyResult<RChartErrorBars> {
        let s = &self.spec;
        let mut bars = RChartErrorBars::new();
        bars.set_type(s.type_.to_rust());
        if let Some(direction) = s.direction {
            bars.set_direction(direction.into());
        }
        if !s.end_cap {
            bars.set_end_cap(false);
        }
        if let Some(mut format) = self.format.build(py)? {
            bars.set_format(&mut format);
        }
        Ok(bars)
    }
}

#[pymethods]
impl ChartErrorBars {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartErrorBars", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartSeries
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartSeriesSpec {
    pub values: Option<RangeSpec>,
    pub categories: Option<RangeSpec>,
    pub name: Option<RangeSpec>,
    pub secondary_axis: bool,
    pub overlap: Option<i8>,
    pub gap: Option<u16>,
    pub smooth: Option<bool>,
    pub invert_if_negative: bool,
    pub invert_if_negative_color: Option<Color>,
    pub delete_from_legend: bool,
}

/// A data series. Create one with `chart.add_series(values=..., categories=..., name=...)`
/// (which returns the live series) or `ChartSeries(...)` + `chart.push_series(series)`.
///
/// Ranges are "=Sheet1!$B$1:$B$5" strings or
/// `(sheet, first_row, first_col, last_row, last_col)` tuples (zero-based).
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartSeries {
    pub spec: ChartSeriesSpec,
    pub format: Slot<ChartFormat>,
    pub marker: Slot<ChartMarker>,
    pub data_label: Slot<ChartDataLabel>,
    pub trendline: Slot<ChartTrendline>,
    pub y_error_bars: Slot<ChartErrorBars>,
    pub x_error_bars: Slot<ChartErrorBars>,
    pub points: Vec<Option<Py<ChartPoint>>>,
    pub custom_data_labels: Vec<Option<Py<ChartDataLabel>>>,
}

chart_props!(ChartSeries
    fields {
        /// The values range (required).
        values: Option<RangeSpec> => check_data_range,
        /// The categories range (X values for scatter charts).
        categories: Option<RangeSpec> => check_data_range,
        /// Series name: text, a formula such as "=Sheet1!$B$1", or a (sheet, row, col) tuple.
        name: Option<RangeSpec>,
        /// Plot the series on the secondary (right/top) axes.
        secondary_axis: bool,
        /// Overlap of bars/columns (-100 to 100).
        overlap: Option<i8> => check_overlap,
        /// Gap between bar/column groups (0 to 500).
        gap: Option<u16> => check_gap,
        /// Smooth the line (line and scatter charts).
        smooth: Option<bool>,
        /// Invert the fill of negative values (bar/column charts).
        invert_if_negative: bool,
        /// Fill colour for inverted negative values (sets invert_if_negative).
        invert_if_negative_color: Option<Color>,
        /// Remove the series from the legend.
        delete_from_legend: bool,
    }
    children {
        /// Line/border and fill of the series.
        format: ChartFormat,
        /// Marker (line, scatter and radar charts).
        marker: ChartMarker,
        /// Data labels for every point.
        data_label: ChartDataLabel,
        /// Trendline.
        trendline: ChartTrendline,
        /// Y error bars (assign to show).
        y_error_bars: ChartErrorBars,
        /// X error bars (assign to show; scatter and bar charts).
        x_error_bars: ChartErrorBars,
    }
    custom {
        /// Per-point formatting: a list of ChartPoint (or dict / None for a default point).
        points,
        /// Per-point fill colours (shortcut for `points` with solid fills).
        point_colors,
        /// Per-point custom data labels: a list of ChartDataLabel (or dict / None for default).
        custom_data_labels,
    }
);

impl ChartSeries {
    fn get_points(&self, py: Python<'_>) -> PyResult<Vec<Option<Py<ChartPoint>>>> {
        Ok(self
            .points
            .iter()
            .map(|p| p.as_ref().map(|p| p.clone_ref(py)))
            .collect())
    }

    fn set_points(
        &mut self,
        _py: Python<'_>,
        value: Option<Vec<Bound<'_, PyAny>>>,
    ) -> PyResult<()> {
        self.points = coerce_list(value.unwrap_or_default())?;
        Ok(())
    }

    fn get_point_colors(&self, py: Python<'_>) -> PyResult<Vec<Option<Color>>> {
        self.points
            .iter()
            .map(|p| match p {
                None => Ok(None),
                Some(p) => p.try_borrow(py)?.solid_color(py),
            })
            .collect()
    }

    fn set_point_colors(
        &mut self,
        py: Python<'_>,
        value: Option<Vec<Option<Color>>>,
    ) -> PyResult<()> {
        self.points = value
            .unwrap_or_default()
            .into_iter()
            .map(|c| c.map(|c| ChartPoint::with_color(py, c)).transpose())
            .collect::<PyResult<_>>()?;
        Ok(())
    }

    fn get_custom_data_labels(&self, py: Python<'_>) -> PyResult<Vec<Option<Py<ChartDataLabel>>>> {
        Ok(self
            .custom_data_labels
            .iter()
            .map(|l| l.as_ref().map(|l| l.clone_ref(py)))
            .collect())
    }

    fn set_custom_data_labels(
        &mut self,
        _py: Python<'_>,
        value: Option<Vec<Bound<'_, PyAny>>>,
    ) -> PyResult<()> {
        self.custom_data_labels = coerce_list(value.unwrap_or_default())?;
        Ok(())
    }

    /// Apply this series to a series created by `Chart::add_series()` (so the
    /// chart type's series defaults are kept unless overridden).
    pub fn apply_to(&self, py: Python<'_>, series: &mut RChartSeries) -> PyResult<()> {
        let s = &self.spec;
        if let Some(values) = &s.values {
            series.set_values(&values.to_rust());
        }
        if let Some(categories) = &s.categories {
            series.set_categories(&categories.to_rust());
        }
        if let Some(name) = &s.name {
            series.set_name(&name.to_rust());
        }
        if s.secondary_axis {
            series.set_secondary_axis(true);
        }
        if let Some(mut format) = self.format.build(py)? {
            series.set_format(&mut format);
        }
        if let Some(marker) = self.marker.build(py)? {
            series.set_marker(&marker);
        }
        if let Some(label) = self.data_label.build(py)? {
            series.set_data_label(&label);
        }
        if !self.custom_data_labels.is_empty() {
            let labels = self
                .custom_data_labels
                .iter()
                .map(|l| match l {
                    None => Ok(RChartDataLabel::default()),
                    Some(l) => l.try_borrow(py)?.build(py),
                })
                .collect::<PyResult<Vec<_>>>()?;
            series.set_custom_data_labels(&labels);
        }
        if !self.points.is_empty() {
            let points = self
                .points
                .iter()
                .map(|p| match p {
                    None => Ok(RChartPoint::new()),
                    Some(p) => p.try_borrow(py)?.build(py),
                })
                .collect::<PyResult<Vec<_>>>()?;
            series.set_points(&points);
        }
        if let Some(trendline) = self.trendline.build(py)? {
            series.set_trendline(&trendline);
        }
        if let Some(bars) = self.y_error_bars.build(py)? {
            series.set_y_error_bars(&bars);
        }
        if let Some(bars) = self.x_error_bars.build(py)? {
            series.set_x_error_bars(&bars);
        }
        if let Some(overlap) = s.overlap {
            series.set_overlap(overlap);
        }
        if let Some(gap) = s.gap {
            series.set_gap(gap);
        }
        if let Some(smooth) = s.smooth {
            series.set_smooth(smooth);
        }
        if s.invert_if_negative {
            series.set_invert_if_negative();
        }
        if let Some(color) = s.invert_if_negative_color {
            series.set_invert_if_negative_color(color.0);
        }
        if s.delete_from_legend {
            series.delete_from_legend(true);
        }
        Ok(())
    }

    /// A new series from the constructor / `add_series()` arguments. Empty
    /// strings are ignored (the old `add_series(values="", ...)` defaults).
    pub fn create(
        py: Python<'_>,
        values: Option<RangeSpec>,
        categories: Option<RangeSpec>,
        name: Option<RangeSpec>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let mut series: ChartSeries = new_with(py, kwargs)?;
        let given = |r: Option<RangeSpec>| r.filter(|r| *r != RangeSpec::Text(String::new()));
        if let Some(values) = given(values) {
            check_data_range("values", &Some(values.clone()))?;
            series.spec.values = Some(values);
        }
        if let Some(categories) = given(categories) {
            check_data_range("categories", &Some(categories.clone()))?;
            series.spec.categories = Some(categories);
        }
        if let Some(name) = given(name) {
            series.spec.name = Some(name);
        }
        Ok(series)
    }

    /// A deep copy, including points and custom labels.
    pub fn deep_copy(&self, py: Python<'_>) -> PyResult<Self> {
        let mut copy = self.props_copy(py)?;
        copy.points = copy_list(py, &self.points)?;
        copy.custom_data_labels = copy_list(py, &self.custom_data_labels)?;
        Ok(copy)
    }
}

#[pymethods]
impl ChartSeries {
    #[new]
    #[pyo3(signature = (values=None, categories=None, name=None, **kwargs))]
    fn py_new(
        py: Python<'_>,
        values: Option<RangeSpec>,
        categories: Option<RangeSpec>,
        name: Option<RangeSpec>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        ChartSeries::create(py, values, categories, name, kwargs)
    }

    /// A deep copy of the series (nested objects are copied too).
    fn copy(&self, py: Python<'_>) -> PyResult<Self> {
        self.deep_copy(py)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let mut parts = self.props_repr(py)?;
        if !self.points.is_empty() {
            parts.push(format!("points=<{} points>", self.points.len()));
        }
        if !self.custom_data_labels.is_empty() {
            parts.push(format!(
                "custom_data_labels=<{} labels>",
                self.custom_data_labels.len()
            ));
        }
        Ok(repr_of("ChartSeries", parts))
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ChartSeries>()?;
    m.add_class::<ChartMarker>()?;
    m.add_class::<ChartDataLabel>()?;
    m.add_class::<ChartPoint>()?;
    m.add_class::<ChartTrendline>()?;
    m.add_class::<ChartErrorBars>()?;
    Ok(())
}
