//! Chart parts: `ChartTitle`, `ChartLegend`, `ChartArea`, `ChartPlotArea`,
//! `ChartDataTable` and `ChartAxis`.

use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict};
use rust_xlsxwriter::{
    ChartArea as RChartArea, ChartAxis as RChartAxis, ChartDataTable as RChartDataTable,
    ChartLegend as RChartLegend, ChartPlotArea as RChartPlotArea, ChartTitle as RChartTitle,
};

use super::enums::{
    ChartAxisCrossing, ChartAxisDateUnitType, ChartAxisDisplayUnitType, ChartAxisLabelAlignment,
    ChartAxisLabelPosition, ChartAxisTickType, ChartLegendPosition,
};
use super::format::{ChartFont, ChartFormat, ChartLayout, ChartLine};
use super::node::{
    Build, ChartDate, Node, Props, RangeSpec, Slot, check_finite, check_log_base,
    check_non_negative, coerce_default, new_with, repr_of,
};
use crate::common::errors::param_err;

/// A chart part that is applied to an object owned by the `rust_xlsxwriter` chart.
pub trait ApplyTo: Node {
    type Target;
    fn apply_to(&self, py: Python<'_>, target: &mut Self::Target) -> PyResult<()>;
}

/// Apply a slot's object (if it was ever created) to `target`.
pub fn apply_slot<T: ApplyTo>(
    py: Python<'_>,
    slot: &Slot<T>,
    target: &mut T::Target,
) -> PyResult<()> {
    slot.with(py, |part| part.apply_to(py, target))?;
    Ok(())
}

fn check_interval(name: &str, value: &Option<u16>) -> PyResult<()> {
    if let Some(0) = value {
        return Err(param_err(format!("{name} must be 1 or more, got 0")));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// ChartTitle
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartTitleSpec {
    pub name: Option<RangeSpec>,
    pub hidden: bool,
    pub overlay: bool,
}

/// A chart title: `chart.title.name = "Sales"`, or a cell reference
/// `("Sheet1", 0, 0)` / `"=Sheet1!$A$1"`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartTitle {
    pub spec: ChartTitleSpec,
    pub format: Slot<ChartFormat>,
    pub font: Slot<ChartFont>,
    pub layout: Slot<ChartLayout>,
}

chart_props!(ChartTitle
    fields {
        /// Title text, a formula such as "=Sheet1!$A$1", or a (sheet, row, col) tuple.
        name: Option<RangeSpec>,
        /// Hide the title (e.g. the automatic title of a single-series chart).
        hidden: bool,
        /// Let the title overlay the plot area.
        overlay: bool,
    }
    children {
        /// Border and fill of the title.
        format: ChartFormat,
        /// Title font.
        font: ChartFont,
        /// Manual position of the title.
        layout: ChartLayout,
    }
    custom {}
);

impl Node for ChartTitle {}

impl ChartTitle {
    /// The title text (or formula) as a string; "" when unset.
    pub fn text(&self) -> String {
        self.spec
            .name
            .as_ref()
            .map(RangeSpec::display)
            .unwrap_or_default()
    }
}

impl ApplyTo for ChartTitle {
    type Target = RChartTitle;

    fn apply_to(&self, py: Python<'_>, title: &mut RChartTitle) -> PyResult<()> {
        let s = &self.spec;
        if let Some(name) = &s.name {
            title.set_name(&name.to_rust());
        }
        if s.hidden {
            title.set_hidden();
        }
        if let Some(mut format) = self.format.build(py)? {
            title.set_format(&mut format);
        }
        if let Some(font) = self.font.build(py)? {
            title.set_font(&font);
        }
        if let Some(layout) = self.layout.build(py)? {
            title.set_layout(&layout);
        }
        if s.overlay {
            title.set_overlay(true);
        }
        Ok(())
    }
}

#[pymethods]
impl ChartTitle {
    #[new]
    #[pyo3(signature = (name=None, **kwargs))]
    fn py_new(
        py: Python<'_>,
        name: Option<RangeSpec>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let mut title: ChartTitle = new_with(py, kwargs)?;
        if name.is_some() {
            title.spec.name = name;
        }
        Ok(title)
    }

    /// The title text ("" when no name is set).
    fn __str__(&self) -> String {
        self.text()
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartTitle", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartLegend
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ChartLegendSpec {
    pub position: ChartLegendPosition,
    pub hidden: bool,
    pub overlay: bool,
    pub delete_entries: Vec<usize>,
}

impl Default for ChartLegendSpec {
    fn default() -> Self {
        ChartLegendSpec {
            position: ChartLegendPosition::Right,
            hidden: false,
            overlay: false,
            delete_entries: Vec::new(),
        }
    }
}

/// The chart legend: `chart.legend.position = ChartLegendPosition.Bottom`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartLegend {
    pub spec: ChartLegendSpec,
    pub layout: Slot<ChartLayout>,
    pub format: Slot<ChartFormat>,
    pub font: Slot<ChartFont>,
}

chart_props!(ChartLegend
    fields {
        /// Legend position (default Right).
        position: ChartLegendPosition,
        /// Hide the legend.
        hidden: bool,
        /// Let the legend overlay the plot area.
        overlay: bool,
        /// Zero-based indexes of legend entries to remove.
        delete_entries: Vec<usize>,
    }
    children {
        /// Manual position/size of the legend.
        layout: ChartLayout,
        /// Border and fill of the legend.
        format: ChartFormat,
        /// Legend font.
        font: ChartFont,
    }
    custom {}
);

impl Node for ChartLegend {}

impl ApplyTo for ChartLegend {
    type Target = RChartLegend;

    fn apply_to(&self, py: Python<'_>, legend: &mut RChartLegend) -> PyResult<()> {
        let s = &self.spec;
        legend.set_position(s.position.into());
        if s.hidden {
            legend.set_hidden();
        }
        if s.overlay {
            legend.set_overlay(true);
        }
        if let Some(layout) = self.layout.build(py)? {
            legend.set_layout(&layout);
        }
        if let Some(mut format) = self.format.build(py)? {
            legend.set_format(&mut format);
        }
        if let Some(font) = self.font.build(py)? {
            legend.set_font(&font);
        }
        if !s.delete_entries.is_empty() {
            legend.delete_entries(&s.delete_entries);
        }
        Ok(())
    }
}

#[pymethods]
impl ChartLegend {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartLegend", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartArea, ChartPlotArea
// ---------------------------------------------------------------------------

/// The whole chart area (background and border): `chart.chart_area.format = ChartSolidFill(color="#F2F2F2")`.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartArea {
    pub format: Slot<ChartFormat>,
}

chart_props!(ChartArea
    fields {}
    children {
        /// Border and fill of the chart area.
        format: ChartFormat,
    }
    custom {}
);

impl Node for ChartArea {}

impl ApplyTo for ChartArea {
    type Target = RChartArea;

    fn apply_to(&self, py: Python<'_>, area: &mut RChartArea) -> PyResult<()> {
        if let Some(mut format) = self.format.build(py)? {
            area.set_format(&mut format);
        }
        Ok(())
    }
}

#[pymethods]
impl ChartArea {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartArea", self.props_repr(py)?))
    }
}

/// The plot area (inside the axes).
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartPlotArea {
    pub format: Slot<ChartFormat>,
    pub layout: Slot<ChartLayout>,
}

chart_props!(ChartPlotArea
    fields {}
    children {
        /// Border and fill of the plot area.
        format: ChartFormat,
        /// Manual position/size of the plot area.
        layout: ChartLayout,
    }
    custom {}
);

impl Node for ChartPlotArea {}

impl ApplyTo for ChartPlotArea {
    type Target = RChartPlotArea;

    fn apply_to(&self, py: Python<'_>, area: &mut RChartPlotArea) -> PyResult<()> {
        if let Some(mut format) = self.format.build(py)? {
            area.set_format(&mut format);
        }
        if let Some(layout) = self.layout.build(py)? {
            area.set_layout(&layout);
        }
        Ok(())
    }
}

#[pymethods]
impl ChartPlotArea {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartPlotArea", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartDataTable
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ChartDataTableSpec {
    pub show_horizontal_borders: bool,
    pub show_vertical_borders: bool,
    pub show_outline_borders: bool,
    pub show_legend_keys: bool,
}

impl Default for ChartDataTableSpec {
    fn default() -> Self {
        ChartDataTableSpec {
            show_horizontal_borders: true,
            show_vertical_borders: true,
            show_outline_borders: true,
            show_legend_keys: false,
        }
    }
}

/// A data table below the chart. Assign one to show it:
/// `chart.data_table = ChartDataTable(show_legend_keys=True)` (or `= True`).
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartDataTable {
    pub spec: ChartDataTableSpec,
    pub format: Slot<ChartFormat>,
    pub font: Slot<ChartFont>,
}

chart_props!(ChartDataTable
    fields {
        /// Show horizontal cell borders (default True).
        show_horizontal_borders: bool,
        /// Show vertical cell borders (default True).
        show_vertical_borders: bool,
        /// Show the outline border (default True).
        show_outline_borders: bool,
        /// Show legend keys next to the series names (default False).
        show_legend_keys: bool,
    }
    children {
        /// Border and fill of the table.
        format: ChartFormat,
        /// Table font.
        font: ChartFont,
    }
    custom {}
);

impl Node for ChartDataTable {
    /// Also accepts True (a default table) and False (no table).
    fn coerce(value: &Bound<'_, PyAny>) -> PyResult<Option<Py<Self>>> {
        if let Ok(flag) = value.cast::<PyBool>() {
            return Ok(if flag.is_true() {
                Some(Py::new(value.py(), ChartDataTable::default())?)
            } else {
                None
            });
        }
        coerce_default::<Self>(value)
    }
}

impl Build for ChartDataTable {
    type Rust = RChartDataTable;

    fn build(&self, py: Python<'_>) -> PyResult<RChartDataTable> {
        let s = &self.spec;
        let mut table = RChartDataTable::new()
            .show_horizontal_borders(s.show_horizontal_borders)
            .show_vertical_borders(s.show_vertical_borders)
            .show_outline_borders(s.show_outline_borders)
            .show_legend_keys(s.show_legend_keys);
        if let Some(mut format) = self.format.build(py)? {
            table = table.set_format(&mut format);
        }
        if let Some(font) = self.font.build(py)? {
            table = table.set_font(&font);
        }
        Ok(table)
    }
}

#[pymethods]
impl ChartDataTable {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(repr_of("ChartDataTable", self.props_repr(py)?))
    }
}

// ---------------------------------------------------------------------------
// ChartAxis
// ---------------------------------------------------------------------------

/// Axis settings. `None` means "the chart type's default" (for example
/// column charts have major gridlines on the Y axis only).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ChartAxisSpec {
    pub name: Option<RangeSpec>,
    pub num_format: Option<String>,
    pub date_axis: Option<bool>,
    pub text_axis: Option<bool>,
    pub automatic_axis: Option<bool>,
    pub crossing: Option<ChartAxisCrossing>,
    pub reverse: bool,
    pub max: Option<f64>,
    pub min: Option<f64>,
    pub max_date: Option<ChartDate>,
    pub min_date: Option<ChartDate>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    pub display_unit_type: Option<ChartAxisDisplayUnitType>,
    pub display_units_visible: Option<bool>,
    pub major_unit_date_type: Option<ChartAxisDateUnitType>,
    pub minor_unit_date_type: Option<ChartAxisDateUnitType>,
    pub label_alignment: Option<ChartAxisLabelAlignment>,
    pub major_gridlines: Option<bool>,
    pub minor_gridlines: Option<bool>,
    pub label_position: Option<ChartAxisLabelPosition>,
    pub position_between_ticks: Option<bool>,
    pub label_interval: Option<u16>,
    pub tick_interval: Option<u16>,
    pub major_tick_type: Option<ChartAxisTickType>,
    pub minor_tick_type: Option<ChartAxisTickType>,
    pub log_base: Option<u16>,
    pub hidden: Option<bool>,
}

/// A chart axis: `chart.x_axis`, `chart.y_axis`, `chart.x2_axis`, `chart.y2_axis`.
///
/// ```python
/// chart.y_axis.name = "Sales"
/// chart.y_axis.min, chart.y_axis.max = 0, 100
/// chart.y_axis.major_gridlines_line.color = "#D9D9D9"
/// chart.x_axis.crossing = ChartAxisCrossing.AxisValue(50)
/// ```
///
/// Properties left as None keep the default of the chart type.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct ChartAxis {
    pub spec: ChartAxisSpec,
    pub name_font: Slot<ChartFont>,
    pub name_format: Slot<ChartFormat>,
    pub format: Slot<ChartFormat>,
    pub font: Slot<ChartFont>,
    pub major_gridlines_line: Slot<ChartLine>,
    pub minor_gridlines_line: Slot<ChartLine>,
    pub label_layout: Slot<ChartLayout>,
}

chart_props!(ChartAxis
    fields {
        /// Axis title: text, a formula such as "=Sheet1!$A$1", or a (sheet, row, col) tuple.
        name: Option<RangeSpec>,
        /// Number format of the tick labels, e.g. "$#,##0".
        num_format: Option<String>,
        /// Make a category axis a date axis.
        date_axis: Option<bool>,
        /// Make a category axis a text axis.
        text_axis: Option<bool>,
        /// Let Excel choose between a text and a date axis.
        automatic_axis: Option<bool>,
        /// Where the perpendicular axis crosses this one.
        crossing: Option<ChartAxisCrossing>,
        /// Plot the values in reverse order.
        reverse: bool,
        /// Major unit (interval between major ticks), 0 or more.
        major_unit: Option<f64> => check_non_negative,
        /// Minor unit (interval between minor ticks), 0 or more.
        minor_unit: Option<f64> => check_non_negative,
        /// Display units (hundreds, thousands, ...) of a value axis.
        display_unit_type: Option<ChartAxisDisplayUnitType>,
        /// Show the display units label.
        display_units_visible: Option<bool>,
        /// Date unit of `major_unit` on a date axis.
        major_unit_date_type: Option<ChartAxisDateUnitType>,
        /// Date unit of `minor_unit` on a date axis.
        minor_unit_date_type: Option<ChartAxisDateUnitType>,
        /// Alignment of the category labels.
        label_alignment: Option<ChartAxisLabelAlignment>,
        /// Show major gridlines.
        major_gridlines: Option<bool>,
        /// Show minor gridlines.
        minor_gridlines: Option<bool>,
        /// Position of the tick labels.
        label_position: Option<ChartAxisLabelPosition>,
        /// Position the axis between (True) or on (False) the tick marks.
        position_between_ticks: Option<bool>,
        /// Show every n-th category label (1 or more).
        label_interval: Option<u16> => check_interval,
        /// Put a tick mark every n categories (1 or more).
        tick_interval: Option<u16> => check_interval,
        /// Major tick mark type.
        major_tick_type: Option<ChartAxisTickType>,
        /// Minor tick mark type.
        minor_tick_type: Option<ChartAxisTickType>,
        /// Logarithmic scale with this base (2 or more).
        log_base: Option<u16> => check_log_base,
        /// Hide the axis.
        hidden: Option<bool>,
    }
    children {
        /// Font of the axis title.
        name_font: ChartFont,
        /// Border and fill of the axis title.
        name_format: ChartFormat,
        /// Line and fill of the axis.
        format: ChartFormat,
        /// Font of the tick labels.
        font: ChartFont,
        /// Line of the major gridlines (setting it shows them).
        major_gridlines_line: ChartLine,
        /// Line of the minor gridlines (setting it shows them).
        minor_gridlines_line: ChartLine,
        /// Manual position of the axis title.
        label_layout: ChartLayout,
    }
    custom {
        /// Maximum value (clears `max_date`).
        max,
        /// Minimum value (clears `min_date`).
        min,
        /// Maximum of a date axis as a date/datetime (clears `max`).
        max_date,
        /// Minimum of a date axis as a date/datetime (clears `min`).
        min_date,
    }
);

impl ChartAxis {
    fn get_max(&self, _py: Python<'_>) -> PyResult<Option<f64>> {
        Ok(self.spec.max)
    }

    fn set_max(&mut self, _py: Python<'_>, value: Option<f64>) -> PyResult<()> {
        check_finite("max", &value)?;
        self.spec.max = value;
        self.spec.max_date = None;
        Ok(())
    }

    fn get_min(&self, _py: Python<'_>) -> PyResult<Option<f64>> {
        Ok(self.spec.min)
    }

    fn set_min(&mut self, _py: Python<'_>, value: Option<f64>) -> PyResult<()> {
        check_finite("min", &value)?;
        self.spec.min = value;
        self.spec.min_date = None;
        Ok(())
    }

    fn get_max_date(&self, _py: Python<'_>) -> PyResult<Option<ChartDate>> {
        Ok(self.spec.max_date)
    }

    fn set_max_date(&mut self, _py: Python<'_>, value: Option<ChartDate>) -> PyResult<()> {
        self.spec.max_date = value;
        self.spec.max = None;
        Ok(())
    }

    fn get_min_date(&self, _py: Python<'_>) -> PyResult<Option<ChartDate>> {
        Ok(self.spec.min_date)
    }

    fn set_min_date(&mut self, _py: Python<'_>, value: Option<ChartDate>) -> PyResult<()> {
        self.spec.min_date = value;
        self.spec.min = None;
        Ok(())
    }

    /// The axis title as a string ("" when unset).
    pub fn name_text(&self) -> String {
        self.spec
            .name
            .as_ref()
            .map(RangeSpec::display)
            .unwrap_or_default()
    }
}

impl Node for ChartAxis {
    fn is_default(&self, py: Python<'_>) -> bool {
        self.props_is_default(py) && self.spec == ChartAxisSpec::default()
    }
}

impl ApplyTo for ChartAxis {
    type Target = RChartAxis;

    fn apply_to(&self, py: Python<'_>, axis: &mut RChartAxis) -> PyResult<()> {
        let s = &self.spec;
        if let Some(name) = &s.name {
            axis.set_name(&name.to_rust());
        }
        if let Some(font) = self.name_font.build(py)? {
            axis.set_name_font(&font);
        }
        if let Some(mut format) = self.name_format.build(py)? {
            axis.set_name_format(&mut format);
        }
        if let Some(mut format) = self.format.build(py)? {
            axis.set_format(&mut format);
        }
        if let Some(font) = self.font.build(py)? {
            axis.set_font(&font);
        }
        if let Some(num_format) = &s.num_format {
            axis.set_num_format(num_format);
        }
        if let Some(enable) = s.date_axis {
            axis.set_date_axis(enable);
        }
        if let Some(enable) = s.text_axis {
            axis.set_text_axis(enable);
        }
        if let Some(enable) = s.automatic_axis {
            axis.set_automatic_axis(enable);
        }
        if let Some(crossing) = s.crossing {
            axis.set_crossing(crossing.0);
        }
        if s.reverse {
            axis.set_reverse();
        }
        if let Some(date) = &s.max_date {
            axis.set_max_date(date.0);
        } else if let Some(max) = s.max {
            axis.set_max(max);
        }
        if let Some(date) = &s.min_date {
            axis.set_min_date(date.0);
        } else if let Some(min) = s.min {
            axis.set_min(min);
        }
        if let Some(unit) = s.major_unit {
            axis.set_major_unit(unit);
        }
        if let Some(unit) = s.minor_unit {
            axis.set_minor_unit(unit);
        }
        if let Some(unit_type) = s.display_unit_type {
            axis.set_display_unit_type(unit_type.into());
        }
        if let Some(visible) = s.display_units_visible {
            axis.set_display_units_visible(visible);
        }
        if let Some(unit_type) = s.major_unit_date_type {
            axis.set_major_unit_date_type(unit_type.into());
        }
        if let Some(unit_type) = s.minor_unit_date_type {
            axis.set_minor_unit_date_type(unit_type.into());
        }
        if let Some(alignment) = s.label_alignment {
            axis.set_label_alignment(alignment.into());
        }
        if let Some(line) = self.major_gridlines_line.build(py)? {
            axis.set_major_gridlines_line(&line);
        }
        if let Some(line) = self.minor_gridlines_line.build(py)? {
            axis.set_minor_gridlines_line(&line);
        }
        if let Some(enable) = s.major_gridlines {
            axis.set_major_gridlines(enable);
        }
        if let Some(enable) = s.minor_gridlines {
            axis.set_minor_gridlines(enable);
        }
        if let Some(position) = s.label_position {
            axis.set_label_position(position.into());
        }
        if let Some(enable) = s.position_between_ticks {
            axis.set_position_between_ticks(enable);
        }
        if let Some(interval) = s.label_interval {
            axis.set_label_interval(interval);
        }
        if let Some(interval) = s.tick_interval {
            axis.set_tick_interval(interval);
        }
        if let Some(tick) = s.major_tick_type {
            axis.set_major_tick_type(tick.into());
        }
        if let Some(tick) = s.minor_tick_type {
            axis.set_minor_tick_type(tick.into());
        }
        if let Some(base) = s.log_base {
            axis.set_log_base(base);
        }
        if let Some(hidden) = s.hidden {
            axis.set_hidden(hidden);
        }
        if let Some(layout) = self.label_layout.build(py)? {
            axis.set_label_layout(&layout);
        }
        Ok(())
    }
}

#[pymethods]
impl ChartAxis {
    #[new]
    #[pyo3(signature = (**kwargs))]
    fn py_new(py: Python<'_>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        new_with(py, kwargs)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let mut parts = self.props_repr(py)?;
        for (name, value) in [("max", self.spec.max), ("min", self.spec.min)] {
            if let Some(v) = value {
                parts.push(format!("{name}={v:?}"));
            }
        }
        for (name, value) in [
            ("max_date", self.spec.max_date),
            ("min_date", self.spec.min_date),
        ] {
            if let Some(v) = value {
                parts.push(format!("{name}={}", v.0));
            }
        }
        Ok(repr_of("ChartAxis", parts))
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ChartTitle>()?;
    m.add_class::<ChartLegend>()?;
    m.add_class::<ChartArea>()?;
    m.add_class::<ChartPlotArea>()?;
    m.add_class::<ChartDataTable>()?;
    m.add_class::<ChartAxis>()?;
    Ok(())
}
