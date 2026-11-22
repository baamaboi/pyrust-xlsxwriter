//! Charts and chartsheets.
//!
//! ```python
//! chart = Chart(ChartType.Column)
//! series = chart.add_series(values="=Sheet1!$B$1:$B$5", categories=("Sheet1", 0, 0, 4, 0), name="Sales")
//! series.format.solid_fill.color = "#4472C4"      # nested objects are live
//! chart.title = "Quarterly sales"                 # same as chart.title.name = ...
//! chart.y_axis.major_gridlines_line.dash_type = ChartLineDashType.Dash
//! ws.insert_chart("D2", chart)
//! ```
//!
//! Python objects only hold settings (see `node`); the `rust_xlsxwriter`
//! chart is built from them when it is inserted, so a chart can be changed
//! and inserted again.
//!
//! * `node`   - slots, the property macro, chart ranges and validators.
//! * `enums`  - chart enums (including the data-carrying ones).
//! * `format` - lines, fills, `ChartFormat`, `ChartFont`, `ChartLayout`.
//! * `parts`  - title, legend, chart/plot area, data table, axes.
//! * `series` - series, markers, data labels, points, trendlines, error bars.

#[macro_use]
pub(crate) mod node;
pub mod enums;
pub mod format;
pub mod parts;
pub mod series;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use rust_xlsxwriter::{Chart as RChart, Worksheet as RWorksheet};

use crate::common::cellref::{Params, split_cell};
use crate::common::errors::{ErrorKind, IntoPyResult, new_err, param_err};
use crate::writer::object_movement::ObjectMovement;
use crate::writer::workbook::Workbook;
use crate::writer::worksheet::Worksheet;

use enums::{ChartEmptyCells, ChartLegendPosition, ChartType};
use format::ChartFormat;
use node::{Props, RangeSpec, Slot, new_with, repr_of};
use parts::{
    ChartArea, ChartAxis, ChartDataTable, ChartLegend, ChartPlotArea, ChartTitle, apply_slot,
};
use series::ChartSeries;

// ---------------------------------------------------------------------------
// Validators for chart-level properties.
// ---------------------------------------------------------------------------

fn check_rotation(name: &str, value: &Option<u16>) -> PyResult<()> {
    if let Some(v) = value
        && *v > 360
    {
        return Err(param_err(format!(
            "{name} must be between 0 and 360, got {v}"
        )));
    }
    Ok(())
}

fn check_hole_size(name: &str, value: &Option<u8>) -> PyResult<()> {
    if let Some(v) = value
        && *v > 90
    {
        return Err(param_err(format!(
            "{name} must be between 0 and 90, got {v}"
        )));
    }
    Ok(())
}

fn check_size(name: &str, value: &u32) -> PyResult<()> {
    if *value == 0 {
        return Err(param_err(format!("{name} must be greater than 0")));
    }
    Ok(())
}

fn check_scale(name: &str, value: &f64) -> PyResult<()> {
    if !value.is_finite() || *value <= 0.0 {
        return Err(param_err(format!(
            "{name} must be greater than 0, got {value}"
        )));
    }
    Ok(())
}

fn check_style(name: &str, value: &u8) -> PyResult<()> {
    if !(1..=48).contains(value) {
        return Err(param_err(format!(
            "{name} must be between 1 and 48, got {value}"
        )));
    }
    Ok(())
}

fn check_alt_text(name: &str, value: &Option<String>) -> PyResult<()> {
    if let Some(text) = value
        && text.encode_utf16().count() > 255
    {
        return Err(param_err(format!(
            "{name} is longer than Excel's limit of 255 characters"
        )));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Chart
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
pub struct ChartSpec {
    pub chart_type: ChartType,
    pub rotation: Option<u16>,
    pub hole_size: Option<u8>,
    pub up_down_bars: bool,
    pub high_low_lines: bool,
    pub drop_lines: bool,
    pub width: u32,
    pub height: u32,
    pub scale_width: f64,
    pub scale_height: f64,
    pub style: u8,
    pub name: Option<String>,
    pub alt_text: Option<String>,
    pub decorative: bool,
    pub object_movement: ObjectMovement,
    pub show_empty_cells_as: Option<ChartEmptyCells>,
    pub show_na_as_empty_cell: bool,
    pub show_hidden_data: bool,
}

impl Default for ChartSpec {
    fn default() -> Self {
        ChartSpec {
            chart_type: ChartType::Column,
            rotation: None,
            hole_size: None,
            up_down_bars: false,
            high_low_lines: false,
            drop_lines: false,
            width: 480,
            height: 288,
            scale_width: 1.0,
            scale_height: 1.0,
            style: 2,
            name: None,
            alt_text: None,
            decorative: false,
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            show_empty_cells_as: None,
            show_na_as_empty_cell: false,
            show_hidden_data: false,
        }
    }
}

/// An Excel chart.
///
/// ```python
/// chart = Chart(ChartType.Line, width=640, style=10)
/// s = chart.add_series(values="=Sheet1!$B$2:$B$7", categories="=Sheet1!$A$2:$A$7", name="Sales")
/// s.marker.type = ChartMarkerType.Circle
/// chart.title = "Sales"
/// chart.legend.position = ChartLegendPosition.Bottom
/// worksheet.insert_chart("E2", chart)
/// ```
///
/// Nested objects (`title`, axes, `legend`, formats, ...) are live: reading
/// them returns the same object each time. Every property can also be passed
/// to the constructor as a keyword argument.
#[pyclass(module = "pyrust_xlsxwriter")]
#[derive(Default)]
pub struct Chart {
    pub spec: ChartSpec,
    pub title: Slot<ChartTitle>,
    pub x_axis: Slot<ChartAxis>,
    pub y_axis: Slot<ChartAxis>,
    pub x2_axis: Slot<ChartAxis>,
    pub y2_axis: Slot<ChartAxis>,
    pub legend: Slot<ChartLegend>,
    pub chart_area: Slot<ChartArea>,
    pub plot_area: Slot<ChartPlotArea>,
    pub up_bar_format: Slot<ChartFormat>,
    pub down_bar_format: Slot<ChartFormat>,
    pub high_low_lines_format: Slot<ChartFormat>,
    pub drop_lines_format: Slot<ChartFormat>,
    pub data_table: Slot<ChartDataTable>,
    pub series: Vec<Py<ChartSeries>>,
    pub combined: Option<Py<Chart>>,
}

chart_props!(Chart
    fields {
        /// The chart type.
        chart_type: ChartType,
        /// Rotation of the first slice of a pie/doughnut chart (0 to 360 degrees).
        rotation: Option<u16> => check_rotation,
        /// Hole size of a doughnut chart in percent (0 to 90).
        hole_size: Option<u8> => check_hole_size,
        /// Width in pixels (default 480).
        width: u32 => check_size,
        /// Height in pixels (default 288).
        height: u32 => check_size,
        /// Width scale factor (default 1.0).
        scale_width: f64 => check_scale,
        /// Height scale factor (default 1.0).
        scale_height: f64 => check_scale,
        /// Excel chart style 1 to 48 (default 2).
        style: u8 => check_style,
        /// Object name of the chart.
        name: Option<String>,
        /// Alternative text for accessibility (up to 255 characters).
        alt_text: Option<String> => check_alt_text,
        /// Mark the chart as decorative (no alt text needed).
        decorative: bool,
        /// How the chart moves/resizes with the cells under it.
        object_movement: ObjectMovement,
        /// How empty cells are plotted.
        show_empty_cells_as: Option<ChartEmptyCells>,
        /// Plot #N/A as an empty cell.
        show_na_as_empty_cell: bool,
        /// Plot data in hidden rows and columns.
        show_hidden_data: bool,
    }
    children {
        /// The X axis (the category axis, or the value axis of bar charts).
        x_axis: ChartAxis,
        /// The Y axis.
        y_axis: ChartAxis,
        /// The secondary X axis.
        x2_axis: ChartAxis,
        /// The secondary Y axis.
        y2_axis: ChartAxis,
        /// The legend.
        legend: ChartLegend,
        /// The chart area (background/border of the whole chart).
        chart_area: ChartArea,
        /// The plot area.
        plot_area: ChartPlotArea,
        /// Format of the up bars (setting it shows up/down bars).
        up_bar_format: ChartFormat,
        /// Format of the down bars (setting it shows up/down bars).
        down_bar_format: ChartFormat,
        /// Format of the high-low lines (setting it shows them).
        high_low_lines_format: ChartFormat,
        /// Format of the drop lines (setting it shows them).
        drop_lines_format: ChartFormat,
        /// Data table below the chart (assign a ChartDataTable or True to show one).
        data_table: ChartDataTable,
    }
    custom {
        /// The chart title. Assign a string or (sheet, row, col) tuple to set its
        /// name, a ChartTitle to replace it, or None to reset it.
        title,
        /// The series, in order (a list; assign a list to replace them).
        series,
        /// A chart combined with this one (see `combine()`), or None. Only its
        /// series are used; axes are set on this chart.
        combined_chart,
        /// Show up/down bars (line and stock charts).
        up_down_bars,
        /// Show high-low lines (line and stock charts).
        high_low_lines,
        /// Show drop lines (line and area charts).
        drop_lines,
        /// Shortcut for `x_axis.name` (as a string).
        x_axis_name,
        /// Shortcut for `y_axis.name` (as a string).
        y_axis_name,
        /// Shortcut for `legend.position`.
        legend_position,
    }
);

impl Chart {
    fn get_title(&self, py: Python<'_>) -> PyResult<Py<ChartTitle>> {
        self.title.get(py)
    }

    fn set_title(&mut self, py: Python<'_>, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        match value {
            Some(v) if v.cast::<PyString>().is_ok() || v.cast::<PyTuple>().is_ok() => {
                let name: RangeSpec = v.extract()?;
                self.title.with_mut(py, |title| {
                    title.spec.name = Some(name);
                    Ok(())
                })
            }
            other => self.title.assign(other.as_ref()),
        }
    }

    fn get_series(&self, py: Python<'_>) -> PyResult<Vec<Py<ChartSeries>>> {
        Ok(self.series.iter().map(|s| s.clone_ref(py)).collect())
    }

    fn set_series(&mut self, _py: Python<'_>, value: Option<Vec<Py<ChartSeries>>>) -> PyResult<()> {
        self.series = value.unwrap_or_default();
        Ok(())
    }

    fn get_combined_chart(&self, py: Python<'_>) -> PyResult<Option<Py<Chart>>> {
        Ok(self.combined.as_ref().map(|c| c.clone_ref(py)))
    }

    fn set_combined_chart(&mut self, py: Python<'_>, value: Option<Py<Chart>>) -> PyResult<()> {
        if let Some(other) = &value {
            // This chart is mutably borrowed while its setter runs, so
            // borrowing `other` fails exactly when it is this chart.
            if other.bind(py).try_borrow().is_err() {
                return Err(param_err("a chart cannot be combined with itself"));
            }
        }
        self.combined = value;
        Ok(())
    }

    fn get_up_down_bars(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.up_down_bars
            || self.up_bar_format.active(py)
            || self.down_bar_format.active(py))
    }

    fn set_up_down_bars(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.up_down_bars = value;
        Ok(())
    }

    fn get_high_low_lines(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.high_low_lines || self.high_low_lines_format.active(py))
    }

    fn set_high_low_lines(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.high_low_lines = value;
        Ok(())
    }

    fn get_drop_lines(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.spec.drop_lines || self.drop_lines_format.active(py))
    }

    fn set_drop_lines(&mut self, _py: Python<'_>, value: bool) -> PyResult<()> {
        self.spec.drop_lines = value;
        Ok(())
    }

    fn get_x_axis_name(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self
            .x_axis
            .with(py, |axis| Ok(axis.name_text()))?
            .unwrap_or_default())
    }

    fn set_x_axis_name(&mut self, py: Python<'_>, value: Option<RangeSpec>) -> PyResult<()> {
        self.x_axis.with_mut(py, |axis| {
            axis.spec.name = value;
            Ok(())
        })
    }

    fn get_y_axis_name(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self
            .y_axis
            .with(py, |axis| Ok(axis.name_text()))?
            .unwrap_or_default())
    }

    fn set_y_axis_name(&mut self, py: Python<'_>, value: Option<RangeSpec>) -> PyResult<()> {
        self.y_axis.with_mut(py, |axis| {
            axis.spec.name = value;
            Ok(())
        })
    }

    fn get_legend_position(&self, py: Python<'_>) -> PyResult<ChartLegendPosition> {
        Ok(self
            .legend
            .with(py, |legend| Ok(legend.spec.position))?
            .unwrap_or(ChartLegendPosition::Right))
    }

    fn set_legend_position(&mut self, py: Python<'_>, value: ChartLegendPosition) -> PyResult<()> {
        self.legend.with_mut(py, |legend| {
            legend.spec.position = value;
            Ok(())
        })
    }

    /// Build the `rust_xlsxwriter` chart. A combined chart's own combined
    /// chart is ignored (Excel supports one level).
    pub fn build(&self, py: Python<'_>, with_combined: bool) -> PyResult<RChart> {
        let s = &self.spec;
        let mut chart = RChart::new(s.chart_type.into());

        for series in &self.series {
            series.try_borrow(py)?.apply_to(py, chart.add_series())?;
        }

        apply_slot(py, &self.title, chart.title())?;
        apply_slot(py, &self.x_axis, chart.x_axis())?;
        apply_slot(py, &self.y_axis, chart.y_axis())?;
        apply_slot(py, &self.x2_axis, chart.x2_axis())?;
        apply_slot(py, &self.y2_axis, chart.y2_axis())?;
        apply_slot(py, &self.legend, chart.legend())?;
        apply_slot(py, &self.chart_area, chart.chart_area())?;
        apply_slot(py, &self.plot_area, chart.plot_area())?;

        if let Some(rotation) = s.rotation {
            chart.set_rotation(rotation);
        }
        if let Some(hole_size) = s.hole_size {
            chart.set_hole_size(hole_size);
        }

        if s.up_down_bars {
            chart.set_up_down_bars(true);
        }
        if let Some(mut format) = self.up_bar_format.build(py)? {
            chart.set_up_bar_format(&mut format);
        }
        if let Some(mut format) = self.down_bar_format.build(py)? {
            chart.set_down_bar_format(&mut format);
        }
        if s.high_low_lines {
            chart.set_high_low_lines(true);
        }
        if let Some(mut format) = self.high_low_lines_format.build(py)? {
            chart.set_high_low_lines_format(&mut format);
        }
        if s.drop_lines {
            chart.set_drop_lines(true);
        }
        if let Some(mut format) = self.drop_lines_format.build(py)? {
            chart.set_drop_lines_format(&mut format);
        }
        if let Some(table) = self.data_table.build(py)? {
            chart.set_data_table(&table);
        }

        chart.set_width(s.width);
        chart.set_height(s.height);
        chart.set_scale_width(s.scale_width);
        chart.set_scale_height(s.scale_height);
        chart.set_style(s.style);
        if let Some(name) = &s.name {
            chart.set_name(name);
        }
        if let Some(alt_text) = &s.alt_text {
            chart.set_alt_text(alt_text);
        }
        if s.decorative {
            chart.set_decorative(true);
        }
        chart.set_object_movement(s.object_movement.into());
        if let Some(option) = s.show_empty_cells_as {
            chart.show_empty_cells_as(option.into());
        }
        if s.show_na_as_empty_cell {
            chart.show_na_as_empty_cell();
        }
        if s.show_hidden_data {
            chart.show_hidden_data();
        }

        if with_combined && let Some(other) = &self.combined {
            let other = other.try_borrow(py)?.build(py, false)?;
            chart.combine(&other);
        }
        Ok(chart)
    }
}

#[pymethods]
impl Chart {
    #[new]
    #[pyo3(signature = (chart_type=ChartType::Column, **kwargs))]
    fn py_new(
        py: Python<'_>,
        chart_type: ChartType,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let mut chart: Chart = new_with(py, None)?;
        chart.spec.chart_type = chart_type;
        chart.apply_kwargs(py, kwargs)?;
        Ok(chart)
    }

    /// Add a series and return it (live: later changes still apply).
    /// `values`/`categories` are "=Sheet1!$A$1:$A$5" strings or
    /// `(sheet, first_row, first_col, last_row, last_col)` tuples; any other
    /// ChartSeries property can be passed as a keyword.
    #[pyo3(signature = (values=None, categories=None, name=None, **kwargs))]
    fn add_series(
        &mut self,
        py: Python<'_>,
        values: Option<RangeSpec>,
        categories: Option<RangeSpec>,
        name: Option<RangeSpec>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Py<ChartSeries>> {
        let series = Py::new(
            py,
            ChartSeries::create(py, values, categories, name, kwargs)?,
        )?;
        self.series.push(series.clone_ref(py));
        Ok(series)
    }

    /// Append an existing series (the object itself, so later changes apply).
    fn push_series(&mut self, series: Py<ChartSeries>) {
        self.series.push(series);
    }

    /// Combine another chart with this one (e.g. a line chart over a column
    /// chart). Only the other chart's series are used: set titles, the legend
    /// and every axis (including `y2_axis` for series on the secondary axis)
    /// on this chart. Pass None to remove a combined chart.
    fn combine(slf: &Bound<'_, Self>, chart: Option<&Bound<'_, Chart>>) -> PyResult<()> {
        if let Some(chart) = chart
            && chart.is(slf)
        {
            return Err(param_err("a chart cannot be combined with itself"));
        }
        slf.borrow_mut().combined = chart.map(|c| c.clone().unbind());
        Ok(())
    }

    /// Check the chart the way inserting it does (at least one series, every
    /// series has values, scatter series have categories, ...).
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        self.build(py, true)?.validate().py()?;
        Ok(())
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let mut parts = vec![format!("ChartType.{:?}", self.spec.chart_type)];
        let title = self.title.with(py, |t| Ok(t.text()))?.unwrap_or_default();
        if !title.is_empty() {
            parts.push(format!("title={title:?}"));
        }
        parts.push(format!("series={}", self.series.len()));
        parts.extend(
            self.props_repr(py)?
                .into_iter()
                .filter(|p| !p.starts_with("chart_type=")),
        );
        Ok(repr_of("Chart", parts))
    }
}

// ---------------------------------------------------------------------------
// Worksheet.insert_chart, Workbook.add_chartsheet
// ---------------------------------------------------------------------------

impl Worksheet {
    fn insert_chart_impl(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
        method: &'static str,
    ) -> PyResult<()> {
        let py = args.py();
        let items = args.as_slice();
        // A chart as the first argument means cell (0, 0) (handy for chartsheets).
        let ((row, col), rest) = match items.first() {
            Some(first) if first.cast::<Chart>().is_ok() => ((0, 0), items),
            _ => split_cell(items)?,
        };
        let p = Params::new(method, rest, kwargs, &["chart", "x_offset", "y_offset"], 1)?;
        let chart: Bound<'_, Chart> = p
            .required_any(0)?
            .cast::<Chart>()
            .map_err(|_| {
                pyo3::exceptions::PyTypeError::new_err(format!(
                    "{method}() argument 'chart' must be a Chart"
                ))
            })?
            .clone();
        let x_offset: u32 = p.get_or(1, 0)?;
        let y_offset: u32 = p.get_or(2, 0)?;
        let chart = chart.try_borrow()?.build(py, true)?;
        self.inner
            .insert_chart_with_offset(row, col, &chart, x_offset, y_offset)
            .py()?;
        Ok(())
    }
}

#[pymethods]
impl Worksheet {
    /// Insert a chart: `insert_chart(row, col, chart, x_offset=0, y_offset=0)`
    /// or `insert_chart("E2", chart, ...)`. On a chartsheet the position is
    /// ignored and `insert_chart(chart)` is enough. The chart is copied, so it
    /// can be changed and inserted again.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_chart(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        self.insert_chart_impl(args, kwargs, "insert_chart")
    }

    /// Same as `insert_chart(row, col, chart, x_offset, y_offset)`.
    #[pyo3(signature = (*args, **kwargs))]
    fn insert_chart_with_offset(
        &mut self,
        args: &Bound<'_, PyTuple>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        self.insert_chart_impl(args, kwargs, "insert_chart_with_offset")
    }
}

#[pymethods]
impl Workbook {
    /// Add a chartsheet (a sheet that holds a single chart) and return it.
    /// Put the chart on it with `chartsheet.insert_chart(chart)`. Keyword
    /// arguments set sheet properties such as `active`, `hidden` or `tab_color`.
    #[pyo3(signature = (name=None, **kwargs))]
    fn add_chartsheet<'py>(
        &mut self,
        py: Python<'py>,
        name: Option<&str>,
        kwargs: Option<&Bound<'py, PyDict>>,
    ) -> PyResult<Bound<'py, Worksheet>> {
        let names: Vec<String> = self
            .worksheets
            .iter()
            .map(|ws| ws.borrow(py).inner.name())
            .collect();
        let used = |candidate: &str| names.iter().any(|n| n.eq_ignore_ascii_case(candidate));
        let name = match name {
            Some(name) => {
                if used(name) {
                    return Err(new_err(
                        ErrorKind::SheetName,
                        format!("worksheet name {name:?} is already used in this workbook"),
                    ));
                }
                name.to_string()
            }
            None => (1..)
                .map(|n| format!("Chart{n}"))
                .find(|candidate| !used(candidate))
                .unwrap_or_default(),
        };
        let mut sheet = Worksheet::from_inner(RWorksheet::new_chartsheet());
        sheet.inner.set_name(&name).py()?;
        sheet.apply_kwargs(kwargs)?;
        let sheet = Bound::new(py, sheet)?;
        self.worksheets.push(sheet.clone().unbind());
        Ok(sheet)
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Chart>()?;
    enums::register(m)?;
    format::register(m)?;
    parts::register(m)?;
    series::register(m)?;
    Ok(())
}
