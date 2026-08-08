use pyo3::prelude::*;
use rust_xlsxwriter::{
    Chart as RChart, ChartLegendPosition as RChartLegendPosition, ChartType as RChartType,
};

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartType {
    Area,
    AreaStacked,
    AreaPercentStacked,
    Bar,
    BarStacked,
    BarPercentStacked,
    Column,
    ColumnStacked,
    ColumnPercentStacked,
    Doughnut,
    Line,
    LineStacked,
    LinePercentStacked,
    Pie,
    Radar,
    Scatter,
    ScatterStraight,
    Stock,
}

impl From<ChartType> for RChartType {
    fn from(ct: ChartType) -> Self {
        match ct {
            ChartType::Area => RChartType::Area,
            ChartType::AreaStacked => RChartType::AreaStacked,
            ChartType::AreaPercentStacked => RChartType::AreaPercentStacked,
            ChartType::Bar => RChartType::Bar,
            ChartType::BarStacked => RChartType::BarStacked,
            ChartType::BarPercentStacked => RChartType::BarPercentStacked,
            ChartType::Column => RChartType::Column,
            ChartType::ColumnStacked => RChartType::ColumnStacked,
            ChartType::ColumnPercentStacked => RChartType::ColumnPercentStacked,
            ChartType::Doughnut => RChartType::Doughnut,
            ChartType::Line => RChartType::Line,
            ChartType::LineStacked => RChartType::LineStacked,
            ChartType::LinePercentStacked => RChartType::LinePercentStacked,
            ChartType::Pie => RChartType::Pie,
            ChartType::Radar => RChartType::Radar,
            ChartType::Scatter => RChartType::Scatter,
            ChartType::ScatterStraight => RChartType::ScatterStraight,
            ChartType::Stock => RChartType::Stock,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartLegendPosition {
    Top,
    Bottom,
    Left,
    Right,
    TopRight,
}

impl From<ChartLegendPosition> for RChartLegendPosition {
    fn from(pos: ChartLegendPosition) -> Self {
        match pos {
            ChartLegendPosition::Top => RChartLegendPosition::Top,
            ChartLegendPosition::Bottom => RChartLegendPosition::Bottom,
            ChartLegendPosition::Left => RChartLegendPosition::Left,
            ChartLegendPosition::Right => RChartLegendPosition::Right,
            ChartLegendPosition::TopRight => RChartLegendPosition::TopRight,
        }
    }
}

#[pyclass]
#[derive(Clone)]
pub struct ChartSeries {
    values: String,
    categories: String,
    name: String,
}

#[pymethods]
impl ChartSeries {
    #[new]
    #[pyo3(signature = (values="", categories="", name=""))]
    pub fn new(values: &str, categories: &str, name: &str) -> Self {
        ChartSeries {
            values: values.to_string(),
            categories: categories.to_string(),
            name: name.to_string(),
        }
    }

    #[getter]
    pub fn values(&self) -> String {
        self.values.clone()
    }

    #[setter]
    pub fn set_values(&mut self, val: Bound<'_, PyAny>) {
        self.values = val.to_string();
    }

    #[getter]
    pub fn categories(&self) -> String {
        self.categories.clone()
    }

    #[setter]
    pub fn set_categories(&mut self, val: Bound<'_, PyAny>) {
        self.categories = val.to_string();
    }

    #[getter]
    pub fn name(&self) -> String {
        self.name.clone()
    }

    #[setter]
    pub fn set_name(&mut self, val: Bound<'_, PyAny>) {
        self.name = val.to_string();
    }
}

#[pyclass]
pub struct Chart {
    pub(crate) inner: RChart,
    chart_type_val: ChartType,
    title_val: String,
    x_axis_name_val: String,
    y_axis_name_val: String,
    width_val: u32,
    height_val: u32,
    style_val: u8,
    legend_position_val: ChartLegendPosition,
}

#[pymethods]
impl Chart {
    #[new]
    #[pyo3(signature = (chart_type=ChartType::Column))]
    pub fn new(chart_type: ChartType) -> Self {
        Chart {
            inner: RChart::new(RChartType::from(chart_type)),
            chart_type_val: chart_type,
            title_val: String::new(),
            x_axis_name_val: String::new(),
            y_axis_name_val: String::new(),
            width_val: 480,
            height_val: 288,
            style_val: 1,
            legend_position_val: ChartLegendPosition::Right,
        }
    }

    #[pyo3(signature = (values="", categories="", name=""))]
    pub fn add_series(&mut self, values: &str, categories: &str, name: &str) {
        let series = self.inner.add_series();
        if !values.is_empty() {
            series.set_values(values);
        }
        if !categories.is_empty() {
            series.set_categories(categories);
        }
        if !name.is_empty() {
            series.set_name(name);
        }
    }

    #[getter]
    pub fn chart_type(&self) -> ChartType {
        self.chart_type_val
    }

    #[getter]
    pub fn title(&self) -> String {
        self.title_val.clone()
    }

    #[setter]
    pub fn set_title(&mut self, val: Bound<'_, PyAny>) {
        let t = val.to_string();
        self.inner.title().set_name(&t);
        self.title_val = t;
    }

    #[getter]
    pub fn x_axis_name(&self) -> String {
        self.x_axis_name_val.clone()
    }

    #[setter]
    pub fn set_x_axis_name(&mut self, val: Bound<'_, PyAny>) {
        let name = val.to_string();
        self.inner.x_axis().set_name(&name);
        self.x_axis_name_val = name;
    }

    #[getter]
    pub fn y_axis_name(&self) -> String {
        self.y_axis_name_val.clone()
    }

    #[setter]
    pub fn set_y_axis_name(&mut self, val: Bound<'_, PyAny>) {
        let name = val.to_string();
        self.inner.y_axis().set_name(&name);
        self.y_axis_name_val = name;
    }

    #[getter]
    pub fn width(&self) -> u32 {
        self.width_val
    }

    #[setter]
    pub fn set_width(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let w = val.extract::<u32>()?;
        self.inner.set_width(w);
        self.width_val = w;
        Ok(())
    }

    #[getter]
    pub fn height(&self) -> u32 {
        self.height_val
    }

    #[setter]
    pub fn set_height(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let h = val.extract::<u32>()?;
        self.inner.set_height(h);
        self.height_val = h;
        Ok(())
    }

    #[getter]
    pub fn style(&self) -> u8 {
        self.style_val
    }

    #[setter]
    pub fn set_style(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let s = val.extract::<u8>()?;
        self.inner.set_style(s);
        self.style_val = s;
        Ok(())
    }

    #[getter]
    pub fn legend_position(&self) -> ChartLegendPosition {
        self.legend_position_val
    }

    #[setter]
    pub fn set_legend_position(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let pos = val.extract::<ChartLegendPosition>()?;
        self.inner
            .legend()
            .set_position(RChartLegendPosition::from(pos));
        self.legend_position_val = pos;
        Ok(())
    }
}
