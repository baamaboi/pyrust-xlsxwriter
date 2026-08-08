use pyo3::prelude::*;
use rust_xlsxwriter::{Sparkline as RSparkline, SparklineType as RSparklineType};

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SparklineType {
    Line,
    Column,
    WinLoss,
}

impl From<SparklineType> for RSparklineType {
    fn from(st: SparklineType) -> Self {
        match st {
            SparklineType::Line => RSparklineType::Line,
            SparklineType::Column => RSparklineType::Column,
            SparklineType::WinLoss => RSparklineType::WinLose,
        }
    }
}

#[pyclass]
pub struct Sparkline {
    pub(crate) inner: RSparkline,
    range_val: String,
    sparkline_type_val: SparklineType,
    high_point_val: bool,
    low_point_val: bool,
    first_point_val: bool,
    last_point_val: bool,
    markers_val: bool,
    negative_points_val: bool,
    style_val: u8,
}

#[pymethods]
impl Sparkline {
    #[new]
    #[pyo3(signature = (range="", sparkline_type=SparklineType::Line))]
    pub fn new(range: &str, sparkline_type: SparklineType) -> Self {
        let mut sl = RSparkline::new();
        if !range.is_empty() {
            sl = sl.set_range(range);
        }
        sl = sl.set_type(RSparklineType::from(sparkline_type));

        Sparkline {
            inner: sl,
            range_val: range.to_string(),
            sparkline_type_val: sparkline_type,
            high_point_val: false,
            low_point_val: false,
            first_point_val: false,
            last_point_val: false,
            markers_val: false,
            negative_points_val: false,
            style_val: 1,
        }
    }

    #[getter]
    pub fn range(&self) -> String {
        self.range_val.clone()
    }

    #[setter]
    pub fn set_range(&mut self, val: Bound<'_, PyAny>) {
        let r = val.to_string();
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.set_range(&r);
        self.range_val = r;
    }

    #[getter]
    pub fn sparkline_type(&self) -> SparklineType {
        self.sparkline_type_val
    }

    #[setter]
    pub fn set_sparkline_type(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let st = val.extract::<SparklineType>()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.set_type(RSparklineType::from(st));
        self.sparkline_type_val = st;
        Ok(())
    }

    #[getter]
    pub fn high_point(&self) -> bool {
        self.high_point_val
    }

    #[setter]
    pub fn set_high_point(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_high_point(b);
        self.high_point_val = b;
        Ok(())
    }

    #[getter]
    pub fn low_point(&self) -> bool {
        self.low_point_val
    }

    #[setter]
    pub fn set_low_point(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_low_point(b);
        self.low_point_val = b;
        Ok(())
    }

    #[getter]
    pub fn first_point(&self) -> bool {
        self.first_point_val
    }

    #[setter]
    pub fn set_first_point(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_first_point(b);
        self.first_point_val = b;
        Ok(())
    }

    #[getter]
    pub fn last_point(&self) -> bool {
        self.last_point_val
    }

    #[setter]
    pub fn set_last_point(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_last_point(b);
        self.last_point_val = b;
        Ok(())
    }

    #[getter]
    pub fn markers(&self) -> bool {
        self.markers_val
    }

    #[setter]
    pub fn set_markers(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_markers(b);
        self.markers_val = b;
        Ok(())
    }

    #[getter]
    pub fn negative_points(&self) -> bool {
        self.negative_points_val
    }

    #[setter]
    pub fn set_negative_points(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.show_negative_points(b);
        self.negative_points_val = b;
        Ok(())
    }

    #[getter]
    pub fn style(&self) -> u8 {
        self.style_val
    }

    #[setter]
    pub fn set_style(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let s = val.extract::<u8>()?;
        let old_inner = std::mem::replace(&mut self.inner, RSparkline::new());
        self.inner = old_inner.set_style(s);
        self.style_val = s;
        Ok(())
    }
}
