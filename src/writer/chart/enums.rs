//! Chart enums.
//!
//! Plain Rust enums become Python enums via `py_enum!`. The three enums whose
//! Rust variants carry data (`ChartAxisCrossing`, `ChartErrorBarsType`,
//! `ChartTrendlineType`) are small frozen classes instead: data-less variants
//! are class attributes (`ChartAxisCrossing.Max`) and data-carrying variants
//! are static constructors (`ChartAxisCrossing.AxisValue(2.5)`). Each exposes
//! `kind` (the variant name) and `value` (its data, or None).

use pyo3::IntoPyObjectExt;
use pyo3::basic::CompareOp;
use pyo3::prelude::*;
use rust_xlsxwriter::{
    ChartAxisCrossing as RChartAxisCrossing, ChartAxisDateUnitType as RChartAxisDateUnitType,
    ChartAxisDisplayUnitType as RChartAxisDisplayUnitType,
    ChartAxisLabelAlignment as RChartAxisLabelAlignment,
    ChartAxisLabelPosition as RChartAxisLabelPosition, ChartAxisTickType as RChartAxisTickType,
    ChartDataLabelPosition as RChartDataLabelPosition, ChartEmptyCells as RChartEmptyCells,
    ChartErrorBarsDirection as RChartErrorBarsDirection, ChartErrorBarsType as RChartErrorBarsType,
    ChartGradientFillType as RChartGradientFillType, ChartLegendPosition as RChartLegendPosition,
    ChartLineDashType as RChartLineDashType, ChartMarkerType as RChartMarkerType,
    ChartPatternFillType as RChartPatternFillType, ChartTrendlineType as RChartTrendlineType,
    ChartType as RChartType,
};

use super::node::RangeSpec;
use crate::common::errors::param_err;

crate::py_enum!(
    /// Chart types.
    ChartType => RChartType {
        Area, AreaStacked, AreaPercentStacked, Bar, BarStacked, BarPercentStacked, Column,
        ColumnStacked, ColumnPercentStacked, Doughnut, Line, LineStacked, LinePercentStacked, Pie,
        Radar, RadarWithMarkers, RadarFilled, Scatter, ScatterStraight, ScatterStraightWithMarkers,
        ScatterSmooth, ScatterSmoothWithMarkers, Stock,
    }
);

crate::py_enum!(
    /// Date unit for the major/minor units of a date axis.
    ChartAxisDateUnitType => RChartAxisDateUnitType { Days, Months, Years }
);

crate::py_enum!(
    /// Display units of a value axis (hundreds, thousands, ...).
    ChartAxisDisplayUnitType => RChartAxisDisplayUnitType {
        #[pyo3(name = "None_")] None, Hundreds, Thousands, TenThousands, HundredThousands,
        Millions, TenMillions, HundredMillions, Billions, Trillions,
    }
);

crate::py_enum!(
    /// Alignment of category axis labels.
    ChartAxisLabelAlignment => RChartAxisLabelAlignment { Center, Left, Right }
);

crate::py_enum!(
    /// Position of axis labels relative to the axis.
    ChartAxisLabelPosition => RChartAxisLabelPosition {
        NextTo, High, Low, #[pyo3(name = "None_")] None,
    }
);

crate::py_enum!(
    /// Axis tick mark types.
    ChartAxisTickType => RChartAxisTickType { #[pyo3(name = "None_")] None, Inside, Outside, Cross }
);

crate::py_enum!(
    /// Position of data labels.
    ChartDataLabelPosition => RChartDataLabelPosition {
        Default, Center, Right, Left, Above, Below, InsideBase, InsideEnd, OutsideEnd, BestFit,
    }
);

crate::py_enum!(
    /// How empty cells are plotted.
    ChartEmptyCells => RChartEmptyCells { Gaps, Zero, Connected }
);

crate::py_enum!(
    /// Direction of error bars.
    ChartErrorBarsDirection => RChartErrorBarsDirection { Both, Minus, Plus }
);

crate::py_enum!(
    /// Gradient fill types.
    ChartGradientFillType => RChartGradientFillType { Linear, Radial, Rectangular, Path }
);

crate::py_enum!(
    /// Legend positions.
    ChartLegendPosition => RChartLegendPosition { Right, Left, Top, Bottom, TopRight }
);

crate::py_enum!(
    /// Line dash types.
    ChartLineDashType => RChartLineDashType {
        Solid, RoundDot, SquareDot, Dash, DashDot, LongDash, LongDashDot, LongDashDotDot,
    }
);

crate::py_enum!(
    /// Marker symbols for line, scatter and radar series.
    ChartMarkerType => RChartMarkerType {
        Square, Diamond, Triangle, X, Star, ShortDash, LongDash, Circle, PlusSign,
    }
);

crate::py_enum!(
    /// Pattern fill types.
    ChartPatternFillType => RChartPatternFillType {
        Dotted5Percent, Dotted10Percent, Dotted20Percent, Dotted25Percent, Dotted30Percent,
        Dotted40Percent, Dotted50Percent, Dotted60Percent, Dotted70Percent, Dotted75Percent,
        Dotted80Percent, Dotted90Percent, DiagonalStripesLightDownwards,
        DiagonalStripesLightUpwards, DiagonalStripesDarkDownwards, DiagonalStripesDarkUpwards,
        DiagonalStripesWideDownwards, DiagonalStripesWideUpwards, VerticalStripesLight,
        HorizontalStripesLight, VerticalStripesNarrow, HorizontalStripesNarrow,
        VerticalStripesDark, HorizontalStripesDark, StripesBackslashes, StripesForwardSlashes,
        HorizontalStripesAlternating, VerticalStripesAlternating, SmallConfetti, LargeConfetti,
        Zigzag, Wave, DiagonalBrick, HorizontalBrick, Weave, Plaid, Divot, DottedGrid,
        DottedDiamond, Shingle, Trellis, Sphere, SmallGrid, LargeGrid, SmallCheckerboard,
        LargeCheckerboard, OutlinedDiamondGrid, SolidDiamondGrid,
    }
);

fn hash_of(value: impl std::hash::Hash) -> u64 {
    use std::hash::Hasher;
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    value.hash(&mut hasher);
    hasher.finish()
}

fn compare<T: PartialEq>(
    this: &T,
    other: Option<T>,
    op: CompareOp,
    py: Python<'_>,
) -> PyResult<Py<PyAny>> {
    let Some(other) = other else {
        return Ok(py.NotImplemented());
    };
    match op {
        CompareOp::Eq => (*this == other).into_py_any(py),
        CompareOp::Ne => (*this != other).into_py_any(py),
        _ => Ok(py.NotImplemented()),
    }
}

fn check_finite(what: &str, value: f64) -> PyResult<f64> {
    if !value.is_finite() {
        return Err(param_err(format!(
            "{what} must be a finite number, got {value}"
        )));
    }
    Ok(value)
}

// ---------------------------------------------------------------------------
// ChartAxisCrossing
// ---------------------------------------------------------------------------

/// Where an axis crosses the perpendicular axis.
///
/// `ChartAxisCrossing.Automatic` (the default), `.Min`, `.Max`,
/// `ChartAxisCrossing.CategoryNumber(3)` (category axes) or
/// `ChartAxisCrossing.AxisValue(2.5)` (value and date axes).
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Copy, PartialEq)]
pub struct ChartAxisCrossing(pub RChartAxisCrossing);

#[pymethods]
impl ChartAxisCrossing {
    /// Excel's automatic crossing position.
    #[classattr]
    #[pyo3(name = "Automatic")]
    fn automatic() -> Self {
        ChartAxisCrossing(RChartAxisCrossing::Automatic)
    }

    /// Cross at the minimum of the perpendicular axis.
    #[classattr]
    #[pyo3(name = "Min")]
    fn min() -> Self {
        ChartAxisCrossing(RChartAxisCrossing::Min)
    }

    /// Cross at the maximum of the perpendicular axis.
    #[classattr]
    #[pyo3(name = "Max")]
    fn max() -> Self {
        ChartAxisCrossing(RChartAxisCrossing::Max)
    }

    /// Cross at a category number (category axes; 1-based, as in Excel).
    #[staticmethod]
    #[pyo3(name = "CategoryNumber")]
    fn category_number(number: u32) -> Self {
        ChartAxisCrossing(RChartAxisCrossing::CategoryNumber(number))
    }

    /// Cross at a value (value and date axes).
    #[staticmethod]
    #[pyo3(name = "AxisValue")]
    fn axis_value(value: f64) -> PyResult<Self> {
        Ok(ChartAxisCrossing(RChartAxisCrossing::AxisValue(
            check_finite("AxisValue", value)?,
        )))
    }

    /// The variant name: "Automatic", "Min", "Max", "CategoryNumber" or "AxisValue".
    #[getter]
    fn kind(&self) -> &'static str {
        match self.0 {
            RChartAxisCrossing::Automatic => "Automatic",
            RChartAxisCrossing::Min => "Min",
            RChartAxisCrossing::Max => "Max",
            RChartAxisCrossing::CategoryNumber(_) => "CategoryNumber",
            RChartAxisCrossing::AxisValue(_) => "AxisValue",
        }
    }

    /// The category number or axis value (None for the other variants).
    #[getter]
    fn value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        match self.0 {
            RChartAxisCrossing::CategoryNumber(n) => n.into_py_any(py),
            RChartAxisCrossing::AxisValue(v) => v.into_py_any(py),
            _ => Ok(py.None()),
        }
    }

    fn __repr__(&self) -> String {
        match self.0 {
            RChartAxisCrossing::CategoryNumber(n) => {
                format!("ChartAxisCrossing.CategoryNumber({n})")
            }
            RChartAxisCrossing::AxisValue(v) => format!("ChartAxisCrossing.AxisValue({v:?})"),
            _ => format!("ChartAxisCrossing.{}", self.kind()),
        }
    }

    fn __richcmp__(&self, other: &Bound<'_, PyAny>, op: CompareOp) -> PyResult<Py<PyAny>> {
        let theirs = other.cast::<ChartAxisCrossing>().ok().map(|o| o.get().0);
        compare(&self.0, theirs, op, other.py())
    }

    fn __hash__(&self) -> u64 {
        match self.0 {
            RChartAxisCrossing::CategoryNumber(n) => hash_of(("CategoryNumber", n)),
            RChartAxisCrossing::AxisValue(v) => hash_of(("AxisValue", v.to_bits())),
            _ => hash_of(self.kind()),
        }
    }
}

// ---------------------------------------------------------------------------
// ChartErrorBarsType
// ---------------------------------------------------------------------------

/// The data of a [`ChartErrorBarsType`].
#[derive(Clone, Debug, PartialEq)]
pub enum ErrorBarsKind {
    FixedValue(f64),
    Percentage(f64),
    StandardDeviation(f64),
    StandardError,
    Custom(RangeSpec, RangeSpec),
}

/// How the size of error bars is calculated.
///
/// `ChartErrorBarsType.StandardError` (the default),
/// `ChartErrorBarsType.FixedValue(2)`, `.Percentage(5)`,
/// `.StandardDeviation(1)`, or `.Custom(plus, minus)` where `plus`/`minus`
/// are ranges such as "=Sheet1!$C$1:$C$5" or (sheet, r1, c1, r2, c2) tuples.
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, PartialEq)]
pub struct ChartErrorBarsType(pub ErrorBarsKind);

impl Default for ChartErrorBarsType {
    fn default() -> Self {
        ChartErrorBarsType(ErrorBarsKind::StandardError)
    }
}

impl ChartErrorBarsType {
    pub fn to_rust(&self) -> RChartErrorBarsType {
        match &self.0 {
            ErrorBarsKind::FixedValue(v) => RChartErrorBarsType::FixedValue(*v),
            ErrorBarsKind::Percentage(v) => RChartErrorBarsType::Percentage(*v),
            ErrorBarsKind::StandardDeviation(v) => RChartErrorBarsType::StandardDeviation(*v),
            ErrorBarsKind::StandardError => RChartErrorBarsType::StandardError,
            ErrorBarsKind::Custom(plus, minus) => {
                RChartErrorBarsType::Custom(plus.to_rust(), minus.to_rust())
            }
        }
    }

    fn kind_name(&self) -> &'static str {
        match self.0 {
            ErrorBarsKind::FixedValue(_) => "FixedValue",
            ErrorBarsKind::Percentage(_) => "Percentage",
            ErrorBarsKind::StandardDeviation(_) => "StandardDeviation",
            ErrorBarsKind::StandardError => "StandardError",
            ErrorBarsKind::Custom(_, _) => "Custom",
        }
    }
}

#[pymethods]
impl ChartErrorBarsType {
    /// Standard error (the default).
    #[classattr]
    #[pyo3(name = "StandardError")]
    fn standard_error() -> Self {
        ChartErrorBarsType(ErrorBarsKind::StandardError)
    }

    /// A fixed value (must be greater than 0).
    #[staticmethod]
    #[pyo3(name = "FixedValue")]
    fn fixed_value(value: f64) -> PyResult<Self> {
        if !value.is_finite() || value <= 0.0 {
            return Err(param_err(format!(
                "FixedValue must be greater than 0, got {value}"
            )));
        }
        Ok(ChartErrorBarsType(ErrorBarsKind::FixedValue(value)))
    }

    /// A percentage of each value (0 or more).
    #[staticmethod]
    #[pyo3(name = "Percentage")]
    fn percentage(value: f64) -> PyResult<Self> {
        if !value.is_finite() || value < 0.0 {
            return Err(param_err(format!(
                "Percentage must be 0 or more, got {value}"
            )));
        }
        Ok(ChartErrorBarsType(ErrorBarsKind::Percentage(value)))
    }

    /// A number of standard deviations (0 or more).
    #[staticmethod]
    #[pyo3(name = "StandardDeviation")]
    fn standard_deviation(value: f64) -> PyResult<Self> {
        if !value.is_finite() || value < 0.0 {
            return Err(param_err(format!(
                "StandardDeviation must be 0 or more, got {value}"
            )));
        }
        Ok(ChartErrorBarsType(ErrorBarsKind::StandardDeviation(value)))
    }

    /// Custom plus/minus values taken from worksheet ranges.
    #[staticmethod]
    #[pyo3(name = "Custom")]
    fn custom(plus: RangeSpec, minus: RangeSpec) -> Self {
        ChartErrorBarsType(ErrorBarsKind::Custom(plus, minus))
    }

    /// The variant name, e.g. "FixedValue" or "StandardError".
    #[getter]
    fn kind(&self) -> &'static str {
        self.kind_name()
    }

    /// The numeric value (None for StandardError and Custom).
    #[getter]
    fn value(&self) -> Option<f64> {
        match self.0 {
            ErrorBarsKind::FixedValue(v)
            | ErrorBarsKind::Percentage(v)
            | ErrorBarsKind::StandardDeviation(v) => Some(v),
            _ => None,
        }
    }

    /// The plus range of a Custom type (None otherwise).
    #[getter]
    fn plus(&self) -> Option<RangeSpec> {
        match &self.0 {
            ErrorBarsKind::Custom(plus, _) => Some(plus.clone()),
            _ => None,
        }
    }

    /// The minus range of a Custom type (None otherwise).
    #[getter]
    fn minus(&self) -> Option<RangeSpec> {
        match &self.0 {
            ErrorBarsKind::Custom(_, minus) => Some(minus.clone()),
            _ => None,
        }
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(match &self.0 {
            ErrorBarsKind::StandardError => "ChartErrorBarsType.StandardError".to_string(),
            ErrorBarsKind::Custom(plus, minus) => format!(
                "ChartErrorBarsType.Custom({}, {})",
                plus.clone().into_bound_py_any(py)?.repr()?,
                minus.clone().into_bound_py_any(py)?.repr()?
            ),
            ErrorBarsKind::FixedValue(v)
            | ErrorBarsKind::Percentage(v)
            | ErrorBarsKind::StandardDeviation(v) => {
                format!("ChartErrorBarsType.{}({v:?})", self.kind_name())
            }
        })
    }

    fn __richcmp__(&self, other: &Bound<'_, PyAny>, op: CompareOp) -> PyResult<Py<PyAny>> {
        let theirs = other
            .cast::<ChartErrorBarsType>()
            .ok()
            .map(|o| o.get().0.clone());
        compare(&self.0, theirs, op, other.py())
    }

    fn __hash__(&self) -> u64 {
        match &self.0 {
            ErrorBarsKind::Custom(plus, minus) => {
                hash_of(("Custom", plus.display(), minus.display()))
            }
            other => hash_of((
                self.kind_name(),
                self.value().unwrap_or(0.0).to_bits(),
                matches!(other, ErrorBarsKind::StandardError),
            )),
        }
    }
}

// ---------------------------------------------------------------------------
// ChartTrendlineType
// ---------------------------------------------------------------------------

/// Trendline types.
///
/// `ChartTrendlineType.None_` (the default), `.Exponential`, `.Linear`,
/// `.Logarithmic`, `.Power`, `ChartTrendlineType.Polynomial(order)` or
/// `ChartTrendlineType.MovingAverage(period)`.
#[pyclass(module = "pyrust_xlsxwriter", frozen, from_py_object)]
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ChartTrendlineType(pub RChartTrendlineType);

impl Default for ChartTrendlineType {
    fn default() -> Self {
        ChartTrendlineType(RChartTrendlineType::None)
    }
}

impl ChartTrendlineType {
    fn kind_name(&self) -> &'static str {
        match self.0 {
            RChartTrendlineType::None => "None_",
            RChartTrendlineType::Exponential => "Exponential",
            RChartTrendlineType::Linear => "Linear",
            RChartTrendlineType::Logarithmic => "Logarithmic",
            RChartTrendlineType::Polynomial(_) => "Polynomial",
            RChartTrendlineType::Power => "Power",
            RChartTrendlineType::MovingAverage(_) => "MovingAverage",
        }
    }
}

#[pymethods]
impl ChartTrendlineType {
    /// No trendline (the default).
    #[classattr]
    #[pyo3(name = "None_")]
    fn none() -> Self {
        ChartTrendlineType(RChartTrendlineType::None)
    }

    /// Exponential fit.
    #[classattr]
    #[pyo3(name = "Exponential")]
    fn exponential() -> Self {
        ChartTrendlineType(RChartTrendlineType::Exponential)
    }

    /// Linear fit.
    #[classattr]
    #[pyo3(name = "Linear")]
    fn linear() -> Self {
        ChartTrendlineType(RChartTrendlineType::Linear)
    }

    /// Logarithmic fit.
    #[classattr]
    #[pyo3(name = "Logarithmic")]
    fn logarithmic() -> Self {
        ChartTrendlineType(RChartTrendlineType::Logarithmic)
    }

    /// Power fit.
    #[classattr]
    #[pyo3(name = "Power")]
    fn power() -> Self {
        ChartTrendlineType(RChartTrendlineType::Power)
    }

    /// Polynomial fit of the given order (2 to 5, the range rust_xlsxwriter accepts).
    #[staticmethod]
    #[pyo3(name = "Polynomial")]
    fn polynomial(order: u8) -> PyResult<Self> {
        if !(2..=5).contains(&order) {
            return Err(param_err(format!(
                "Polynomial order must be between 2 and 5, got {order}"
            )));
        }
        Ok(ChartTrendlineType(RChartTrendlineType::Polynomial(order)))
    }

    /// Moving average over the given period (2 or 3, the range rust_xlsxwriter accepts).
    #[staticmethod]
    #[pyo3(name = "MovingAverage")]
    fn moving_average(period: u8) -> PyResult<Self> {
        if !(2..=3).contains(&period) {
            return Err(param_err(format!(
                "MovingAverage period must be 2 or 3, got {period}"
            )));
        }
        Ok(ChartTrendlineType(RChartTrendlineType::MovingAverage(
            period,
        )))
    }

    /// The variant name, e.g. "Linear" or "Polynomial" ("None_" for no trendline).
    #[getter]
    fn kind(&self) -> &'static str {
        self.kind_name()
    }

    /// The polynomial order or moving-average period (None otherwise).
    #[getter]
    fn value(&self) -> Option<u8> {
        match self.0 {
            RChartTrendlineType::Polynomial(v) | RChartTrendlineType::MovingAverage(v) => Some(v),
            _ => None,
        }
    }

    fn __repr__(&self) -> String {
        match self.value() {
            Some(v) => format!("ChartTrendlineType.{}({v})", self.kind_name()),
            None => format!("ChartTrendlineType.{}", self.kind_name()),
        }
    }

    fn __richcmp__(&self, other: &Bound<'_, PyAny>, op: CompareOp) -> PyResult<Py<PyAny>> {
        let theirs = other.cast::<ChartTrendlineType>().ok().map(|o| o.get().0);
        compare(&self.0, theirs, op, other.py())
    }

    fn __hash__(&self) -> u64 {
        hash_of((self.kind_name(), self.value()))
    }
}

impl std::fmt::Debug for ChartAxisCrossing {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.__repr__())
    }
}

impl std::fmt::Debug for ChartErrorBarsType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ChartErrorBarsType({:?})", self.0)
    }
}

impl std::fmt::Debug for ChartTrendlineType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.__repr__())
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ChartType>()?;
    m.add_class::<ChartAxisCrossing>()?;
    m.add_class::<ChartAxisDateUnitType>()?;
    m.add_class::<ChartAxisDisplayUnitType>()?;
    m.add_class::<ChartAxisLabelAlignment>()?;
    m.add_class::<ChartAxisLabelPosition>()?;
    m.add_class::<ChartAxisTickType>()?;
    m.add_class::<ChartDataLabelPosition>()?;
    m.add_class::<ChartEmptyCells>()?;
    m.add_class::<ChartErrorBarsDirection>()?;
    m.add_class::<ChartErrorBarsType>()?;
    m.add_class::<ChartGradientFillType>()?;
    m.add_class::<ChartLegendPosition>()?;
    m.add_class::<ChartLineDashType>()?;
    m.add_class::<ChartMarkerType>()?;
    m.add_class::<ChartPatternFillType>()?;
    m.add_class::<ChartTrendlineType>()?;
    Ok(())
}
