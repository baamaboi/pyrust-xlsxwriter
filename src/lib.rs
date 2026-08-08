use num_bigint::BigInt;
use pyo3::prelude::*;

pub mod chart;
pub mod conditional_format;
pub mod errors;
pub mod format;
pub mod macros;
pub mod sparkline;
pub mod workbook;
pub mod worksheet;

#[pyfunction]
fn add_big_ints(a: BigInt, b: BigInt) -> BigInt {
    return a + b;
}

#[pymodule]
pub mod _pyrust_xlsxwriter {
    use super::chart;
    use super::conditional_format;
    use super::format;
    use super::macros;
    use super::sparkline;
    use super::workbook;
    use super::worksheet;
    use pyo3::prelude::*;
    use rust_decimal::Decimal;

    #[pymodule_export]
    use super::add_big_ints;

    #[pyfunction]
    fn add_decimals(a: Decimal, b: Decimal) -> Decimal {
        return a + b;
    }

    #[pymodule_init]
    fn init(m: &Bound<'_, PyModule>) -> PyResult<()> {
        m.add_class::<workbook::Workbook>()?;
        m.add_class::<worksheet::Worksheet>()?;
        m.add_class::<workbook::Color>()?;
        m.add_class::<workbook::PaperSize>()?;
        m.add_class::<format::Format>()?;
        m.add_class::<format::FormatAlign>()?;
        m.add_class::<format::FormatBorder>()?;
        m.add_class::<format::FormatUnderline>()?;
        m.add_class::<format::FormatScript>()?;
        m.add_class::<format::FormatPattern>()?;

        // Chart
        m.add_class::<chart::Chart>()?;
        m.add_class::<chart::ChartType>()?;
        m.add_class::<chart::ChartLegendPosition>()?;
        m.add_class::<chart::ChartSeries>()?;

        // Sparkline
        m.add_class::<sparkline::Sparkline>()?;
        m.add_class::<sparkline::SparklineType>()?;

        // Conditional Format
        m.add_class::<conditional_format::ConditionalFormat2ColorScale>()?;
        m.add_class::<conditional_format::ConditionalFormat3ColorScale>()?;
        m.add_class::<conditional_format::ConditionalFormatDataBar>()?;

        // Macros / VBA
        m.add_class::<macros::VbaProject>()?;

        Ok(())
    }
}
