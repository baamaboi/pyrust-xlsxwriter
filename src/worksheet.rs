use pyo3::prelude::*;
use pyo3::types::PyString;

use crate::chart::Chart;
use crate::conditional_format::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatDataBar,
};
use crate::errors::Error;
use crate::format::Format;
use crate::sparkline::Sparkline;
use crate::workbook::{Color, PaperSize};
use rust_xlsxwriter::Worksheet as RWorksheet;

#[pyclass]
pub struct Worksheet {
    pub(crate) inner: RWorksheet,
    active_val: bool,
    selected_val: bool,
    hidden_val: bool,
    tab_color_val: Option<Py<PyAny>>,
    default_row_height_val: Option<f64>,
    paper_size_val: Option<Py<PyAny>>,
    header_val: String,
    footer_val: String,
    print_center_horizontally_val: bool,
    print_center_vertically_val: bool,
    screen_gridlines_val: bool,
    print_gridlines_val: bool,
    print_headings_val: bool,
    zoom_val: u16,
    right_to_left_val: bool,
}

impl Worksheet {
    fn write_item_internal(
        &mut self,
        row: u32,
        col: u16,
        value: Bound<'_, PyAny>,
        format: Option<&rust_xlsxwriter::Format>,
    ) -> PyResult<()> {
        if let Ok(val) = value.extract::<f64>() {
            if let Some(fmt) = format {
                self.inner
                    .write_number_with_format(row, col, val, fmt)
                    .map_err(Error::RustXlsxError)?;
            } else {
                self.inner
                    .write_number(row, col, val)
                    .map_err(Error::RustXlsxError)?;
            }
        } else if let Ok(val) = value.extract::<bool>() {
            if let Some(fmt) = format {
                self.inner
                    .write_boolean_with_format(row, col, val, fmt)
                    .map_err(Error::RustXlsxError)?;
            } else {
                self.inner
                    .write_boolean(row, col, val)
                    .map_err(Error::RustXlsxError)?;
            }
        } else if let Ok(pystr) = value.cast::<PyString>() {
            let val_str = pystr.to_str()?;
            if let Some(fmt) = format {
                self.inner
                    .write_string_with_format(row, col, val_str, fmt)
                    .map_err(Error::RustXlsxError)?;
            } else {
                self.inner
                    .write_string(row, col, val_str)
                    .map_err(Error::RustXlsxError)?;
            }
        } else {
            let val_str = value.to_string();
            if let Some(fmt) = format {
                self.inner
                    .write_string_with_format(row, col, &val_str, fmt)
                    .map_err(Error::RustXlsxError)?;
            } else {
                self.inner
                    .write_string(row, col, &val_str)
                    .map_err(Error::RustXlsxError)?;
            }
        }
        Ok(())
    }
}

#[pymethods]
impl Worksheet {
    #[new]
    #[pyo3(signature = (
        name = "",
        active = false,
        selected = false,
        hidden = false,
        tab_color = None,
        default_row_height = None,
        paper_size = None,
        header = "",
        footer = "",
        print_center_horizontally = false,
        print_center_vertically = false,
        screen_gridlines = true,
        print_gridlines = false,
        print_headings = false,
        zoom = 100,
        right_to_left = false,
    ))]
    pub fn new(
        name: &str,
        active: bool,
        selected: bool,
        hidden: bool,
        tab_color: Option<Bound<'_, PyAny>>,
        default_row_height: Option<f64>,
        paper_size: Option<Bound<'_, PyAny>>,
        header: &str,
        footer: &str,
        print_center_horizontally: bool,
        print_center_vertically: bool,
        screen_gridlines: bool,
        print_gridlines: bool,
        print_headings: bool,
        zoom: u16,
        right_to_left: bool,
    ) -> PyResult<Self> {
        let mut worksheet = RWorksheet::new();
        if !name.is_empty() {
            worksheet.set_name(name).map_err(Error::RustXlsxError)?;
        }

        if active {
            worksheet.set_active(true);
        }

        if selected {
            worksheet.set_selected(true);
        }

        if hidden {
            worksheet.set_hidden(true);
        }

        let tab_color_val = if let Some(ref color_obj) = tab_color {
            if let Ok(c) = color_obj.extract::<Color>() {
                worksheet.set_tab_color(c);
                Some(color_obj.clone().unbind())
            } else if let Ok(color_str) = color_obj.extract::<&str>() {
                worksheet.set_tab_color(color_str);
                Some(color_obj.clone().unbind())
            } else if let Ok((r, g, b)) = color_obj.extract::<(u8, u8, u8)>() {
                let hex_str = format!("#{:02X}{:02X}{:02X}", r, g, b);
                worksheet.set_tab_color(hex_str.as_str());
                Some(color_obj.clone().unbind())
            } else {
                return Err(pyo3::exceptions::PyTypeError::new_err(
                    "tab_color must be a Color enum, string (name/hex), or (r, g, b) tuple",
                ));
            }
        } else {
            None
        };

        if let Some(height) = default_row_height {
            worksheet.set_default_row_height(height);
        }

        let paper_size_val = if let Some(ref paper_obj) = paper_size {
            if let Ok(ps) = paper_obj.extract::<PaperSize>() {
                worksheet.set_paper_size(ps as u8);
                Some(paper_obj.clone().unbind())
            } else if let Ok(val) = paper_obj.extract::<u8>() {
                worksheet.set_paper_size(val);
                Some(paper_obj.clone().unbind())
            } else {
                return Err(pyo3::exceptions::PyTypeError::new_err(
                    "paper_size must be a PaperSize enum or integer",
                ));
            }
        } else {
            None
        };

        if !header.is_empty() {
            worksheet.set_header(header);
        }

        if !footer.is_empty() {
            worksheet.set_footer(footer);
        }

        if print_center_horizontally {
            worksheet.set_print_center_horizontally(true);
        }

        if print_center_vertically {
            worksheet.set_print_center_vertically(true);
        }

        if !screen_gridlines {
            worksheet.set_screen_gridlines(false);
        }

        if print_gridlines {
            worksheet.set_print_gridlines(true);
        }

        if print_headings {
            worksheet.set_print_headings(true);
        }

        if zoom != 100 {
            worksheet.set_zoom(zoom);
        }

        if right_to_left {
            worksheet.set_right_to_left(true);
        }

        Ok(Worksheet {
            inner: worksheet,
            active_val: active,
            selected_val: selected,
            hidden_val: hidden,
            tab_color_val,
            default_row_height_val: default_row_height,
            paper_size_val,
            header_val: header.to_string(),
            footer_val: footer.to_string(),
            print_center_horizontally_val: print_center_horizontally,
            print_center_vertically_val: print_center_vertically,
            screen_gridlines_val: screen_gridlines,
            print_gridlines_val: print_gridlines,
            print_headings_val: print_headings,
            zoom_val: zoom,
            right_to_left_val: right_to_left,
        })
    }

    #[getter]
    pub fn name(&self) -> String {
        self.inner.name()
    }

    #[setter]
    pub fn set_name(&mut self, name: &str) -> PyResult<()> {
        self.inner.set_name(name).map_err(Error::RustXlsxError)?;
        Ok(())
    }

    #[getter]
    pub fn active(&self) -> bool {
        self.active_val
    }

    #[setter]
    pub fn set_active(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_active(b);
        self.active_val = b;
        Ok(())
    }

    #[getter]
    pub fn selected(&self) -> bool {
        self.selected_val
    }

    #[setter]
    pub fn set_selected(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_selected(b);
        self.selected_val = b;
        Ok(())
    }

    #[getter]
    pub fn hidden(&self) -> bool {
        self.hidden_val
    }

    #[setter]
    pub fn set_hidden(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_hidden(b);
        self.hidden_val = b;
        Ok(())
    }

    #[getter]
    pub fn tab_color(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.tab_color_val.as_ref().map(|c| c.clone_ref(py))
    }

    #[setter]
    pub fn set_tab_color(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.inner.set_tab_color(rust_xlsxwriter::Color::Default);
            self.tab_color_val = None;
            return Ok(());
        }

        if let Ok(c) = val.extract::<Color>() {
            self.inner.set_tab_color(c);
            self.tab_color_val = Some(val.unbind());
            Ok(())
        } else if let Ok(color_str) = val.extract::<&str>() {
            self.inner.set_tab_color(color_str);
            self.tab_color_val = Some(val.unbind());
            Ok(())
        } else if let Ok((r, g, b)) = val.extract::<(u8, u8, u8)>() {
            let hex_str = format!("#{:02X}{:02X}{:02X}", r, g, b);
            self.inner.set_tab_color(hex_str.as_str());
            self.tab_color_val = Some(val.unbind());
            Ok(())
        } else {
            Err(pyo3::exceptions::PyTypeError::new_err(
                "tab_color must be a Color enum, string (name/hex), or (r, g, b) tuple",
            ))
        }
    }

    #[getter]
    pub fn default_row_height(&self) -> Option<f64> {
        self.default_row_height_val
    }

    #[setter]
    pub fn set_default_row_height(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.default_row_height_val = None;
            return Ok(());
        }
        let height = val.extract::<f64>()?;
        self.inner.set_default_row_height(height);
        self.default_row_height_val = Some(height);
        Ok(())
    }

    #[getter]
    pub fn paper_size(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        self.paper_size_val.as_ref().map(|p| p.clone_ref(py))
    }

    #[setter]
    pub fn set_paper_size(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        if val.is_none() {
            self.paper_size_val = None;
            return Ok(());
        }

        if let Ok(ps) = val.extract::<PaperSize>() {
            self.inner.set_paper_size(ps as u8);
            self.paper_size_val = Some(val.unbind());
            Ok(())
        } else if let Ok(pval) = val.extract::<u8>() {
            self.inner.set_paper_size(pval);
            self.paper_size_val = Some(val.unbind());
            Ok(())
        } else {
            Err(pyo3::exceptions::PyTypeError::new_err(
                "paper_size must be a PaperSize enum or integer",
            ))
        }
    }

    #[getter]
    pub fn header(&self) -> String {
        self.header_val.clone()
    }

    #[setter]
    pub fn set_header(&mut self, val: Bound<'_, PyAny>) {
        let h = val.to_string();
        self.inner.set_header(&h);
        self.header_val = h;
    }

    #[getter]
    pub fn footer(&self) -> String {
        self.footer_val.clone()
    }

    #[setter]
    pub fn set_footer(&mut self, val: Bound<'_, PyAny>) {
        let f = val.to_string();
        self.inner.set_footer(&f);
        self.footer_val = f;
    }

    #[getter]
    pub fn print_center_horizontally(&self) -> bool {
        self.print_center_horizontally_val
    }

    #[setter]
    pub fn set_print_center_horizontally(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_print_center_horizontally(b);
        self.print_center_horizontally_val = b;
        Ok(())
    }

    #[getter]
    pub fn print_center_vertically(&self) -> bool {
        self.print_center_vertically_val
    }

    #[setter]
    pub fn set_print_center_vertically(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_print_center_vertically(b);
        self.print_center_vertically_val = b;
        Ok(())
    }

    #[getter]
    pub fn screen_gridlines(&self) -> bool {
        self.screen_gridlines_val
    }

    #[setter]
    pub fn set_screen_gridlines(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_screen_gridlines(b);
        self.screen_gridlines_val = b;
        Ok(())
    }

    #[getter]
    pub fn print_gridlines(&self) -> bool {
        self.print_gridlines_val
    }

    #[setter]
    pub fn set_print_gridlines(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_print_gridlines(b);
        self.print_gridlines_val = b;
        Ok(())
    }

    #[getter]
    pub fn print_headings(&self) -> bool {
        self.print_headings_val
    }

    #[setter]
    pub fn set_print_headings(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_print_headings(b);
        self.print_headings_val = b;
        Ok(())
    }

    #[getter]
    pub fn zoom(&self) -> u16 {
        self.zoom_val
    }

    #[setter]
    pub fn set_zoom(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let z = val.extract::<u16>()?;
        self.inner.set_zoom(z);
        self.zoom_val = z;
        Ok(())
    }

    #[getter]
    pub fn right_to_left(&self) -> bool {
        self.right_to_left_val
    }

    #[setter]
    pub fn set_right_to_left(&mut self, val: Bound<'_, PyAny>) -> PyResult<()> {
        let b = val.is_truthy()?;
        self.inner.set_right_to_left(b);
        self.right_to_left_val = b;
        Ok(())
    }

    #[pyo3(signature = (row, col, value, format=None))]
    pub fn write(
        &mut self,
        row: u32,
        col: u16,
        value: Bound<'_, PyAny>,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        self.write_item_internal(row, col, value, format.as_ref().map(|f| &f.inner))
    }

    #[pyo3(signature = (row, col, data, format=None))]
    pub fn write_row(
        &mut self,
        row: u32,
        col: u16,
        data: Bound<'_, PyAny>,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        let items: Vec<Bound<'_, PyAny>> = data.extract()?;
        let fmt_inner = format.as_ref().map(|f| &f.inner);
        for (i, item) in items.into_iter().enumerate() {
            let c = col + i as u16;
            self.write_item_internal(row, c, item, fmt_inner)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, data, format=None))]
    pub fn write_column(
        &mut self,
        row: u32,
        col: u16,
        data: Bound<'_, PyAny>,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        let items: Vec<Bound<'_, PyAny>> = data.extract()?;
        let fmt_inner = format.as_ref().map(|f| &f.inner);
        for (i, item) in items.into_iter().enumerate() {
            let r = row + i as u32;
            self.write_item_internal(r, col, item, fmt_inner)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, number, format=None))]
    pub fn write_number(
        &mut self,
        row: u32,
        col: u16,
        number: f64,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        if let Some(fmt) = format {
            self.inner
                .write_number_with_format(row, col, number, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        } else {
            self.inner
                .write_number(row, col, number)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, string, format=None))]
    pub fn write_string(
        &mut self,
        row: u32,
        col: u16,
        string: &str,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        if let Some(fmt) = format {
            self.inner
                .write_string_with_format(row, col, string, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        } else {
            self.inner
                .write_string(row, col, string)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, value, format=None))]
    pub fn write_boolean(
        &mut self,
        row: u32,
        col: u16,
        value: bool,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        if let Some(fmt) = format {
            self.inner
                .write_boolean_with_format(row, col, value, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        } else {
            self.inner
                .write_boolean(row, col, value)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, formula, format=None))]
    pub fn write_formula(
        &mut self,
        row: u32,
        col: u16,
        formula: &str,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        if let Some(fmt) = format {
            self.inner
                .write_formula_with_format(row, col, formula, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        } else {
            self.inner
                .write_formula(row, col, formula)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (first_row, first_col, last_row, last_col, string, format=None))]
    pub fn merge_range(
        &mut self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        string: &str,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        if let Some(fmt) = format {
            self.inner
                .merge_range(first_row, first_col, last_row, last_col, string, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        } else {
            let empty_fmt = rust_xlsxwriter::Format::new();
            self.inner
                .merge_range(first_row, first_col, last_row, last_col, string, &empty_fmt)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (col, width, format=None))]
    pub fn set_column_width(
        &mut self,
        col: u16,
        width: f64,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        self.inner
            .set_column_width(col, width)
            .map_err(Error::RustXlsxError)?;
        if let Some(fmt) = format {
            self.inner
                .set_column_format(col, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (first_col, last_col, width, format=None))]
    pub fn set_column_range_width(
        &mut self,
        first_col: u16,
        last_col: u16,
        width: f64,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        self.inner
            .set_column_range_width(first_col, last_col, width)
            .map_err(Error::RustXlsxError)?;
        if let Some(fmt) = format {
            self.inner
                .set_column_range_format(first_col, last_col, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, height, format=None))]
    pub fn set_row_height(
        &mut self,
        row: u32,
        height: f64,
        format: Option<PyRef<'_, Format>>,
    ) -> PyResult<()> {
        self.inner
            .set_row_height(row, height)
            .map_err(Error::RustXlsxError)?;
        if let Some(fmt) = format {
            self.inner
                .set_row_format(row, &fmt.inner)
                .map_err(Error::RustXlsxError)?;
        }
        Ok(())
    }

    pub fn set_landscape(&mut self) {
        self.inner.set_landscape();
    }

    pub fn set_portrait(&mut self) {
        self.inner.set_portrait();
    }

    pub fn set_fit_to_pages(&mut self, width: u16, height: u16) {
        self.inner.set_print_fit_to_pages(width, height);
    }

    pub fn autofit(&mut self) {
        self.inner.autofit();
    }

    pub fn freeze_panes(&mut self, row: u32, col: u16) -> PyResult<()> {
        self.inner
            .set_freeze_panes(row, col)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn set_vba_name(&mut self, name: &str) -> PyResult<()> {
        self.inner
            .set_vba_name(name)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn insert_chart(&mut self, row: u32, col: u16, chart: PyRef<'_, Chart>) -> PyResult<()> {
        self.inner
            .insert_chart(row, col, &chart.inner)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn insert_chart_with_offset(
        &mut self,
        row: u32,
        col: u16,
        chart: PyRef<'_, Chart>,
        x_offset: u32,
        y_offset: u32,
    ) -> PyResult<()> {
        self.inner
            .insert_chart_with_offset(row, col, &chart.inner, x_offset, y_offset)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn add_sparkline(
        &mut self,
        row: u32,
        col: u16,
        sparkline: PyRef<'_, Sparkline>,
    ) -> PyResult<()> {
        self.inner
            .add_sparkline(row, col, &sparkline.inner)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn add_conditional_format_2color_scale(
        &mut self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cond_fmt: PyRef<'_, ConditionalFormat2ColorScale>,
    ) -> PyResult<()> {
        self.inner
            .add_conditional_format(first_row, first_col, last_row, last_col, &cond_fmt.inner)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn add_conditional_format_3color_scale(
        &mut self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cond_fmt: PyRef<'_, ConditionalFormat3ColorScale>,
    ) -> PyResult<()> {
        self.inner
            .add_conditional_format(first_row, first_col, last_row, last_col, &cond_fmt.inner)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn add_conditional_format_data_bar(
        &mut self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cond_fmt: PyRef<'_, ConditionalFormatDataBar>,
    ) -> PyResult<()> {
        self.inner
            .add_conditional_format(first_row, first_col, last_row, last_col, &cond_fmt.inner)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }
}
