use pyo3::prelude::*;

use crate::errors::Error;
use crate::worksheet::Worksheet;
use rust_xlsxwriter::Workbook as RWorkbook;
use rust_xlsxwriter::Worksheet as RWorksheet;

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Color {
    Default,
    Automatic,
    Black,
    Blue,
    Brown,
    Cyan,
    Gray,
    Green,
    Lime,
    Magenta,
    Navy,
    Orange,
    Pink,
    Purple,
    Red,
    Silver,
    White,
    Yellow,
}

impl From<Color> for rust_xlsxwriter::Color {
    fn from(c: Color) -> Self {
        match c {
            Color::Default => rust_xlsxwriter::Color::Default,
            Color::Automatic => rust_xlsxwriter::Color::Automatic,
            Color::Black => rust_xlsxwriter::Color::Black,
            Color::Blue => rust_xlsxwriter::Color::Blue,
            Color::Brown => rust_xlsxwriter::Color::Brown,
            Color::Cyan => rust_xlsxwriter::Color::Cyan,
            Color::Gray => rust_xlsxwriter::Color::Gray,
            Color::Green => rust_xlsxwriter::Color::Green,
            Color::Lime => rust_xlsxwriter::Color::Lime,
            Color::Magenta => rust_xlsxwriter::Color::Magenta,
            Color::Navy => rust_xlsxwriter::Color::Navy,
            Color::Orange => rust_xlsxwriter::Color::Orange,
            Color::Pink => rust_xlsxwriter::Color::Pink,
            Color::Purple => rust_xlsxwriter::Color::Purple,
            Color::Red => rust_xlsxwriter::Color::Red,
            Color::Silver => rust_xlsxwriter::Color::Silver,
            Color::White => rust_xlsxwriter::Color::White,
            Color::Yellow => rust_xlsxwriter::Color::Yellow,
        }
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PaperSize {
    PrinterDefault = 0,
    Letter = 1,
    LetterSmall = 2,
    Tabloid = 3,
    Ledger = 4,
    Legal = 5,
    Statement = 6,
    Executive = 7,
    A3 = 8,
    A4 = 9,
    A4Small = 10,
    A5 = 11,
    B4 = 12,
    B5 = 13,
    Folio = 14,
    Quarto = 15,
    Note = 18,
    Envelope9 = 19,
    Envelope10 = 20,
    Envelope11 = 21,
    Envelope12 = 22,
    Envelope14 = 23,
    EnvelopeDL = 27,
    EnvelopeC3 = 28,
    EnvelopeC4 = 29,
    EnvelopeC5 = 30,
    EnvelopeC6 = 31,
}

impl PaperSize {
    pub fn from_u8(val: u8) -> Option<Self> {
        match val {
            0 => Some(PaperSize::PrinterDefault),
            1 => Some(PaperSize::Letter),
            2 => Some(PaperSize::LetterSmall),
            3 => Some(PaperSize::Tabloid),
            4 => Some(PaperSize::Ledger),
            5 => Some(PaperSize::Legal),
            6 => Some(PaperSize::Statement),
            7 => Some(PaperSize::Executive),
            8 => Some(PaperSize::A3),
            9 => Some(PaperSize::A4),
            10 => Some(PaperSize::A4Small),
            11 => Some(PaperSize::A5),
            12 => Some(PaperSize::B4),
            13 => Some(PaperSize::B5),
            14 => Some(PaperSize::Folio),
            15 => Some(PaperSize::Quarto),
            18 => Some(PaperSize::Note),
            19 => Some(PaperSize::Envelope9),
            20 => Some(PaperSize::Envelope10),
            21 => Some(PaperSize::Envelope11),
            22 => Some(PaperSize::Envelope12),
            23 => Some(PaperSize::Envelope14),
            27 => Some(PaperSize::EnvelopeDL),
            28 => Some(PaperSize::EnvelopeC3),
            29 => Some(PaperSize::EnvelopeC4),
            30 => Some(PaperSize::EnvelopeC5),
            31 => Some(PaperSize::EnvelopeC6),
            _ => None,
        }
    }

    pub fn to_u8(self) -> u8 {
        self as u8
    }
}

#[pyclass]
pub struct Workbook {
    #[allow(dead_code)]
    pub(crate) inner: RWorkbook,
    worksheets: Vec<Py<Worksheet>>,
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Self {
        Workbook {
            inner: RWorkbook::new(),
            worksheets: Vec::new(),
        }
    }

    fn save(&mut self, py: Python, path: &str) -> PyResult<()> {
        let mut save_wb = RWorkbook::new();

        for py_sheet in &self.worksheets {
            let mut sheet_ref = py_sheet.bind(py).borrow_mut();
            let real_sheet = std::mem::replace(&mut sheet_ref.inner, RWorksheet::new());
            save_wb.push_worksheet(real_sheet);
        }

        let save_result = save_wb.save(path);

        let recovered_sheets = save_wb.worksheets_mut();

        for (py_sheet, rec_sheet) in self.worksheets.iter().zip(recovered_sheets.into_iter()) {
            let mut sheet_ref = py_sheet.bind(py).borrow_mut();
            let real_sheet = std::mem::replace(rec_sheet, RWorksheet::new());
            sheet_ref.inner = real_sheet;
        }

        save_result.map_err(Error::RustXlsxError)?;
        Ok(())
    }

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
    fn add_worksheet<'py>(
        &mut self,
        py: Python<'py>,
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
    ) -> PyResult<Bound<'py, Worksheet>> {
        let new_sheet = Worksheet::new(
            name,
            active,
            selected,
            hidden,
            tab_color,
            default_row_height,
            paper_size,
            header,
            footer,
            print_center_horizontally,
            print_center_vertically,
            screen_gridlines,
            print_gridlines,
            print_headings,
            zoom,
            right_to_left,
        )?;
        let py_ref = Bound::new(py, new_sheet)?;
        self.worksheets.push(py_ref.clone().unbind());
        Ok(py_ref)
    }

    fn push_worksheet(&mut self, _py: Python, worksheet: Py<Worksheet>) {
        self.worksheets.push(worksheet);
    }

    pub fn add_vba_project(&mut self, path: &str) -> PyResult<()> {
        self.inner
            .add_vba_project(path)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn add_vba_project_with_signature(
        &mut self,
        project_path: &str,
        signature_path: &str,
    ) -> PyResult<()> {
        self.inner
            .add_vba_project_with_signature(project_path, signature_path)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }

    pub fn set_vba_name(&mut self, name: &str) -> PyResult<()> {
        self.inner
            .set_vba_name(name)
            .map_err(Error::RustXlsxError)?;
        Ok(())
    }
}
