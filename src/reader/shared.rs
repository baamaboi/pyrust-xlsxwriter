//! State shared by a `ReadWorkbook` and the objects it hands out.
//!
//! calamine's readers need `&mut self`, so the open reader lives in a
//! `Mutex`. It is only ever locked with the GIL released (inside
//! `py.detach`), so a thread waiting for it never blocks Python, and a panic
//! inside calamine (malformed input) is caught and reported as `ReadError`.
//!
//! Sheet data is parsed on first access and cached as plain Rust values
//! (`SheetData`); the parts that are rarely needed (formulas, styles, merged
//! ranges, hyperlinks) are loaded on their own first access. No Python
//! objects are stored here, so there are no reference cycles. While the
//! workbook is open every parsed sheet stays cached; `close()` drops the
//! cache, so a sheet then lives only as long as a Python object uses it.

use std::collections::HashMap;
use std::io::Read;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::{Arc, Mutex, OnceLock, PoisonError, Weak};

use calamine::{
    Data, Hyperlink, Range, Reader, SheetType, SheetVisible, Sheets, StyleRange,
    WorkbookProperties, XlsxError,
};
use pyo3::prelude::*;
use quick_xml::XmlVersion;
use quick_xml::events::Event;

use crate::common::cellref::parse_range;
use crate::reader::source::{self, Failure, FileKind, Origin, Source};

/// Name, type and visibility of one sheet.
pub struct SheetMeta {
    pub name: String,
    pub kind: &'static str,
    pub visibility: &'static str,
}

/// Which cells have an explicit (non-default) style, plus calamine's styles.
pub struct SheetStyles {
    pub range: StyleRange,
    /// `(row << 16) | col` of every cell with a non-zero style id, sorted.
    styled: Vec<u64>,
}

impl SheetStyles {
    fn key(row: u32, col: u32) -> u64 {
        ((row as u64) << 16) | col as u64
    }

    /// The style of a cell, or None for the default style.
    pub fn get(&self, row: u32, col: u32) -> Option<&calamine::Style> {
        if self.styled.binary_search(&Self::key(row, col)).is_ok() {
            self.range.get((row, col))
        } else {
            None
        }
    }
}

/// One worksheet's parsed contents.
pub struct SheetData {
    pub values: Range<Data>,
    formulas: OnceLock<Range<String>>,
    styles: OnceLock<SheetStyles>,
    merged: OnceLock<Vec<CellRange32>>,
    hyperlinks: OnceLock<Vec<Hyperlink>>,
}

/// A range with u32 columns (calamine's coordinates).
pub type CellRange32 = (u32, u32, u32, u32);

/// A table definition (xlsx only).
pub struct TableMeta {
    pub name: String,
    pub sheet_index: usize,
    /// The whole table, header and totals rows included.
    pub range: CellRange32,
    pub header_rows: u32,
    pub totals_rows: u32,
    pub columns: Vec<String>,
}

/// The open reader plus how to reopen the source; dropped by `close()`.
struct Open {
    sheets: Sheets<Source>,
    origin: Origin,
}

pub struct Shared {
    reader: Mutex<Option<Open>>,
    pub kind: FileKind,
    pub is_1904: bool,
    pub sheets: Vec<SheetMeta>,
    pub defined_names: Vec<(String, String)>,
    /// Each parsed sheet, found again through a weak reference.
    data: Vec<Mutex<Weak<SheetData>>>,
    /// Keeps every parsed sheet alive while the workbook is open.
    open_data: Mutex<Vec<Arc<SheetData>>>,
    properties: OnceLock<WorkbookProperties>,
    tables: OnceLock<Vec<Arc<TableMeta>>>,
}

fn sheet_kind(typ: SheetType) -> &'static str {
    match typ {
        SheetType::WorkSheet => "worksheet",
        SheetType::ChartSheet => "chartsheet",
        SheetType::DialogSheet => "dialogsheet",
        SheetType::MacroSheet => "macrosheet",
        SheetType::Vba => "vba",
    }
}

fn sheet_visibility(visible: SheetVisible) -> &'static str {
    match visible {
        SheetVisible::Visible => "visible",
        SheetVisible::Hidden => "hidden",
        SheetVisible::VeryHidden => "very_hidden",
    }
}

/// Run `f`, turning a panic into a `ReadError`-style failure.
fn guarded<T>(what: &str, f: impl FnOnce() -> Result<T, Failure>) -> Result<T, Failure> {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(result) => result,
        Err(_) => Err(Failure::Read(format!(
            "the spreadsheet reader failed unexpectedly while reading {what} (the file may be damaged)"
        ))),
    }
}

impl Shared {
    /// Open a workbook (call with the GIL released).
    pub fn open(origin: Origin) -> Result<Shared, Failure> {
        guarded("the workbook", || {
            let source = origin.open()?;
            let (kind, sheets) = source::open(source, &origin)?;
            let is_1904 = match &sheets {
                Sheets::Xlsx(r) => r.has_1904_epoch(),
                Sheets::Xls(r) => r.has_1904_epoch(),
                Sheets::Xlsb(r) => r.has_1904_epoch(),
                Sheets::Ods(_) => false,
            };
            let metas: Vec<SheetMeta> = sheets
                .sheets_metadata()
                .iter()
                .map(|s| SheetMeta {
                    name: s.name.clone(),
                    kind: sheet_kind(s.typ),
                    visibility: sheet_visibility(s.visible),
                })
                .collect();
            let defined_names = sheets.defined_names().to_vec();
            let data = metas.iter().map(|_| Mutex::new(Weak::new())).collect();
            Ok(Shared {
                reader: Mutex::new(Some(Open { sheets, origin })),
                kind,
                is_1904,
                sheets: metas,
                defined_names,
                data,
                open_data: Mutex::new(Vec::new()),
                properties: OnceLock::new(),
                tables: OnceLock::new(),
            })
        })
    }

    /// Run `f` on the open reader with the GIL released.
    pub fn with_reader<T, F>(&self, py: Python<'_>, what: &str, f: F) -> PyResult<T>
    where
        T: Send,
        F: FnOnce(&mut Sheets<Source>) -> Result<T, Failure> + Send,
    {
        self.with_open(py, what, |open| f(&mut open.sheets))
    }

    fn with_open<T, F>(&self, py: Python<'_>, what: &str, f: F) -> PyResult<T>
    where
        T: Send,
        F: FnOnce(&mut Open) -> Result<T, Failure> + Send,
    {
        py.detach(|| {
            let mut guard = self
                .reader
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            let open = guard.as_mut().ok_or_else(|| {
                Failure::Read(format!("cannot read {what}: the workbook is closed"))
            })?;
            guarded(what, || f(open))
        })
        .map_err(Failure::into_py)
    }

    /// Release the file (or the in-memory copy) and the sheet cache. Sheets
    /// still held by Python objects stay usable; the others are freed.
    pub fn close(&self, py: Python<'_>) {
        py.detach(|| {
            let reader = self.reader.lock().unwrap_or_else(PoisonError::into_inner).take();
            let sheets = std::mem::take(&mut *self.open_data.lock().unwrap_or_else(PoisonError::into_inner));
            drop((reader, sheets));
        });
    }

    pub fn is_closed(&self) -> bool {
        match self.reader.try_lock() {
            Ok(reader) => reader.is_none(),
            Err(std::sync::TryLockError::Poisoned(p)) => p.into_inner().is_none(),
            // Another thread is reading from it right now.
            Err(std::sync::TryLockError::WouldBlock) => false,
        }
    }

    fn sheet_name(&self, index: usize) -> &str {
        &self.sheets[index].name
    }

    fn read_failure(&self, index: usize, part: &str, error: impl std::fmt::Display) -> Failure {
        Failure::Read(format!(
            "cannot read the {part} of sheet {:?}: {error}",
            self.sheet_name(index)
        ))
    }

    /// The parsed values of sheet `index` (loaded on first use).
    pub fn sheet(&self, py: Python<'_>, index: usize) -> PyResult<Arc<SheetData>> {
        if let Some(data) = self.data[index].lock().unwrap_or_else(PoisonError::into_inner).upgrade() {
            return Ok(data);
        }
        let name = self.sheet_name(index).to_string();
        let values = self.with_reader(py, &format!("sheet {name:?}"), |r| {
            r.worksheet_range(&name)
                .map_err(|e| self.read_failure(index, "cells", e))
        })?;
        let data = Arc::new(SheetData {
            values,
            formulas: OnceLock::new(),
            styles: OnceLock::new(),
            merged: OnceLock::new(),
            hyperlinks: OnceLock::new(),
        });
        let mut slot = self.data[index].lock().unwrap_or_else(PoisonError::into_inner);
        if let Some(existing) = slot.upgrade() {
            // Another thread parsed it meanwhile.
            return Ok(existing);
        }
        *slot = Arc::downgrade(&data);
        // shortcut: a sheet parsed while close() runs stays cached until the workbook is dropped; harmless, fix only if it matters.
        self.open_data.lock().unwrap_or_else(PoisonError::into_inner).push(data.clone());
        Ok(data)
    }

    pub fn formulas<'a>(
        &self,
        py: Python<'_>,
        index: usize,
        data: &'a SheetData,
    ) -> PyResult<&'a Range<String>> {
        if let Some(f) = data.formulas.get() {
            return Ok(f);
        }
        let name = self.sheet_name(index).to_string();
        let formulas = self.with_reader(py, &format!("the formulas of sheet {name:?}"), |r| {
            r.worksheet_formula(&name)
                .map_err(|e| self.read_failure(index, "formulas", e))
        })?;
        Ok(data.formulas.get_or_init(|| formulas))
    }

    pub fn styles<'a>(
        &self,
        py: Python<'_>,
        index: usize,
        data: &'a SheetData,
    ) -> PyResult<&'a SheetStyles> {
        if let Some(s) = data.styles.get() {
            return Ok(s);
        }
        let name = self.sheet_name(index).to_string();
        let styles = self.with_reader(py, &format!("the styles of sheet {name:?}"), |r| {
            let Sheets::Xlsx(x) = r else {
                return Ok(SheetStyles {
                    range: StyleRange::empty(),
                    styled: Vec::new(),
                });
            };
            let range = match x.worksheet_style(&name) {
                Ok(range) => range,
                Err(XlsxError::NotAWorksheet(_)) => StyleRange::empty(),
                Err(e) => return Err(self.read_failure(index, "styles", e)),
            };
            let mut styled = Vec::new();
            if !range.is_empty() {
                let mut cells = x
                    .worksheet_cells_reader(&name)
                    .map_err(|e| self.read_failure(index, "styles", e))?;
                while let Some((row, col, id)) = cells
                    .next_style_id()
                    .map_err(|e| self.read_failure(index, "styles", e))?
                {
                    if id != 0 {
                        styled.push(SheetStyles::key(row, col));
                    }
                }
                styled.sort_unstable();
                styled.dedup();
            }
            Ok(SheetStyles { range, styled })
        })?;
        Ok(data.styles.get_or_init(|| styles))
    }

    pub fn merged<'a>(
        &self,
        py: Python<'_>,
        index: usize,
        data: &'a SheetData,
    ) -> PyResult<&'a [CellRange32]> {
        if let Some(m) = data.merged.get() {
            return Ok(m);
        }
        let name = self.sheet_name(index).to_string();
        let is_worksheet = self.sheets[index].kind == "worksheet";
        let merged =
            self.with_reader(py, &format!("the merged ranges of sheet {name:?}"), |r| {
                let dims = match r {
                    _ if !is_worksheet => Vec::new(),
                    Sheets::Xlsx(x) => x
                        .merge_cells_by_sheet_name(&name)
                        .map_err(|e| self.read_failure(index, "merged ranges", e))?,
                    Sheets::Xls(x) => x
                        .merge_cells_by_sheet_name(&name)
                        .map_err(|e| self.read_failure(index, "merged ranges", e))?,
                    _ => Vec::new(),
                };
                Ok(dims
                    .iter()
                    .map(|d| (d.start.0, d.start.1, d.end.0, d.end.1))
                    .collect::<Vec<_>>())
            })?;
        Ok(data.merged.get_or_init(|| merged))
    }

    pub fn hyperlinks<'a>(
        &self,
        py: Python<'_>,
        index: usize,
        data: &'a SheetData,
    ) -> PyResult<&'a [Hyperlink]> {
        if let Some(h) = data.hyperlinks.get() {
            return Ok(h);
        }
        let name = self.sheet_name(index).to_string();
        let is_worksheet = self.sheets[index].kind == "worksheet";
        let links = self.with_reader(
            py,
            &format!("the hyperlinks of sheet {name:?}"),
            |r| match r {
                Sheets::Xlsx(x) if is_worksheet => x
                    .hyperlinks_by_sheet_name(&name)
                    .map_err(|e| self.read_failure(index, "hyperlinks", e)),
                _ => Ok(Vec::new()),
            },
        )?;
        Ok(data.hyperlinks.get_or_init(|| links))
    }

    pub fn properties(&self, py: Python<'_>) -> PyResult<&WorkbookProperties> {
        if let Some(p) = self.properties.get() {
            return Ok(p);
        }
        let props = self.with_reader(py, "the document properties", |r| {
            r.workbook_properties()
                .map_err(|e| Failure::Read(format!("cannot read the document properties: {e}")))
        })?;
        Ok(self.properties.get_or_init(|| props))
    }

    pub fn pictures(&self, py: Python<'_>) -> PyResult<Vec<(String, Vec<u8>)>> {
        self.with_reader(py, "the pictures", |r| Ok(r.pictures().unwrap_or_default()))
    }

    /// `(module name, source code)` pairs, sorted by name; None without a VBA project.
    pub fn vba_modules(&self, py: Python<'_>) -> PyResult<Option<Vec<(String, String)>>> {
        self.with_reader(py, "the VBA project", |r| {
            let failure = |e: &dyn std::fmt::Display| {
                Failure::Read(format!("cannot read the VBA project: {e}"))
            };
            let Some(project) = r.vba_project().map_err(|e| failure(&e))? else {
                return Ok(None);
            };
            let mut names: Vec<String> = project
                .get_module_names()
                .into_iter()
                .map(String::from)
                .collect();
            names.sort();
            let mut modules = Vec::with_capacity(names.len());
            for name in names {
                let code = project.get_module(&name).map_err(|e| failure(&e))?;
                modules.push((name, code));
            }
            Ok(Some(modules))
        })
    }

    /// The tables of an xlsx workbook, by sheet then by table id.
    pub fn tables(&self, py: Python<'_>) -> PyResult<&[Arc<TableMeta>]> {
        if let Some(t) = self.tables.get() {
            return Ok(t);
        }
        let tables = if self.kind != FileKind::Xlsx {
            Vec::new()
        } else {
            self.with_open(py, "the tables", |open| {
                let Sheets::Xlsx(x) = &mut open.sheets else {
                    return Ok(Vec::new());
                };
                x.load_tables()
                    .map_err(|e| Failure::Read(format!("cannot read the tables: {e}")))?;
                let mut owners: HashMap<String, usize> = HashMap::new();
                for (index, sheet) in self.sheets.iter().enumerate() {
                    for name in x.table_names_in_sheet(&sheet.name) {
                        owners.insert(name.clone(), index);
                    }
                }
                if owners.is_empty() {
                    return Ok(Vec::new());
                }
                let mut tables = read_table_parts(&open.origin)?
                    .into_iter()
                    .filter_map(|(id, mut meta)| {
                        meta.sheet_index = *owners.get(&meta.name)?;
                        Some((meta.sheet_index, id, meta))
                    })
                    .collect::<Vec<_>>();
                tables.sort_by_key(|(sheet, id, _)| (*sheet, *id));
                Ok(tables
                    .into_iter()
                    .map(|(_, _, meta)| Arc::new(meta))
                    .collect())
            })?
        };
        Ok(self.tables.get_or_init(|| tables))
    }
}

fn xml_failure(part: &str, e: impl std::fmt::Display) -> Failure {
    Failure::Read(format!("cannot read {part}: {e}"))
}

/// Parse every `xl/tables/*.xml` part: `(id, TableMeta)` with `sheet_index` unset.
fn read_table_parts(origin: &Origin) -> Result<Vec<(u32, TableMeta)>, Failure> {
    let source = origin.open()?;
    let mut archive = zip::ZipArchive::new(source).map_err(|e| xml_failure("the tables", e))?;
    let names: Vec<String> = archive
        .file_names()
        .flatten()
        .filter(|n| {
            let lower = n.to_ascii_lowercase();
            lower.starts_with("xl/tables/") && lower.ends_with(".xml") && !lower.contains("/_rels/")
        })
        .map(|n| n.into_owned())
        .collect();
    let mut tables = Vec::with_capacity(names.len());
    for name in names {
        let mut text = String::new();
        archive
            .by_name(&name)
            .map_err(|e| xml_failure(&name, e))?
            .read_to_string(&mut text)
            .map_err(|e| xml_failure(&name, e))?;
        if let Some(table) = parse_table_xml(&text).map_err(|e| xml_failure(&name, e))? {
            tables.push(table);
        }
    }
    Ok(tables)
}

/// Parse one table part. Returns None if it has no usable name or range.
fn parse_table_xml(text: &str) -> Result<Option<(u32, TableMeta)>, quick_xml::Error> {
    let mut reader = quick_xml::Reader::from_str(text);
    let mut id = 0u32;
    let mut name = None;
    let mut display_name = None;
    let mut reference = None;
    let mut header_rows = 1u32;
    let mut totals_rows = 0u32;
    let mut columns = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                "table" => {
                    for attr in e.attributes().flatten() {
                        let value = attr.normalized_value(XmlVersion::Implicit1_0)?.into_owned();
                        match attr.key.local_name().as_ref() {
                            "id" => id = value.parse().unwrap_or(0),
                            "name" => name = Some(value),
                            "displayName" => display_name = Some(value),
                            "ref" => reference = Some(value),
                            "headerRowCount" => header_rows = value.parse().unwrap_or(1),
                            "totalsRowCount" => totals_rows = value.parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                }
                "tableColumn" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == "name" {
                            columns
                                .push(attr.normalized_value(XmlVersion::Implicit1_0)?.into_owned());
                        }
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    let (Some(name), Some((r1, c1, r2, c2))) = (
        display_name.or(name),
        reference.as_deref().and_then(parse_range),
    ) else {
        return Ok(None);
    };
    let meta = TableMeta {
        name,
        sheet_index: 0,
        range: (r1, c1 as u32, r2, c2 as u32),
        header_rows,
        totals_rows,
        columns,
    };
    Ok(Some((id, meta)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_xml() {
        let xml = r#"<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="3" name="Table3" displayName="Sales &amp; Co" ref="B2:D6" totalsRowCount="1"><autoFilter ref="B2:D5"/><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Q1"/><tableColumn id="3" name="Q2"/></tableColumns></table>"#;
        let (id, meta) = parse_table_xml(xml).unwrap().unwrap();
        assert_eq!(id, 3);
        assert_eq!(meta.name, "Sales & Co");
        assert_eq!(meta.range, (1, 1, 5, 3));
        assert_eq!((meta.header_rows, meta.totals_rows), (1, 1));
        assert_eq!(meta.columns, ["Region", "Q1", "Q2"]);
        assert!(parse_table_xml("<table id='1'/>").unwrap().is_none());
    }
}
