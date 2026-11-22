//! Parsed references: sheet prefixes, A1 areas, and the arithmetic used to
//! shift and translate them.

use std::fmt;

use super::edit::StructuralEdit;
use crate::common::cellref::{MAX_COLS, MAX_ROWS, col_to_name, name_to_col};

/// One cell corner of a reference. Zero-based; `*_abs` are the `$` markers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRef {
    pub row: u32,
    pub col: u16,
    pub row_abs: bool,
    pub col_abs: bool,
}

/// One end of a whole-column reference (`$A` in `$A:C`). Zero-based.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ColRef {
    pub col: u16,
    pub abs: bool,
}

/// One end of a whole-row reference (`$1` in `$1:3`). Zero-based.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RowRef {
    pub row: u32,
    pub abs: bool,
}

/// The A1 part of a reference, after any sheet prefix.
///
/// The two corners/ends are kept in the order written (Excel always writes
/// them top-left first, but reversed input is handled).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Area {
    /// `A1`, `$B$2`.
    Cell(CellRef),
    /// `A1:B2`.
    Range(CellRef, CellRef),
    /// `A:C`, `$A:$A`.
    Cols(ColRef, ColRef),
    /// `1:3`, `$2:$2`.
    Rows(RowRef, RowRef),
    /// `#REF!`: an invalid reference (e.g. `Sheet2!#REF!`).
    RefError,
}

/// The sheet qualifier in front of a reference or name, e.g. `Sheet1!`,
/// `'My Sheet'!`, `Sheet1:Sheet3!`, `[1]Sheet1!`, `[1]!`, or `#REF!` (a
/// reference to a deleted sheet, as in `#REF!A1`).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct SheetPrefix {
    /// The prefix exactly as written, including quotes and the final `!`.
    pub text: String,
    /// External workbook part as written, e.g. `[1]` or `C:\dir\[Book.xlsx]`.
    pub workbook: Option<String>,
    /// The (first) sheet name, unquoted and unescaped. Empty for a
    /// workbook-level external name (`[1]!Name`); `#REF!` for a deleted sheet.
    pub sheet: String,
    /// The last sheet of a 3D reference (`Sheet3` in `Sheet1:Sheet3!A1`).
    pub last_sheet: Option<String>,
}

/// A cell/range reference: optional sheet prefix plus the area.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Reference {
    pub sheet: Option<SheetPrefix>,
    pub area: Area,
}

/// A defined name (or LAMBDA parameter / unknown identifier) with its
/// optional sheet prefix (`Sheet1!LocalName`, `[1]!ExternalName`).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NameRef {
    pub sheet: Option<SheetPrefix>,
    pub name: String,
}

// ---------------------------------------------------------------------------
// Sheet prefixes and quoting.
// ---------------------------------------------------------------------------

impl SheetPrefix {
    /// A prefix for a sheet of this workbook, quoted only if needed.
    pub fn local(sheet: &str) -> SheetPrefix {
        SheetPrefix {
            text: render_prefix(sheet, None),
            workbook: None,
            sheet: sheet.to_string(),
            last_sheet: None,
        }
    }

    /// Refers to another workbook (`[1]Sheet1!`).
    pub fn is_external(&self) -> bool {
        self.workbook.is_some()
    }

    /// A 3D span of sheets (`Sheet1:Sheet3!`).
    pub fn is_3d(&self) -> bool {
        self.last_sheet.is_some()
    }

    /// The `#REF!` prefix Excel leaves after the sheet was deleted.
    pub fn is_deleted_sheet(&self) -> bool {
        self.text.starts_with('#')
    }

    /// A plain single sheet of this workbook: the only kind that structural
    /// edits apply to.
    pub fn is_local_sheet(&self) -> bool {
        !self.is_external() && !self.is_3d() && !self.is_deleted_sheet()
    }

    /// Parse the text inside quotes (already unescaped): `Sheet 1`,
    /// `Sheet 1:Sheet 3`, `[1]Sheet 1`, `C:\dir\[Book.xlsx]Sheet1`.
    pub(crate) fn from_parts(text: String, content: &str) -> SheetPrefix {
        // Sheet names cannot contain brackets, so `[...]` marks a workbook.
        let (workbook, rest) = match (content.find('['), content.rfind(']')) {
            (Some(open), Some(end)) if open < end => {
                (Some(content[..=end].to_string()), &content[end + 1..])
            }
            // A workbook path with no sheet ('C:\dir\Book.xlsx'!Name); sheet
            // names cannot contain '\' or '/'.
            _ if content.contains(['\\', '/']) => (Some(content.to_string()), ""),
            _ => (None, content),
        };
        let (sheet, last_sheet) = match rest.split_once(':') {
            Some((first, last)) => (first.to_string(), Some(last.to_string())),
            None => (rest.to_string(), None),
        };
        SheetPrefix { text, workbook, sheet, last_sheet }
    }

    /// The same prefix with sheet `old` renamed to `new` (either end of a 3D
    /// span), or `None` if it does not name `old`. External and deleted-sheet
    /// prefixes are never renamed.
    pub(crate) fn renamed(&self, old: &str, new: &str) -> Option<SheetPrefix> {
        if self.is_external() || self.is_deleted_sheet() {
            return None;
        }
        let first_hit = same_sheet(&self.sheet, old);
        let last_hit = self.last_sheet.as_deref().is_some_and(|last| same_sheet(last, old));
        if !first_hit && !last_hit {
            return None;
        }
        let sheet = if first_hit { new.to_string() } else { self.sheet.clone() };
        let last_sheet = match &self.last_sheet {
            Some(_) if last_hit => Some(new.to_string()),
            other => other.clone(),
        };
        Some(SheetPrefix {
            text: render_prefix(&sheet, last_sheet.as_deref()),
            workbook: None,
            sheet,
            last_sheet,
        })
    }
}

/// `Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`, `'Sheet 1:Sheet3'!`.
fn render_prefix(sheet: &str, last: Option<&str>) -> String {
    let quote = sheet_name_needs_quotes(sheet) || last.is_some_and(sheet_name_needs_quotes);
    let body = match last {
        Some(last) => format!("{sheet}:{last}"),
        None => sheet.to_string(),
    };
    if quote { format!("'{}'!", body.replace('\'', "''")) } else { format!("{body}!") }
}

/// Excel sheet names compare case-insensitively.
pub(crate) fn same_sheet(a: &str, b: &str) -> bool {
    a.chars().flat_map(char::to_lowercase).eq(b.chars().flat_map(char::to_lowercase))
}

/// Whether `name` must be written as `'name'` in a formula: anything other
/// than letters, digits, `_` and `.`; a leading digit or `.`; names that read
/// as an A1 cell (`A1`, `XFD5`), start with an R1C1 reference (`R1C1`, `R2x`,
/// `C3`, `RC`), or are `TRUE`/`FALSE`. Quoting more than strictly needed is
/// harmless, so non-alphanumeric symbols (emoji etc.) are always quoted.
pub fn sheet_name_needs_quotes(name: &str) -> bool {
    if name.is_empty() {
        return true;
    }
    if !name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.') {
        return true;
    }
    if name.starts_with(|c: char| c.is_ascii_digit() || c == '.') {
        return true;
    }
    let upper = name.to_ascii_uppercase();
    if upper == "TRUE" || upper == "FALSE" || upper == "R" || upper == "C" || upper == "RC" {
        return true;
    }
    // Split into the letters before the first digit and the rest.
    let Some(split) = upper.find(|c: char| c.is_ascii_digit()) else {
        return false;
    };
    let (letters, rest) = upper.split_at(split);
    // A1 style: 1-3 ASCII letters then only digits, within the sheet limits.
    if (1..=3).contains(&letters.len())
        && letters.chars().all(|c| c.is_ascii_alphabetic())
        && rest.chars().all(|c| c.is_ascii_digit())
        && name_to_col(letters).is_some()
        && parse_row_number(rest).is_some()
    {
        return true;
    }
    // R1C1 style prefix: R<row>..., C<col>..., RC<col>...
    let digits: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
    let number = digits.parse::<u64>().unwrap_or(0);
    match letters {
        "R" => (1..=MAX_ROWS as u64).contains(&number),
        "C" | "RC" => (1..=MAX_COLS as u64).contains(&number),
        _ => false,
    }
}

/// `name` as it must appear before `!` in a formula: quoted (with `'`
/// doubled) when [`sheet_name_needs_quotes`], else unchanged.
pub fn quote_sheet_name(name: &str) -> String {
    if sheet_name_needs_quotes(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

// ---------------------------------------------------------------------------
// Parsing A1 parts.
// ---------------------------------------------------------------------------

/// What one `:`-free word of A1 text is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Part {
    Cell(CellRef),
    Col(ColRef),
    Row(RowRef),
}

/// "1".."1048576" (no sign, at most 7 digits) to a zero-based row.
fn parse_row_number(digits: &str) -> Option<u32> {
    if digits.is_empty() || digits.len() > 7 || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if (1..=MAX_ROWS).contains(&row) { Some(row - 1) } else { None }
}

/// Classify one word as a cell (`$A$1`), column (`$A`) or row (`$1`).
pub(crate) fn classify_part(word: &str) -> Option<Part> {
    let bytes = word.as_bytes();
    let mut i = 0;
    let first_abs = bytes.first() == Some(&b'$');
    if first_abs {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let letters = &word[letters_start..i];
    if letters.is_empty() {
        // Whole row: "$1", "12".
        return parse_row_number(&word[i..]).map(|row| Part::Row(RowRef { row, abs: first_abs }));
    }
    if letters.len() > 3 {
        return None;
    }
    let col = name_to_col(letters)?;
    if i == bytes.len() {
        return Some(Part::Col(ColRef { col, abs: first_abs }));
    }
    let row_abs = bytes[i] == b'$';
    if row_abs {
        i += 1;
    }
    let row = parse_row_number(&word[i..])?;
    Some(Part::Cell(CellRef { row, col, row_abs, col_abs: first_abs }))
}

/// Combine two parts joined by `:` into an area (same kinds only).
pub(crate) fn join_parts(first: Part, last: Part) -> Option<Area> {
    match (first, last) {
        (Part::Cell(a), Part::Cell(b)) => Some(Area::Range(a, b)),
        (Part::Col(a), Part::Col(b)) => Some(Area::Cols(a, b)),
        (Part::Row(a), Part::Row(b)) => Some(Area::Rows(a, b)),
        _ => None,
    }
}

impl Area {
    /// Parse A1 text without a sheet prefix: `A1`, `$A$1:B2`, `A:C`, `$1:3`
    /// or `#REF!`. Case-insensitive.
    pub fn parse(text: &str) -> Option<Area> {
        if text.eq_ignore_ascii_case("#REF!") {
            return Some(Area::RefError);
        }
        match text.split_once(':') {
            Some((a, b)) => join_parts(classify_part(a)?, classify_part(b)?),
            None => match classify_part(text)? {
                Part::Cell(cell) => Some(Area::Cell(cell)),
                _ => None,
            },
        }
    }

    pub fn is_error(&self) -> bool {
        matches!(self, Area::RefError)
    }

    /// Normalised zero-based bounds `(first_row, first_col, last_row,
    /// last_col)`; whole columns span all rows and vice versa. `None` for
    /// `#REF!`.
    pub fn bounds(&self) -> Option<(u32, u16, u32, u16)> {
        match *self {
            Area::Cell(c) => Some((c.row, c.col, c.row, c.col)),
            Area::Range(a, b) => {
                Some((a.row.min(b.row), a.col.min(b.col), a.row.max(b.row), a.col.max(b.col)))
            }
            Area::Cols(a, b) => Some((0, a.col.min(b.col), MAX_ROWS - 1, a.col.max(b.col))),
            Area::Rows(a, b) => Some((a.row.min(b.row), 0, a.row.max(b.row), MAX_COLS - 1)),
            Area::RefError => None,
        }
    }

    /// Apply a structural edit to this area, assuming it is on the edited
    /// sheet (the sheet in `edit` is not looked at). `None` means the area
    /// was deleted or pushed off the sheet (it becomes `#REF!`).
    pub fn shift(&self, edit: &StructuralEdit) -> Option<Area> {
        let (axis, rows) = AxisEdit::from(edit);
        if axis.count == 0 {
            return Some(*self);
        }
        let max_rows = MAX_ROWS as u64;
        let max_cols = MAX_COLS as u64;
        match *self {
            Area::RefError => Some(Area::RefError),
            Area::Cell(c) => {
                let mut c = c;
                if rows {
                    c.row = axis.apply(c.row as u64, c.row as u64, max_rows, false)?.0 as u32;
                } else {
                    c.col = axis.apply(c.col as u64, c.col as u64, max_cols, false)?.0 as u16;
                }
                Some(Area::Cell(c))
            }
            Area::Range(a, b) => {
                let (mut a, mut b) = (a, b);
                if rows {
                    let (ra, rb) = axis.apply_pair(a.row as u64, b.row as u64, max_rows)?;
                    a.row = ra as u32;
                    b.row = rb as u32;
                } else {
                    let (ca, cb) = axis.apply_pair(a.col as u64, b.col as u64, max_cols)?;
                    a.col = ca as u16;
                    b.col = cb as u16;
                }
                Some(Area::Range(a, b))
            }
            Area::Cols(a, b) => {
                if rows {
                    return Some(*self);
                }
                let (ca, cb) = axis.apply_pair(a.col as u64, b.col as u64, max_cols)?;
                Some(Area::Cols(ColRef { col: ca as u16, ..a }, ColRef { col: cb as u16, ..b }))
            }
            Area::Rows(a, b) => {
                if !rows {
                    return Some(*self);
                }
                let (ra, rb) = axis.apply_pair(a.row as u64, b.row as u64, max_rows)?;
                Some(Area::Rows(RowRef { row: ra as u32, ..a }, RowRef { row: rb as u32, ..b }))
            }
        }
    }

    /// Move the relative (non-`$`) parts by the given deltas, as copy/paste
    /// does. `None` if any part ends up off the sheet.
    pub fn translate(&self, row_delta: i64, col_delta: i64) -> Option<Area> {
        fn mv(value: u64, abs: bool, delta: i64, max: u64) -> Option<u64> {
            if abs {
                return Some(value);
            }
            let moved = value as i64 + delta;
            if moved < 0 || moved as u64 >= max { None } else { Some(moved as u64) }
        }
        let cell = |c: CellRef| -> Option<CellRef> {
            Some(CellRef {
                row: mv(c.row as u64, c.row_abs, row_delta, MAX_ROWS as u64)? as u32,
                col: mv(c.col as u64, c.col_abs, col_delta, MAX_COLS as u64)? as u16,
                ..c
            })
        };
        let col = |c: ColRef| -> Option<ColRef> {
            Some(ColRef { col: mv(c.col as u64, c.abs, col_delta, MAX_COLS as u64)? as u16, ..c })
        };
        let row = |r: RowRef| -> Option<RowRef> {
            Some(RowRef { row: mv(r.row as u64, r.abs, row_delta, MAX_ROWS as u64)? as u32, ..r })
        };
        match *self {
            Area::Cell(c) => Some(Area::Cell(cell(c)?)),
            Area::Range(a, b) => Some(Area::Range(cell(a)?, cell(b)?)),
            Area::Cols(a, b) => Some(Area::Cols(col(a)?, col(b)?)),
            Area::Rows(a, b) => Some(Area::Rows(row(a)?, row(b)?)),
            Area::RefError => Some(Area::RefError),
        }
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{}{}{}{}",
            if self.col_abs { "$" } else { "" },
            col_to_name(self.col),
            if self.row_abs { "$" } else { "" },
            self.row + 1
        )
    }
}

impl fmt::Display for ColRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", if self.abs { "$" } else { "" }, col_to_name(self.col))
    }
}

impl fmt::Display for RowRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", if self.abs { "$" } else { "" }, self.row + 1)
    }
}

/// Canonical upper-case A1 text (`$A$1:B2`, `A:C`, `1:3`, `#REF!`).
impl fmt::Display for Area {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Area::Cell(c) => write!(f, "{c}"),
            Area::Range(a, b) => write!(f, "{a}:{b}"),
            Area::Cols(a, b) => write!(f, "{a}:{b}"),
            Area::Rows(a, b) => write!(f, "{a}:{b}"),
            Area::RefError => f.write_str("#REF!"),
        }
    }
}

/// The prefix as written followed by the canonical area.
impl fmt::Display for Reference {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(prefix) = &self.sheet {
            f.write_str(&prefix.text)?;
        }
        write!(f, "{}", self.area)
    }
}

// ---------------------------------------------------------------------------
// One-axis edit arithmetic.
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy)]
struct AxisEdit {
    insert: bool,
    at: u64,
    count: u64,
}

impl AxisEdit {
    /// The edit on its axis, plus whether that axis is rows.
    fn from(edit: &StructuralEdit) -> (AxisEdit, bool) {
        match *edit {
            StructuralEdit::InsertRows { at, count, .. } => {
                (AxisEdit { insert: true, at: at as u64, count: count as u64 }, true)
            }
            StructuralEdit::DeleteRows { at, count, .. } => {
                (AxisEdit { insert: false, at: at as u64, count: count as u64 }, true)
            }
            StructuralEdit::InsertCols { at, count, .. } => {
                (AxisEdit { insert: true, at: at as u64, count: count as u64 }, false)
            }
            StructuralEdit::DeleteCols { at, count, .. } => {
                (AxisEdit { insert: false, at: at as u64, count: count as u64 }, false)
            }
        }
    }

    /// Apply to the interval `lo..=hi` on an axis of `max` positions. A range
    /// whose end is pushed off is clipped; a single position (or a range
    /// whose start is pushed off) is lost.
    fn apply(&self, lo: u64, hi: u64, max: u64, is_range: bool) -> Option<(u64, u64)> {
        if self.insert {
            let (new_lo, new_hi) = if lo >= self.at {
                (lo + self.count, hi + self.count)
            } else if hi >= self.at {
                (lo, hi + self.count)
            } else {
                return Some((lo, hi));
            };
            if new_lo >= max {
                return None;
            }
            if new_hi >= max {
                return if is_range { Some((new_lo, max - 1)) } else { None };
            }
            Some((new_lo, new_hi))
        } else {
            let end = self.at + self.count - 1; // last deleted position
            if hi < self.at {
                Some((lo, hi))
            } else if lo > end {
                Some((lo - self.count, hi - self.count))
            } else if lo >= self.at && hi <= end {
                None
            } else {
                let new_lo = if lo < self.at { lo } else { self.at };
                let new_hi = if hi > end { hi - self.count } else { self.at - 1 };
                Some((new_lo, new_hi))
            }
        }
    }

    /// Apply to the two ends of a range, which may be written reversed.
    fn apply_pair(&self, a: u64, b: u64, max: u64) -> Option<(u64, u64)> {
        if a <= b {
            self.apply(a, b, max, true)
        } else {
            let (lo, hi) = self.apply(b, a, max, true)?;
            Some((hi, lo))
        }
    }
}
