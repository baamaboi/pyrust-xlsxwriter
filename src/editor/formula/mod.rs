//! Excel formula engine for the editor: tokenizing formulas and rewriting the
//! references inside them.
//!
//! Pure Rust, no Python involved. Formulas are handled in the form they are
//! stored in xlsx files (A1 notation, `,` argument separators, `_xlfn.`
//! prefixes); a leading `=` is optional and kept if present, and the
//! array-formula wrapper `{=...}` is recognised too.
//!
//! # API
//!
//! * [`tokenize`] splits a formula into [`Token`]s and [`render`] joins them
//!   back. Every character of the input belongs to exactly one token, so
//!   `render(&tokenize(f)?) == f` for every formula `f`, including its
//!   whitespace, case and quoting.
//! * [`references`] lists the cell/range references in a formula as
//!   [`Reference`] values (sheet prefix + [`Area`]).
//! * [`shift_formula`] applies a [`StructuralEdit`] (rows/columns inserted or
//!   deleted on one sheet) to a formula; [`shift_range_list`] does the same
//!   for space-separated range lists (`sqref`, merge ranges); [`Area::shift`]
//!   does it for one parsed area (e.g. a `<dimension>` or table `ref`).
//! * [`translate_formula`] moves a formula as copy/paste does (used to expand
//!   shared formulas to each cell).
//! * [`rename_sheet`] updates sheet qualifiers after a sheet is renamed.
//! * [`quote_sheet_name`] / [`sheet_name_needs_quotes`] apply Excel's rules
//!   for when a sheet name must be quoted in a formula.
//!
//! The rewriting functions return the input unchanged (byte for byte) when
//! nothing in it is affected; only the tokens that change are re-rendered
//! (in canonical upper-case A1 form).
//!
//! # Tokens
//!
//! Cell references, ranges (`A1:B2`), whole columns (`A:C`) and whole rows
//! (`1:3`) are single [`TokenKind::Reference`] tokens that also carry their
//! sheet prefix (`Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`, `[1]Sheet1!`).
//! A `:` only joins two parts into one reference when both are the same kind
//! of A1 reference and directly adjacent; otherwise (`A1:INDEX(...)`,
//! `Name1:Name2`, `Sheet1!A1:Sheet1!B2`) it is an operator token and each side
//! is its own token. Defined names ([`TokenKind::Name`]) keep their sheet
//! prefix too. Structured references (`Table1[[#This Row],[Col]]`, `[@Col]`)
//! are opaque [`TokenKind::StructuredRef`] tokens. Inside array constants
//! (`{1,2;"a",TRUE}`) only literals and separators are recognised; anything
//! word-like there is [`TokenKind::Unknown`], never a reference. Words
//! followed directly by `(` are functions even when they look like cells
//! (`LOG10(`, `ATAN2(`); `1E+3` is a number; `R1C1`-looking words are names.
//! Whitespace is kept as [`TokenKind::Whitespace`], or
//! [`TokenKind::Intersection`] when it sits between two reference operands.
//! Characters that fit nothing are [`TokenKind::Unknown`] tokens (so malformed
//! input still round-trips); only an unterminated string, quoted sheet name or
//! `[` bracket is a [`FormulaError`].
//!
//! # Structural edit semantics (Excel's)
//!
//! * Only references that point at the edited sheet change. Unqualified
//!   references belong to the formula's own sheet; sheet names compare
//!   case-insensitively. Relative and `$absolute` references shift alike.
//! * Inserting at or above/left of a reference moves it; inserting strictly
//!   inside a range grows it; inserting below/right of it changes nothing.
//! * Deleting rows/columns that hold a single-cell reference, or all of a
//!   range, turns it into `#REF!` (a sheet prefix stays: `Sheet2!#REF!`, as
//!   Excel writes it; a spill `#` after it is dropped). Deleting part of a
//!   range shrinks it.
//! * Whole-column references (`A:C`) ignore row edits and whole-row
//!   references (`1:3`) ignore column edits; on their own axis they shift,
//!   grow and shrink like ranges.
//! * A reference whose (first) cell is pushed past the last row/column
//!   becomes `#REF!`. A range whose start stays on the sheet but whose end is
//!   pushed off is clipped at the sheet edge (Excel keeps `A2:A1048576`
//!   anchored to the last row).
//! * External references (`[1]Sheet1!A1`, `'C:\x\[Book.xlsx]Sheet1'!A1`) are
//!   never changed by structural edits or renames.
//! * 3D references (`Sheet1:Sheet3!A1`) are left unchanged by structural
//!   edits: an edit on one sheet does not move the other sheets of the span,
//!   and Excel likewise only adjusts 3D references when the whole group of
//!   sheets is edited together. [`rename_sheet`] does rename either end.
//! * References after a deleted-sheet prefix (`#REF!A1`) are never shifted.
//! * Structured references, defined names, text in strings and array
//!   constants are never changed.
//!
//! [`translate_formula`] follows copy/paste instead: only the relative parts
//! move (on every sheet, including external and 3D references, which is what
//! shared-formula expansion needs), `$absolute` parts stay, and a reference
//! with any part pushed off the sheet becomes `#REF!`.

mod edit;
mod reference;
mod tokenize;

#[cfg(test)]
mod tests;

use std::fmt;

pub use edit::{StructuralEdit, rename_sheet, shift_formula, shift_range_list, translate_formula};
pub use reference::{
    Area, CellRef, ColRef, NameRef, Reference, RowRef, SheetPrefix, quote_sheet_name,
    sheet_name_needs_quotes,
};
pub use tokenize::{Token, TokenKind, render, tokenize};

/// A formula that could not be tokenized (unterminated string, quoted sheet
/// name or bracket).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaError {
    /// What went wrong.
    pub message: String,
    /// Byte offset in the formula where the problem starts.
    pub position: usize,
}

impl FormulaError {
    pub(crate) fn new(message: impl Into<String>, position: usize) -> Self {
        FormulaError { message: message.into(), position }
    }
}

impl fmt::Display for FormulaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{} at position {}", self.message, self.position)
    }
}

impl std::error::Error for FormulaError {}

/// The cell/range references in `formula`, in order of appearance (each
/// with its sheet prefix, if any). Defined names, structured references and
/// text inside strings or array constants are not included.
pub fn references(formula: &str) -> Result<Vec<Reference>, FormulaError> {
    Ok(tokenize(formula)?
        .into_iter()
        .filter_map(|token| match token.kind {
            TokenKind::Reference(reference) => Some(reference),
            _ => None,
        })
        .collect())
}
