//! Rewriting formulas: structural edits, copy/paste translation and sheet
//! renames.

use super::reference::{Area, Reference, same_sheet};
use super::tokenize::{Token, TokenKind, render, tokenize};
use super::FormulaError;

/// Rows or columns inserted into / deleted from one sheet. Positions are
/// zero-based: `InsertRows { at: 4, count: 2 }` inserts two rows before row
/// 5 (so the old row 5 becomes row 7); `DeleteRows { at: 4, count: 2 }`
/// deletes rows 5 and 6.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum StructuralEdit {
    InsertRows { sheet: String, at: u32, count: u32 },
    DeleteRows { sheet: String, at: u32, count: u32 },
    InsertCols { sheet: String, at: u16, count: u16 },
    DeleteCols { sheet: String, at: u16, count: u16 },
}

impl StructuralEdit {
    /// The sheet being edited.
    pub fn sheet(&self) -> &str {
        match self {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet,
        }
    }

    /// Whether an edit applies to `sheet` (case-insensitive, like Excel).
    pub fn is_on_sheet(&self, sheet: &str) -> bool {
        same_sheet(self.sheet(), sheet)
    }
}

/// Rewrite the references in a formula held on `formula_sheet` for `edit`.
///
/// See the module docs for the exact semantics. The formula is returned
/// unchanged (byte for byte) when no reference is affected.
pub fn shift_formula(
    formula: &str,
    formula_sheet: &str,
    edit: &StructuralEdit,
) -> Result<String, FormulaError> {
    rewrite(formula, |reference| {
        let on_edited_sheet = match &reference.sheet {
            None => edit.is_on_sheet(formula_sheet),
            Some(prefix) => prefix.is_local_sheet() && edit.is_on_sheet(&prefix.sheet),
        };
        if on_edited_sheet { Some(reference.area.shift(edit)) } else { None }
    })
}

/// Move a formula by `row_delta` rows and `col_delta` columns as copy/paste
/// does (used to expand a shared formula from its anchor cell to another
/// cell). Relative parts move, `$absolute` parts stay; references on every
/// sheet (also external and 3D ones) move, since copying is not tied to a
/// sheet. A reference with any part pushed off the sheet becomes `#REF!`.
pub fn translate_formula(
    formula: &str,
    row_delta: i64,
    col_delta: i64,
) -> Result<String, FormulaError> {
    if row_delta == 0 && col_delta == 0 {
        // Still validate the formula, for a consistent error contract.
        tokenize(formula)?;
        return Ok(formula.to_string());
    }
    rewrite(formula, |reference| Some(reference.area.translate(row_delta, col_delta)))
}

/// Update the sheet qualifiers that name sheet `old` (case-insensitively;
/// either end of a 3D reference) to `new`, quoting or unquoting as needed:
/// `Sheet1!A1` -> `'My Data'!A1` -> `Data!A1`. Applies to references and to
/// sheet-scoped names (`Sheet1!MyName`). External references and text in
/// strings are not touched. `old` and `new` are plain (unquoted) names.
pub fn rename_sheet(formula: &str, old: &str, new: &str) -> Result<String, FormulaError> {
    let mut tokens = tokenize(formula)?;
    let mut changed = false;
    for token in &mut tokens {
        let prefix = match &mut token.kind {
            TokenKind::Reference(reference) => &mut reference.sheet,
            TokenKind::Name(name) => &mut name.sheet,
            _ => continue,
        };
        let Some(old_prefix) = prefix.as_ref() else { continue };
        let Some(new_prefix) = old_prefix.renamed(old, new) else { continue };
        token.text = format!("{}{}", new_prefix.text, &token.text[old_prefix.text.len()..]);
        *prefix = Some(new_prefix);
        changed = true;
    }
    Ok(if changed { render(&tokens) } else { formula.to_string() })
}

/// Apply `edit` to a space-separated list of areas such as a conditional
/// format / data validation `sqref` or a merge range (`"A1:B2 D4 F:F"`).
///
/// `edit_applies` says whether the list belongs to the edited sheet (range
/// lists carry no sheet name); when false the list is returned unchanged.
/// Deleted areas are dropped, and `None` is returned when every area was
/// deleted. Areas that change are written in canonical form, and a range
/// that shrinks to one cell is written as that cell (`B2`, as Excel writes
/// `sqref`). Items that are not A1 areas are kept as they are. When nothing
/// changes the input is returned unchanged.
pub fn shift_range_list(sqref: &str, edit_applies: bool, edit: &StructuralEdit) -> Option<String> {
    if !edit_applies {
        return Some(sqref.to_string());
    }
    let mut items = Vec::new();
    let mut changed = false;
    let mut seen_any = false;
    for item in sqref.split_whitespace() {
        seen_any = true;
        let Some(area) = Area::parse(item).filter(|area| !area.is_error()) else {
            items.push(item.to_string());
            continue;
        };
        match area.shift(edit) {
            Some(new) if new == area => items.push(item.to_string()),
            Some(new) => {
                changed = true;
                let new = match new {
                    Area::Range(a, b) if a == b => Area::Cell(a),
                    other => other,
                };
                items.push(new.to_string());
            }
            None => changed = true,
        }
    }
    if !changed {
        return Some(sqref.to_string());
    }
    if items.is_empty() && seen_any {
        return None;
    }
    Some(items.join(" "))
}

/// Tokenize, let `map` decide each reference's new area (`None` = leave the
/// token alone, `Some(None)` = becomes `#REF!`), and render the result.
fn rewrite(
    formula: &str,
    map: impl Fn(&Reference) -> Option<Option<Area>>,
) -> Result<String, FormulaError> {
    let tokens = tokenize(formula)?;
    let mut out: Vec<Token> = Vec::with_capacity(tokens.len());
    let mut changed = false;
    let mut drop_spill = false;
    for mut token in tokens {
        if std::mem::take(&mut drop_spill) && token.kind == TokenKind::Spill {
            // `A5#` whose anchor was deleted: Excel shows plain `#REF!`.
            continue;
        }
        if let TokenKind::Reference(reference) = &mut token.kind
            && !reference.area.is_error()
            && let Some(new_area) = map(reference)
        {
            let new_area = new_area.unwrap_or(Area::RefError);
            if new_area != reference.area {
                drop_spill = new_area.is_error();
                reference.area = new_area;
                token.text = reference.to_string();
                changed = true;
            }
        }
        out.push(token);
    }
    Ok(if changed { render(&out) } else { formula.to_string() })
}
