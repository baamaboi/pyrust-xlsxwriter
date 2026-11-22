//! `ObjectMovement`: how images, charts, shapes, buttons and notes move and
//! resize with the cells under them. Shared by several feature areas.

use pyo3::prelude::*;
use rust_xlsxwriter::ObjectMovement as RObjectMovement;

crate::py_enum!(
    /// How an object moves/resizes when the cells under it are resized.
    ObjectMovement => RObjectMovement {
        MoveAndSizeWithCells,
        MoveButDontSizeWithCells,
        DontMoveOrSizeWithCells,
        MoveAndSizeWithCellsAfter,
    }
);

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ObjectMovement>()
}
