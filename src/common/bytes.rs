//! Read-only bytes shared without copying.
//!
//! A workbook or image handed over as Python `bytes` is read in place: the
//! `bytes` object is kept alive and Rust reads its buffer directly. Bytes Rust
//! reads itself (a file from disk, a rendered workbook) are moved in, not
//! copied. Cloning only bumps a reference count, so the zip reader, calamine
//! and the editor can all hold the same buffer.

use std::ops::Deref;
use std::sync::Arc;

use pyo3::prelude::*;
use pyo3::pybacked::PyBackedBytes;
use pyo3::types::PyBytes;

/// Immutable bytes that are cheap to clone and never copied once held.
#[derive(Clone)]
pub struct SharedBytes(Arc<dyn AsRef<[u8]> + Send + Sync>);

impl SharedBytes {
    /// Take ownership of bytes Rust already holds (no copy).
    pub fn from_vec(data: Vec<u8>) -> Self {
        SharedBytes(Arc::new(data))
    }

    /// Share a Python `bytes` object in place; a `bytearray` is copied once,
    /// because it could change while we read it.
    pub fn from_py(data: &Bound<'_, PyAny>) -> PyResult<Self> {
        Ok(SharedBytes(Arc::new(data.extract::<PyBackedBytes>()?)))
    }

    /// True if both share the same buffer (a cheap equality shortcut).
    pub fn same_buffer(&self, other: &SharedBytes) -> bool {
        std::ptr::addr_eq(Arc::as_ptr(&self.0), Arc::as_ptr(&other.0))
    }
}

impl Deref for SharedBytes {
    type Target = [u8];

    fn deref(&self) -> &[u8] {
        (*self.0).as_ref()
    }
}

impl std::fmt::Debug for SharedBytes {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SharedBytes({} bytes)", self.len())
    }
}

impl AsRef<[u8]> for SharedBytes {
    fn as_ref(&self) -> &[u8] {
        self
    }
}

/// Write `data` to a binary file object in 1 MiB pieces, so the whole file is
/// never duplicated into one Python `bytes` object.
pub fn write_to_file(file: &Bound<'_, PyAny>, data: &[u8]) -> PyResult<()> {
    for chunk in data.chunks(1 << 20) {
        file.call_method1("write", (PyBytes::new(file.py(), chunk),))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn clones_share_the_buffer() {
        let data = vec![1u8, 2, 3];
        let address = data.as_ptr();
        let shared = SharedBytes::from_vec(data);
        let copy = shared.clone();
        assert_eq!(shared.as_ptr(), address, "the Vec was moved in, not copied");
        assert!(shared.same_buffer(&copy));
        assert!(!shared.same_buffer(&SharedBytes::from_vec(vec![1, 2, 3])));
        assert_eq!(&*copy, &[1, 2, 3]);
    }
}
