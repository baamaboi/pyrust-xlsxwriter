use pyo3::prelude::*;

#[pyclass(from_py_object)]
#[derive(Clone)]
pub struct VbaProject {
    path_val: String,
    signature_path_val: Option<String>,
}

#[pymethods]
impl VbaProject {
    #[new]
    #[pyo3(signature = (path="", signature_path=None))]
    pub fn new(path: &str, signature_path: Option<&str>) -> Self {
        VbaProject {
            path_val: path.to_string(),
            signature_path_val: signature_path.map(|s| s.to_string()),
        }
    }

    #[getter]
    pub fn path(&self) -> String {
        self.path_val.clone()
    }

    #[setter]
    pub fn set_path(&mut self, val: Bound<'_, PyAny>) {
        self.path_val = val.to_string();
    }

    #[getter]
    pub fn signature_path(&self) -> Option<String> {
        self.signature_path_val.clone()
    }

    #[setter]
    pub fn set_signature_path(&mut self, val: Bound<'_, PyAny>) {
        if val.is_none() {
            self.signature_path_val = None;
        } else {
            self.signature_path_val = Some(val.to_string());
        }
    }
}
