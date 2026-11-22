//! Declarative helpers that keep the (very large) binding surface mechanical.
//!
//! The binding follows a "spec first" design: every Python object keeps its
//! settings in a plain `spec` struct, which is the single source of truth, and
//! rebuilds the underlying `rust_xlsxwriter` object from it in `on_change()`.
//! That makes getters trivial, makes "turning a setting off" work for free, and
//! gives the reader/editor engines a plain description of the object.

/// Define a Python enum (a simple `#[pyclass]` enum) together with lossless
/// conversions to and from the equivalent Rust enum.
///
/// ```ignore
/// py_enum!(
///     /// Docs.
///     FormatScript => RFormatScript { None, Superscript, Subscript }
/// );
/// // A variant whose Rust name differs: `WinLoss => WinLose`.
/// // A variant named `None` must be renamed for Python (`None` is a keyword):
/// //     #[pyo3(name = "None_")] None
/// ```
#[macro_export]
macro_rules! py_enum {
    (
        $(#[$meta:meta])*
        $name:ident => $rust:ident { $( $(#[$vmeta:meta])* $variant:ident $(=> $rv:ident)? ),* $(,)? }
    ) => {
        $(#[$meta])*
        #[pyo3::pyclass(module = "pyrust_xlsxwriter", eq, eq_int, hash, frozen, from_py_object)]
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        pub enum $name {
            $( $(#[$vmeta])* $variant ),*
        }

        impl From<$name> for $rust {
            fn from(value: $name) -> Self {
                match value {
                    $( $name::$variant => $crate::__py_enum_variant!($rust, $variant $(, $rv)?), )*
                }
            }
        }

        impl $name {
            /// Map a Rust value back to the Python enum (None for unbound variants).
            #[allow(unreachable_patterns, dead_code)]
            pub(crate) fn from_rust(value: &$rust) -> Option<Self> {
                match value {
                    $( $crate::__py_enum_variant!($rust, $variant $(, $rv)?) => Some($name::$variant), )*
                    _ => None,
                }
            }
        }
    };
}

#[doc(hidden)]
#[macro_export]
macro_rules! __py_enum_variant {
    ($rust:ident, $v:ident) => {
        $rust::$v
    };
    ($rust:ident, $v:ident, $rv:ident) => {
        $rust::$rv
    };
}

/// Generate Python properties backed by `self.spec`, plus an `apply_kwargs()`
/// helper so that constructors can accept every property as a keyword
/// argument (`Format(bold=True, font_color="#FF0000")`).
///
/// * `fields { name: Type, ... }` read/write `self.spec.name` directly.
/// * `custom { name: Type, ... }` delegate to hand-written
///   `fn get_name(&self) -> Type` and `fn set_name(&mut self, Type) -> PyResult<()>`.
///
/// After every change the type's `fn on_change(&mut self) -> PyResult<()>` is
/// called, which normally rebuilds the wrapped Rust object.
#[macro_export]
macro_rules! py_props {
    (
        $ty:ident
        $( fields { $( $prop:ident : $t:ty ),* $(,)? } )?
        $( custom { $( $cprop:ident : $ct:ty ),* $(,)? } )?
    ) => {
        paste::paste! {
            #[pyo3::pymethods]
            impl $ty {
                $($(
                    #[getter($prop)]
                    fn [<pyget_ $prop>](&self) -> $t {
                        self.spec.$prop.clone()
                    }

                    #[setter($prop)]
                    fn [<pyset_ $prop>](&mut self, value: $t) -> pyo3::PyResult<()> {
                        self.spec.$prop = value;
                        self.on_change()
                    }
                )*)?
                $($(
                    #[getter($cprop)]
                    fn [<pyget_ $cprop>](&self) -> $ct {
                        self.[<get_ $cprop>]()
                    }

                    #[setter($cprop)]
                    fn [<pyset_ $cprop>](&mut self, value: $ct) -> pyo3::PyResult<()> {
                        self.[<set_ $cprop>](value)?;
                        self.on_change()
                    }
                )*)?
            }

            impl $ty {
                /// Names of every keyword accepted by `apply_kwargs()`.
                #[allow(dead_code)]
                pub(crate) const PROPERTY_NAMES: &'static [&'static str] = &[
                    $($( stringify!($prop), )*)?
                    $($( stringify!($cprop), )*)?
                ];

                /// Apply `**kwargs` as property assignments, then call `on_change()`.
                #[allow(dead_code)]
                pub(crate) fn apply_kwargs(
                    &mut self,
                    kwargs: Option<&pyo3::Bound<'_, pyo3::types::PyDict>>,
                ) -> pyo3::PyResult<()> {
                    use pyo3::types::PyDictMethods;
                    use pyo3::types::PyAnyMethods;
                    if let Some(kwargs) = kwargs {
                        for (key, value) in kwargs.iter() {
                            let key: String = key.extract()?;
                            self.apply_one(&key, &value)?;
                        }
                    }
                    self.on_change()
                }

                /// Set one property by name without calling `on_change()`.
                #[allow(dead_code, unused_variables)]
                pub(crate) fn apply_one(
                    &mut self,
                    key: &str,
                    value: &pyo3::Bound<'_, pyo3::PyAny>,
                ) -> pyo3::PyResult<()> {
                    use pyo3::types::PyAnyMethods;
                    match key {
                        $($( stringify!($prop) => { self.spec.$prop = value.extract::<$t>()?; } )*)?
                        $($( stringify!($cprop) => { self.[<set_ $cprop>](value.extract::<$ct>()?)?; } )*)?
                        _ => {
                            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                                "{}() got an unexpected keyword argument '{}'",
                                stringify!($ty),
                                key
                            )))
                        }
                    }
                    Ok(())
                }
            }
        }
    };
}
