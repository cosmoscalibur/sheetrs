//! Conversions between Rust and JavaScript values, shared by the tools.

#[cfg(any(feature = "query", feature = "edit"))]
use serde::Deserialize;
use serde::Serialize;
#[cfg(any(feature = "query", feature = "edit"))]
use serde::de::DeserializeOwned;
use wasm_bindgen::prelude::*;

/// A Rust value as a JavaScript one: maps become plain objects (not `Map`), and a missing value is
/// `null`.
pub(crate) fn to_js<T: Serialize + ?Sized>(value: &T) -> Result<JsValue, JsValue> {
    value
        .serialize(
            &serde_wasm_bindgen::Serializer::new()
                .serialize_maps_as_objects(true)
                .serialize_missing_as_null(true),
        )
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

#[cfg(any(feature = "query", feature = "edit"))]
/// A JavaScript value as a Rust one, for an object of options: a key that is not in `known` is an error
/// that names it (a typo is not silently ignored). `what` names the value in the message.
pub(crate) fn from_js_object<T: DeserializeOwned>(
    value: JsValue,
    what: &str,
    known: &[&str],
) -> Result<T, JsValue> {
    if value.is_object() {
        for key in js_sys::Object::keys(&js_sys::Object::from(value.clone())) {
            let key = key.as_string().unwrap_or_default();
            if !known.contains(&key.as_str()) {
                return Err(JsValue::from_str(&format!(
                    "{what} is not valid: unknown key `{key}` (the keys are {})",
                    known.join(", ")
                )));
            }
        }
    }
    from_js(value, what)
}

#[cfg(any(feature = "query", feature = "edit"))]
/// A JavaScript value as a Rust one; `what` names it in the message when it is not valid.
pub(crate) fn from_js<T: DeserializeOwned>(value: JsValue, what: &str) -> Result<T, JsValue> {
    serde_wasm_bindgen::from_value(value)
        .map_err(|e| JsValue::from_str(&format!("{what} is not valid: {e}")))
}

#[cfg(feature = "bundled-runtime")]
/// An object with the fields given (`undefined` values are left out).
pub(crate) fn object(fields: &[(&str, JsValue)]) -> Result<JsValue, JsValue> {
    let object = js_sys::Object::new();
    for (name, value) in fields {
        if !value.is_undefined() {
            js_sys::Reflect::set(&object, &JsValue::from_str(name), value)?;
        }
    }
    Ok(object.into())
}

#[cfg(feature = "bundled-runtime")]
/// What the compiler said while it read a workbook (texts) as warning records with the stable code
/// `BUILD_DIAGNOSTIC`.
pub(crate) fn compiler_diagnostics(texts: &[String]) -> Vec<sheetrs_protocol::Diagnostic> {
    texts
        .iter()
        .map(|text| sheetrs_protocol::Diagnostic::warning("BUILD_DIAGNOSTIC", text.clone()))
        .collect()
}

#[cfg(feature = "edit")]
/// The bytes of a workbook as a `Uint8Array` (a copy: the memory of the module can move).
pub(crate) fn bytes(data: &[u8]) -> JsValue {
    js_sys::Uint8Array::from(data).into()
}

#[cfg(any(feature = "query", feature = "edit"))]
/// One text or a list of them.
#[derive(Debug, Deserialize)]
#[serde(untagged)]
pub(crate) enum Targets {
    One(String),
    Many(Vec<String>),
}

#[cfg(any(feature = "query", feature = "edit"))]
impl Targets {
    pub(crate) fn into_vec(self) -> Vec<String> {
        match self {
            Self::One(one) => vec![one],
            Self::Many(many) => many,
        }
    }
}

#[cfg(any(feature = "query", feature = "edit"))]
/// The targets of a call: one text, or a list of them (`undefined` is none).
pub(crate) fn targets(value: JsValue, what: &str) -> Result<Vec<String>, JsValue> {
    if value.is_undefined() || value.is_null() {
        return Ok(Vec::new());
    }
    Ok(from_js::<Targets>(value, what)?.into_vec())
}
