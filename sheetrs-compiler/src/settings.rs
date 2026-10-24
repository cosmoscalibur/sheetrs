//! The settings sheet: the workbook's own declaration of inputs, outputs and
//! metadata.
//!
//! The first row holds the headers `key`, `value`, `description` and
//! `meta_type` (any order, other columns ignored). A row whose `meta_type` is
//! `input` or `output` gives the name `key` to the range in `value`: the name
//! is the JSON key of that input or output. `value` starts with the sheet
//! (`Sheet1!B2:C9`, `'My Sheet'!A1`, a bare sheet name for its used range) or is
//! a defined name. For any other `meta_type`, `key` is the name of a metadata
//! field and `value` its literal value. Everything is listed by the executable's
//! `--meta`.

use crate::addr;
use crate::compile::{CompileError, resolve_target};
use crate::ir::Meta;
use crate::model::Model;
use crate::run::value_to_json;
use crate::value::Value;
use std::collections::{BTreeMap, HashSet};

pub const DEFAULT_SHEET: &str = "settings";

#[derive(Debug, Clone, Default)]
pub struct SheetSettings {
    pub sheet: u16,
    /// `(name, reference)` in sheet order.
    pub inputs: Vec<(String, String)>,
    pub outputs: Vec<(String, String)>,
    pub meta: Vec<Meta>,
}

fn text(v: &Value) -> String {
    match v {
        Value::Text(t) => t.trim().to_string(),
        Value::Number(n) if n.fract() == 0.0 && n.abs() < 1e15 => format!("{}", *n as i64),
        Value::Number(n) => n.to_string(),
        Value::Bool(b) => b.to_string().to_uppercase(),
        _ => String::new(),
    }
}

/// The range a settings row points at. It must start with the sheet (or be a
/// sheet or a defined name): an address with no sheet depends on which sheet is
/// first, and the settings sheet is not the place to guess.
fn resolve_reference(model: &Model, reference: &str) -> Result<crate::addr::Rect, CompileError> {
    let r = reference.trim();
    if r.is_empty() {
        return Err(CompileError("`value` must hold the range".into()));
    }
    if let Some(id) = model.sheet_id(r) {
        let used = model.cells.keys().filter(|a| a.sheet == id).fold(
            None,
            |acc: Option<(u32, u32)>, a| {
                Some(acc.map_or((a.row, a.col), |(r, c)| (r.max(a.row), c.max(a.col))))
            },
        );
        return used
            .map(|(rows, cols)| crate::addr::Rect::new(id, 0, 0, rows, cols))
            .ok_or_else(|| CompileError(format!("sheet `{r}` is empty")));
    }
    let (sheet, _) = addr::split_sheet(r);
    let is_name = model.names.keys().any(|k| addr::same_name(k, r));
    if sheet.is_none() && !is_name {
        return Err(CompileError(format!(
            "`{r}` must start with the sheet (`Sheet1!{r}`), or be a sheet or a defined name"
        )));
    }
    resolve_target(model, r)
}

/// Read the settings sheet called `name`. A missing sheet is an error only when
/// `required` (the caller named it), since the default name is optional.
pub fn read(
    model: &Model,
    name: &str,
    required: bool,
) -> Result<Option<SheetSettings>, CompileError> {
    let Some(sheet) = model.sheet_id(name) else {
        return if required {
            Err(CompileError(format!("settings sheet `{name}` not found")))
        } else {
            Ok(None)
        };
    };
    let mut rows: BTreeMap<u32, BTreeMap<u32, &Value>> = BTreeMap::new();
    for (at, c) in &model.cells {
        if at.sheet == sheet {
            rows.entry(at.row).or_default().insert(at.col, &c.value);
        }
    }
    let at = |m: &str| format!("settings sheet `{name}`: {m}");
    let Some(header) = rows.get(&0) else {
        return Err(CompileError(at(
            "the first row must hold the headers key, value, description, meta_type",
        )));
    };
    let col = |h: &str| {
        header
            .iter()
            .find(|(_, v)| text(v).eq_ignore_ascii_case(h))
            .map(|(c, _)| *c)
    };
    let (Some(kc), Some(tc)) = (col("key"), col("meta_type")) else {
        return Err(CompileError(at(
            "the headers `key` and `meta_type` are required in the first row",
        )));
    };
    let (vc, dc) = (col("value"), col("description"));

    let mut out = SheetSettings {
        sheet,
        ..SheetSettings::default()
    };
    let mut seen: HashSet<(String, String)> = HashSet::new();
    for (&row, cells) in rows.iter().skip(1) {
        let get = |c: u32| cells.get(&c).map_or_else(String::new, |v| text(v));
        let (key, meta_type) = (get(kc), get(tc).to_lowercase());
        if key.is_empty() && meta_type.is_empty() {
            continue;
        }
        let line = row + 1;
        if key.is_empty() || meta_type.is_empty() {
            return Err(CompileError(at(&format!(
                "row {line} needs both `key` and `meta_type`"
            ))));
        }
        let value = vc.and_then(|c| cells.get(&c));
        let mut shown = value.map_or(serde_json::Value::Null, |v| value_to_json(v));
        if matches!(meta_type.as_str(), "input" | "output") {
            let reference = value.map(|v| text(v)).unwrap_or_default();
            let rect = resolve_reference(model, &reference)
                .map_err(|e| CompileError(at(&format!("row {line}: {e}"))))?;
            if rect.sheet == sheet {
                return Err(CompileError(at(&format!(
                    "row {line}: `{reference}` lies in the settings sheet"
                ))));
            }
            shown = serde_json::Value::String(addr::qualified(&model.sheets, &rect));
            let list = if meta_type == "input" {
                &mut out.inputs
            } else {
                &mut out.outputs
            };
            list.push((key.clone(), reference));
        }
        if !seen.insert((meta_type.clone(), key.clone())) {
            return Err(CompileError(at(&format!(
                "row {line}: `{key}` is already listed as {meta_type}"
            ))));
        }
        out.meta.push(Meta {
            key,
            value: shown,
            description: dc.map_or_else(String::new, get),
            meta_type,
        });
    }
    Ok(Some(out))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::Ns;

    fn model() -> Model {
        let mut m = Model::new(Ns::Xlsx);
        let data = m.add_sheet("Data");
        let s = m.add_sheet("settings");
        m.set(data, "A1", Value::Number(1.0));
        m.set(data, "B2", Value::Number(2.0));
        for (c, h) in ["key", "value", "description", "meta_type"]
            .iter()
            .enumerate()
        {
            m.set(
                s,
                &format!("{}1", (b'A' + c as u8) as char),
                Value::Text((*h).into()),
            );
        }
        let rows = [
            ("rate", "Data!A1", "interest rate", "input"),
            ("result", "Data!B1:B3", "result", "output"),
            ("version", "1.2", "model version", "info"),
        ];
        for (i, (k, v, d, t)) in rows.iter().enumerate() {
            let r = i + 2;
            m.set(s, &format!("A{r}"), Value::Text((*k).into()));
            m.set(s, &format!("B{r}"), Value::Text((*v).into()));
            m.set(s, &format!("C{r}"), Value::Text((*d).into()));
            m.set(s, &format!("D{r}"), Value::Text((*t).into()));
        }
        m
    }

    #[test]
    fn names_ranges_and_carries_metadata() {
        let s = read(&model(), "SETTINGS", false).unwrap().unwrap();
        assert_eq!(s.inputs, [("rate".to_string(), "Data!A1".to_string())]);
        assert_eq!(s.outputs[0].0, "result");
        assert_eq!(s.meta.len(), 3);
        // Inputs and outputs list their canonical range as the value.
        assert_eq!(s.meta[0].key, "rate");
        assert_eq!(s.meta[0].value, serde_json::json!("Data!A1"));
        assert_eq!(s.meta[0].description, "interest rate");
        // Other types carry the literal.
        assert_eq!(s.meta[2].meta_type, "info");
        assert_eq!(s.meta[2].value, serde_json::json!("1.2"));
    }

    #[test]
    fn references_start_with_the_sheet() {
        let mut m = model();
        m.names.insert("Total".into(), "Data!B2".into());
        let with = |m: &mut Model, v: &str| {
            m.set(1, "B3", Value::Text(v.into()));
            read(m, "settings", false)
        };
        // A bare address is refused; a sheet name or a defined name is not.
        let e = with(&mut m, "B1:B3").unwrap_err().0;
        assert!(e.contains("must start with the sheet"), "{e}");
        assert!(with(&mut m, "Total").is_ok());
        assert!(with(&mut m, "Data").is_ok());
        assert!(with(&mut m, "Data!$B$1:$B$3").is_ok());
        assert!(with(&mut m, "Nowhere!A1").is_err());
        assert!(with(&mut m, "").is_err());
    }

    #[test]
    fn absent_sheet_is_optional_unless_named() {
        assert!(read(&model(), "other", false).unwrap().is_none());
        assert!(read(&model(), "other", true).is_err());
    }

    #[test]
    fn bad_rows_fail_fast() {
        let mut m = model();
        m.set(1, "A5", Value::Text("rate".into()));
        m.set(1, "B5", Value::Text("Data!B2".into()));
        m.set(1, "D5", Value::Text("input".into()));
        let e = read(&m, "settings", false).unwrap_err().0;
        assert!(e.contains("row 5") && e.contains("already listed"), "{e}");
        // The same name may be an input and an output.
        let mut m = model();
        m.set(1, "A5", Value::Text("rate".into()));
        m.set(1, "B5", Value::Text("Data!B2".into()));
        m.set(1, "D5", Value::Text("output".into()));
        assert!(read(&m, "settings", false).is_ok());
    }
}
