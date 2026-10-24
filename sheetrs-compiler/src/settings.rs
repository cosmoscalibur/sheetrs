//! The settings sheet: the workbook's own declaration of inputs, outputs and
//! metadata.
//!
//! The first row holds the headers `key`, `value`, `description` and
//! `meta_type` (any order, other columns ignored). A row whose `meta_type` is
//! `input` or `output` declares a range in `key` (`Sheet1!B2:C9`, `A1`, or a
//! defined name); any other `meta_type` is free metadata (`version`, `author`…).
//! Everything is listed by the executable's `--meta`.

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
    /// Targets as written, in sheet order.
    pub inputs: Vec<String>,
    pub outputs: Vec<String>,
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
        let key = match meta_type.as_str() {
            "input" | "output" => {
                let rect = resolve_target(model, &key)
                    .map_err(|e| CompileError(at(&format!("row {line}: {e}"))))?;
                if rect.sheet == sheet {
                    return Err(CompileError(at(&format!(
                        "row {line}: `{key}` lies in the settings sheet"
                    ))));
                }
                if meta_type == "input" {
                    out.inputs.push(key.clone());
                } else {
                    out.outputs.push(key.clone());
                }
                addr::qualified(&model.sheets, &rect)
            }
            _ => key,
        };
        if !seen.insert((meta_type.clone(), key.clone())) {
            return Err(CompileError(at(&format!(
                "row {line}: `{key}` is already listed as {meta_type}"
            ))));
        }
        out.meta.push(Meta {
            key,
            value: vc
                .and_then(|c| cells.get(&c))
                .map_or(serde_json::Value::Null, |v| value_to_json(v)),
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
            ("Data!A1", "", "rate", "input"),
            ("Data!B1:B3", "", "result", "output"),
            ("version", "1.2", "model version", "info"),
        ];
        for (i, (k, v, d, t)) in rows.iter().enumerate() {
            let r = i + 2;
            m.set(s, &format!("A{r}"), Value::Text((*k).into()));
            if !v.is_empty() {
                m.set(s, &format!("B{r}"), Value::Text((*v).into()));
            }
            m.set(s, &format!("C{r}"), Value::Text((*d).into()));
            m.set(s, &format!("D{r}"), Value::Text((*t).into()));
        }
        m
    }

    #[test]
    fn reads_bindings_and_metadata() {
        let s = read(&model(), "SETTINGS", false).unwrap().unwrap();
        assert_eq!(s.inputs, ["Data!A1"]);
        assert_eq!(s.outputs, ["Data!B1:B3"]);
        assert_eq!(s.meta.len(), 3);
        assert_eq!(s.meta[2].meta_type, "info");
        assert_eq!(s.meta[2].value, serde_json::json!("1.2"));
        assert_eq!(s.meta[0].description, "rate");
    }

    #[test]
    fn absent_sheet_is_optional_unless_named() {
        assert!(read(&model(), "other", false).unwrap().is_none());
        assert!(read(&model(), "other", true).is_err());
    }

    #[test]
    fn bad_rows_fail_fast() {
        let mut m = model();
        m.set(1, "A5", Value::Text("Data!A1".into()));
        m.set(1, "D5", Value::Text("input".into()));
        let e = read(&m, "settings", false).unwrap_err().0;
        assert!(e.contains("row 5") && e.contains("already listed"), "{e}");
        let mut m = model();
        m.set(1, "A5", Value::Text("Nowhere!Z1".into()));
        m.set(1, "D5", Value::Text("output".into()));
        assert!(read(&m, "settings", false).is_err());
    }
}
