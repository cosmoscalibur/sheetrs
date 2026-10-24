//! Running a [`Program`]: JSON in, JSON out, with run-time option overrides.

use crate::addr::{CellRef, Rect};
use crate::engine::Engine;
use crate::ir::{Binding, Program, Queryable};
use crate::options::EngineOptions;
use crate::value::{ErrorKind, Value};
use serde_json::{Map, Number, Value as Json};
use std::sync::Arc;

#[derive(Debug)]
pub struct RunError(pub String);

impl std::fmt::Display for RunError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for RunError {}

fn fail<T>(msg: impl Into<String>) -> Result<T, RunError> {
    Err(RunError(msg.into()))
}

/// Run-time choices. Only what a file cannot say and a build did not fix is
/// open: the clock and the random seed. The date system, iteration, search
/// settings and dialect are the file's, and the locale is fixed when the
/// program is built (folding already used it).
#[derive(Debug, Clone, Default)]
pub struct Overrides {
    pub now: Option<f64>,
    pub rand_seed: Option<u64>,
}

impl Program {
    pub fn effective_options(&self, o: &Overrides) -> EngineOptions {
        let mut opts = self.options.clone();
        opts.now = o.now.or(opts.now);
        opts.rand_seed = o.rand_seed.or(opts.rand_seed);
        opts
    }

    /// Evaluate with JSON inputs, returning a JSON object keyed by output name.
    pub fn execute(&self, input: &Json, o: &Overrides) -> Result<Json, RunError> {
        self.execute_with(input, o, &[])
    }

    /// Resolve a range typed by a caller (`Sheet1!B2:C9`) against the ranges the
    /// program can answer.
    pub fn resolve_query(&self, text: &str) -> Result<Rect, RunError> {
        let (sheet, tail) = crate::addr::split_sheet(text.trim());
        let sheet_id = match &sheet {
            Some(n) => self
                .sheets
                .iter()
                .position(|s| crate::addr::same_name(s, n)),
            None => Some(0),
        };
        let (Some(s), Some(area)) = (sheet_id, crate::addr::parse_area(tail)) else {
            return fail(format!("`{text}` is not a range of this workbook"));
        };
        let rect = crate::addr::bind(area, s as u16);
        let inside = |b: &Binding| {
            let o = &b.rect;
            o.sheet == rect.sheet
                && o.r0 <= rect.r0
                && o.c0 <= rect.c0
                && o.r1 >= rect.r1
                && o.c1 >= rect.c1
        };
        let allowed = match &self.queryable {
            Queryable::Any { excluded } => !excluded.contains(&rect.sheet),
            Queryable::Named(extra) => self.outputs.iter().chain(extra).any(inside),
            Queryable::Declared => self.outputs.iter().any(inside),
        };
        if allowed {
            Ok(rect)
        } else {
            fail(format!(
                "`{text}` cannot be queried: this program was built with outputs = {} (see --meta)",
                match &self.queryable {
                    Queryable::Declared => "declared",
                    Queryable::Named(_) => "named",
                    Queryable::Any { .. } => "any",
                }
            ))
        }
    }

    /// Like [`execute`](Self::execute), also answering `queries`: ranges keyed by
    /// their canonical text, subject to the program's [`Queryable`] level.
    pub fn execute_with(
        &self,
        input: &Json,
        o: &Overrides,
        queries: &[String],
    ) -> Result<Json, RunError> {
        let query_rects = queries
            .iter()
            .map(|q| self.resolve_query(q))
            .collect::<Result<Vec<_>, _>>()?;
        let mut engine = Engine::new(self, self.effective_options(o));
        let empty = Map::new();
        let obj = match input {
            Json::Object(m) => m,
            Json::Null => &empty,
            _ => return fail("the input must be a JSON object"),
        };
        for key in obj.keys() {
            if !self.inputs.iter().any(|b| &b.name == key) {
                let known: Vec<&str> = self.inputs.iter().map(|b| b.name.as_str()).collect();
                return fail(format!(
                    "unknown input `{key}` (declared inputs: {})",
                    known.join(", ")
                ));
            }
        }
        for b in &self.inputs {
            if let Some(j) = obj.get(&b.name) {
                for (at, v) in bind_input(b, j)? {
                    engine.set_input(at, v);
                }
            }
        }
        engine.run();
        let mut out = Map::new();
        for b in &self.outputs {
            out.insert(b.name.clone(), read_output(&mut engine, &b.rect));
        }
        for r in query_rects {
            let name = crate::addr::qualified(&self.sheets, &r);
            if !out.contains_key(&name) {
                out.insert(name, read_output(&mut engine, &r));
            }
        }
        Ok(Json::Object(out))
    }
}

fn scalar_in(j: &Json, name: &str) -> Result<Value, RunError> {
    Ok(match j {
        Json::Null => Value::Empty,
        Json::Bool(b) => Value::Bool(*b),
        Json::Number(n) => Value::Number(n.as_f64().unwrap_or(f64::NAN)),
        Json::String(s) => Value::Text(Arc::from(s.as_str())),
        Json::Object(m) => match m
            .get("error")
            .and_then(Json::as_str)
            .and_then(ErrorKind::parse)
        {
            Some(e) => Value::Error(e),
            None => {
                return fail(format!(
                    "input `{name}`: object values must be {{\"error\": \"#N/A\"}}"
                ));
            }
        },
        Json::Array(_) => return fail(format!("input `{name}`: unexpected nested array")),
    })
}

fn bind_input(b: &Binding, j: &Json) -> Result<Vec<(CellRef, Value)>, RunError> {
    let r = &b.rect;
    let (rows, cols) = (r.rows() as usize, r.cols() as usize);
    let at = |i: usize, k: usize| CellRef::new(r.sheet, r.r0 + i as u32, r.c0 + k as u32);
    if rows * cols == 1 {
        let inner = match j {
            Json::Array(a) if a.len() == 1 => match &a[0] {
                Json::Array(b) if b.len() == 1 => &b[0],
                x => x,
            },
            x => x,
        };
        return Ok(vec![(at(0, 0), scalar_in(inner, &b.name)?)]);
    }
    let Json::Array(items) = j else {
        return fail(format!(
            "input `{}` expects an array ({rows}×{cols})",
            b.name
        ));
    };
    let mut out = Vec::with_capacity(rows * cols);
    if items.iter().all(Json::is_array) && !items.is_empty() {
        if items.len() != rows {
            return fail(format!(
                "input `{}` expects {rows} row(s), got {}",
                b.name,
                items.len()
            ));
        }
        for (i, row) in items.iter().enumerate() {
            let Json::Array(row) = row else {
                unreachable!()
            };
            if row.len() != cols {
                return fail(format!(
                    "input `{}` row {} expects {cols} value(s), got {}",
                    b.name,
                    i + 1,
                    row.len()
                ));
            }
            for (k, v) in row.iter().enumerate() {
                out.push((at(i, k), scalar_in(v, &b.name)?));
            }
        }
    } else if rows == 1 || cols == 1 {
        if items.len() != rows * cols {
            return fail(format!(
                "input `{}` expects {} value(s), got {}",
                b.name,
                rows * cols,
                items.len()
            ));
        }
        for (i, v) in items.iter().enumerate() {
            let (ri, ci) = if rows == 1 { (0, i) } else { (i, 0) };
            out.push((at(ri, ci), scalar_in(v, &b.name)?));
        }
    } else {
        return fail(format!(
            "input `{}` expects an array of {rows} rows",
            b.name
        ));
    }
    Ok(out)
}

pub fn value_to_json(v: &Value) -> Json {
    match v {
        Value::Empty | Value::Array(_) | Value::Ref(_) => Json::Null,
        Value::Bool(b) => Json::Bool(*b),
        Value::Text(t) => Json::String(t.to_string()),
        Value::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                Json::Number(Number::from(*n as i64))
            } else {
                Number::from_f64(*n).map_or(Json::Null, Json::Number)
            }
        }
        Value::Error(e) => {
            let mut m = Map::new();
            m.insert("error".into(), Json::String(e.as_str().into()));
            Json::Object(m)
        }
    }
}

fn read_output(e: &mut Engine, r: &Rect) -> Json {
    if r.is_single() {
        return value_to_json(&e.get(r.top_left()));
    }
    Json::Array(
        (r.r0..=r.r1)
            .map(|row| {
                Json::Array(
                    (r.c0..=r.c1)
                        .map(|col| value_to_json(&e.get(CellRef::new(r.sheet, row, col))))
                        .collect(),
                )
            })
            .collect(),
    )
}
