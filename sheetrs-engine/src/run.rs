//! Running a [`Program`]: JSON in, JSON out, with run-time option overrides.

use crate::addr::{CellRef, Rect};
use crate::engine::{ConstSource, Engine};
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
/// settings are the file's, and the locale is fixed when the
/// program is built (folding already used it).
#[derive(Debug, Clone, Default)]
pub struct Overrides {
    pub now: Option<f64>,
    pub rand_seed: Option<u64>,
    pub host: crate::options::Host,
}

/// Orientation of the arrays returned for ranges (Google Sheets' `majorDimension`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum MajorDimension {
    /// An array of rows.
    #[default]
    Rows,
    /// An array of columns.
    Columns,
}

impl MajorDimension {
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_ascii_uppercase().as_str() {
            "ROWS" => Some(Self::Rows),
            "COLUMNS" => Some(Self::Columns),
            _ => None,
        }
    }
}

/// A requested range, resolved and checked.
struct Target {
    /// The defined name or output name it was asked by.
    name: Option<String>,
    /// `None`: an open-ended range (`A:A`) over an empty sheet.
    rect: Option<Rect>,
}

impl Program {
    pub fn effective_options(&self, o: &Overrides) -> EngineOptions {
        let mut opts = self.options.clone();
        opts.now = o.now.or(opts.now);
        opts.rand_seed = o.rand_seed.or(opts.rand_seed);
        opts.host = o.host.clone().over(opts.host);
        opts
    }

    /// Evaluate with JSON inputs (an object keyed by input name), returning the
    /// declared outputs as an object keyed by output name.
    pub fn execute(&self, inputs: &Json, o: &Overrides) -> Result<Json, RunError> {
        self.execute_over(None, inputs, o)
    }

    /// [`execute`](Self::execute) for a program whose constants are kept in `base`.
    pub fn execute_over(
        &self,
        base: Option<&dyn ConstSource>,
        inputs: &Json,
        o: &Overrides,
    ) -> Result<Json, RunError> {
        let mut engine = self.start(base, inputs, o)?;
        Ok(self.declared_outputs(&mut engine))
    }

    /// Evaluate with JSON inputs and read `ranges`, in the manner of the batch
    /// read of a spreadsheet API: an array with one `{range, values}` per request
    /// (plus `name` when asked by a name), in request order. See
    /// [`Queryable`] for what may be asked.
    pub fn query(
        &self,
        inputs: &Json,
        o: &Overrides,
        ranges: &[String],
        dim: MajorDimension,
    ) -> Result<Json, RunError> {
        self.query_over(None, inputs, o, ranges, dim)
    }

    /// [`query`](Self::query) for a program whose constants are kept in `base`.
    pub fn query_over(
        &self,
        base: Option<&dyn ConstSource>,
        inputs: &Json,
        o: &Overrides,
        ranges: &[String],
        dim: MajorDimension,
    ) -> Result<Json, RunError> {
        let targets = ranges
            .iter()
            .map(|t| self.resolve(base, t))
            .collect::<Result<Vec<_>, _>>()?;
        let mut engine = self.start(base, inputs, o)?;
        let mut out = Vec::with_capacity(targets.len());
        for t in targets {
            let mut m = Map::new();
            if let Some(n) = t.name {
                m.insert("name".into(), Json::String(n));
            }
            m.insert(
                "range".into(),
                Json::String(
                    t.rect
                        .map_or_else(String::new, |r| crate::addr::qualified(&self.sheets, &r)),
                ),
            );
            m.insert(
                "values".into(),
                t.rect
                    .map_or_else(|| Json::Array(Vec::new()), |r| matrix(&mut engine, &r, dim)),
            );
            out.push(Json::Object(m));
        }
        Ok(Json::Array(out))
    }

    /// Answer one request: `{"inputs": {...}, "ranges": [...], "majorDimension":
    /// "ROWS"}`, all optional. Without `ranges` the answer is
    /// `{"outputs": {...}}`, the outputs declared at build time; with them,
    /// `{"valueRanges": [...]}` for exactly those ranges. A program built with
    /// `outputs = any` has no default answer and requires `ranges`.
    pub fn handle(&self, request: &Json, o: &Overrides) -> Result<Json, RunError> {
        self.handle_over(None, request, o)
    }

    /// [`handle`](Self::handle) for a program whose constants are kept in `base`.
    pub fn handle_over(
        &self,
        base: Option<&dyn ConstSource>,
        request: &Json,
        o: &Overrides,
    ) -> Result<Json, RunError> {
        let empty = Map::new();
        let req = match request {
            Json::Object(m) => m,
            Json::Null => &empty,
            _ => return fail("the request must be a JSON object"),
        };
        for key in req.keys() {
            if !matches!(key.as_str(), "inputs" | "ranges" | "majorDimension") {
                return fail(format!(
                    "unknown key `{key}` in the request (expected inputs, ranges, majorDimension; \
                     input values go under `inputs`)"
                ));
            }
        }
        let inputs = req.get("inputs").unwrap_or(&Json::Null);
        let dim = match req.get("majorDimension") {
            None => MajorDimension::Rows,
            Some(j) => j
                .as_str()
                .and_then(MajorDimension::parse)
                .ok_or_else(|| RunError("majorDimension must be \"ROWS\" or \"COLUMNS\"".into()))?,
        };
        let ranges: Option<Vec<String>> = match req.get("ranges") {
            None => None,
            Some(Json::Array(a)) if !a.is_empty() => Some(
                a.iter()
                    .map(|v| v.as_str().map(str::to_string))
                    .collect::<Option<_>>()
                    .ok_or_else(|| RunError("`ranges` must be an array of strings".into()))?,
            ),
            Some(_) => return fail("`ranges` must be a non-empty array of strings"),
        };
        match ranges {
            Some(r) => {
                Ok(serde_json::json!({ "valueRanges": self.query_over(base, inputs, o, &r, dim)? }))
            }
            None if matches!(self.queryable, Queryable::Any { .. }) => fail(
                "this program was built with outputs = any: the request must list `ranges` (see --meta)",
            ),
            None => Ok(serde_json::json!({ "outputs": self.execute_over(base, inputs, o)? })),
        }
    }

    fn start<'p>(
        &'p self,
        base: Option<&'p dyn ConstSource>,
        input: &Json,
        o: &Overrides,
    ) -> Result<Engine<'p>, RunError> {
        let opts = self.effective_options(o);
        let mut engine = match base {
            Some(b) => Engine::with_base(self, opts, b),
            None => Engine::new(self, opts),
        };
        let empty = Map::new();
        let obj = match input {
            Json::Object(m) => m,
            Json::Null => &empty,
            _ => return fail("`inputs` must be a JSON object"),
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
        Ok(engine)
    }

    fn declared_outputs(&self, engine: &mut Engine) -> Json {
        let mut out = Map::new();
        for b in &self.outputs {
            out.insert(b.name.clone(), read_output(engine, &b.rect));
        }
        Json::Object(out)
    }

    /// Resolve what a caller asked for against what the program can answer: a
    /// declared output's name, a defined name (levels `named` and `any`), or a
    /// range (`Sheet1!B2:C9`, `Sheet1!A:A`; open-ended ones end at the used area).
    fn resolve(&self, base: Option<&dyn ConstSource>, text: &str) -> Result<Target, RunError> {
        let text = text.trim();
        let same = crate::addr::same_name;
        let names: &[Binding] = match &self.queryable {
            Queryable::Declared => &[],
            Queryable::Named(n) | Queryable::Any { names: n, .. } => n,
        };
        if let Some(b) = self
            .outputs
            .iter()
            .chain(names)
            .find(|b| same(&b.name, text))
        {
            return Ok(Target {
                name: Some(b.name.clone()),
                rect: Some(b.rect),
            });
        }
        let (sheet, tail) = crate::addr::split_sheet(text);
        let sheet_id = match &sheet {
            Some(n) => self.sheets.iter().position(|s| same(s, n)),
            None => Some(0),
        };
        let (Some(s), Some(area)) = (sheet_id, crate::addr::parse_area(tail)) else {
            return fail(format!(
                "`{text}` is not a range or a name of this workbook"
            ));
        };
        let mut rect = crate::addr::bind(area, s as u16);
        if rect.r1 + 1 >= crate::addr::MAX_ROWS || rect.c1 + 1 >= crate::addr::MAX_COLS {
            // `A:A`, `3:5`: as far as the sheet has data.
            let (mut rows, mut cols) = (0, 0);
            for c in self.cells.iter().filter(|c| c.at.sheet == rect.sheet) {
                rows = rows.max(c.at.row + 1);
                cols = cols.max(c.at.col + 1);
            }
            if let Some(b) = base {
                let (r, c) = b.extent(rect.sheet);
                rows = rows.max(r);
                cols = cols.max(c);
            }
            let r1 = rect.r1.min(rows.saturating_sub(1));
            let c1 = rect.c1.min(cols.saturating_sub(1));
            if rows == 0 || cols == 0 || r1 < rect.r0 || c1 < rect.c0 {
                return Ok(Target {
                    name: None,
                    rect: None,
                });
            }
            rect = Rect::new(rect.sheet, rect.r0, rect.c0, r1, c1);
        }
        let inside = |b: &Binding| {
            let o = &b.rect;
            o.sheet == rect.sheet
                && o.r0 <= rect.r0
                && o.c0 <= rect.c0
                && o.r1 >= rect.r1
                && o.c1 >= rect.c1
        };
        let allowed = match &self.queryable {
            Queryable::Any { excluded, .. } => !excluded.contains(&rect.sheet),
            Queryable::Named(extra) => self.outputs.iter().chain(extra).any(inside),
            Queryable::Declared => self.outputs.iter().any(inside),
        };
        if allowed {
            Ok(Target {
                name: None,
                rect: Some(rect),
            })
        } else {
            fail(format!(
                "`{text}` cannot be read: this program was built with outputs = {} (see --meta)",
                match &self.queryable {
                    Queryable::Declared => "declared",
                    Queryable::Named(_) => "named",
                    Queryable::Any { .. } => "any",
                }
            ))
        }
    }
}

/// `values` of a range: always an array of rows (or of columns), whatever its size.
fn matrix(e: &mut Engine, r: &Rect, dim: MajorDimension) -> Json {
    let cell =
        |row: u32, col: u32, e: &mut Engine| value_to_json(&e.get(CellRef::new(r.sheet, row, col)));
    match dim {
        MajorDimension::Rows => Json::Array(
            (r.r0..=r.r1)
                .map(|row| Json::Array((r.c0..=r.c1).map(|col| cell(row, col, e)).collect()))
                .collect(),
        ),
        MajorDimension::Columns => Json::Array(
            (r.c0..=r.c1)
                .map(|col| Json::Array((r.r0..=r.r1).map(|row| cell(row, col, e)).collect()))
                .collect(),
        ),
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
