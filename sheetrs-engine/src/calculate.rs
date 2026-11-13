//! Running a [`Program`]: JSON in, JSON out, with run-time option overrides.

use crate::addr::{CellRef, Rect};
use crate::engine::{ConstSource, Engine};
use crate::error::{Error, Result};
use crate::ir::{Binding, Program};
use crate::options::EngineOptions;
use crate::value::{ErrorKind, Value};
use serde_json::{Map, Number, Value as Json};
use std::sync::Arc;

/// Say what is wrong with a request that does not have the shape of
/// [`sheetrs_protocol::Request`], naming the field.
fn explain_request(request: &Json, serde_says: &str) -> Error {
    let Json::Object(req) = request else {
        return Error::Request("the request must be a JSON object".into());
    };
    for key in req.keys() {
        if !matches!(key.as_str(), "inputs" | "ranges" | "majorDimension") {
            return Error::Request(format!(
                "unknown key `{key}` in the request (expected inputs, ranges, majorDimension; \
                 input values go under `inputs`)"
            ));
        }
    }
    if req.contains_key("majorDimension") {
        let ok = req
            .get("majorDimension")
            .and_then(Json::as_str)
            .and_then(MajorDimension::parse)
            .is_some();
        if !ok {
            return Error::Request("majorDimension must be \"ROWS\" or \"COLUMNS\"".into());
        }
    }
    match req.get("ranges") {
        Some(Json::Array(a)) if a.iter().all(Json::is_string) => {}
        Some(Json::Array(_)) => {
            return Error::Request("`ranges` must be an array of strings".into());
        }
        Some(_) => return Error::Request("`ranges` must be a non-empty array of strings".into()),
        None => {}
    }
    if matches!(req.get("inputs"), Some(i) if !i.is_object() && !i.is_null()) {
        return Error::Request("`inputs` must be a JSON object".into());
    }
    Error::Request(format!("invalid request: {serde_says}"))
}

fn fail<T>(msg: impl Into<String>) -> Result<T> {
    Err(Error::Request(msg.into()))
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

pub use sheetrs_protocol::{MajorDimension, ValueRange, ValueRanges};

/// A requested range, resolved and checked.
struct Target {
    /// The defined name or output name it was asked by.
    name: Option<String>,
    /// `None`: an open-ended range (`A:A`) over an empty sheet.
    rect: Option<Rect>,
}

impl Program {
    #[must_use]
    pub fn effective_options(&self, o: &Overrides) -> EngineOptions {
        let mut opts = self.options.clone();
        opts.now = o.now.or(opts.now);
        opts.rand_seed = o.rand_seed.or(opts.rand_seed);
        opts.host = o.host.clone().over(opts.host);
        opts
    }

    /// Evaluate with JSON inputs (an object keyed by input name), returning the
    /// declared outputs as an object keyed by output name.
    ///
    /// # Errors
    ///
    /// When an input is not one of the program, or the program fails.
    pub fn calculate(&self, inputs: &Json, o: &Overrides) -> Result<Json> {
        self.calculate_over(None, inputs, o)
    }

    /// [`calculate`](Self::calculate) for a program whose constants are kept in `base`.
    ///
    /// # Errors
    ///
    /// When an input is not one of the program, or the program fails.
    pub fn calculate_over(
        &self,
        base: Option<&dyn ConstSource>,
        inputs: &Json,
        o: &Overrides,
    ) -> Result<Json> {
        let mut engine = self.start(base, inputs, o)?;
        Ok(Json::Object(self.declared_outputs(&mut engine)))
    }

    /// Evaluate with JSON inputs and read `ranges`, in the manner of the batch
    /// read of a spreadsheet API: an array with one `{range, values}` per request
    /// (plus `name` when asked by a name), in request order. See
    /// [`Program::queryable`](crate::ir::Program) for what may be asked.
    ///
    /// # Errors
    ///
    /// When an input or an output is not one of the program, or the program fails.
    pub fn query(
        &self,
        inputs: &Json,
        o: &Overrides,
        ranges: &[String],
        dim: MajorDimension,
    ) -> Result<Json> {
        self.query_over(None, inputs, o, ranges, dim)
    }

    /// [`query`](Self::query) for a program whose constants are kept in `base`.
    ///
    /// # Errors
    ///
    /// When an input or an output is not one of the program, or the program fails.
    pub fn query_over(
        &self,
        base: Option<&dyn ConstSource>,
        inputs: &Json,
        o: &Overrides,
        ranges: &[String],
        dim: MajorDimension,
    ) -> Result<Json> {
        let targets = ranges
            .iter()
            .map(|t| self.resolve(base, t))
            .collect::<Result<Vec<_>, _>>()?;
        let mut engine = self.start(base, inputs, o)?;
        let ranges = self.read_targets(&mut engine, targets, dim);
        Ok(Json::Array(
            ranges
                .iter()
                .map(|range| serde_json::to_value(range).unwrap_or(Json::Null))
                .collect(),
        ))
    }

    /// The `{range, values}` of each target, in order.
    fn read_targets(
        &self,
        engine: &mut Engine,
        targets: Vec<Target>,
        dim: MajorDimension,
    ) -> Vec<ValueRange> {
        targets
            .into_iter()
            .map(|t| ValueRange {
                name: t.name,
                range: t
                    .rect
                    .map_or_else(String::new, |r| crate::addr::qualified(&self.sheets, &r)),
                values: t
                    .rect
                    .map_or_else(Vec::new, |r| matrix_rows(engine, &r, dim)),
            })
            .collect()
    }

    /// Evaluate with the inputs of `inputs` (an object keyed by input name) and read `rects`, rows
    /// first: one matrix of values per rectangle, in order. It is how the workbook answer takes the
    /// values of its formula cells, whatever the outputs the program declares.
    ///
    /// # Errors
    ///
    /// When an input is not one of the program.
    pub fn read_rects(
        &self,
        base: Option<&dyn ConstSource>,
        inputs: &Json,
        o: &Overrides,
        rects: &[Rect],
    ) -> Result<Vec<Json>> {
        let mut engine = self.start(base, inputs, o)?;
        Ok(rects
            .iter()
            .map(|r| matrix(&mut engine, r, MajorDimension::Rows))
            .collect())
    }

    /// Answer one request: `{"inputs": {...}, "ranges": [...], "majorDimension":
    /// "ROWS"}`, all optional. The answer is always `{"valueRanges": [...]}`: for exactly the
    /// `ranges` asked, or, without them, one entry (with its `name`) for each output declared at
    /// build time.
    ///
    /// # Errors
    ///
    /// When the request is not valid or the program fails.
    pub fn handle(&self, request: &Json, o: &Overrides) -> Result<ValueRanges> {
        self.handle_over(None, request, o)
    }

    /// [`handle`](Self::handle) for a program whose constants are kept in `base`.
    ///
    /// # Errors
    ///
    /// When the request is not valid or the program fails.
    pub fn handle_over(
        &self,
        base: Option<&dyn ConstSource>,
        request: &Json,
        o: &Overrides,
    ) -> Result<ValueRanges> {
        Ok(self.handle_with(base, request, o, &[])?.0)
    }

    /// [`handle_over`](Self::handle_over) and, from the same calculation, the values of `rects`
    /// (rows first, one matrix per rectangle): the JSON answer and the workbook answer of a program
    /// that carries its template come from one run.
    ///
    /// # Errors
    ///
    /// When the request is not valid or the program fails.
    pub fn handle_with(
        &self,
        base: Option<&dyn ConstSource>,
        request: &Json,
        o: &Overrides,
        rects: &[Rect],
    ) -> Result<(ValueRanges, Vec<Json>)> {
        let request = match request {
            Json::Null => sheetrs_protocol::Request::default(),
            other => serde_json::from_value(other.clone())
                .map_err(|e| explain_request(other, &e.to_string()))?,
        };
        let inputs = request.inputs.map_or(Json::Null, Json::Object);
        let inputs = &inputs;
        let dim = request.major_dimension.unwrap_or_default();
        let ranges = match request.ranges {
            Some(r) if r.is_empty() => {
                return fail("`ranges` must be a non-empty array of strings");
            }
            r => r,
        };
        let targets = match ranges.as_ref() {
            Some(r) => r
                .iter()
                .map(|t| self.resolve(base, t))
                .collect::<Result<Vec<_>, _>>()?,
            // No range asked: the outputs the program declares, each with its name.
            None => self
                .outputs
                .iter()
                .map(|b| Target {
                    name: Some(b.name.clone()),
                    rect: Some(b.rect),
                })
                .collect(),
        };
        let mut engine = self.start(base, inputs, o)?;
        let answer = ValueRanges {
            value_ranges: self.read_targets(&mut engine, targets, dim),
        };
        let extra = rects
            .iter()
            .map(|r| matrix(&mut engine, r, MajorDimension::Rows))
            .collect();
        Ok((answer, extra))
    }

    fn start<'p>(
        &'p self,
        base: Option<&'p dyn ConstSource>,
        input: &Json,
        o: &Overrides,
    ) -> Result<Engine<'p>> {
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
        // Each key is an input binding (declared or deduced), an alias (a defined name that lies
        // inside the inputs) or a range inside the inputs.
        let mut bound: Vec<(Binding, &Json)> = Vec::with_capacity(obj.len());
        for (key, j) in obj {
            bound.push((self.resolve_input(key, &engine)?, j));
        }
        for (b, j) in bound {
            for (at, v) in bind_input(&b, j)? {
                engine.set_input(at, v);
            }
        }
        engine.calculate();
        Ok(engine)
    }

    fn declared_outputs(&self, engine: &mut Engine) -> Map<String, Json> {
        let mut out = Map::new();
        for b in &self.outputs {
            out.insert(b.name.clone(), read_output(engine, &b.rect));
        }
        out
    }

    /// Resolve what a caller asked for against what the program can answer: a
    /// output's name, a defined name that lies inside the outputs, or a
    /// range (`Sheet1!B2:C9`, `Sheet1!A:A`; open-ended ones end at the used area).
    fn resolve(&self, base: Option<&dyn ConstSource>, text: &str) -> Result<Target> {
        let text = text.trim();
        let same = crate::addr::same_name;
        let names: &[Binding] = &self.queryable.names;
        if let Some(b) = self.outputs.iter().find(|b| same(&b.name, text)) {
            return Ok(Target {
                name: Some(b.name.clone()),
                rect: Some(b.rect),
            });
        }
        // An alias: a defined name of the workbook; it is read if it lies inside the outputs.
        let alias = names.iter().find(|b| same(&b.name, text));
        let Some(mut rect) = (match alias {
            Some(b) => Some(b.rect),
            None => self.parse_range(text),
        }) else {
            return fail(format!(
                "`{text}` is not a range or a name of this workbook"
            ));
        };
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
        if self.outputs.iter().any(inside) {
            Ok(Target {
                name: alias.map(|b| b.name.clone()),
                rect: Some(rect),
            })
        } else {
            fail(format!(
                "`{text}` cannot be read: it is not inside the outputs of this program (see --meta)"
            ))
        }
    }

    /// A range written as text (`Sheet1!B2:C9`, `Sheet1!A:A`, a bare address on the first sheet).
    fn parse_range(&self, text: &str) -> Option<Rect> {
        let same = crate::addr::same_name;
        let (sheet, tail) = crate::addr::split_sheet(text);
        let sheet_id = match &sheet {
            Some(n) => self.sheets.iter().position(|s| same(s, n)),
            None => Some(0),
        }?;
        let area = crate::addr::parse_area(tail)?;
        Some(crate::addr::bind(area, sheet_id as u16))
    }

    /// The binding a key of the `inputs` of a request stands for: the binding of that name, an alias
    /// (a defined name of the workbook) or a range, the last two only where every cell is an input.
    fn resolve_input(&self, key: &str, engine: &Engine) -> Result<Binding> {
        if let Some(b) = self.inputs.iter().find(|b| b.name == key) {
            return Ok(b.clone());
        }
        let b = self.input_range(key)?;
        let rect = b.rect;
        for row in rect.r0..=rect.r1 {
            for col in rect.c0..=rect.c1 {
                if !engine.is_input(CellRef::new(rect.sheet, row, col)) {
                    return fail(format!(
                        "`{key}` is not an input: the cell {}!{} is not one (see --meta)",
                        self.sheets[usize::from(rect.sheet)],
                        crate::addr::a1(row, col)
                    ));
                }
            }
        }
        Ok(b)
    }

    /// The range an entry of the `inputs` of a request stands for: a binding of the program, an
    /// alias (a defined name of the workbook) or a range.
    ///
    /// # Errors
    ///
    /// When the key is none of those.
    pub fn input_range(&self, key: &str) -> Result<Binding> {
        let same = crate::addr::same_name;
        if let Some(b) = self.inputs.iter().find(|b| b.name == key) {
            return Ok(b.clone());
        }
        let rect = self
            .queryable
            .names
            .iter()
            .find(|b| same(&b.name, key))
            .map(|b| b.rect)
            .or_else(|| self.parse_range(key));
        let Some(rect) =
            rect.filter(|r| r.r1 + 1 < crate::addr::MAX_ROWS && r.c1 + 1 < crate::addr::MAX_COLS)
        else {
            let known: Vec<&str> = self
                .inputs
                .iter()
                .map(|b| b.name.as_str())
                .take(8)
                .collect();
            return fail(format!(
                "unknown input `{key}` (inputs: {}{})",
                known.join(", "),
                if self.inputs.len() > known.len() {
                    ", …"
                } else {
                    ""
                }
            ));
        };
        Ok(Binding {
            name: key.to_string(),
            rect,
        })
    }
}

/// `values` of a range: always an array of rows (or of columns), whatever its size.
fn matrix(e: &mut Engine, r: &Rect, dim: MajorDimension) -> Json {
    Json::Array(
        matrix_rows(e, r, dim)
            .into_iter()
            .map(Json::Array)
            .collect(),
    )
}

/// The values of a rectangle as arrays of rows (or of columns).
fn matrix_rows(e: &mut Engine, r: &Rect, dim: MajorDimension) -> Vec<Vec<Json>> {
    let cell =
        |row: u32, col: u32, e: &mut Engine| value_to_json(&e.get(CellRef::new(r.sheet, row, col)));
    match dim {
        MajorDimension::Rows => (r.r0..=r.r1)
            .map(|row| (r.c0..=r.c1).map(|col| cell(row, col, e)).collect())
            .collect(),
        MajorDimension::Columns => (r.c0..=r.c1)
            .map(|col| (r.r0..=r.r1).map(|row| cell(row, col, e)).collect())
            .collect(),
    }
}

fn scalar_in(j: &Json, name: &str) -> Result<Value> {
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

fn bind_input(b: &Binding, j: &Json) -> Result<Vec<(CellRef, Value)>> {
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
            // An error only `LibreOffice` has keeps its code next to the name Excel would give it.
            if matches!(e, ErrorKind::Libre(_))
                && let Some(code) = e.libre_code()
            {
                m.insert("code".into(), Json::String(code));
            }
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
