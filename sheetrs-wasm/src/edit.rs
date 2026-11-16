//! The operations of a `Book`: they accumulate in its plan, and `save` runs them.

use crate::book::{Book, failure};
use crate::js::{Targets, bytes, from_js, from_js_object, targets, to_js};
use serde::Deserialize;
use serde_json::Value as Json;
use sheetrs_edit::{BookEdit, Calculation, ColumnRange, RowRange, ValueScope};
use sheetrs_io::write::Operation;
use sheetrs_protocol::{Find, Mode, Replace, Request, Scope, SearchIn};
use std::collections::BTreeMap;
use wasm_bindgen::prelude::*;

fn parse_all<T: std::str::FromStr>(items: &[String]) -> Result<Vec<T>, JsValue>
where
    T::Err: std::fmt::Display,
{
    items
        .iter()
        .map(|item| {
            item.parse()
                .map_err(|e| failure("Edit error", format!("`{item}`: {e}")))
        })
        .collect()
}

/// The options of `replace`: the search of a query (`mode`, `matchCase`, `entireCells`, `searchIn`,
/// `formatted`, `filtered`) and the `scope`: a sheet (`Data`) or a range (`Data!A1:C9`), or a list
/// of them; the whole workbook when absent. A replacement always looks inside formulas, as
/// `LibreOffice` does.
#[derive(Debug, Default, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
#[allow(clippy::struct_excessive_bools)] // the options of a search, as the protocol has them
struct ReplaceArgs {
    mode: Mode,
    match_case: bool,
    entire_cells: bool,
    search_in: SearchIn,
    formatted: bool,
    filtered: bool,
    scope: Option<Targets>,
}

/// The replacements: a list of `[text, with]` pairs, or an object `{text: with}`.
#[derive(Debug, Deserialize)]
#[serde(untagged)]
enum Pairs {
    List(Vec<(String, String)>),
    Map(BTreeMap<String, String>),
}

/// What a run did besides writing the workbook.
#[wasm_bindgen]
pub struct Report {
    inner: sheetrs_edit::Report,
}

#[wasm_bindgen]
impl Report {
    /// The cells a `replace` changed.
    #[wasm_bindgen(getter)]
    pub fn replaced(&self) -> f64 {
        self.inner.replaced as f64
    }

    /// The cells written (the results a program computed into a template).
    #[wasm_bindgen(getter)]
    pub fn written(&self) -> f64 {
        self.inner.written as f64
    }

    /// The cells a `replace` found and left: `{sheet, cell, reason}`.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter)]
    pub fn skipped(&self) -> Result<JsValue, JsValue> {
        let skipped: Vec<_> = self
            .inner
            .skipped
            .iter()
            .map(|s| serde_json::json!({"sheet": s.sheet, "cell": s.cell, "reason": s.reason}))
            .collect();
        to_js(&skipped)
    }

    /// What the run did that is not what was asked for and is not an error: records
    /// `{level, code, message, context}`.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter)]
    pub fn diagnostics(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.diagnostics)
    }

    /// The answer of a `calculate` step, when the run was asked for it (`valueRanges()`).
    #[wasm_bindgen(getter, js_name = valueRanges)]
    pub fn value_ranges(&self) -> Option<ValueRanges> {
        self.inner
            .value_ranges
            .clone()
            .map(|inner| ValueRanges { inner })
    }

    /// `{replaced, written, skipped, diagnostics, valueRanges}` as an object.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        let object = js_sys::Object::new();
        for (name, value) in [
            ("replaced", JsValue::from_f64(self.replaced())),
            ("written", JsValue::from_f64(self.written())),
            ("skipped", self.skipped()?),
            ("diagnostics", self.diagnostics()?),
            (
                "valueRanges",
                self.inner
                    .value_ranges
                    .as_ref()
                    .map_or(Ok(JsValue::NULL), to_js)?,
            ),
        ] {
            js_sys::Reflect::set(&object, &JsValue::from_str(name), &value)?;
        }
        Ok(object.into())
    }
}

/// The value ranges of a calculation.
#[wasm_bindgen]
pub struct ValueRanges {
    inner: sheetrs_protocol::ValueRanges,
}

#[wasm_bindgen]
impl ValueRanges {
    /// The answer of the protocol, `{valueRanges: [{range, name, values}]}`, as an object.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner)
    }

    /// The values by name, `{name: rows}`: each range by its name, or by its range when it has
    /// none.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(js_name = toObject)]
    pub fn to_object(&self) -> Result<JsValue, JsValue> {
        let entries: serde_json::Map<String, Json> =
            self.inner.to_entries(None).into_iter().collect();
        to_js(&entries)
    }
}

#[wasm_bindgen]
impl Book {
    /// The plan as it runs: a copy of the `Book` whose `calculate` steps keep their answer when
    /// `all` says so or when they name the ranges they answer.
    fn answering(&self, all: bool) -> sheetrs_io::Book {
        let mut copy = self.inner.clone();
        copy.clear_operations();
        for operation in self.inner.operations().iter().cloned() {
            copy.push(match operation {
                Operation::Calculate(calculation) => {
                    let answer = all || calculation.ranges.is_some();
                    Operation::Calculate(Calculation {
                        answer,
                        ..calculation
                    })
                }
                other => other,
            });
        }
        copy
    }

    /// Delete sheets by name (*Delete Sheet*): a name or a list of them.
    ///
    /// # Errors
    ///
    /// A message when the argument is not a text or a list of texts.
    #[wasm_bindgen(js_name = deleteSheets)]
    pub fn delete_sheets(&mut self, names: JsValue) -> Result<(), JsValue> {
        let names = targets(names, "`names`")?;
        if !names.is_empty() {
            self.inner.delete_sheets(names);
        }
        Ok(())
    }

    /// Delete defined names (or database ranges): a name or a list of them.
    ///
    /// # Errors
    ///
    /// A message when the argument is not a text or a list of texts.
    #[wasm_bindgen(js_name = deleteNames)]
    pub fn delete_names(&mut self, names: JsValue) -> Result<(), JsValue> {
        let names = targets(names, "`names`")?;
        if !names.is_empty() {
            self.inner.delete_names(names);
        }
        Ok(())
    }

    /// Replace the formulas by the values they have (*Data ▸ Calculate ▸ Formula to Value*):
    /// `Sheet`, `Sheet!B:D` or `Sheet!3:10`, or a list of them.
    ///
    /// # Errors
    ///
    /// A message when a scope is not valid.
    #[wasm_bindgen(js_name = formulaToValue)]
    pub fn formula_to_value(&mut self, scopes: JsValue) -> Result<(), JsValue> {
        let scopes = parse_all::<ValueScope>(&targets(scopes, "`scopes`")?)?;
        if !scopes.is_empty() {
            self.inner.formula_to_value(scopes);
        }
        Ok(())
    }

    /// Delete rows, in the rows of the workbook as it is: `Sheet!3:10` or `Sheet!3`, or a list of
    /// them.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = deleteRows)]
    pub fn delete_rows(&mut self, rows: JsValue) -> Result<(), JsValue> {
        let rows = parse_all::<RowRange>(&targets(rows, "`rows`")?)?;
        if !rows.is_empty() {
            self.inner.delete_rows(rows);
        }
        Ok(())
    }

    /// Delete columns, in the columns of the workbook as it is: `Sheet!B:D` or `Sheet!B`, or a list
    /// of them.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = deleteColumns)]
    pub fn delete_columns(&mut self, columns: JsValue) -> Result<(), JsValue> {
        let columns = parse_all::<ColumnRange>(&targets(columns, "`columns`")?)?;
        if !columns.is_empty() {
            self.inner.delete_columns(columns);
        }
        Ok(())
    }

    /// Hide rows, in the rows of the workbook as it is: `Sheet!3:10` or `Sheet!3`, or a list of
    /// them. They stay in the file.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = hideRows)]
    pub fn hide_rows(&mut self, rows: JsValue) -> Result<(), JsValue> {
        let rows = parse_all::<RowRange>(&targets(rows, "`rows`")?)?;
        if !rows.is_empty() {
            self.inner.hide_rows(rows);
        }
        Ok(())
    }

    /// Show rows, whatever hid them.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = showRows)]
    pub fn show_rows(&mut self, rows: JsValue) -> Result<(), JsValue> {
        let rows = parse_all::<RowRange>(&targets(rows, "`rows`")?)?;
        if !rows.is_empty() {
            self.inner.show_rows(rows);
        }
        Ok(())
    }

    /// Hide columns: `Sheet!B:D` or `Sheet!B`, or a list of them.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = hideColumns)]
    pub fn hide_columns(&mut self, columns: JsValue) -> Result<(), JsValue> {
        let columns = parse_all::<ColumnRange>(&targets(columns, "`columns`")?)?;
        if !columns.is_empty() {
            self.inner.hide_columns(columns);
        }
        Ok(())
    }

    /// Show columns.
    ///
    /// # Errors
    ///
    /// A message when a range is not valid.
    #[wasm_bindgen(js_name = showColumns)]
    pub fn show_columns(&mut self, columns: JsValue) -> Result<(), JsValue> {
        let columns = parse_all::<ColumnRange>(&targets(columns, "`columns`")?)?;
        if !columns.is_empty() {
            self.inner.show_columns(columns);
        }
        Ok(())
    }

    /// Hide sheets by name (one sheet must stay visible): a name or a list of them.
    ///
    /// # Errors
    ///
    /// A message when the argument is not a text or a list of texts.
    #[wasm_bindgen(js_name = hideSheets)]
    pub fn hide_sheets(&mut self, names: JsValue) -> Result<(), JsValue> {
        let names = targets(names, "`names`")?;
        if !names.is_empty() {
            self.inner.hide_sheets(names);
        }
        Ok(())
    }

    /// Show sheets by name: a name or a list of them.
    ///
    /// # Errors
    ///
    /// A message when the argument is not a text or a list of texts.
    #[wasm_bindgen(js_name = showSheets)]
    pub fn show_sheets(&mut self, names: JsValue) -> Result<(), JsValue> {
        let names = targets(names, "`names`")?;
        if !names.is_empty() {
            self.inner.show_sheets(names);
        }
        Ok(())
    }

    /// Find and replace text, as `LibreOffice` does. `replaces` is a list of `[text, with]` pairs or
    /// an object `{text: with}`; each pair is a replacement of its own, in order, with the options
    /// of `options` (`mode`, `matchCase`, `entireCells`, `searchIn`, `formatted`, `filtered`,
    /// `scope`).
    ///
    /// # Errors
    ///
    /// A message when an argument is not valid.
    pub fn replace(&mut self, replaces: JsValue, options: JsValue) -> Result<(), JsValue> {
        let pairs = match from_js::<Pairs>(replaces, "The replacements")? {
            Pairs::List(list) => list,
            Pairs::Map(map) => map.into_iter().collect(),
        };
        let args: ReplaceArgs = if options.is_undefined() || options.is_null() {
            ReplaceArgs::default()
        } else {
            from_js_object(
                options,
                "The options",
                &[
                    "mode",
                    "matchCase",
                    "entireCells",
                    "searchIn",
                    "formatted",
                    "filtered",
                    "scope",
                ],
            )?
        };
        let scopes: Vec<_> = args
            .scope
            .map(Targets::into_vec)
            .unwrap_or_default()
            .iter()
            .map(|s| Scope::from(s.as_str()))
            .collect();
        let replaces: Vec<Replace> = pairs
            .into_iter()
            .map(|(text, with)| Replace {
                find: Find {
                    text,
                    mode: args.mode,
                    match_case: args.match_case,
                    entire_cells: args.entire_cells,
                    search_in: args.search_in,
                    include_formulas: true,
                    formatted: args.formatted,
                    filtered: args.filtered,
                },
                with,
            })
            .collect();
        if !replaces.is_empty() {
            self.inner.replace(scopes, replaces);
        }
        Ok(())
    }

    /// Calculate with the engine, which reproduces `LibreOffice` (*Data ▸ Calculate*): an operation
    /// like the others. `request` is `{inputs, ranges, majorDimension}`, the request of `sheetc
    /// calculate`: `inputs` are written into their cells first (`{"Data!B2": 1500}`), every formula is
    /// calculated, and `ranges` are the ranges `valueRanges()` answers (`majorDimension`: `"ROWS"` or
    /// `"COLUMNS"`); with none, the outputs the workbook declares or its formula cells. `save()` writes
    /// the workbook with fresh results. Where the engine cannot calculate it, it is written marked
    /// to be recalculated when it is opened, and a diagnostic says so.
    ///
    /// # Errors
    ///
    /// A message when the request is not valid.
    pub fn calculate(&mut self, request: JsValue) -> Result<(), JsValue> {
        let request: Request = if request.is_undefined() || request.is_null() {
            Request::default()
        } else {
            from_js_object(
                request,
                "The request",
                &["inputs", "ranges", "majorDimension"],
            )?
        };
        self.inner.calculate(Calculation {
            inputs: request.inputs.unwrap_or_default(),
            ranges: request.ranges,
            major: request.major_dimension.unwrap_or_default(),
            answer: false,
        });
        Ok(())
    }

    /// How many operations are pending.
    #[wasm_bindgen(getter)]
    pub fn pending(&self) -> usize {
        self.inner.operations().len()
    }

    /// What `save` will do: the passes over the file, one operation per line.
    pub fn explain(&self) -> String {
        self.inner.explain()
    }

    /// What the last `save` or `valueRanges` did besides writing the workbook; `undefined` before.
    #[wasm_bindgen(getter)]
    pub fn report(&mut self) -> Option<Report> {
        self.report.as_ref().map(|r| Report {
            inner: r.inner.clone(),
        })
    }

    /// The cut: runs the operations and returns the bytes of the workbook (a `Uint8Array`), in the
    /// format of the `Book`. The plan is not consumed, so the same operations can be saved again; to
    /// go on from the result, make a `Book` of it. With no operations it returns the bytes as they
    /// are.
    ///
    /// # Errors
    ///
    /// A message when an operation is not valid or addresses a sheet that does not exist, or the
    /// writer cannot make a change faithfully.
    pub fn save(&mut self) -> Result<JsValue, JsValue> {
        if self.inner.operations().is_empty() {
            let data = self
                .inner
                .bytes()
                .map(|b| bytes(b))
                .ok_or_else(|| JsValue::from_str("Edit error: the Book has no bytes"))?;
            return Ok(data);
        }
        let mut workbook = Vec::new();
        // A `calculate` that names its ranges is asked for an answer as well as for the workbook.
        let report = self
            .answering(false)
            .write_to(&mut workbook)
            .map_err(|e| failure("Edit error", format!("{e:#}")))?;
        self.emit(&report.diagnostics)?;
        self.report = Some(Report { inner: report });
        Ok(bytes(&workbook))
    }

    /// The value ranges of the calculation: the ranges `calculate` asked for, or the outputs of the
    /// workbook. It runs the operations (the workbook goes through memory and is written nowhere).
    ///
    /// # Errors
    ///
    /// A message when there is no `calculate` step or the engine cannot calculate the workbook.
    #[wasm_bindgen(js_name = valueRanges)]
    pub fn value_ranges(&mut self) -> Result<ValueRanges, JsValue> {
        if !self
            .inner
            .operations()
            .iter()
            .any(|o| matches!(o, Operation::Calculate(_)))
        {
            return Err(JsValue::from_str(
                "Edit error: there is no calculate step: valueRanges() gives the value ranges of a calculation, so call calculate() first",
            ));
        }
        let answer = self.answering(true);
        let report = answer
            .write_to(std::io::sink())
            .map_err(|e| failure("Edit error", format!("{e:#}")))?;
        self.emit(&report.diagnostics)?;
        let value_ranges = report
            .value_ranges
            .clone()
            .ok_or_else(|| JsValue::from_str("Edit error: the calculation gave no value ranges"))?;
        self.report = Some(Report { inner: report });
        Ok(ValueRanges {
            inner: value_ranges,
        })
    }
}
