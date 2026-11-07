//! A workbook as the template of an answer.
//!
//! [`fill_template`] runs a compiled program and writes what it computed into a copy of the workbook:
//! the formulas, the styles and everything else of the template stay as they are, the cached result
//! of every formula is the one the program computed, and the inputs that were given are written in
//! the cells they belong to. What comes out is the template with the numbers of this run, in the
//! formats of its cells, and it is the whole answer: no range is returned apart from it.
//!
//! The template is the workbook the program was made from: the compiler embeds it in the executable
//! (`sheetc build --embed-template`) and offers no other. The cells are addressed by position, so a
//! program written into a workbook of another layout would put results where they do not belong, and
//! nothing could check that. The program must have been built with `outputs = any`: only then does it
//! compute every formula of the workbook; one that prunes would leave the results it did not compute
//! as the template had them, and an answer that mixes the two is not one.
//!
//! The cells of an array formula other than its first keep the results the template cached.

use crate::Editor;
use crate::error::{Error, Result};
use crate::replace::Report;
use crate::writer::CachedResult;
use crate::writes::{CellWrites, Constant};
use serde_json::{Map, Value as Json};
use sheetrs_engine::Overrides;
use sheetrs_engine::engine::ConstSource;
use sheetrs_engine::ir::{Program, Queryable};
use sheetrs_formula::addr::{a1, same_name};
use sheetrs_formula::print::sheet_text;
use sheetrs_io::reader::{Cell, CellSink, Format, Properties, Selection, Source, Window};
use std::collections::BTreeMap;
use std::path::Path;

/// The formula cells of the template, by sheet index.
#[derive(Default)]
struct Formulas {
    cells: Vec<(u16, u32, u32)>,
}

impl CellSink for Formulas {
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        if cell.formula.is_some() {
            self.cells.push((sheet, at.0, at.1));
        }
    }
}

/// An error of the program written as the format of the file has it: Excel's names in an XLSX, and
/// `LibreOffice`'s own codes in an ODS (`#CIRC!` is `Err:522`; what has no name in Excel is
/// `#VALUE!` in an XLSX, as `LibreOffice` writes it).
fn error_text(code: &str, format: Format) -> String {
    const EXCEL: [&str; 7] = [
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    ];
    if EXCEL.contains(&code) {
        return code.to_string();
    }
    match (code, format) {
        ("#CIRC!", Format::Ods) => "Err:522".to_string(),
        _ => "#VALUE!".to_string(),
    }
}

fn result_of(value: &Json, format: Format) -> CachedResult {
    match value {
        Json::Null => CachedResult::Text(String::new()),
        Json::Bool(b) => CachedResult::Boolean(*b),
        Json::Number(n) => CachedResult::Number(n.as_f64().unwrap_or(f64::NAN)),
        Json::String(s) => CachedResult::Text(s.clone()),
        Json::Object(m) => CachedResult::Error(error_text(
            m.get("error").and_then(Json::as_str).unwrap_or("#VALUE!"),
            format,
        )),
        Json::Array(_) => CachedResult::Error(error_text("#VALUE!", format)),
    }
}

fn constant_of(value: &Json) -> Constant {
    match value {
        Json::Bool(b) => Constant::Boolean(*b),
        Json::Number(n) => Constant::Number(n.as_f64().unwrap_or(f64::NAN)),
        Json::String(s) => Constant::Text(s.clone()),
        _ => Constant::Empty,
    }
}

/// Runs `program` with the inputs of `request` (`{"inputs": {...}}`) and saves to `output` a copy of
/// `template` that holds its results and its inputs.
///
/// Returns the report of the write: how many cells were written.
///
/// # Errors
///
/// When the program was not built with `outputs = any`, the request holds more than inputs, the run
/// fails, or the template cannot be read or the copy written.
pub fn fill_template(
    template: &Source,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    output: &Path,
) -> Result<Report> {
    let fail = |m: String| Error::Template(m);
    if !matches!(program.queryable, Queryable::Any { .. }) {
        return Err(fail(
            "the program computes only part of the workbook: build it with outputs = any to write the answer into a template"
                .into(),
        ));
    }
    let inputs = match request {
        Json::Null => Map::new(),
        Json::Object(m) if m.keys().all(|k| k == "inputs") => match m.get("inputs") {
            None | Some(Json::Null) => Map::new(),
            Some(Json::Object(i)) => i.clone(),
            Some(_) => return Err(fail("`inputs` must be a JSON object".into())),
        },
        _ => {
            return Err(fail(
                "the request for a template holds only `inputs` (the answer is the workbook)"
                    .into(),
            ));
        }
    };

    // What the template has: the sheets and the cells that hold a formula.
    let mut formulas = Formulas::default();
    let selection = Selection::new(Properties {
        value: false,
        formula: true,
        number_format: false,
    })
    .window_everywhere(Window::ALL);
    let sheets = template.read_selection_into(&selection, &mut formulas)?;
    let format = template.format();

    // What the program answers for them: the bounding box of the cells to write in each sheet.
    let program_sheet = |name: &str| program.sheets.iter().position(|s| same_name(s, name));
    let mut boxes: BTreeMap<String, (u32, u32, u32, u32)> = BTreeMap::new();
    let mut grow = |name: &str, r0: u32, c0: u32, r1: u32, c1: u32| {
        boxes
            .entry(name.to_string())
            .and_modify(|b| {
                *b = (b.0.min(r0), b.1.min(c0), b.2.max(r1), b.3.max(c1));
            })
            .or_insert((r0, c0, r1, c1));
    };
    let mut written_formulas: Vec<(String, u32, u32)> = Vec::new();
    for &(index, row, col) in &formulas.cells {
        let Some(info) = sheets.iter().find(|s| s.index == index) else {
            continue;
        };
        if program_sheet(&info.name).is_none() {
            continue;
        }
        grow(&info.name, row, col, row, col);
        written_formulas.push((info.name.clone(), row, col));
    }
    let mut written_inputs: Vec<(String, u32, u32)> = Vec::new();
    for binding in program
        .inputs
        .iter()
        .filter(|b| inputs.contains_key(&b.name))
    {
        let name = &program.sheets[usize::from(binding.rect.sheet)];
        let r = binding.rect;
        grow(name, r.r0, r.c0, r.r1, r.c1);
        for row in r.r0..=r.r1 {
            for col in r.c0..=r.c1 {
                written_inputs.push((name.clone(), row, col));
            }
        }
    }
    if boxes.is_empty() {
        return Err(fail("there is no formula or input to write".into()));
    }
    let names: Vec<&String> = boxes.keys().collect();
    let ranges: Vec<String> = boxes
        .iter()
        .map(|(name, (r0, c0, r1, c1))| {
            format!("{}!{}:{}", sheet_text(name), a1(*r0, *c0), a1(*r1, *c1))
        })
        .collect();
    let mut asked = Map::new();
    asked.insert("inputs".into(), Json::Object(inputs));
    asked.insert(
        "ranges".into(),
        Json::Array(ranges.into_iter().map(Json::String).collect()),
    );
    let answer = program
        .handle_over(base, &Json::Object(asked), overrides)
        .map_err(|e| fail(e.to_string()))?;
    let value_ranges = answer
        .get("valueRanges")
        .and_then(Json::as_array)
        .ok_or_else(|| fail("the program gave no ranges".into()))?;

    let at = |name: &str, row: u32, col: u32| -> Option<&Json> {
        let i = names.iter().position(|n| n.as_str() == name)?;
        let (r0, c0, _, _) = boxes[name];
        value_ranges
            .get(i)?
            .get("values")?
            .as_array()?
            .get((row - r0) as usize)?
            .as_array()?
            .get((col - c0) as usize)
    };
    let mut writes = CellWrites::new();
    for (name, row, col) in written_formulas {
        if let Some(v) = at(&name, row, col) {
            writes.result(name, row, col, result_of(v, format));
        }
    }
    for (name, row, col) in written_inputs {
        if let Some(v) = at(&name, row, col) {
            writes.constant(name, row, col, constant_of(v));
        }
    }
    Editor::new(template).write_cells(writes).save(output)
}
