//! A workbook as the template of an answer.
//!
//! [`fill_template`] runs a compiled program and writes what it computed into a copy of the workbook:
//! the formulas, the styles and everything else of the template stay as they are, the cached result
//! of every formula inside the outputs of the program is the one it computed (the other formulas
//! keep the result the template stored), and the inputs that were given are written in the cells
//! they belong to. What comes out is the template with the numbers of this run, in the
//! formats of its cells, and it is the whole answer: no range is returned apart from it.
//!
//! The template is the workbook the program was made from: the compiler embeds it in the executable
//! (`sheetc build --embed-template`, or the workbook that `sheetc calculate -o` writes) and offers no other. The workbook is
//! its own template, so nothing about it is discovered from the file: the program knows where every
//! formula to write is (`Program::formula_cells`, filled by the compiler in the workbook modes with the
//! formula cells inside the outputs, whatever became of the formula: lowered, folded, kept as the value
//! the file cached, or inside an array formula) and the inputs it has, and the file is read once, by the writer
//! that rewrites the cells. A program built in a mode that prunes or gives values would leave results
//! as the file had them, and an answer that mixes the two is not one: `fill_template` refuses it.
//!
//! Every formula cell of the outputs is written, the cells of an array formula included.

use crate::Written;
use crate::error::{Error, Result};
use serde_json::{Map, Value as Json};
use sheetrs_engine::Overrides;
use sheetrs_engine::addr::Rect;
use sheetrs_engine::calculate::ValueRanges;
use sheetrs_engine::engine::ConstSource;
use sheetrs_engine::ir::Program;
use sheetrs_io::Book;
use sheetrs_io::reader::Format;
use sheetrs_io::write::{CachedResult, CellWrites, Constant, WorkbookModifications};
use std::path::Path;

/// An error of the program written as the format of the file has it: Excel's names in an XLSX, and
/// `LibreOffice`'s own codes in an ODS (`Err:502`, `Err:523`, the circular reference `#CIRC!` as
/// `Err:522`); an XLSX holds what has no name in Excel as `#VALUE!`, as `LibreOffice` writes it.
/// `libre` is the code the program gave next to the name, when it has one.
fn error_text(name: &str, libre: Option<&str>, format: Format) -> String {
    const EXCEL: [&str; 7] = [
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    ];
    match (name, libre, format) {
        ("#CIRC!", _, Format::Ods) => "Err:522".to_string(),
        (_, Some(code), Format::Ods) => code.to_string(),
        // The error `LibreOffice` writes in an XLSX for its `Err:5xx` (`Err:520` is `#N/A`).
        (_, Some(code), _)
            if code
                .strip_prefix("Err:")
                .is_some_and(|n| n.parse::<u16>().is_ok()) =>
        {
            let number = code["Err:".len()..].parse::<u16>().unwrap_or(0);
            sheetrs_io::write::xlsx_error(number).to_string()
        }
        (name, _, _) if EXCEL.contains(&name) => name.to_string(),
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
            m.get("code").and_then(Json::as_str),
            format,
        )),
        Json::Array(_) => CachedResult::Error(error_text("#VALUE!", None, format)),
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

/// The cells an input of a request sets and the value of each, as the program takes them: one value
/// for a single cell, rows of values (or a flat list for a row or a column) for a range.
fn input_cells(rect: &sheetrs_engine::addr::Rect, given: &Json) -> Vec<(u16, u32, u32, Json)> {
    let (rows, cols) = (rect.rows() as usize, rect.cols() as usize);
    let cell = |i: usize, k: usize, v: &Json| {
        (
            rect.sheet,
            rect.r0 + u32::try_from(i).unwrap_or(0),
            rect.c0 + u32::try_from(k).unwrap_or(0),
            v.clone(),
        )
    };
    if rows * cols == 1 {
        let inner = match given {
            Json::Array(a) if a.len() == 1 => match &a[0] {
                Json::Array(b) if b.len() == 1 => &b[0],
                x => x,
            },
            x => x,
        };
        return vec![cell(0, 0, inner)];
    }
    let Json::Array(items) = given else {
        return Vec::new();
    };
    let mut out = Vec::with_capacity(rows * cols);
    if !items.is_empty() && items.iter().all(Json::is_array) {
        for (i, row) in items.iter().enumerate() {
            if let Json::Array(row) = row {
                out.extend(row.iter().enumerate().map(|(k, v)| cell(i, k, v)));
            }
        }
    } else if rows == 1 || cols == 1 {
        for (i, v) in items.iter().enumerate() {
            out.push(if rows == 1 {
                cell(0, i, v)
            } else {
                cell(i, 0, v)
            });
        }
    }
    out
}

/// Runs `program` with the inputs of `request` (`{"inputs": {...}}`) and saves to `output` a copy of
/// `template` that holds its results and its inputs.
///
/// Returns how many cells were written ([`Written`]).
///
/// # Errors
///
/// When the program was not built in a workbook mode, the request holds more than inputs, the run
/// fails, the program gave no value for a formula of the workbook, or the copy cannot be written.
pub fn fill_template(
    template: &Book,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    output: &Path,
) -> Result<Written> {
    let filled = fill(template, program, base, request, overrides, false)?;
    write_cells(filled.writes, template, output)
}

/// [`fill_template`] that also gives the answer of the program to the same request (its
/// [`ValueRanges`]: the `ranges` the request asks for, or the declared outputs), from the same
/// calculation. The request is then the whole of the protocol, not only `inputs`.
///
/// # Errors
///
/// As [`fill_template`], and when the request is not a valid one for the program.
pub fn fill_template_with_answer(
    template: &Book,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    output: &Path,
) -> Result<(Written, ValueRanges)> {
    let filled = fill(template, program, base, request, overrides, true)?;
    let report = write_cells(filled.writes, template, output)?;
    Ok((report, answer_of(filled.answer)))
}

/// [`fill_template`] that writes the bytes of the workbook to `out`, in memory (no file is written;
/// `out` is a `Vec<u8>` where there is no file system).
///
/// # Errors
///
/// As [`fill_template`], without the errors of writing a file, and when `out` fails.
pub fn fill_template_to(
    template: &Book,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    out: impl std::io::Write,
) -> Result<Written> {
    let filled = fill(template, program, base, request, overrides, false)?;
    write_cells_to(filled.writes, template, out)
}

/// [`fill_template_with_answer`] that writes the bytes of the workbook to `out`, in memory.
///
/// # Errors
///
/// As [`fill_template_to`], and when the request is not a valid one for the program.
pub fn fill_template_to_with_answer(
    template: &Book,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    out: impl std::io::Write,
) -> Result<(Written, ValueRanges)> {
    let filled = fill(template, program, base, request, overrides, true)?;
    let report = write_cells_to(filled.writes, template, out)?;
    Ok((report, answer_of(filled.answer)))
}

/// The answer of a run that was asked for it (no value ranges when it was not).
fn answer_of(answer: Option<ValueRanges>) -> ValueRanges {
    answer.unwrap_or_default()
}

/// What a template is filled with, and the JSON answer of the same calculation when it was asked.
struct Filled {
    writes: CellWrites,
    answer: Option<ValueRanges>,
}

/// What the writer is asked for the cells of a template: the replacements, and nothing else (the
/// results written are taken as current, so the workbook is not marked to be recalculated).
fn modifications(writes: CellWrites) -> WorkbookModifications {
    WorkbookModifications {
        replacements: writes.into_sheets(),
        ..WorkbookModifications::default()
    }
}

fn write_cells(writes: CellWrites, template: &Book, output: &Path) -> Result<Written> {
    let cells = writes.len();
    template.modify(output, &modifications(writes))?;
    Ok(Written { cells })
}

fn write_cells_to(
    writes: CellWrites,
    template: &Book,
    mut out: impl std::io::Write,
) -> Result<Written> {
    let cells = writes.len();
    let bytes = template.modify_to_bytes(&modifications(writes))?;
    out.write_all(&bytes)
        .map_err(sheetrs_io::write::Error::from)?;
    Ok(Written { cells })
}

fn fill(
    template: &Book,
    program: &Program,
    base: Option<&dyn ConstSource>,
    request: &Json,
    overrides: &Overrides,
    with_answer: bool,
) -> Result<Filled> {
    let fail = |m: String| Error::Template(m);
    if !program.mode.is_workbook() {
        return Err(fail(
            "the program gives values, not a workbook: build it with --embed-template to write the answer into its workbook"
                .into(),
        ));
    }
    let inputs = match request {
        Json::Null => Map::new(),
        Json::Object(m) if with_answer || m.keys().all(|k| k == "inputs") => {
            match m.get("inputs") {
                None | Some(Json::Null) => Map::new(),
                Some(Json::Object(i)) => i.clone(),
                Some(_) => return Err(fail("`inputs` must be a JSON object".into())),
            }
        }
        _ => {
            return Err(fail(
                "the request for a template holds only `inputs` (the answer is the workbook; ask for the values as well to give more)"
                    .into(),
            ));
        }
    };

    if program.mode == sheetrs_engine::ir::Mode::Recalculate && !inputs.is_empty() {
        return Err(fail(
            "the editor's recalculation takes no input: calculate with inputs and ranges instead"
                .into(),
        ));
    }

    let format = template.format();

    // Where the values of the formulas are read: the outputs of the program (declared, or the
    // formula cells when nothing is declared). Every formula to write lies inside one of them.
    let rects: Vec<Rect> = program.outputs.iter().map(|b| b.rect).collect();
    // The inputs given, written back as given (an input is a constant of the template, not an
    // output of the program, so it is not read from the answer).
    let mut written_inputs: Vec<(u16, u32, u32, Json)> = Vec::new();
    for (key, given) in &inputs {
        // A key is an input, an alias (a defined name) or a range of the workbook, as in a request.
        let binding = program.input_range(key).map_err(|e| fail(e.to_string()))?;
        written_inputs.extend(input_cells(&binding.rect, given));
    }
    if program.formula_cells.is_empty() && written_inputs.is_empty() {
        return Err(fail("there is no formula or input to write".into()));
    }
    let (answer, matrices) = if with_answer {
        let (answer, matrices) = program
            .handle_with(base, request, overrides, &rects)
            .map_err(|e| fail(e.to_string()))?;
        (Some(answer), matrices)
    } else {
        let matrices = program
            .read_rects(base, &Json::Object(inputs), overrides, &rects)
            .map_err(|e| fail(e.to_string()))?;
        (None, matrices)
    };

    let mut writes = CellWrites::new();
    // The runs come by sheet, row and column; the rectangles that cover the row of a run are the ones
    // that start in or before it and have not ended.
    let mut by_start: Vec<usize> = (0..rects.len()).collect();
    by_start.sort_by_key(|&i| (rects[i].sheet, rects[i].r0));
    let mut next = 0;
    let mut open: Vec<usize> = Vec::new();
    let mut at_sheet = None;
    for run in &program.formula_cells {
        if at_sheet != Some(run.sheet) {
            at_sheet = Some(run.sheet);
            open.clear();
            while next < by_start.len() && rects[by_start[next]].sheet < run.sheet {
                next += 1;
            }
        }
        while next < by_start.len()
            && rects[by_start[next]].sheet == run.sheet
            && rects[by_start[next]].r0 <= run.row
        {
            open.push(by_start[next]);
            next += 1;
        }
        open.retain(|&i| rects[i].r1 >= run.row);
        let covering = open.iter().copied().find(|&i| {
            let r = &rects[i];
            r.c0 <= run.first && run.last <= r.c1
        });
        let name = &program.sheets[usize::from(run.sheet)];
        for col in run.first..=run.last {
            let v = covering
                .and_then(|i| {
                    let r = &rects[i];
                    matrices
                        .get(i)?
                        .as_array()?
                        .get((run.row - r.r0) as usize)?
                        .as_array()?
                        .get((col - r.c0) as usize)
                })
                .ok_or_else(|| {
                    fail("the program gave no value for a formula of the workbook".into())
                })?;
            writes.result(name.as_str(), run.row, col, result_of(v, format));
        }
    }
    for (sheet, row, col, v) in written_inputs {
        writes.constant(
            program.sheets[usize::from(sheet)].as_str(),
            row,
            col,
            constant_of(&v),
        );
    }
    Ok(Filled { writes, answer })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_error_is_written_as_the_format_can_hold_it() {
        // An ODS keeps the code of LibreOffice; an XLSX has none and holds `#VALUE!`.
        assert_eq!(
            error_text("#VALUE!", Some("Err:502"), Format::Ods),
            "Err:502"
        );
        assert_eq!(
            error_text("#VALUE!", Some("Err:502"), Format::Xlsx),
            "#VALUE!"
        );
        assert_eq!(error_text("#CIRC!", None, Format::Ods), "Err:522");
        assert_eq!(error_text("#CIRC!", None, Format::Xlsx), "#VALUE!");
        assert_eq!(error_text("#DIV/0!", None, Format::Ods), "#DIV/0!");
        assert_eq!(error_text("#N/A", None, Format::Xlsx), "#N/A");
    }
}
