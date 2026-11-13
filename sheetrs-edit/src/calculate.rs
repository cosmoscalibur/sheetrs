//! The engine, as the plan uses it: a `calculate` step, and the recalculation that precedes a
//! conversion of formulas to values on stale results.
//!
//! Both are the same thing, the compiler's program for the workbook run in memory (the semantics of
//! `LibreOffice`): [`recalculate`] is a calculation with nothing written and nothing asked for.

use serde_json::{Map, Value as Json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::ir::{CellRun, Mode};
use sheetrs_io::Book;
use sheetrs_io::write::Calculation;
use sheetrs_protocol::ValueRanges;
use std::fmt;

/// What the engine gives back for a [`Calculation`].
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct Calculated {
    /// The bytes of the workbook with the inputs written and every formula calculated, in its own
    /// format.
    pub workbook: Vec<u8>,
    /// The value ranges of the calculation, when the [`Calculation`] asked for them.
    pub answer: Option<ValueRanges>,
}

/// Why a [`Calculation`] did not give a result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum CalculateError {
    /// The engine cannot calculate this workbook (a function it does not have, a reference it does
    /// not read): a spreadsheet application can, so the plan goes on and marks the workbook to be
    /// recalculated when it is opened, and says so in a diagnostic ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)).
    Unsupported(String),
    /// The calculation was asked for something that is not valid (an input that is no cell, a range
    /// that does not exist): an error of the person who asked, not of the engine.
    Failed(String),
}

impl fmt::Display for CalculateError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Unsupported(reason) | Self::Failed(reason) => f.write_str(reason),
        }
    }
}

/// The inputs written and every formula calculated by the engine, the workbook written back in its
/// format and, when the answer is asked for, the values of the ranges. A workbook the compiler
/// cannot compile (a function the engine does not have, a reference it does not read) is
/// [`Unsupported`](CalculateError::Unsupported); a request that is not valid for the workbook (an
/// input that is no cell) is [`Failed`](CalculateError::Failed).
pub(crate) fn calculate(
    source: &Book,
    calculation: &Calculation,
) -> Result<Calculated, CalculateError> {
    let options = Options::default();
    // Inputs, ranges or an answer are the embedded mode (the workbook is its own template, inputs
    // and outputs declared or deduced); with none of them the workbook is calculated as it is.
    let embedded =
        !calculation.inputs.is_empty() || calculation.ranges.is_some() || calculation.answer;
    let mode = if embedded {
        Mode::Embedded
    } else {
        Mode::Recalculate
    };
    let mut request = Map::new();
    if !calculation.inputs.is_empty() {
        request.insert("inputs".into(), Json::Object(calculation.inputs.clone()));
    }
    if let Some(ranges) = &calculation.ranges {
        request.insert("ranges".into(), Json::from(ranges.clone()));
    }
    let request = if embedded {
        Json::Object(request)
    } else {
        Json::Null
    };
    // Whether the workbook can be calculated at all is told by building its program; what the
    // request asks of it is told by running it.
    let (program, _) = source
        .compile(&options, mode)
        .map_err(|e| CalculateError::Unsupported(e.to_string()))?;
    // A workbook with no formula and no input has nothing to calculate: it is as it is, which is not
    // the engine failing.
    if program.formula_cells.is_empty() && calculation.inputs.is_empty() && !calculation.answer {
        let workbook = match (source.bytes(), source.path()) {
            (Some(data), _) => data.to_vec(),
            (None, Some(path)) => std::fs::read(path)
                .map_err(|e| CalculateError::Failed(format!("cannot read the workbook: {e}")))?,
            (None, None) => unreachable!("a book is a file or bytes"),
        };
        return Ok(Calculated {
            workbook,
            answer: None,
        });
    }
    let overrides = options.overrides();
    let failed = |e: sheetrs_answer::Error| CalculateError::Failed(e.to_string());
    let mut workbook = Vec::new();
    let answer = if calculation.answer {
        let (_, answer) = sheetrs_answer::fill_template_to_with_answer(
            source,
            &program,
            None,
            &request,
            &overrides,
            &mut workbook,
        )
        .map_err(failed)?;
        Some(answer.with_major(calculation.major))
    } else {
        sheetrs_answer::fill_template_to(
            source,
            &program,
            None,
            &request,
            &overrides,
            &mut workbook,
        )
        .map_err(failed)?;
        None
    };
    Ok(Calculated { workbook, answer })
}

/// The most cells the recalculation declares as the outputs of a pruned program.
const DECLARED_LIMIT: usize = 20_000;

/// The cells as the declaration of a program that has no inputs and has them as its outputs, in
/// runs of one row (`'Sheet'!B3:D3`).
fn declared_outputs(cells: &[(String, u32, u32)]) -> Option<String> {
    let mut sorted: Vec<&(String, u32, u32)> = cells.iter().collect();
    sorted.sort_unstable();
    sorted.dedup();
    let mut ranges: Vec<String> = Vec::new();
    let mut i = 0;
    while i < sorted.len() {
        let (sheet, row, first) = sorted[i];
        let mut last = *first;
        let mut j = i + 1;
        while j < sorted.len()
            && &sorted[j].0 == sheet
            && sorted[j].1 == *row
            && sorted[j].2 == last + 1
        {
            last += 1;
            j += 1;
        }
        let quoted = format!("'{}'", sheet.replace('\'', "''"));
        let start = sheetrs_formula::addr::a1(*row, *first);
        ranges.push(if last == *first {
            format!("{quoted}!{start}")
        } else {
            format!("{quoted}!{start}:{}", sheetrs_formula::addr::a1(*row, last))
        });
        i = j;
    }
    let outputs: Vec<Json> = ranges.into_iter().map(Json::String).collect();
    let mut declared = Map::new();
    declared.insert("inputs".into(), Json::Array(Vec::new()));
    declared.insert("outputs".into(), Json::Array(outputs));
    serde_json::to_string(&Json::Object(declared)).ok()
}

/// Every formula of the workbook calculated again, as it is: the bytes of the workbook with fresh
/// cached results, in its own format. With `cells` (`(sheet, row, column)`, 0-based, in the
/// workbook), only those formulas are written with their new result and the others keep the one
/// the file stores. The error is why that cannot be done.
pub(crate) fn recalculate(
    source: &Book,
    cells: Option<&[(String, u32, u32)]>,
) -> Result<Vec<u8>, String> {
    let Some(cells) = cells else {
        return calculate(source, &Calculation::default())
            .map(|calculated| calculated.workbook)
            .map_err(|error| error.to_string());
    };
    // The cells to calculate are the outputs of the program, declared: the compiler keeps only
    // what they need, so the engine does not evaluate the formulas that nothing stale reads. A long
    // list of cells is not worth declaring (the declaration is as heavy as the pass it saves): the
    // whole workbook is calculated and only those cells are written.
    if cells.len() <= DECLARED_LIMIT
        && let Some(declared) = declared_outputs(cells)
    {
        let options = Options {
            declarations: Some(declared),
            ..Options::default()
        };
        if let Ok((program, _)) = source.compile(&options, Mode::Embedded) {
            let mut workbook = Vec::new();
            sheetrs_answer::fill_template_to(
                source,
                &program,
                None,
                &Json::Null,
                &options.overrides(),
                &mut workbook,
            )
            .map_err(|e| e.to_string())?;
            return Ok(workbook);
        }
    }
    let options = Options::default();
    let (mut program, _) = source
        .compile(&options, Mode::Recalculate)
        .map_err(|e| e.to_string())?;
    let mut at: Vec<CellRef> = cells
        .iter()
        .filter_map(|(sheet, row, column)| {
            let index = program.sheets.iter().position(|name| name == sheet)?;
            Some(CellRef::new(u16::try_from(index).ok()?, *row, *column))
        })
        .collect();
    at.sort_unstable();
    at.dedup();
    // The formulas the compiler lists are the ones the engine writes: a cell asked for that is not
    // one of them (a formula the compiler could not take) keeps what the file holds.
    let listed = |cell: &CellRef| {
        let runs = &program.formula_cells;
        let i =
            runs.partition_point(|r| (r.sheet, r.row, r.last) < (cell.sheet, cell.row, cell.col));
        runs.get(i).is_some_and(|r| {
            r.sheet == cell.sheet && r.row == cell.row && (r.first..=r.last).contains(&cell.col)
        })
    };
    at.retain(listed);
    program.formula_cells = CellRun::runs(at);
    let mut workbook = Vec::new();
    sheetrs_answer::fill_template_to(
        source,
        &program,
        None,
        &Json::Null,
        &options.overrides(),
        &mut workbook,
    )
    .map_err(|e| e.to_string())?;
    Ok(workbook)
}
