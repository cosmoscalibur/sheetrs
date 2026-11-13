//! The engine, as the plan uses it: a `calculate` step, and the recalculation that precedes a
//! conversion of formulas to values on stale results.
//!
//! Both are the same thing, the compiler's program for the workbook run in memory (the semantics of
//! `LibreOffice`): [`recalculate`] is a calculation with nothing written and nothing asked for.

use serde_json::{Map, Value as Json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::ir::Mode;
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

/// Every formula of the workbook calculated again, as it is: the bytes of the workbook with fresh
/// cached results, in its own format. The error is why that cannot be done.
pub(crate) fn recalculate(source: &Book) -> Result<Vec<u8>, String> {
    calculate(source, &Calculation::default())
        .map(|calculated| calculated.workbook)
        .map_err(|error| error.to_string())
}
