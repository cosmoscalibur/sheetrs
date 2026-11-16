//! The engine, as the plan uses it: a `calculate` step, and the recalculation that precedes a
//! conversion of formulas to values on stale results.
//!
//! Both are the same thing, the compiler's program for the workbook run in memory (the semantics of
//! `LibreOffice`): [`recalculate`] is a calculation with nothing written and nothing asked for.

use serde_json::{Map, Value as Json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_compiler::{Cache, FromCacheNote, Only};
use sheetrs_engine::ir::Mode;
use sheetrs_io::write::{Calculation, ValueScope, ValueTarget, WorkbookModifications};
use sheetrs_io::{Book, Volatile};
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
    /// What the engine left as the file caches it (see [`FromCacheNote`]).
    pub from_cache: Vec<FromCacheNote>,
}

/// Why a [`Calculation`] did not give a result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum CalculateError {
    /// The engine cannot calculate this workbook (a function it does not have, a reference it does
    /// not read): a spreadsheet application can, so the plan goes on and marks the workbook to be
    /// recalculated when it is opened, and says so in a diagnostic ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)).
    Unsupported(String),
    /// The workbook depends on something the engine cannot calculate (a macro, an extension) in a cell the run
    /// calculates, and only `LibreOffice` can give its result: the error names the cause, and the plan stops, as
    /// it has not the right to mark the workbook to be recalculated instead.
    Depends(String),
    /// The workbook carries macros and secure mode is off: the build refuses it (`sheetrs_compiler::Error::Macros`), in every
    /// layer. The plan stops: it is not the engine that cannot, and `LibreOffice` is not left to do it.
    Refused(String),
    /// The calculation was asked for something that is not valid (an input that is no cell, a range
    /// that does not exist): an error of the person who asked, not of the engine.
    Failed(String),
}

impl fmt::Display for CalculateError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Unsupported(reason)
            | Self::Depends(reason)
            | Self::Refused(reason)
            | Self::Failed(reason) => f.write_str(reason),
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
    cache: Option<&Cache>,
) -> Result<Calculated, CalculateError> {
    // The clock and the seed reach the engine as the options of the build and of the run, as the
    // `--now` and `--seed` of `sheetc` do: the step's own, else the ones given to the book.
    let mut options = Options::default().with_book(source);
    // The macros of a workbook are never run: the build refuses a workbook that carries them unless secure mode is on, and then
    // what depends on them keeps the cache until something dirties it (and is an error cell when the run calculates it).
    options.config.now = calculation.now.or(options.config.now);
    options.config.seed = calculation.seed.or(options.config.seed);
    // No cone of changes means every formula is calculated (the run asks for it, the file says its
    // results are stale, or the changes cannot be placed).
    match cache {
        Some(cache) => options.cache = cache.clone(),
        None => options.config.calculate_all = Some(true),
    }
    // The inputs given are the declared inputs of the build (unless the workbook declares its own): what they
    // do not reach is folded as the result the workbook stores.
    options.request_inputs = calculation.inputs.keys().cloned().collect();
    // What it declares of the outputs: the workbook is the answer (every formula is calculated and the inputs are
    // all it declares), or the ranges it asks are the outputs of the build (see `Options::request_outputs`).
    options.request_outputs = Some(calculation.ranges.clone().unwrap_or_default());
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
    let (program, built) = source.compile(&options, mode).map_err(|e| cannot(&e))?;
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
            from_cache: built.from_cache,
        });
    }
    // The workbook is its own template here (the embedded mode). A range outside the outputs it
    // declares is an error (`OutsideOutputs`): it is not an answer of `calculate`, it is read from the
    // workbook that is written, with `query`.
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
    Ok(Calculated {
        workbook,
        answer,
        from_cache: built.from_cache,
    })
}

/// Every formula of the workbook calculated again, as it is: the bytes of the workbook with fresh
/// cached results, in its own format, with the clock and the seed `volatile` fixes. With a `cache`, only
/// the formulas it names and the ones that depend on them: the others keep the result the workbook
/// caches. The error is why that cannot be done.
pub(crate) fn recalculate(
    source: &Book,
    volatile: Volatile,
    cache: Option<&Cache>,
) -> Result<Calculated, CalculateError> {
    let calculation = Calculation {
        now: volatile.now,
        seed: volatile.seed,
        ..Calculation::default()
    };
    calculate(source, &calculation, cache)
}

/// What the error of the compiler says about a workbook the engine cannot calculate: a calculated cell that depends
/// on a macro or an extension is [`Depends`](CalculateError::Depends) (the run stops), a workbook with macros built with secure
/// mode off is [`Refused`](CalculateError::Refused) (the run stops), any other cause is
/// [`Unsupported`](CalculateError::Unsupported).
pub(crate) fn cannot(error: &sheetrs_compiler::Error) -> CalculateError {
    match error {
        sheetrs_compiler::Error::Macros(_) => CalculateError::Refused(error.to_string()),
        sheetrs_compiler::Error::Unsupported { .. } => CalculateError::Depends(error.to_string()),
        _ => CalculateError::Unsupported(error.to_string()),
    }
}

/// Why a conversion with the engine did not give a workbook.
#[derive(Debug)]
pub(crate) enum ConvertError {
    /// The engine cannot calculate what the conversion asks for (a function it does not have, a
    /// reference it does not read).
    Engine(String),
    /// The workbook depends on something the engine cannot calculate in a cell the run calculates: see
    /// [`CalculateError::Depends`].
    Depends(String),
    /// The workbook carries macros and secure mode is off: see [`CalculateError::Refused`].
    Refused(String),
    /// The writer of the workbook failed (a sheet that is not there, a deletion that cannot be made).
    Write(sheetrs_io::write::Error),
    /// The cells could not be written.
    Other(String),
}

impl From<sheetrs_io::write::Error> for ConvertError {
    fn from(error: sheetrs_io::write::Error) -> Self {
        Self::Write(error)
    }
}

/// A pass over the file that converts formulas to values with the engine, in the one step: the cells
/// the conversion asks for are what the engine is asked for (it calculates them and what they read, and
/// nothing else of the workbook), and the writer, as it applies `pass` (the conversions and the
/// deletions, against the workbook as it is), writes each of them as the value the engine calculated
/// and not the one the file cached. The bytes of the workbook that comes out, in its own format.
///
/// A workbook with no formula inside the cells asked for has nothing to calculate: `pass` is applied
/// as it is. The error is why the engine cannot, as in [`calculate`].
pub(crate) fn convert(
    source: &Book,
    pass: &WorkbookModifications,
    volatile: Volatile,
    cache: Option<&Cache>,
) -> Result<(Vec<u8>, Vec<FromCacheNote>), ConvertError> {
    let mut options = Options::default().with_book(source);
    // The macros of a workbook are never run: the build refuses a workbook that carries them unless secure mode is on, and then
    // what depends on them keeps the cache until something dirties it (and is an error cell when the run calculates it).
    options.config.now = volatile.now.or(options.config.now);
    options.config.seed = volatile.seed.or(options.config.seed);
    // The cache is what the passes before changed (the formulas that reach calculate, the rest keep what the file caches, and
    // `LibreOffice` calculates some formulas on load); without it every formula is calculated.
    match cache {
        Some(cache) => options.cache = cache.clone(),
        None => options.config.calculate_all = Some(true),
    }
    // A sheet the pass deletes takes its conversion with it: nothing is asked of it.
    let named: Vec<Only> = pass
        .formula_to_value
        .iter()
        .filter(|scope| {
            !pass
                .delete_sheets
                .as_ref()
                .is_some_and(|gone| gone.contains(&scope.sheet))
        })
        .map(only_of)
        .collect();
    let write_as_it_is = |from_cache| Ok((source.modify_to_bytes(pass)?, from_cache));
    if named.is_empty() {
        return write_as_it_is(Vec::new());
    }
    // `calculate_all` of an XLSX calculates the cells the conversion names and what they read, and nothing else. The other runs write
    // every result they calculate: the cache, the formulas the plan reaches and those `LibreOffice` calculates on load, which are
    // not only the ones the conversion names; an ODS with `calculate_all`, whose startup calculates every formula.
    if source.calculate_all() && source.format() != sheetrs_io::reader::Format::Ods {
        options.only = named;
    }
    let (program, built) =
        source
            .compile(&options, Mode::Recalculate)
            .map_err(|e| match cannot(&e) {
                CalculateError::Depends(reason) => ConvertError::Depends(reason),
                CalculateError::Refused(reason) => ConvertError::Refused(reason),
                other => ConvertError::Engine(other.to_string()),
            })?;
    if program.formula_cells.is_empty() {
        return write_as_it_is(built.from_cache);
    }
    let mut workbook = Vec::new();
    sheetrs_answer::fill_template_to_also(
        source,
        &program,
        None,
        &Json::Null,
        &options.overrides(),
        pass,
        &mut workbook,
    )
    .map_err(|e| match e {
        sheetrs_answer::Error::Write(e) => ConvertError::Write(e),
        e => ConvertError::Other(e.to_string()),
    })?;
    Ok((workbook, built.from_cache))
}

/// The cells a conversion names, as what the engine is asked for.
fn only_of(scope: &ValueScope) -> Only {
    let (rows, columns) = match scope.target {
        ValueTarget::Sheet => (None, None),
        ValueTarget::Columns(first, last) => (None, Some((first, last))),
        ValueTarget::Rows(first, last) => (Some((first, last)), None),
    };
    Only {
        sheet: scope.sheet.clone(),
        rows,
        columns,
    }
}
