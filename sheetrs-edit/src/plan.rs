//! The planner: what runs the operations a [`Book`] holds, and the passes over the file that do it.
//!
//! **The operations are not a sequence.** Each one is written against the workbook as it is when
//! it is handed over, in its rows, its columns and its sheets, and the planner decides the order
//! they run in, so that nobody has to think about how the workbook looks after each step. The
//! order the operations were pushed in changes nothing: `delete_rows("S!3")` and
//! `delete_rows("S!4")` are the rows 3 and 4 of the workbook, whichever came first and however they
//! were split into calls. The rule is one canonical order (see [`canonical`]):
//!
//! 1. the cells written ([`Operation::WriteCells`]) and the replacements
//!    ([`Operation::Replace`]), each decided against the workbook as the ones before it in this
//!    list leave it (the operations of one kind keep the order they were pushed in: a sweep of
//!    replacements is a sequence of texts, and the next one sees what the previous one left);
//! 2. the calculation ([`Operation::Calculate`]): the inputs are written and every formula is
//!    calculated;
//! 3. the conversion of formulas to values (it reads the results the workbook caches, so it comes
//!    after what could have left them stale, and when something did, the formulas are calculated
//!    again first), the deletions of rows and columns, of defined names and of sheets, all in one
//!    pass over the file: they are written against that workbook, whose rows, columns and sheets are
//!    those of the input.
//!
//! Step by step is a different thing: one cut per step (`write`, or save, between operations) gives
//! the workbook the first step leaves as the input of the next.
//!
//! One pass applies what its operations ask in the order the writer does, all against the input
//! of that pass: formulas converted to values first, then, in each sheet, the rows and then the
//! columns deleted from the last to the first, then the named ranges deleted, then the sheets
//! deleted. A reference that spans sheets (`Sheet1:Sheet3!A1`) is not supported: issue #75.
//!
//! A conversion reads the results the workbook caches, and a deletion (or a removal, or a
//! replacement) leaves the cached results of the formulas that read what changed stale: their text
//! may not change, and so their value is out of date. An application that calculates as it goes
//! shows the new results, so the conversion that follows gives those. The writer streams the file
//! and reads no results of its own; a conversion after such a change is a pass of its own preceded by
//! another, [`Operation::Recalculate`], which the engine (the compiler's program, with the semantics
//! of `LibreOffice`) runs over the workbook: every formula calculated again, so that stale values are
//! never written; when the engine cannot, the run stops, since the values to convert do not exist.
//! The same holds when the workbook says it is stale when the run starts (an XLSX marked to
//! recalculate when it is opened; an ODS with no cached result in its formulas).
//!
//! **The mark to recalculate is a consequence of the engine failing, never a decision taken
//! before.** A plan that leaves results stale (rows, columns, sheets or names deleted, text replaced)
//! ends with a pass that calculates every formula again with the engine: the workbook is written
//! with the fresh results and no mark. The result a file cached is not kept as a goal (see
//! `docs/approved-differences.md`): what a cell holds is the value its formula gives now. Only when
//! the engine cannot calculate the workbook (a function it does not have, a reference it does not
//! read) is the workbook marked to be recalculated when it is opened (`fullCalcOnLoad` in an XLSX,
//! the cached results dropped in an ODS) and the [`Report`] says why
//! ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)): the spreadsheet program that opens it
//! does the recalculation then. Nothing assigns or clears the mark beforehand, and there is no way to
//! ask for it or to avoid it. A run that only converts formulas to the values they have changes no
//! other result and has nothing to calculate. A workbook that came marked and was calculated by the
//! engine loses the mark, since its results are fresh.
//!
//! The plan is a description, not a model of the workbook: it holds the operations (a few bands
//! and names), and each pass streams the file once, rewriting each cell as it goes by. The memory
//! of a pass does not depend on the size of the workbook.

use crate::calculate::CalculateError;
use crate::diagnostics::not_calculated;
use crate::error::{Context, Error, Result};
use crate::replace::Report;
use sheetrs_io::Book;
use sheetrs_io::reader::Format;
use sheetrs_io::write::{
    self as writer, Bands, Calculation, ColumnRange, Operation, RowRange, ValueScope, ValueTarget,
    WorkbookModifications,
};
use std::collections::BTreeMap;
use std::path::Path;

/// What a run is asked besides its operations (nothing, but in the tests: the engine of the
/// recalculation can be replaced).
#[derive(Debug, Clone, Default)]
pub(crate) struct Engine {
    /// Replaces the engine in the recalculation of the passes, in the tests that are about how the
    /// operations group into passes and not about what a recalculation gives.
    #[cfg(test)]
    pub recalculate: Option<StandIn>,
}

#[cfg(test)]
pub(crate) type StandIn = fn(&Book) -> std::result::Result<Vec<u8>, String>;

impl Engine {
    /// Every formula of `workbook` calculated again, by the engine.
    #[cfg_attr(not(test), allow(clippy::unused_self))] // the tests give the engine a stand-in
    fn recalculate(&self, workbook: &Book, operation: &str) -> Result<Vec<u8>> {
        #[cfg(test)]
        if let Some(recalculate) = self.recalculate {
            return recalculate(workbook).map_err(|reason| Error::Recalculation {
                operation: operation.to_string(),
                reason,
            });
        }
        crate::calculate::recalculate(workbook).map_err(|reason| Error::Recalculation {
            operation: operation.to_string(),
            reason,
        })
    }
}

fn target_key(target: ValueTarget) -> (u8, u32, u32) {
    match target {
        ValueTarget::Sheet => (0, 0, 0),
        ValueTarget::Columns(first, last) => (1, first, last),
        ValueTarget::Rows(first, last) => (2, first, last),
    }
}

/// The deleted bands of each sheet as ranges: the ones that overlap or touch are one, in the order
/// of the sheets and of the positions.
fn bands<T>(
    ranges: impl IntoIterator<Item = (String, u32, u32)>,
    make: impl Fn(String, u32, u32) -> T,
) -> Vec<T> {
    let mut by_sheet: BTreeMap<String, Vec<(u32, u32)>> = BTreeMap::new();
    for (sheet, first, last) in ranges {
        by_sheet.entry(sheet).or_default().push((first, last));
    }
    by_sheet
        .into_iter()
        .flat_map(|(sheet, ranges)| {
            Bands::new(ranges)
                .ranges()
                .map(|(first, last)| make(sheet.clone(), first, last))
                .collect::<Vec<_>>()
        })
        .collect()
}

/// The operations in the order they run in, whatever order they were pushed in: the cells written
/// and the replacements first, then the calculations (each kind in the order it was pushed), then one
/// conversion, one deletion of rows, one of columns, one of names and one of sheets, each the
/// union of what was asked, in the order of the sheets and the positions. See the module
/// documentation.
pub(crate) fn canonical(operations: &[Operation]) -> Vec<Operation> {
    let mut converts: Vec<ValueScope> = Vec::new();
    let mut rows = Vec::new();
    let mut columns = Vec::new();
    let mut names: Vec<String> = Vec::new();
    let mut sheets: Vec<String> = Vec::new();
    let (mut writes, mut replaces, mut calculations) = (Vec::new(), Vec::new(), Vec::new());
    for operation in operations {
        match operation {
            Operation::FormulaToValue(scopes) => converts.extend(scopes.iter().cloned()),
            Operation::DeleteRows(ranges) => {
                rows.extend(ranges.iter().map(|r| (r.sheet.clone(), r.first, r.last)));
            }
            Operation::DeleteColumns(ranges) => {
                columns.extend(ranges.iter().map(|r| (r.sheet.clone(), r.first, r.last)));
            }
            Operation::DeleteNames(n) => names.extend(n.iter().cloned()),
            Operation::DeleteSheets(n) => sheets.extend(n.iter().cloned()),
            Operation::WriteCells(_) => writes.push(operation.clone()),
            Operation::Replace(_) => replaces.push(operation.clone()),
            Operation::Calculate(_) => calculations.push(operation.clone()),
            // What the planner puts where a conversion needs it: never written by hand.
            Operation::Recalculate => {}
        }
    }
    names.sort();
    names.dedup();
    sheets.sort();
    sheets.dedup();
    converts.sort_by_key(|scope| (scope.sheet.clone(), target_key(scope.target)));
    converts.dedup();
    let mut canonical = writes;
    canonical.extend(replaces);
    canonical.extend(calculations);
    for operation in [
        Operation::FormulaToValue(converts),
        Operation::DeleteRows(bands(rows, |sheet, first, last| RowRange {
            sheet,
            first,
            last,
        })),
        Operation::DeleteColumns(bands(columns, |sheet, first, last| ColumnRange {
            sheet,
            first,
            last,
        })),
        Operation::DeleteNames(names),
        Operation::DeleteSheets(sheets),
    ] {
        if !operation.is_empty() {
            canonical.push(operation);
        }
    }
    canonical
}

/// Whether the cached results the workbook ends with are stale after the operations: one that leaves
/// them so (rows, columns, sheets or names deleted, text replaced) comes after the last time they
/// were calculated (a `calculate` step, or a conversion of formulas to values, which calculates them
/// first when they were stale). The run then ends with the engine calculating every formula again
/// (see the module documentation); nothing is decided about the mark.
pub(crate) fn leaves_stale(operations: &[Operation]) -> bool {
    let mut stale = false;
    for operation in &canonical(operations) {
        if matches!(
            operation,
            Operation::FormulaToValue(_) | Operation::Calculate(_)
        ) {
            stale = false;
        }
        stale |= leaves_results_stale(operation);
    }
    stale
}

/// Whether some operation is a `calculate` step.
pub(crate) fn has_calculate(operations: &[Operation]) -> bool {
    operations
        .iter()
        .any(|operation| matches!(operation, Operation::Calculate(_)))
}

/// Whether some operation reads the results the workbook caches (a conversion of formulas to
/// values), so that a workbook whose results are stale needs them calculated first.
pub(crate) fn reads_results(operations: &[Operation]) -> bool {
    operations
        .iter()
        .any(|operation| matches!(operation, Operation::FormulaToValue(_)))
}

/// The passes over the file that run the operations: the operations of each, in the canonical
/// order and against the input of that pass. Their number is what the run costs in reads and writes
/// of the file. The last is the engine calculating again what the operations left stale, when they
/// did.
pub(crate) fn passes(operations: &[Operation]) -> Vec<Vec<Operation>> {
    passes_from(operations, false)
}

/// As [`passes`], for a workbook whose cached results are stale when the run starts (`stale`) or
/// not.
///
/// A conversion of formulas to values reads the results the workbook caches. When what runs before
/// it (or the file itself) left them stale, they are calculated again first: a pass of its own that
/// no operation asks for ([`Operation::Recalculate`]). That is what running the operations one at
/// a time gives in an application that calculates as it goes.
pub(crate) fn passes_from(operations: &[Operation], mut stale: bool) -> Vec<Vec<Operation>> {
    let mut passes = Vec::new();
    let mut last: Vec<Operation> = Vec::new();
    for operation in canonical(operations) {
        match operation {
            Operation::Calculate(_) => {
                // It calculates the workbook as the operations before it leave it: a pass of its
                // own.
                passes.push(vec![operation]);
                stale = false;
            }
            Operation::Replace(_) | Operation::WriteCells(_) => {
                // A replacement reads what the operations before it leave and decides the cells it
                // changes; a write addresses the cells as they are at that point: a pass of its own.
                stale |= leaves_results_stale(&operation);
                passes.push(vec![operation]);
            }
            Operation::FormulaToValue(_) => {
                if stale {
                    passes.push(vec![Operation::Recalculate]);
                    stale = false;
                }
                last.push(operation);
            }
            Operation::Recalculate => {}
            _ => {
                stale = true;
                last.push(operation);
            }
        }
    }
    if !last.is_empty() {
        passes.push(last);
    }
    if stale {
        // What the operations left stale is calculated again by the engine, as the last thing.
        passes.push(vec![Operation::Recalculate]);
    }
    passes
}

/// What running `operations` would do, as text, one operation per line and grouped by the pass over
/// the file that runs it, for a person to read. It reads no workbook ([`BookEdit::explain`](crate::BookEdit::explain) says it
/// for the operations of a book). The operations of one kind keep their order, and the planner
/// puts the kinds in an order of its own.
#[must_use]
pub fn explain(operations: &[Operation]) -> String {
    use std::fmt::Write as _;
    let passes = passes(operations);
    let mut out = String::new();
    for (i, pass) in passes.iter().enumerate() {
        let _ = writeln!(out, "pass {} of {}:", i + 1, passes.len());
        for operation in pass {
            let _ = writeln!(out, "  {operation}");
        }
    }
    if has_calculate(operations) || leaves_stale(operations) {
        out.push_str(
            "if the engine cannot calculate the workbook: warn, and mark it to be recalculated when opened\n",
        );
    }
    out
}

/// As [`explain`], for the operations `book` holds: the workbook is looked at to tell whether its
/// results are stale when the run starts, which puts a calculation before a conversion. A workbook
/// that cannot be read is explained as [`explain`] does.
pub(crate) fn explain_book(book: &Book) -> String {
    use std::fmt::Write as _;
    let operations = book.operations();
    let Ok(passes) = passes_for(book, operations) else {
        return explain(operations);
    };
    let mut out = String::new();
    for (i, pass) in passes.iter().enumerate() {
        let _ = writeln!(out, "pass {} of {}:", i + 1, passes.len());
        for operation in pass {
            let _ = writeln!(out, "  {operation}");
        }
    }
    if has_calculate(operations)
        || passes.last().map(Vec::as_slice) == Some(&[Operation::Recalculate])
    {
        out.push_str(
            "if the engine cannot calculate the workbook: warn, and mark it to be recalculated when opened\n",
        );
    }
    out
}

/// Runs the operations of `book` and writes `output`, which must not be the input file. Each pass
/// reads and writes the file once, as a stream; a run of several passes keeps the file between
/// them in a temporary folder, and so does a book of bytes (the writer works on files). A failed
/// run leaves no `output` behind. No operations is an error: there is nothing to do, and a copy of
/// the file is not what was asked for.
///
/// # Errors
///
/// When there are no operations, when a file cannot be read or written, or when a pass fails (a
/// sheet that does not exist, a formula with a 3D reference, a change the writer cannot make
/// faithfully).
pub(crate) fn execute(book: &Book, output: impl AsRef<Path>, engine: &Engine) -> Result<Report> {
    let operations = book.operations();
    if operations.is_empty() {
        return Err(Error::EmptyPlan);
    }
    let output = output.as_ref();
    if let Some(path) = book.path()
        && same_file(path, output)
    {
        return Err(writer::Error::OutputIsInput(output.to_path_buf()).into());
    }
    let passes = passes_for(book, operations)?;
    let extension = book.format().extension();
    let folder = (passes.len() > 1 || book.path().is_none())
        .then(tempfile::tempdir)
        .transpose()
        .context("Cannot create a temporary folder")?;
    let mut current = match (book.path(), book.bytes(), &folder) {
        (Some(path), _, _) => path.to_path_buf(),
        (None, Some(data), Some(folder)) => {
            let path = folder.path().join(format!("input.{extension}"));
            std::fs::write(&path, data).context("Cannot write the workbook to edit")?;
            path
        }
        _ => unreachable!("a book of bytes has a folder"),
    };
    let mut report = Report::default();
    // Whether the engine failed: only then is the workbook marked to be recalculated.
    let mut failed = false;
    let marked = WorkbookModifications {
        recalculate_on_open: true,
        ..WorkbookModifications::default()
    };
    for (i, pass) in passes.iter().enumerate() {
        let last = i + 1 == passes.len();
        let target = match (&folder, last) {
            (Some(folder), false) => folder.path().join(format!("pass{i}.{extension}")),
            _ => output.to_path_buf(),
        };
        let said = || format!("Failed in pass {} of {}", i + 1, passes.len());
        let calculated = match pass.as_slice() {
            [Operation::Recalculate] => Some(if last && failed {
                None
            } else {
                recalculate_pass(
                    &Book::from_file(&current)?,
                    last,
                    &first_conversion(operations),
                    engine,
                    &mut report,
                )
                .with_context(said)?
            }),
            [Operation::Calculate(calculation)] => Some(
                calculate_pass(calculation, &Book::from_file(&current)?, &mut report)
                    .with_context(said)?,
            ),
            _ => None,
        };
        if let Some(calculated) = calculated {
            if let Some(bytes) = calculated {
                std::fs::write(&target, bytes).context("Cannot write the calculated workbook")?;
                current = target;
            } else {
                // The engine cannot: what is there goes on, and the last pass marks it.
                failed = true;
                if last {
                    writer::modify_workbook(&current, &target, &marked).with_context(said)?;
                    current = target;
                }
            }
            continue;
        }
        let mut mods = modifications_for(
            pass,
            &Book::from_file(&current)?,
            last && failed,
            &mut report,
        )
        .with_context(said)?;
        mods.clear_recalculation_mark = false;
        writer::modify_workbook(&current, &target, &mods).with_context(said)?;
        current = target;
    }
    Ok(report)
}

/// Whether two paths are one file (the output must never be the input: the passes write the result
/// as they go).
fn same_file(a: &Path, b: &Path) -> bool {
    match (std::fs::canonicalize(a), std::fs::canonicalize(b)) {
        (Ok(a), Ok(b)) => a == b,
        _ => a == b,
    }
}

/// Runs the operations of `book` and gives the result as the bytes of the workbook: the same passes
/// as [`execute`], with the workbook kept in memory between them (no file is read, except the one
/// a book from a file names, and none is written). This is what runs where there is no file
/// system, as in a browser.
///
/// # Errors
///
/// As [`execute`], without the errors of writing files.
pub(crate) fn execute_in_memory(book: &Book, engine: &Engine) -> Result<(Vec<u8>, Report)> {
    let operations = book.operations();
    if operations.is_empty() {
        return Err(Error::EmptyPlan);
    }
    let passes = passes_for(book, operations)?;
    let format = book.format();
    let mut current = if let Some(path) = book.path() {
        Book::from_bytes(std::fs::read(path).context("Cannot read the workbook to edit")?)?
    } else {
        let mut copy = book.clone();
        copy.clear_operations();
        copy
    };
    let mut report = Report::default();
    let mut result = Vec::new();
    // Whether the engine failed: only then is the workbook marked to be recalculated.
    let mut failed = false;
    let marked = WorkbookModifications {
        recalculate_on_open: true,
        ..WorkbookModifications::default()
    };
    for (i, pass) in passes.iter().enumerate() {
        let last = i + 1 == passes.len();
        let said = || format!("Failed in pass {} of {}", i + 1, passes.len());
        let calculated = match pass.as_slice() {
            [Operation::Recalculate] => Some(if last && failed {
                None
            } else {
                recalculate_pass(
                    &current,
                    last,
                    &first_conversion(operations),
                    engine,
                    &mut report,
                )
                .with_context(said)?
            }),
            [Operation::Calculate(calculation)] => {
                Some(calculate_pass(calculation, &current, &mut report).with_context(said)?)
            }
            _ => None,
        };
        let data = std::sync::Arc::clone(
            current
                .bytes()
                .expect("the workbook of a pass is in memory"),
        );
        if let Some(calculated) = calculated {
            if let Some(bytes) = calculated {
                result = bytes;
            } else {
                failed = true;
                if last {
                    result =
                        writer::modify_workbook_bytes(format, &data, &marked).with_context(said)?;
                } else {
                    continue;
                }
            }
            current = Book::from_bytes(result.clone())?;
            continue;
        }
        let mods =
            modifications_for(pass, &current, last && failed, &mut report).with_context(said)?;
        result = writer::modify_workbook_bytes(format, &data, &mods).with_context(said)?;
        if !last {
            current = Book::from_bytes(result.clone())?;
        }
    }
    Ok((result, report))
}

/// The workbook as the engine leaves it, with the mark to recalculate taken off when it came with
/// one: its results are fresh now.
fn fresh(bytes: Vec<u8>, format: Format) -> Result<Vec<u8>> {
    if format == Format::Xlsx && crate::stale::input_is_stale(&Book::from_bytes(bytes.clone())?)? {
        let mods = WorkbookModifications {
            clear_recalculation_mark: true,
            ..WorkbookModifications::default()
        };
        return Ok(writer::modify_workbook_bytes(format, &bytes, &mods)?);
    }
    Ok(bytes)
}

/// Calculates every formula of `workbook` again, for the results the operations left stale. The
/// bytes of the workbook with fresh results, or `None` when the engine cannot and this is the last
/// pass (`last`): nothing needs the values then, so the run says so ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)) and
/// the workbook is marked to be recalculated when it is opened by whoever calls this. A conversion
/// of formulas to values after it needs the values: the engine failing is an error then.
fn recalculate_pass(
    workbook: &Book,
    last: bool,
    operation: &str,
    engine: &Engine,
    report: &mut Report,
) -> Result<Option<Vec<u8>>> {
    match engine.recalculate(workbook, operation) {
        Ok(bytes) => Ok(Some(fresh(bytes, workbook.format())?)),
        Err(Error::Recalculation { reason, .. }) if last => {
            // The recalculation that ends a run is the planner's own; a conversion that asks for it names itself.
            let operation = if operation.is_empty() {
                "recalculate"
            } else {
                operation
            };
            report.diagnostics.push(not_calculated(operation, &reason));
            Ok(None)
        }
        Err(error) => Err(error),
    }
}

/// Runs a `calculate` step on `workbook`: the bytes of the workbook calculated, or `None` when the
/// engine cannot calculate it and the value ranges were not asked for: then the run says so
/// ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)) and the workbook is marked to be recalculated when it is opened by
/// whoever calls this, so the recalculation that is needed is done by the spreadsheet program that
/// opens it. With the value ranges asked for (`answer`) there is no answer to give, so it is an error.
fn calculate_pass(
    calculation: &Calculation,
    workbook: &Book,
    report: &mut Report,
) -> Result<Option<Vec<u8>>> {
    match crate::calculate::calculate(workbook, calculation) {
        Ok(done) => {
            if calculation.answer {
                report.value_ranges = done.answer;
            }
            Ok(Some(fresh(done.workbook, workbook.format())?))
        }
        Err(CalculateError::Unsupported(reason)) if !calculation.answer => {
            report
                .diagnostics
                .push(not_calculated("calculate", &reason));
            Ok(None)
        }
        Err(CalculateError::Unsupported(reason) | CalculateError::Failed(reason)) => {
            Err(Error::Calculation { reason })
        }
    }
}

/// The passes for a run on `book`: when the operations convert formulas to values, whether the
/// workbook already holds stale results is looked at, and the results are calculated again before
/// the conversion if so.
fn passes_for(book: &Book, operations: &[Operation]) -> Result<Vec<Vec<Operation>>> {
    let stale_start = reads_results(operations) && crate::stale::input_is_stale(book)?;
    Ok(passes_from(operations, stale_start))
}

/// The first conversion of formulas to values, as text, for what is said about it.
fn first_conversion(operations: &[Operation]) -> String {
    canonical(operations)
        .iter()
        .find(|operation| matches!(operation, Operation::FormulaToValue(_)))
        .map(ToString::to_string)
        .unwrap_or_default()
}

/// What the writer is asked for one pass over `book`: the operations of the pass, with the
/// replacements a search or a write decided against that workbook.
fn modifications_for(
    pass: &[Operation],
    book: &Book,
    mark: bool,
    report: &mut Report,
) -> Result<WorkbookModifications> {
    let mut mods = modifications_of(pass);
    match pass {
        [Operation::Replace(spec)] => {
            let (edits, done) = crate::replace::decide(book, spec)?;
            mods.replacements = edits;
            report.absorb(done);
        }
        [Operation::WriteCells(writes)] => {
            mods.replacements.clone_from(writes.sheets());
            report.written += writes.len();
        }
        _ => {}
    }
    mods.recalculate_on_open = mark;
    Ok(mods)
}

/// Whether the results the workbook caches are out of date after `operation`: it changes what
/// formulas read or the formulas themselves. A conversion changes no result, and a write is made by
/// a program that computed what depends on it.
fn leaves_results_stale(operation: &Operation) -> bool {
    !matches!(
        operation,
        Operation::FormulaToValue(_)
            | Operation::WriteCells(_)
            | Operation::Recalculate
            | Operation::Calculate(_)
    )
}

/// The operations of one pass as the writer takes them.
fn modifications_of(pass: &[Operation]) -> WorkbookModifications {
    let mut mods = WorkbookModifications::default();
    for operation in pass {
        match operation {
            Operation::DeleteSheets(names) => mods
                .delete_sheets
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::DeleteNames(names) => mods
                .delete_names
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::FormulaToValue(scopes) => {
                mods.formula_to_value.extend(scopes.iter().cloned());
            }
            Operation::DeleteRows(rows) => {
                mods.delete_rows
                    .extend(rows.iter().cloned().map(ValueScope::from));
            }
            Operation::DeleteColumns(columns) => {
                mods.delete_columns
                    .extend(columns.iter().cloned().map(ValueScope::from));
            }
            // Decided against the input of its pass, when the pass runs, or run by the editor.
            Operation::Replace(_)
            | Operation::WriteCells(_)
            | Operation::Recalculate
            | Operation::Calculate(_) => {}
        }
    }
    mods
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_io::write::ReplaceSpec;
    use sheetrs_protocol::{Find, Mode, Replace, SearchIn};

    fn scopes(specs: &[&str]) -> Vec<ValueScope> {
        specs.iter().map(|s| s.parse().unwrap()).collect()
    }

    fn rows(specs: &[&str]) -> Vec<RowRange> {
        specs.iter().map(|s| s.parse().unwrap()).collect()
    }

    fn columns(specs: &[&str]) -> Vec<ColumnRange> {
        specs.iter().map(|s| s.parse().unwrap()).collect()
    }

    fn described(passes: &[Vec<Operation>]) -> Vec<Vec<String>> {
        passes
            .iter()
            .map(|pass| pass.iter().map(ToString::to_string).collect())
            .collect()
    }

    fn replacement() -> Operation {
        Operation::Replace(ReplaceSpec {
            scopes: Vec::new(),
            replace: Replace {
                find: Find {
                    text: "a".into(),
                    mode: Mode::default(),
                    match_case: false,
                    entire_cells: false,
                    search_in: SearchIn::default(),
                    include_formulas: true,
                    formatted: false,
                    filtered: false,
                },
                with: "b".into(),
            },
        })
    }

    #[test]
    fn rows_are_deleted_with_row_ranges_and_columns_with_column_ranges() {
        // The kinds are types: a range of the wrong kind cannot be built.
        assert!("S!B:D".parse::<RowRange>().is_err());
        assert!("S".parse::<RowRange>().is_err());
        assert!("S!3:10".parse::<ColumnRange>().is_err());
        let row: RowRange = "'My Sheet'!3:10".parse().unwrap();
        assert_eq!(
            (row.sheet.as_str(), row.first, row.last),
            ("My Sheet", 2, 9)
        );
        assert_eq!(row.to_string(), "My Sheet!3:10");
        assert_eq!(row.to_string().parse::<RowRange>().unwrap(), row);
        let quoted: RowRange = "'R!D'!3:10".parse().unwrap();
        assert_eq!(quoted.to_string(), "'R!D'!3:10");
        assert_eq!("S!B".parse::<ColumnRange>().unwrap().to_string(), "S!B:B");
        assert_eq!(ValueScope::from(row).to_string(), "My Sheet!3:10");
    }

    #[test]
    fn an_operation_with_nothing_in_it_is_not_added() {
        let mut book =
            Book::from_bytes(include_bytes!("../../tests/lint_book.xlsx").as_slice()).unwrap();
        book.push(Operation::DeleteSheets(Vec::new()))
            .push(Operation::FormulaToValue(Vec::new()))
            .push(Operation::DeleteRows(Vec::new()))
            .push(Operation::DeleteColumns(Vec::new()));
        assert!(book.operations().is_empty());
    }

    #[test]
    fn the_deletions_are_the_rows_of_the_workbook_in_whatever_order_and_split_they_come() {
        // Rows 3, 4 and 10 of the workbook, however they are given; the ones that touch are one band.
        let expected = [["delete rows T!3:4, T!10:10"]];
        for given in [
            vec![rows(&["T!3"]), rows(&["T!4"]), rows(&["T!10"])],
            vec![rows(&["T!10"]), rows(&["T!4"]), rows(&["T!3"])],
            vec![rows(&["T!4", "T!10"]), rows(&["T!3"])],
            vec![rows(&["T!3:4"]), rows(&["T!3", "T!10"])],
        ] {
            let operations: Vec<Operation> = given.into_iter().map(Operation::DeleteRows).collect();
            // The last pass is the recalculation of what the deletions left stale.
            let described = described(&passes(&operations));
            assert_eq!(described[..1], expected);
            assert_eq!(described.len(), 2);
        }
    }

    #[test]
    fn deletions_names_and_sheets_share_a_pass_and_a_conversion_before_them_too() {
        let operations = [
            Operation::DeleteSheets(vec!["V".into()]),
            Operation::DeleteNames(vec!["N".into()]),
            Operation::DeleteRows(rows(&["U!2:3"])),
            Operation::DeleteColumns(columns(&["T!D"])),
            Operation::FormulaToValue(scopes(&["F"])),
        ];
        // One pass of the file for all of them, then the recalculation of what they left stale.
        assert_eq!(
            described(&passes(&operations))[0],
            [
                "formula to value in F",
                "delete rows U!2:3",
                "delete columns T!D:D",
                "delete names N",
                "delete sheets V",
            ]
        );
    }

    #[test]
    fn a_conversion_after_a_replacement_calculates_the_results_again_in_a_pass_of_its_own() {
        // Pushed after the conversion, the replacement still runs before it.
        let operations = [
            Operation::FormulaToValue(scopes(&["F"])),
            Operation::DeleteRows(rows(&["T!1"])),
            replacement(),
        ];
        let described = described(&passes(&operations));
        assert_eq!(described.len(), 4);
        assert!(described[0][0].starts_with("replace"));
        assert_eq!(
            described[1],
            ["recalculate every formula (the results are stale)"]
        );
        assert_eq!(described[2], ["formula to value in F", "delete rows T!1:1"]);
        // And what the deletion leaves stale is calculated again at the end.
        assert_eq!(described[3], described[1]);
        assert!(explain(&operations).starts_with("pass 1 of 4:"));
    }

    #[test]
    fn what_the_operations_leave_stale_is_calculated_again_as_the_last_pass() {
        let ends_with_recalculation = |operations: &[Operation]| {
            passes(operations)
                .last()
                .is_some_and(|pass| pass == &[Operation::Recalculate])
        };
        // Deleting leaves formulas whose text did not change with a stale result.
        assert!(ends_with_recalculation(&[Operation::DeleteRows(rows(&[
            "T!3"
        ]))]));
        assert!(ends_with_recalculation(&[Operation::DeleteColumns(
            columns(&["T!B"])
        )]));
        assert!(ends_with_recalculation(&[Operation::DeleteSheets(vec![
            "U".into()
        ])]));
        assert!(ends_with_recalculation(&[Operation::DeleteNames(vec![
            "N".into()
        ])]));
        assert!(ends_with_recalculation(&[replacement()]));
        // Converting formulas to the values they have changes no other cell.
        assert!(!leaves_stale(&[Operation::FormulaToValue(scopes(&["F"]))]));
        assert!(!leaves_stale(&[]));
        // A calculation calculates them, unless a deletion runs after it.
        assert!(!leaves_stale(&[Operation::Calculate(
            Calculation::default()
        )]));
        assert!(leaves_stale(&[
            Operation::DeleteRows(rows(&["T!3"])),
            Operation::Calculate(Calculation::default()),
        ]));
    }

    #[test]
    fn nothing_decides_the_mark_before_the_engine_runs() {
        // No pass of the plan carries the mark: it is a recalculation the engine either does or
        // fails at, and the explanation says what happens then.
        let operations = [
            Operation::DeleteRows(rows(&["T!3"])),
            Operation::FormulaToValue(scopes(&["F"])),
        ];
        let passes = passes(&operations);
        assert_eq!(passes.len(), 2);
        let text = explain(&operations);
        assert!(
            text.contains("pass 2 of 2:\n  recalculate every formula"),
            "{text}"
        );
        assert!(
            text.ends_with("if the engine cannot calculate the workbook: warn, and mark it to be recalculated when opened\n"),
            "{text}"
        );
        let converting = [Operation::FormulaToValue(scopes(&["F"]))];
        assert!(!explain(&converting).contains("recalculate"));
        assert!(passes_from(&[], false).is_empty());
    }

    #[test]
    fn a_book_with_no_operations_does_not_execute() {
        let book =
            Book::from_bytes(include_bytes!("../../tests/lint_book.xlsx").as_slice()).unwrap();
        assert!(matches!(
            execute(&book, "out.xlsx", &Engine::default()),
            Err(Error::EmptyPlan)
        ));
    }

    #[test]
    fn the_errors_say_what_is_wrong_with_a_type_and_not_only_with_a_text() {
        use crate::ScopeError;
        assert!(matches!(
            "S!B:D".parse::<RowRange>(),
            Err(ScopeError::NotRows { .. })
        ));
        assert!(matches!(
            "S!3".parse::<ColumnRange>(),
            Err(ScopeError::NotColumns { .. })
        ));
        assert!(matches!(
            "!A".parse::<ValueScope>(),
            Err(ScopeError::MissingSheet { .. })
        ));
        assert!(matches!(
            "'S".parse::<ValueScope>(),
            Err(ScopeError::Unterminated { .. })
        ));
        assert!(matches!(
            "'S'x".parse::<ValueScope>(),
            Err(ScopeError::TrailingText { .. })
        ));
    }
}
