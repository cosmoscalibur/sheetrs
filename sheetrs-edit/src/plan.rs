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
//! 3. the changes of visibility (rows, columns and sheets hidden or shown): a pass of their own,
//!    against the input, so that they name the rows and columns of the workbook as it is and a
//!    deletion that follows removes hidden rows like any other;
//! 4. the conversion of formulas to values (of what the formulas calculate, see below), the deletions
//!    of rows and columns, of defined names and of sheets, all in one pass over the file: they are
//!    written against that workbook, whose rows, columns and sheets are those of the input.
//!
//! Step by step is a different thing: one cut per step (`write`, or save, between operations) gives
//! the workbook the first step leaves as the input of the next.
//!
//! One pass applies what its operations ask in the order the writer does, all against the input
//! of that pass: formulas converted to values first, then, in each sheet, the rows and then the
//! columns deleted from the last to the first, then the named ranges deleted, then the sheets
//! deleted. A reference that spans sheets (`Sheet1:Sheet3!A1`) is not supported: issue #75.
//!
//! **A conversion converts what the formulas calculate**, not what the file caches (the result a file
//! caches is not a goal of the project, see `docs/approved-differences.md`; `LibreOffice` itself
//! calculates some formulas when it loads a file and keeps the cached results of the others). It is one
//! step in every mode: where the engine runs for it (always for an ODS, and for an XLSX whose results
//! are stale or that calculates every formula) it calculates in the pass that writes the conversion. With
//! `calculate_all` of an XLSX the cells the conversion names are the outputs of the engine (it calculates
//! them and what they read: the compiler prunes the workbook to that, [`sheetrs_compiler::Only`]); in the
//! cache mode the outputs are the formulas the plan reaches and the ones `LibreOffice` calculates on load,
//! and the cells the conversion names that the engine did not calculate are converted with the result the
//! file caches, which the same pass reads. The writer, in the same pass over the file that applies the
//! deletions, writes each calculated cell as the value the engine gave, in place of the formula
//! ([`sheetrs_io::write::CellEdit::Cached`] in a cell the pass converts). There is no pass of its own for the
//! calculation and none for the conversion: the file is read once more, by the engine. The engine
//! calculates the workbook as the passes before leave it, so a replacement or a calculation that comes
//! first is seen, and the conversion is of the workbook as it is before the deletions of its own pass
//! move any cell. A cell that depends on what the engine cannot calculate (another workbook, an extension, or a
//! macro in secure mode) keeps the result the file caches unless the run calculates it, in which case an add-in that is
//! not installed is `#NAME?` and a function only `LibreOffice` has is an error that names the cause; a workbook with macros
//! is refused by the build that runs the engine ([`Error::Refused`]) unless secure mode is on
//! (`sheetrs_compiler::Unsupported`; the run says so with [`EXTERNAL_FROM_CACHE`](crate::EXTERNAL_FROM_CACHE) and
//! its two relatives); any other cause that the engine cannot calculate it is an error for a conversion, which
//! needs the values.
//!
//! **The startup of an ODS.** `LibreOffice` calculates some formulas of an ODS when it opens it (the formulas
//! that cache no result, an error, a result with no type or of some types in a cell with no format; a few
//! functions), so the plan of an ODS starts with that calculation, in the mode of the run: the roots of the load
//! by default, every formula with `calculate_all`, which is all that differs. It is the first pass that runs the
//! engine (a calculation, a conversion, the recalculation that ends the plan), and a pass of its own, the first,
//! when the plan has none (hiding columns). The compiler finds the roots in the read that compiles the
//! workbook for that pass; the planner looks at nothing in the file before. After the startup each operation
//! applies its own rule (a deletion reaches a cone, a conversion writes what the calculation gave, hiding rows
//! reaches its cells). An XLSX has no startup: it keeps the functions `LibreOffice` calculates on load only.
//!
//! A deletion (or a removal, or a replacement) leaves the cached results of the formulas that read what
//! changed stale: their text may not change, and so their value is out of date. The writer streams the
//! file and reads no results of its own, so the plan ends with the engine calculating every formula
//! again ([`Operation::Recalculate`]), with the semantics of `LibreOffice`: every formula calculated
//! again, so that stale values are never written. What the engine gives is what `LibreOffice` shows when
//! the workbook it leaves is saved and opened again (re-read): a reference that went is a bare `#REF!` in
//! the text, the error value of an ODS and the unknown name of an XLSX, and the engine reads it so; the
//! The same holds when the workbook says it is stale when the run starts (an XLSX marked to
//! recalculate when it is opened; an ODS with no cached result in its formulas) and a run that uses the
//! cache converts.
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
//! **The cache is the default** (what `LibreOffice` does when it opens a file and when a cell changes:
//! it calculates what has no result and what the change reaches, and keeps the rest), and
//! `calculate_all` (`Book::with_calculate_all(true)`, `--calculate-all`, `calculate_all=True`,
//! `calculateAll = true`) is the opt-in that emulates `calculateAll()`: every formula is calculated
//! again. By default the recalculation calculates the formulas the plan reaches and what
//! depends on them, and every other formula keeps, byte for byte, the result the file caches. The
//! option changes nothing about when the engine runs, only how much of the workbook it calculates. The
//! default is also what makes a conversion read the results the file caches: the engine of the pass
//! that writes the conversion calculates the cone of the plan, and the cells the conversion names that
//! it did not calculate keep the results the file holds; with `calculate_all` the engine calculates what
//! the conversion names (every formula in an ODS), and those are the results it converts. A formula is reached when the writer
//! left it with no cached result (it read a deleted cell or a range that shrank; the replaced
//! formula), when it reads a cell a replacement wrote, when it is volatile, when it is a `SUBTOTAL`
//! or `AGGREGATE` and rows were hidden or shown, when a deletion moved the cell of a `ROW` or `COLUMN` or what they are
//! given, when it reads through a defined name whose target a deletion moved, shrank or removed (or any name when one was deleted),
//! and every formula when a sheet was deleted: what `LibreOffice` recalculates after the same edit (see
//! `sheetrs_compiler::Cache`); the engine itself is what finds who depends on those. A workbook that reads
//! cells at run time (`INDIRECT`, `OFFSET`), one that the file says is stale (an XLSX marked to be recalculated when it is opened),
//! and one where the replacements and the deletions meet, are calculated
//! whole: the result is then that of a run with `calculate_all`. A run that does not need the
//! engine at all (the cone is empty) writes the workbook as the passes left it.
//!
//! The plan is a description, not a model of the workbook: it holds the operations (a few bands
//! and names), and each pass streams the file once, rewriting each cell as it goes by. The memory
//! of a pass does not depend on the size of the workbook.

use crate::calculate::{CalculateError, ConvertError};
use crate::diagnostics::{from_cache, not_calculated, not_calculated_on_load};
use crate::error::{Context, Error, Result};
use crate::replace::Report;
use sheetrs_compiler::{Cache, FromCacheNote};
use sheetrs_io::reader::Format;
use sheetrs_io::write::{
    self as writer, Bands, Calculation, ColumnRange, Operation, RowRange, SheetEdits, ValueScope,
    ValueTarget, VisibilityChange, WorkbookModifications,
};
use sheetrs_io::{Book, RowState, Volatile};
use std::collections::{BTreeMap, HashMap};
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
    fn recalculate(
        &self,
        workbook: &Book,
        operation: &str,
        volatile: Volatile,
        cache: Option<&Cache>,
    ) -> Result<(Vec<u8>, Vec<FromCacheNote>)> {
        #[cfg(test)]
        if let Some(recalculate) = self.recalculate {
            return recalculate(workbook)
                .map(|bytes| (bytes, Vec::new()))
                .map_err(|reason| Error::Recalculation {
                    operation: operation.to_string(),
                    reason,
                });
        }
        crate::calculate::recalculate(workbook, volatile, cache)
            .map(|done| (done.workbook, done.from_cache))
            .map_err(|error| match error {
                CalculateError::Depends(reason) => Error::Unsupported {
                    operation: operation.to_string(),
                    reason,
                },
                CalculateError::Refused(reason) => Error::Refused {
                    operation: operation.to_string(),
                    reason,
                },
                error => Error::Recalculation {
                    operation: operation.to_string(),
                    reason: error.to_string(),
                },
            })
    }

    /// The pass `pass` over `workbook`, whose conversions of formulas to values take the results the
    /// engine calculates for the cells they name (see [`crate::calculate::convert`]).
    #[cfg_attr(not(test), allow(clippy::unused_self))] // the tests give the engine a stand-in
    fn convert(
        &self,
        workbook: &Book,
        pass: &WorkbookModifications,
        operation: &str,
        volatile: Volatile,
        cache: Option<&Cache>,
    ) -> Result<(Vec<u8>, Vec<FromCacheNote>)> {
        #[cfg(test)]
        if let Some(recalculate) = self.recalculate {
            // The stand-in leaves the results as the file has them: the conversion is of those.
            let bytes = recalculate(workbook).map_err(|reason| Error::Recalculation {
                operation: operation.to_string(),
                reason,
            })?;
            return Ok((
                writer::modify_workbook_bytes(workbook.format(), &bytes, pass)?,
                Vec::new(),
            ));
        }
        crate::calculate::convert(workbook, pass, volatile, cache).map_err(|error| match error {
            ConvertError::Engine(reason) => Error::Recalculation {
                operation: operation.to_string(),
                reason,
            },
            ConvertError::Depends(reason) => Error::Unsupported {
                operation: operation.to_string(),
                reason,
            },
            ConvertError::Refused(reason) => Error::Refused {
                operation: operation.to_string(),
                reason,
            },
            ConvertError::Write(error) => Error::Write(error),
            ConvertError::Other(reason) => Error::Calculation { reason },
        })
    }
}

/// What a run that keeps the results the file caches (no [`Book::calculate_all`]) has changed since the
/// results of the workbook were last known to be right: the recalculation calculates what this
/// reaches and leaves every other formula as the file has it.
struct Touched {
    cached: bool,
    /// `None`: every formula is calculated (the run does not use the cache, or the file says its
    /// results are stale, or the changes cannot be placed).
    cache: Option<Cache>,
}

/// Replacements beyond this many rectangles in a sheet are one rectangle that holds them all.
const MOST_CHANGES: usize = 64;

impl Touched {
    fn start(book: &Book) -> Result<Self> {
        let cached = !book.calculate_all();
        // An XLSX marked to be recalculated when it is opened caches nothing to trust. An ODS has no
        // such mark: a formula that has no result is calculated like any other that the plan reaches.
        let marked = cached && book.format() == Format::Xlsx && crate::stale::input_is_stale(book)?;
        Ok(Self {
            cached,
            cache: (cached && !marked).then(Cache::default),
        })
    }

    /// The workbook was calculated: its results are right, whatever came before.
    fn calculated(&mut self) {
        self.cache = self.cached.then(Cache::default);
    }

    /// A pass ran, with these replacements in it. `names` are the defined names of the workbook the pass
    /// read (name to definition), asked for only when the pass deletes rows, columns or names.
    #[cfg(test)]
    fn record(
        &mut self,
        pass: &[Operation],
        replacements: &HashMap<String, SheetEdits>,
        names: impl FnOnce() -> HashMap<String, String>,
    ) {
        self.record_with(pass, replacements, names, |_| None);
    }

    /// As `record`, with `hidden`: the rows the input hides in the sheets asked for (`None` when they cannot be read, and then
    /// the rows a pass hides or shows all count as changed). `LibreOffice` dirties the formulas of the rows whose visibility
    /// changes, so hiding a row that is hidden or showing one that is visible reaches nothing.
    fn record_with(
        &mut self,
        pass: &[Operation],
        replacements: &HashMap<String, SheetEdits>,
        names: impl FnOnce() -> HashMap<String, String>,
        hidden: impl FnOnce(&[(&str, u32, u32)]) -> Option<HashMap<String, RowState>>,
    ) {
        let Some(cache) = &mut self.cache else {
            return;
        };
        let mut names_deleted = false;
        let mut visibility: Vec<&VisibilityChange> = Vec::new();
        for operation in pass {
            match operation {
                Operation::SetVisibility(changes) => visibility.extend(changes),
                Operation::DeleteRows(rows) => cache.rows_deleted.extend(
                    rows.iter()
                        .map(|range| (range.sheet.clone(), range.first, range.last)),
                ),
                Operation::DeleteColumns(columns) => cache.columns_deleted.extend(
                    columns
                        .iter()
                        .map(|range| (range.sheet.clone(), range.first, range.last)),
                ),
                Operation::DeleteSheets(_) => cache.sheets_deleted = true,
                Operation::DeleteNames(deleted) => {
                    // The workbook that is calculated no longer has them: the formulas that read through them are found by their names.
                    names_deleted = true;
                    cache.stale_names.extend(deleted.iter().cloned());
                }
                _ => {}
            }
        }
        cache
            .rows_hidden
            .extend(rows_that_change(&visibility, hidden));
        // A deleted name recompiles every formula that reads through one (the compiler knows the names of the workbook); a
        // deletion of rows or columns, the ones whose name has a target it moved or shrank, which needs the definitions before it.
        cache.names_deleted |= names_deleted;
        if !(cache.rows_deleted.is_empty() && cache.columns_deleted.is_empty()) {
            cache.stale_names = sheetrs_compiler::compile::names_reached(
                &names(),
                &cache.rows_deleted,
                &cache.columns_deleted,
            );
        }
        for (sheet, edits) in replacements {
            // Runs of cells along a row, one rectangle each.
            let mut rects: Vec<(u32, u32, u32, u32)> = Vec::new();
            for &(row, col) in edits.keys() {
                match rects.last_mut() {
                    Some(last) if last.0 == row && last.2 == row && last.3 + 1 == col => {
                        last.3 = col;
                    }
                    _ => rects.push((row, col, row, col)),
                }
            }
            if rects.len() > MOST_CHANGES {
                let hull = rects.iter().fold(rects[0], |h, r| {
                    (h.0.min(r.0), h.1.min(r.1), h.2.max(r.2), h.3.max(r.3))
                });
                rects = vec![hull];
            }
            cache.changed.extend(
                rects
                    .into_iter()
                    .map(|(r0, c0, r1, c1)| (sheet.clone(), r0, c0, r1, c1)),
            );
        }
        // Where a change and a deletion meet, the cells of the change are not where they were.
        let layout = cache.sheets_deleted
            || !(cache.rows_deleted.is_empty() && cache.columns_deleted.is_empty());
        if layout && !cache.changed.is_empty() {
            self.cache = None;
        }
    }
}

/// The rows whose visibility the `changes` of a pass change, as `(sheet, first, last)`: each change is applied in order to the
/// rows the input hides, and a row is one that changed when some change flipped it (a row hidden and shown again counts, as it
/// does for `LibreOffice`; hiding a hidden row or showing a visible one flips nothing) and it holds a value or a formula: the
/// formulas that read a row are dirtied by its cells (OBSERVED, `probe` of `visibility.*`: a row with no cell, hidden or shown, calculates
/// nothing).
fn rows_that_change(
    changes: &[&VisibilityChange],
    hidden: impl FnOnce(&[(&str, u32, u32)]) -> Option<HashMap<String, RowState>>,
) -> Vec<(String, u32, u32)> {
    let rows: Vec<(&str, u32, u32, bool)> = changes
        .iter()
        .filter_map(|change| match change {
            VisibilityChange::Rows { range, visible } => {
                Some((range.sheet.as_str(), range.first, range.last, *visible))
            }
            _ => None,
        })
        .collect();
    if rows.is_empty() {
        return Vec::new();
    }
    let spans: Vec<(&str, u32, u32)> = rows
        .iter()
        .map(|&(sheet, first, last, _)| (sheet, first, last))
        .collect();
    let Some(mut state) = hidden(&spans) else {
        return rows
            .iter()
            .map(|&(sheet, first, last, _)| (sheet.to_string(), first, last))
            .collect();
    };
    let filled: HashMap<String, Vec<(u32, u32)>> = state
        .iter()
        .map(|(sheet, rows)| (sheet.clone(), rows.filled.clone()))
        .collect();
    let mut state: HashMap<String, Vec<(u32, u32)>> = state
        .drain()
        .map(|(sheet, rows)| (sheet, rows.hidden))
        .collect();
    let mut changed: HashMap<&str, Vec<(u32, u32)>> = HashMap::new();
    for (sheet, first, last, visible) in rows {
        let now = state.entry(sheet.to_string()).or_default();
        let flipped = if visible {
            intersect(now, (first, last))
        } else {
            subtract(&[(first, last)], now)
        };
        *now = if visible {
            subtract(now, &[(first, last)])
        } else {
            union(now, &[(first, last)])
        };
        let all = changed.entry(sheet).or_default();
        *all = union(all, &flipped);
    }
    for (sheet, bands) in &mut changed {
        let with_cells = filled.get(*sheet).map_or(&[][..], Vec::as_slice);
        *bands = bands
            .iter()
            .flat_map(|&band| intersect(with_cells, band))
            .collect();
    }
    let mut out: Vec<(String, u32, u32)> = changed
        .into_iter()
        .flat_map(|(sheet, bands)| {
            bands
                .into_iter()
                .map(move |(first, last)| (sheet.to_string(), first, last))
        })
        .collect();
    out.sort();
    out
}

/// The bands of `a` and `b` as one sorted list of bands that neither overlap nor touch.
fn union(a: &[(u32, u32)], b: &[(u32, u32)]) -> Vec<(u32, u32)> {
    let mut all: Vec<(u32, u32)> = a.iter().chain(b).copied().collect();
    all.sort_unstable();
    let mut out: Vec<(u32, u32)> = Vec::new();
    for (first, last) in all {
        match out.last_mut() {
            Some(prev) if first <= prev.1.saturating_add(1) => prev.1 = prev.1.max(last),
            _ => out.push((first, last)),
        }
    }
    out
}

/// What of the bands `a` is not in the bands `b`.
fn subtract(a: &[(u32, u32)], b: &[(u32, u32)]) -> Vec<(u32, u32)> {
    let mut out = Vec::new();
    for &(first, last) in a {
        let mut from = first;
        let mut done = false;
        for &(b_first, b_last) in &union(b, &[]) {
            if b_last < from {
                continue;
            }
            if b_first > last {
                break;
            }
            if b_first > from {
                out.push((from, b_first - 1));
            }
            if b_last >= last {
                done = true;
                break;
            }
            from = b_last + 1;
        }
        if !done {
            out.push((from, last));
        }
    }
    out
}

/// What of the bands `a` is in the band `b`.
fn intersect(a: &[(u32, u32)], b: (u32, u32)) -> Vec<(u32, u32)> {
    a.iter()
        .filter_map(|&(first, last)| {
            let (first, last) = (first.max(b.0), last.min(b.1));
            (first <= last).then_some((first, last))
        })
        .collect()
}

/// The defined names of `book`, name to definition (none when they cannot be read: the engine reads them again).
fn input_names(book: &Book) -> HashMap<String, String> {
    book.defined_name_definitions().unwrap_or_default()
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
    let mut visibility: Vec<VisibilityChange> = Vec::new();
    for operation in operations {
        match operation {
            Operation::SetVisibility(changes) => visibility.extend(changes.iter().cloned()),
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
        // The rows, columns and sheets are those of the input: it runs before any deletion.
        Operation::SetVisibility(visibility),
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
    passes_from(operations, false, true, false)
}

/// The passes of a plan, and what the run needs to know about them besides the operations they hold.
pub(crate) struct Plan {
    pub passes: Vec<Vec<Operation>>,
    /// The first pass is the startup of an ODS that the plan adds because no pass of it runs the engine (see
    /// [`plan_from`]): a recalculation of what `LibreOffice` calculates when it opens the file.
    pub startup_pass: bool,
    /// The conversion of formulas to values is calculated by the engine in the pass that writes it (see [`plan_from`]).
    pub engine_converts: bool,
}

impl Plan {
    /// Whether `pass` converts formulas to values with the engine.
    fn converts_with_engine(&self, pass: &[Operation]) -> bool {
        self.engine_converts
            && pass
                .iter()
                .any(|operation| matches!(operation, Operation::FormulaToValue(_)))
    }

    /// The index of the first pass that runs the engine.
    fn first_engine_pass(&self) -> Option<usize> {
        self.passes
            .iter()
            .position(|pass| runs_the_engine(pass))
            .filter(|_| !self.startup_pass || self.passes.len() > 1)
    }
}

/// Whether `pass` runs the engine: a calculation, the recalculation of the planner, or a conversion (calculated by the
/// engine in the pass that writes it whenever the run has an engine to start with: see [`plan_from`]).
fn runs_the_engine(pass: &[Operation]) -> bool {
    pass.iter().any(|operation| {
        matches!(
            operation,
            Operation::Recalculate | Operation::Calculate(_) | Operation::FormulaToValue(_)
        )
    })
}

/// As [`passes`], for a workbook whose cached results are stale when the run starts (`stale`) or not; `calculate_all` says
/// whether the run calculates every formula, and `startup` whether the workbook is an ODS.
pub(crate) fn passes_from(
    operations: &[Operation],
    stale: bool,
    calculate_all: bool,
    startup: bool,
) -> Vec<Vec<Operation>> {
    plan_from(operations, stale, calculate_all, startup).passes
}

/// The plan of `operations` (see [`Plan`]).
///
/// **A conversion of formulas to values is one step in every mode.** Where the engine runs for the conversion (always for an ODS;
/// for an XLSX when the results the file caches are stale or the run calculates every formula) it calculates in the pass that
/// writes the conversion: the cells the conversion names with `calculate_all` of an XLSX, and otherwise (the cache) the formulas the
/// plan reaches and the ones `LibreOffice` calculates on load, which are the ones the engine is given; the rest of the cells the
/// conversion names are converted with the result the file caches, which the same pass reads. No calculation of the workbook comes
/// first: a conversion needs the values the engine gives, not a file that holds them (a second read and write of the whole file).
///
/// **The startup of an ODS.** `LibreOffice` calculates some formulas of an ODS when it opens it (the formulas that cache no result, an
/// error, a result with no type or of a type in a cell with no format; a few functions, in both formats), so the plan of an ODS
/// starts with that calculation, in the mode of the run: the roots of the load by default, every formula with `calculate_all`. It is the
/// first pass that runs the engine, and when the plan has none (nothing leaves results stale, no calculation, no conversion) it is a
/// pass of its own, the first. The roots are found in the read that compiles the workbook for it: nothing scans the file before.
pub(crate) fn plan_from(
    operations: &[Operation],
    mut stale: bool,
    calculate_all: bool,
    startup: bool,
) -> Plan {
    let mut passes = Vec::new();
    let mut last: Vec<Operation> = Vec::new();
    let mut convert_stale = false;
    for operation in canonical(operations) {
        match operation {
            Operation::Calculate(_) => {
                // It calculates the workbook as the operations before it leave it: a pass of its
                // own.
                passes.push(vec![operation]);
                stale = false;
            }
            Operation::SetVisibility(_) => {
                // Hidden rows change what SUBTOTAL and AGGREGATE add up, and the changes are in
                // the coordinates of the input: a pass of its own, before the deletions.
                stale |= leaves_results_stale(&operation);
                passes.push(vec![operation]);
            }
            Operation::Replace(_) | Operation::WriteCells(_) => {
                // A replacement reads what the operations before it leave and decides the cells it
                // changes; a write addresses the cells as they are at that point: a pass of its own.
                stale |= leaves_results_stale(&operation);
                passes.push(vec![operation]);
            }
            Operation::FormulaToValue(_) => {
                // The engine that calculates the conversion calculates, in the same pass, what is stale (the cache) or everything
                // (an ODS with `calculate_all`): nothing is left for a recalculation before it.
                convert_stale = stale;
                if !calculate_all || startup {
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
    let startup_pass = startup && !passes.iter().any(|pass| runs_the_engine(pass));
    if startup_pass {
        passes.insert(0, vec![Operation::Recalculate]);
    }
    Plan {
        passes,
        startup_pass,
        engine_converts: calculate_all || startup || convert_stale,
    }
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
/// results are stale when the run starts (an XLSX marked to recalculate, or one with functions `LibreOffice` calculates on load),
/// which makes the engine calculate a conversion. An ODS starts with what `LibreOffice` calculates when it opens it, which comes
/// first in the text. A workbook that cannot be read is explained as [`explain`] does.
pub(crate) fn explain_book(book: &Book) -> String {
    use std::fmt::Write as _;
    let operations = book.operations();
    let Ok(plan) = passes_for(book, operations) else {
        return explain(operations);
    };
    let mut out = String::new();
    if book.format() == Format::Ods {
        let what = if book.calculate_all() {
            "every formula is calculated (`calculate_all`)"
        } else {
            "the formulas LibreOffice calculates when it opens the file are calculated, and nothing that reads them"
        };
        let carried = match (plan.startup_pass, plan.first_engine_pass()) {
            (true, _) => "in a pass of its own, the first".to_string(),
            (false, Some(i)) => format!("in pass {}, with what the plan asks of the engine", i + 1),
            (false, None) => "in no pass".to_string(),
        };
        let _ = writeln!(out, "startup of an ODS: {what}; {carried}");
    }
    for (i, pass) in plan.passes.iter().enumerate() {
        let _ = writeln!(out, "pass {} of {}:", i + 1, plan.passes.len());
        for operation in pass {
            let _ = writeln!(out, "  {operation}");
        }
    }
    if plan
        .passes
        .iter()
        .any(|pass| plan.converts_with_engine(pass))
    {
        out.push_str(if book.calculate_all() {
            "the formulas converted to values are calculated by the engine in the pass that writes them\n"
        } else {
            "the formulas converted to values are calculated by the engine in the pass that writes them: the ones the plan reaches and the ones LibreOffice calculates on load; the others are converted with the results the file caches\n"
        });
    }
    if !book.calculate_all() && plan.passes.iter().any(|pass| runs_the_engine(pass)) {
        out.push_str(
            "the cache (the default): the formulas the plan reaches, and what depends on them, are calculated; the others keep the results the file caches\n",
        );
    }
    if has_calculate(operations)
        || plan.passes.last().map(Vec::as_slice) == Some(&[Operation::Recalculate])
            && !plan.startup_pass
    {
        out.push_str(
            "if the engine cannot calculate the workbook: warn, and mark it to be recalculated when opened\n",
        );
    }
    out
}

/// What the volatile functions give in every calculation of the plan of `book`: the `now` and the
/// `seed` of its last `calculate` step that sets them, else the ones given to the book
/// ([`Book::with_volatile`]). The recalculation the planner adds after what leaves results stale
/// runs with them too, so that a plan that fixes the clock gives the same workbook whichever pass
/// calculates.
fn plan_volatile(book: &Book) -> Volatile {
    let given = book.volatile();
    let step = |pick: fn(&Calculation) -> bool| {
        book.operations()
            .iter()
            .rev()
            .find_map(|operation| match operation {
                Operation::Calculate(calculation) if pick(calculation) => Some(calculation),
                _ => None,
            })
    };
    Volatile {
        now: step(|c| c.now.is_some()).and_then(|c| c.now).or(given.now),
        seed: step(|c| c.seed.is_some())
            .and_then(|c| c.seed)
            .or(given.seed),
    }
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
    let plan = passes_for(book, operations)?;
    let passes = &plan.passes;
    let only_on_load = only_on_load(book, operations);
    let volatile = plan_volatile(book);
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
    let mut touched = Touched::start(book)?;
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
                    &Book::from_file(&current)?
                        .with_locale_of(book)
                        .with_calculate_all_of(book)
                        .with_secure_mode_of(book),
                    only_on_load,
                    plan.startup_pass && i == 0,
                    last,
                    &first_conversion(operations),
                    engine,
                    volatile,
                    &touched,
                    &mut report,
                )
                .with_context(said)?
            }),
            [Operation::Calculate(calculation)] => Some(
                calculate_pass(
                    calculation,
                    &Book::from_file(&current)?
                        .with_locale_of(book)
                        .with_calculate_all_of(book)
                        .with_secure_mode_of(book),
                    volatile,
                    &touched,
                    &mut report,
                )
                .with_context(said)?,
            ),
            _ => None,
        };
        if let Some(calculated) = calculated {
            if let Some(bytes) = calculated {
                touched.calculated();
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
        let input = Book::from_file(&current)?
            .with_locale_of(book)
            .with_calculate_all_of(book)
            .with_secure_mode_of(book);
        let mut mods =
            modifications_for(pass, &input, last && failed, &mut report).with_context(said)?;
        mods.clear_recalculation_mark = false;
        // The engine calculates the cells to convert and the writer writes them as values, with the
        // deletions, in this one pass.
        let converted = if plan.converts_with_engine(pass) {
            convert_pass(
                engine,
                &input,
                &mods,
                &first_conversion(operations),
                volatile,
                touched.cache.as_ref(),
                &mut report,
            )
            .with_context(said)
            .map(Some)?
        } else {
            None
        };
        if converted.is_some() && (!book.calculate_all() || book.format() == Format::Ods) {
            // The engine calculated what was stale (the cache) or every formula (an ODS) in this pass.
            touched.calculated();
        }
        touched.record_with(
            pass,
            &mods.replacements,
            || input_names(&input),
            |spans| input.row_state(spans).ok(),
        );
        match converted {
            Some(bytes) => {
                std::fs::write(&target, bytes).context("Cannot write the converted workbook")?;
            }
            None => writer::modify_workbook(&current, &target, &mods).with_context(said)?,
        }
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
    let plan = passes_for(book, operations)?;
    let passes = &plan.passes;
    let only_on_load = only_on_load(book, operations);
    let volatile = plan_volatile(book);
    let format = book.format();
    let mut current = if let Some(path) = book.path() {
        Book::from_bytes(std::fs::read(path).context("Cannot read the workbook to edit")?)?
            .with_locale_of(book)
            .with_calculate_all_of(book)
            .with_secure_mode_of(book)
    } else {
        let mut copy = book.clone();
        copy.clear_operations();
        copy
    };
    let mut report = Report::default();
    let mut touched = Touched::start(book)?;
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
                    only_on_load,
                    plan.startup_pass && i == 0,
                    last,
                    &first_conversion(operations),
                    engine,
                    volatile,
                    &touched,
                    &mut report,
                )
                .with_context(said)?
            }),
            [Operation::Calculate(calculation)] => Some(
                calculate_pass(calculation, &current, volatile, &touched, &mut report)
                    .with_context(said)?,
            ),
            _ => None,
        };
        let data = std::sync::Arc::clone(
            current
                .bytes()
                .expect("the workbook of a pass is in memory"),
        );
        if let Some(calculated) = calculated {
            if let Some(bytes) = calculated {
                touched.calculated();
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
            current = Book::from_bytes(result.clone())?
                .with_locale_of(book)
                .with_calculate_all_of(book)
                .with_secure_mode_of(book);
            continue;
        }
        let mods =
            modifications_for(pass, &current, last && failed, &mut report).with_context(said)?;
        let converted = if plan.converts_with_engine(pass) {
            convert_pass(
                engine,
                &current,
                &mods,
                &first_conversion(operations),
                volatile,
                touched.cache.as_ref(),
                &mut report,
            )
            .with_context(said)
            .map(Some)?
        } else {
            None
        };
        if converted.is_some() && (!book.calculate_all() || format == Format::Ods) {
            touched.calculated();
        }
        touched.record_with(
            pass,
            &mods.replacements,
            || input_names(&current),
            |spans| current.row_state(spans).ok(),
        );
        result = match converted {
            Some(bytes) => bytes,
            None => writer::modify_workbook_bytes(format, &data, &mods).with_context(said)?,
        };
        if !last {
            current = Book::from_bytes(result.clone())?
                .with_locale_of(book)
                .with_calculate_all_of(book)
                .with_secure_mode_of(book);
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

/// The pass `mods` over `workbook` with the conversions of formulas to values calculated by the engine
/// (see [`Engine::convert`]; `cache` is what the passes before changed, `None` when every formula is calculated): the bytes
/// of the workbook, and the notes of what the engine could not calculate and left as the file caches it. The engine failing
/// is an error: the values to convert do not exist.
#[allow(clippy::too_many_arguments)] // the state of one pass of the plan
fn convert_pass(
    engine: &Engine,
    workbook: &Book,
    mods: &WorkbookModifications,
    operation: &str,
    volatile: Volatile,
    cache: Option<&Cache>,
    report: &mut Report,
) -> Result<Vec<u8>> {
    let (bytes, from_cache_notes) = engine.convert(workbook, mods, operation, volatile, cache)?;
    report.diagnostics.extend(
        from_cache_notes
            .iter()
            .map(|note| from_cache(operation, note)),
    );
    Ok(bytes)
}

/// Calculates every formula of `workbook` again, for the results the operations left stale. The
/// bytes of the workbook with fresh results, or `None` when the engine cannot and this is the last
/// pass (`last`): nothing needs the values then, so the run says so ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)) and
/// the workbook is marked to be recalculated when it is opened by whoever calls this. A conversion
/// of formulas to values after it needs the values: the engine failing is an error then.
#[allow(clippy::too_many_arguments)] // the state of one pass of the plan
fn recalculate_pass(
    workbook: &Book,
    only_on_load: bool,
    startup: bool,
    last: bool,
    operation: &str,
    engine: &Engine,
    volatile: Volatile,
    touched: &Touched,
    report: &mut Report,
) -> Result<Option<Vec<u8>>> {
    match engine.recalculate(workbook, operation, volatile, touched.cache.as_ref()) {
        Ok((bytes, notes)) => {
            let label = if operation.is_empty() {
                "recalculate"
            } else {
                operation
            };
            report
                .diagnostics
                .extend(notes.iter().map(|note| from_cache(label, note)));
            Ok(Some(fresh(bytes, workbook.format())?))
        }
        // The startup of an ODS that nothing else in the plan needs the engine for: what `LibreOffice` calculates on load is not
        // calculated by the engine here (a function it lacks, a reference it does not read, a macro), and the file goes on as it is,
        // with the results it caches; nothing the plan did left a result stale, so nothing marks the workbook.
        Err(Error::Unsupported { reason, .. } | Error::Recalculation { reason, .. }) if startup => {
            report.diagnostics.push(not_calculated_on_load(&reason));
            Ok(Some(match workbook.bytes() {
                Some(bytes) => bytes.to_vec(),
                None => std::fs::read(workbook.path().expect("a book is a file or bytes"))
                    .context("Cannot read the workbook")?,
            }))
        }
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
        // The recalculation an ODS gets only because `LibreOffice` calculates some of its formulas on load (nothing the plan did
        // left a result stale) cannot be done for a workbook that holds what the engine cannot calculate (another workbook, macros,
        // an extension): those cells keep the result the file caches, and the file goes on as it is.
        Err(Error::Unsupported { .. }) if only_on_load => Ok(Some(match workbook.bytes() {
            Some(bytes) => bytes.to_vec(),
            None => std::fs::read(workbook.path().expect("a book is a file or bytes"))
                .context("Cannot read the workbook")?,
        })),
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
    volatile: Volatile,
    touched: &Touched,
    report: &mut Report,
) -> Result<Option<Vec<u8>>> {
    // The step's own clock and seed win; what it leaves open is the plan's.
    let calculation = &Calculation {
        now: calculation.now.or(volatile.now),
        seed: calculation.seed.or(volatile.seed),
        ..calculation.clone()
    };
    match crate::calculate::calculate(workbook, calculation, touched.cache.as_ref()) {
        Ok(done) => {
            if calculation.answer {
                report.value_ranges = done.answer;
            }
            report.diagnostics.extend(
                done.from_cache
                    .iter()
                    .map(|note| from_cache("calculate", note)),
            );
            Ok(Some(fresh(done.workbook, workbook.format())?))
        }
        Err(CalculateError::Unsupported(reason)) if !calculation.answer => {
            report
                .diagnostics
                .push(not_calculated("calculate", &reason));
            Ok(None)
        }
        Err(CalculateError::Refused(reason)) => Err(Error::Refused {
            operation: "calculate".to_string(),
            reason,
        }),
        Err(
            CalculateError::Unsupported(reason)
            | CalculateError::Depends(reason)
            | CalculateError::Failed(reason),
        ) => Err(Error::Calculation { reason }),
    }
}

/// The plan for a run on `book` (see [`plan_from`]): an XLSX that says its results are stale (marked to be recalculated when it
/// is opened) or that has functions `LibreOffice` calculates on load starts with them stale, which a conversion is calculated
/// for; an ODS starts with the startup of the load, found by the read that compiles it (nothing scans the file here).
fn passes_for(book: &Book, operations: &[Operation]) -> Result<Plan> {
    let ods = book.format() == Format::Ods;
    // `LibreOffice` calculates some formulas of an XLSX when it opens it (see `stale`): the run that keeps the cache calculates them
    // whatever the operations are (writing results is the exception: they are taken as current). A plan that leaves results stale and
    // converts nothing ends with a recalculation of the engine that calculates them as roots, so the file is not looked at for them.
    let ends_with_recalculation = leaves_stale(operations) && !reads_results(operations);
    let stale_start = !book.calculate_all()
        && !ods
        && (reads_results(operations) && crate::stale::input_is_stale(book)?
            || !ends_with_recalculation
                && operations
                    .iter()
                    .any(|operation| !matches!(operation, Operation::WriteCells(_)))
                && crate::stale::calculates_on_load(book)?);
    Ok(plan_from(
        operations,
        stale_start,
        book.calculate_all(),
        ods,
    ))
}

/// Whether the recalculation of the plan is there only because `LibreOffice` calculates some formulas when it loads the file:
/// nothing the operations do leaves a result stale (a conversion of formulas to values reads the results the load calculated).
fn only_on_load(book: &Book, operations: &[Operation]) -> bool {
    !book.calculate_all()
        && !leaves_stale(operations)
        && !crate::stale::input_is_stale(book).unwrap_or(true)
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
    // Rows hidden or shown change what SUBTOTAL and AGGREGATE read; columns and sheets are read
    // by no formula.
    if let Operation::SetVisibility(changes) = operation {
        return changes
            .iter()
            .any(|change| matches!(change, VisibilityChange::Rows { .. }));
    }
    // The empty search text matches nothing: the replacement changes no cell.
    if let Operation::Replace(spec) = operation {
        return !spec.replace.find.matches_nothing();
    }
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
            Operation::SetVisibility(changes) => mods.visibility.extend(changes.iter().cloned()),
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
    fn a_conversion_after_a_replacement_is_calculated_while_it_is_written() {
        // Pushed after the conversion, the replacement still runs before it.
        let operations = [
            Operation::FormulaToValue(scopes(&["F"])),
            Operation::DeleteRows(rows(&["T!1"])),
            replacement(),
        ];
        let described = described(&passes(&operations));
        assert_eq!(described.len(), 3);
        assert!(described[0][0].starts_with("replace"));
        // The engine calculates the cells of `F` while the pass writes them as values: no calculation of
        // the whole workbook comes first.
        assert_eq!(described[1], ["formula to value in F", "delete rows T!1:1"]);
        // What the replacement and the deletion leave stale is calculated again at the end.
        assert_eq!(
            described[2],
            ["recalculate every formula (the results are stale)"]
        );
        assert!(explain(&operations).starts_with("pass 1 of 3:"));
    }

    #[test]
    fn a_conversion_with_the_cache_is_one_pass_that_calculates_the_replacement() {
        let operations = [
            Operation::FormulaToValue(scopes(&["F"])),
            Operation::DeleteRows(rows(&["T!1"])),
            replacement(),
        ];
        // The engine that calculates the conversion calculates what the replacement left stale in the same pass: no calculation of
        // the workbook comes first, whichever the mode. The deletion of the pass leaves results stale again, which the last pass calculates.
        let described = described(&passes_from(&operations, false, false, false));
        assert_eq!(described.len(), 3);
        assert!(described[0][0].starts_with("replace"));
        assert_eq!(described[1], ["formula to value in F", "delete rows T!1:1"]);
        assert_eq!(
            described[2],
            ["recalculate every formula (the results are stale)"]
        );
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
        assert!(passes_from(&[], false, true, false).is_empty());
    }

    #[test]
    fn a_conversion_is_one_pass_in_every_mode() {
        // The results the file caches are stale (an XLSX marked to recalculate, a deletion before it): no pass comes before.
        let operations = [Operation::FormulaToValue(scopes(&["F"]))];
        for (stale, calculate_all) in [(false, true), (true, false), (false, false)] {
            assert_eq!(
                passes_from(&operations, stale, calculate_all, false),
                vec![operations.to_vec()],
                "stale {stale}, calculate_all {calculate_all}"
            );
        }
        // The engine runs for the conversion when the file is stale, with `calculate_all` and always for an ODS; a conversion of an
        // XLSX whose results are right and that keeps the cache converts what the file caches, with no engine.
        let engine = |stale, calculate_all, ods| {
            plan_from(&operations, stale, calculate_all, ods).engine_converts
        };
        assert!(
            engine(true, false, false) && engine(false, true, false) && engine(false, false, true)
        );
        assert!(!engine(false, false, false));
    }

    #[test]
    fn an_ods_load_is_a_pass_only_when_no_other_pass_runs_the_engine() {
        let hide = [Operation::SetVisibility(vec![VisibilityChange::Columns {
            range: "T!B".parse().unwrap(),
            visible: false,
        }])];
        // Nothing leaves a result stale, nothing calculates, nothing converts: the startup is the first pass.
        let plan = plan_from(&hide, false, false, true);
        assert!(plan.startup_pass);
        assert_eq!(plan.passes.len(), 2);
        assert_eq!(plan.passes[0], vec![Operation::Recalculate]);
        // In the other plans it is the pass that runs the engine anyway: a deletion's recalculation, a conversion, a calculation.
        for operations in [
            vec![Operation::DeleteRows(rows(&["T!3"]))],
            vec![Operation::FormulaToValue(scopes(&["F"]))],
            vec![Operation::Calculate(Calculation::default())],
        ] {
            let plan = plan_from(&operations, false, false, true);
            assert!(!plan.startup_pass, "{operations:?}");
            assert_eq!(
                plan.passes,
                passes_from(&operations, false, false, false)
                    .into_iter()
                    .collect::<Vec<_>>()
            );
        }
        // An XLSX has no startup.
        assert!(!plan_from(&hide, false, false, false).startup_pass);
        // The mode changes the roots of the startup, not its place in the plan.
        assert_eq!(
            plan_from(&hide, false, true, true).passes,
            plan_from(&hide, false, false, true).passes
        );
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

    /// What a run that uses the cache has to tell the engine (`Touched`): the changes the passes made,
    /// as the engine reads them.
    mod touched {
        use super::*;
        use sheetrs_io::write::CellEdit;

        fn tracking() -> Touched {
            Touched {
                cached: true,
                cache: Some(Cache::default()),
            }
        }

        fn rows(sheet: &str, first: u32, last: u32) -> RowRange {
            RowRange {
                sheet: sheet.to_string(),
                first,
                last,
            }
        }

        fn columns(sheet: &str, first: u32, last: u32) -> ColumnRange {
            ColumnRange {
                sheet: sheet.to_string(),
                first,
                last,
            }
        }

        /// The cache after `pass` ran with `replacements`.
        fn after(pass: &[Operation], replacements: &[(&str, &[(u32, u32)])]) -> Cache {
            let mut touched = tracking();
            touched.record(pass, &edits(replacements), HashMap::new);
            touched.cache.expect("the cache is still used")
        }

        fn edits(replacements: &[(&str, &[(u32, u32)])]) -> HashMap<String, SheetEdits> {
            replacements
                .iter()
                .map(|(sheet, cells)| {
                    let cells = cells
                        .iter()
                        .map(|at| (*at, CellEdit::Text("x".into())))
                        .collect();
                    ((*sheet).to_string(), cells)
                })
                .collect()
        }

        /// The rows a pass flips, with `hidden` and `filled` as the bands of the input (sheet `S`).
        fn flipped(
            changes: &[(u32, u32, bool)],
            hidden: &[(u32, u32)],
            filled: &[(u32, u32)],
        ) -> Vec<(u32, u32)> {
            let changes: Vec<VisibilityChange> = changes
                .iter()
                .map(|&(first, last, visible)| VisibilityChange::Rows {
                    range: rows("S", first, last),
                    visible,
                })
                .collect();
            let changes: Vec<&VisibilityChange> = changes.iter().collect();
            rows_that_change(&changes, |_| {
                Some(HashMap::from([(
                    "S".to_string(),
                    RowState {
                        hidden: hidden.to_vec(),
                        filled: filled.to_vec(),
                    },
                )]))
            })
            .into_iter()
            .map(|(_, first, last)| (first, last))
            .collect()
        }

        #[test]
        fn hiding_a_hidden_row_or_showing_a_visible_one_changes_nothing() {
            let all = [(0, 20)];
            assert_eq!(flipped(&[(2, 2, false)], &[(2, 2)], &all), vec![]);
            assert_eq!(flipped(&[(5, 6, true)], &[(2, 2)], &all), vec![]);
        }

        #[test]
        fn a_band_that_overlaps_the_hidden_rows_changes_only_the_others() {
            let all = [(0, 20)];
            assert_eq!(
                flipped(&[(1, 5, false)], &[(2, 3)], &all),
                vec![(1, 1), (4, 5)]
            );
            assert_eq!(flipped(&[(3, 8, true)], &[(2, 4)], &all), vec![(3, 4)]);
        }

        #[test]
        fn a_row_hidden_and_shown_again_in_one_plan_changed() {
            let all = [(0, 20)];
            assert_eq!(
                flipped(&[(3, 3, false), (3, 3, true)], &[], &all),
                vec![(3, 3)]
            );
            // Shown and hidden again, and hidden then shown while already hidden: flipped twice.
            assert_eq!(
                flipped(&[(3, 3, true), (3, 3, false)], &[(3, 3)], &all),
                vec![(3, 3)]
            );
            assert_eq!(
                flipped(&[(3, 3, false), (3, 3, true)], &[(3, 3)], &all),
                vec![(3, 3)]
            );
            // Hidden twice flips once.
            assert_eq!(
                flipped(&[(3, 3, false), (3, 4, false)], &[], &all),
                vec![(3, 4)]
            );
        }

        #[test]
        fn a_row_with_no_cell_calculates_nothing() {
            assert_eq!(
                flipped(&[(2, 8, false)], &[], &[(0, 3), (7, 7)]),
                vec![(2, 3), (7, 7)]
            );
            assert_eq!(flipped(&[(10, 12, false)], &[], &[(0, 3)]), vec![]);
        }

        #[test]
        fn the_rows_are_all_noted_when_the_workbook_cannot_say_which_are_hidden() {
            let change = VisibilityChange::Rows {
                range: rows("S", 1, 2),
                visible: true,
            };
            assert_eq!(
                rows_that_change(&[&change], |_| None),
                vec![("S".to_string(), 1, 2)]
            );
        }

        #[test]
        fn hiding_or_showing_rows_is_noted_and_other_visibility_is_not() {
            let rows_change = |visible| VisibilityChange::Rows {
                range: rows("S", 1, 2),
                visible,
            };
            for visible in [true, false] {
                let cache = after(&[Operation::SetVisibility(vec![rows_change(visible)])], &[]);
                assert_eq!(cache.rows_hidden.len(), 1, "rows, visible {visible}");
            }
            let others = vec![
                VisibilityChange::Columns {
                    range: columns("S", 1, 2),
                    visible: false,
                },
                VisibilityChange::Sheet {
                    name: "S".into(),
                    visible: false,
                },
            ];
            let cache = after(&[Operation::SetVisibility(others)], &[]);
            assert_eq!(cache, Cache::default());
        }

        #[test]
        fn deletions_are_noted_with_the_bands_they_cut() {
            let cache = after(
                &[Operation::DeleteRows(vec![
                    rows("A", 0, 0),
                    rows("B", 3, 4),
                ])],
                &[],
            );
            assert_eq!(
                cache.rows_deleted,
                [("A".to_string(), 0, 0), ("B".to_string(), 3, 4)]
            );
            assert!(cache.columns_deleted.is_empty() && !cache.sheets_deleted);

            let cache = after(&[Operation::DeleteColumns(vec![columns("C", 1, 1)])], &[]);
            assert_eq!(cache.columns_deleted, [("C".to_string(), 1, 1)]);
            assert!(cache.rows_deleted.is_empty() && !cache.sheets_deleted);

            let cache = after(&[Operation::DeleteSheets(vec!["Gone".into()])], &[]);
            assert!(cache.sheets_deleted);
            assert!(cache.rows_deleted.is_empty() && cache.columns_deleted.is_empty());
        }

        /// What `LibreOffice` recalculates when a name goes: every formula that reads through a name, whichever (the compiler
        /// has the names of the workbook, so the planner reads none).
        #[test]
        fn a_deleted_name_reaches_the_formulas_of_every_name() {
            let mut touched = tracking();
            touched.record(
                &[Operation::DeleteNames(vec!["rate".into()])],
                &HashMap::new(),
                || panic!("the names are not read for this"),
            );
            let cache = touched.cache.expect("the cache is still used");
            assert!(cache.names_deleted);
            // The deleted name is no longer in the workbook the engine reads: it is told by its name.
            assert_eq!(cache.stale_names, ["rate"]);
            assert!(cache.rows_deleted.is_empty() && !cache.sheets_deleted);
        }

        /// A deletion of rows or columns reaches the names whose target it moved or shrank, and no other.
        #[test]
        fn a_deletion_reaches_the_names_whose_target_it_moves() {
            let mut touched = tracking();
            touched.record(
                &[Operation::DeleteRows(vec![rows("Data", 2, 2)])],
                &HashMap::new(),
                || {
                    HashMap::from([
                        ("above".to_string(), "Data!$A$1".to_string()),
                        ("below".to_string(), "Data!$A$5".to_string()),
                        ("elsewhere".to_string(), "Other!$A$5".to_string()),
                    ])
                },
            );
            let cache = touched.cache.expect("the cache is still used");
            assert_eq!(cache.stale_names, ["below"]);
            assert!(!cache.names_deleted);
        }

        #[test]
        fn what_changes_no_result_is_not_noted() {
            let pass = vec![
                Operation::FormulaToValue(scopes(&["S"])),
                Operation::Calculate(Calculation::default()),
                Operation::Recalculate,
            ];
            assert_eq!(after(&pass, &[]), Cache::default());
        }

        #[test]
        fn a_calculation_leaves_nothing_to_reach() {
            let mut touched = tracking();
            let delete = [Operation::DeleteRows(vec![rows("A", 0, 0)])];
            touched.record(&delete, &edits(&[("A", &[(0, 0)])]), HashMap::new);
            // The engine calculated every formula again: its results are right, whatever came before.
            touched.calculated();
            assert_eq!(touched.cache, Some(Cache::default()));

            // A run that does not use the cache has no cache to start from again.
            let mut off = Touched {
                cached: false,
                cache: None,
            };
            off.calculated();
            assert_eq!(off.cache, None);
        }

        #[test]
        fn a_run_with_no_cache_notes_nothing() {
            let mut off = Touched {
                cached: true,
                cache: None,
            };
            off.record(
                &[Operation::DeleteRows(vec![rows("A", 0, 0)])],
                &edits(&[("A", &[(0, 0)])]),
                HashMap::new,
            );
            assert_eq!(off.cache, None);
        }

        #[test]
        fn adjacent_cells_of_a_row_are_one_rectangle() {
            let cells = [(2, 1), (2, 2), (2, 3), (2, 5), (3, 5), (4, 1), (4, 2)];
            let cache = after(&[], &[("S", &cells)]);
            let rectangles: Vec<_> = cache.changed.iter().map(|c| (c.1, c.2, c.3, c.4)).collect();
            // The run of three, the lone cell after a gap, the cell below it (rows do not merge) and
            // the run of two.
            assert_eq!(
                rectangles,
                [(2, 1, 2, 3), (2, 5, 2, 5), (3, 5, 3, 5), (4, 1, 4, 2)]
            );
            assert!(cache.changed.iter().all(|c| c.0 == "S"));
        }

        #[test]
        fn more_than_the_most_changes_in_a_sheet_are_the_rectangle_that_holds_them() {
            // 64 isolated cells stay as they are, one more makes them one rectangle.
            let cells = |count: u32| -> Vec<(u32, u32)> {
                (0..count).map(|i| (2 * i + 1, 3 + i % 2)).collect()
            };
            let sixty_four = after(&[], &[("S", &cells(64))]);
            assert_eq!(sixty_four.changed.len(), 64);
            let sixty_five = after(&[], &[("S", &cells(65))]);
            assert_eq!(sixty_five.changed, [("S".to_string(), 1, 3, 129, 4)]);
            // Each sheet is counted alone.
            let two = after(&[], &[("S", &cells(64)), ("T", &cells(64))]);
            assert_eq!(two.changed.len(), 128);
        }

        #[test]
        fn a_change_and_a_deletion_together_leave_no_cache() {
            let mut touched = tracking();
            touched.record(&[], &edits(&[("S", &[(0, 0)])]), HashMap::new);
            assert!(touched.cache.is_some());
            touched.record(
                &[Operation::DeleteColumns(vec![columns("S", 5, 5)])],
                &HashMap::new(),
                HashMap::new,
            );
            assert_eq!(touched.cache, None);

            // In the same pass, and the other way round.
            let mut touched = tracking();
            touched.record(
                &[Operation::DeleteSheets(vec!["Gone".into()])],
                &edits(&[("S", &[(0, 0)])]),
                HashMap::new,
            );
            assert_eq!(touched.cache, None);
        }
    }
}
