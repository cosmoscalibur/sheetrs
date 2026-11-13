//! Which formulas have a stale result after a plan of deletions, and which of them need the engine.
//!
//! A deletion leaves the cached result of a formula out of date only when the formula reads
//! something the deletion changes: a deleted cell, a range that loses cells, a deleted sheet, or
//! another formula whose result is stale for one of those reasons. Every other formula keeps the
//! result the file stores, and nothing has to be calculated for it. The planner used to decide
//! from the type of the operation (any deletion, any replacement) and calculated the whole
//! workbook; this module decides from the formulas.
//!
//! The analysis reads the workbook as the engine does (the compiler's program, in the
//! coordinates of the input) and goes through its formulas in the order the engine calculates them,
//! the ones they read first:
//!
//! - A formula **converted to a value** in the same plan is not a formula any more, so its result
//!   cannot be stale: the conversion reads the values of the input.
//! - A formula that **reads nothing the plan changes** is left alone.
//! - A stale formula whose value is `#REF!` by its own reference, whatever the rest of the
//!   workbook holds, is given that error directly ([`Known::assigned`]): a reference to a deleted
//!   cell is the error, and the functions that pass an error on (arithmetic, `PMT`, `SUM`, text
//!   functions...) give it as their result when it is the first thing they meet. A formula that reads
//!   a formula given the error is given it too, on the same terms.
//! - Any other stale formula (a range that shrinks, a function that traps or chooses by errors such
//!   as `IFERROR` or `IF`, a function that reads positions such as `ROW` or `INDIRECT`) is left to
//!   the engine ([`Known::engine`]), in the coordinates of the output.
//!
//! What an unaffected cell gives is known from the file: the cached result of a formula the plan
//! does not touch is valid, so an operand whose cached value is a number, a boolean or empty is
//! never an error, and a stale formula that combines it with a deleted reference is that error.
//! The cells whose cached value is a text or an error, and the array formulas, are not told apart:
//! they go to the engine.
//!
//! When the workbook cannot be analyzed (the engine cannot compile it, it has cycles, or it holds
//! formulas the compiler keeps as the values the file cached), the answer is
//! [`Verdict::Unknown`] and the plan calculates the whole workbook, as it does for what it cannot
//! tell.

use sheetrs_compiler::Model;
use sheetrs_compiler::memory::{Options, compile_keeping};
use sheetrs_engine::addr::{CellRef, Rect};
use sheetrs_engine::ir::{BinOp, CellKind, Expr, Mode};
use sheetrs_engine::value::Value;
use sheetrs_io::Book;
use sheetrs_io::write::{ColumnRange, RowRange, ValueScope};
use std::collections::{BTreeSet, HashMap};

/// What a plan deletes and converts, in the rows, columns and sheets of the input.
#[derive(Debug, Default, Clone)]
pub(crate) struct Changes {
    pub rows: Vec<RowRange>,
    pub columns: Vec<ColumnRange>,
    pub sheets: Vec<String>,
    pub converted: Vec<ValueScope>,
}

/// Whether the plan changes nothing the definition of a name reads: it points to no sheet the plan
/// changes, or to cells (`Sheet!$A$1:$B$5`) that no deleted row or column touches. A definition that
/// is not a plain reference (a formula, a union) is not told apart.
fn name_is_left_alone(definition: &str, changes: &Changes) -> bool {
    let touched = |sheet: &str| {
        changes.sheets.iter().any(|s| s == sheet)
            || changes.rows.iter().any(|r| r.sheet == sheet)
            || changes.columns.iter().any(|c| c.sheet == sheet)
    };
    let mentions = changes
        .sheets
        .iter()
        .map(String::as_str)
        .chain(changes.rows.iter().map(|r| r.sheet.as_str()))
        .chain(changes.columns.iter().map(|c| c.sheet.as_str()))
        .any(|sheet| definition.contains(sheet));
    if !mentions {
        return true;
    }
    let text = definition.strip_prefix('=').unwrap_or(definition);
    let (Some(sheet), tail) = sheetrs_formula::addr::split_sheet(text) else {
        return false;
    };
    let Some(area) = sheetrs_formula::addr::parse_area(tail) else {
        return false;
    };
    if !touched(&sheet) {
        return true;
    }
    if changes.sheets.contains(&sheet) {
        return false;
    }
    let rows = changes
        .rows
        .iter()
        .any(|r| r.sheet == sheet && r.first <= area.r1 && area.r0 <= r.last);
    let columns = changes
        .columns
        .iter()
        .any(|c| c.sheet == sheet && c.first <= area.c1 && area.c0 <= c.last);
    !rows && !columns
}

/// The cells the plan leaves with a stale result, split by who gives them the new one.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct Known {
    /// Formulas whose value is `#REF!`: `(sheet, row, column)`, 0-based, in the output.
    pub assigned: Vec<(String, u32, u32)>,
    /// Formulas the engine calculates: `(sheet, row, column)`, 0-based, in the output.
    pub engine: Vec<(String, u32, u32)>,
}

/// The cells of the input whose cached value can give another error than the one a deleted
/// reference gives, or none at all: a text, an error, or an array formula, in the program's
/// numbering of the sheets. Every other cell is a number, a boolean or empty.
fn unsure_cells(model: &Model) -> BTreeSet<(u16, u32, u32)> {
    model
        .cells
        .iter()
        .filter(|(_, cell)| cell.array || matches!(cell.value, Value::Text(_) | Value::Error(_)))
        .map(|(at, _)| (at.sheet, at.row, at.col))
        .collect()
}

/// The answer of the analysis.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Verdict {
    /// The workbook cannot be analyzed: what it leaves stale is not known.
    Unknown,
    Known(Known),
}

/// Disjoint sorted bands of deleted indexes and how many are deleted before each.
#[derive(Debug, Default, Clone)]
struct Cuts {
    bands: Vec<(u32, u32)>,
    before: Vec<u32>,
}

impl Cuts {
    fn new(mut ranges: Vec<(u32, u32)>) -> Self {
        ranges.sort_unstable();
        let mut bands: Vec<(u32, u32)> = Vec::with_capacity(ranges.len());
        for (first, last) in ranges {
            match bands.last_mut() {
                Some(band) if first <= band.1.saturating_add(1) => band.1 = band.1.max(last),
                _ => bands.push((first, last)),
            }
        }
        let mut count = 0;
        let before = bands
            .iter()
            .map(|&(first, last)| {
                let n = count;
                count += last - first + 1;
                n
            })
            .collect();
        Self { bands, before }
    }

    /// The band that has `index`.
    fn band_of(&self, index: u32) -> Option<usize> {
        let i = self.bands.partition_point(|&(first, _)| first <= index);
        let i = i.checked_sub(1)?;
        (index <= self.bands[i].1).then_some(i)
    }

    fn contains(&self, index: u32) -> bool {
        self.band_of(index).is_some()
    }

    /// Whether one band holds all of `first..=last`.
    fn covers(&self, first: u32, last: u32) -> bool {
        self.band_of(first).is_some_and(|i| last <= self.bands[i].1)
    }

    /// Whether some band holds part of `first..=last`.
    fn touches(&self, first: u32, last: u32) -> bool {
        let i = self.bands.partition_point(|&(f, _)| f <= last);
        i.checked_sub(1).is_some_and(|i| self.bands[i].1 >= first)
    }

    /// Where `index` is once the bands are gone.
    fn moved(&self, index: u32) -> u32 {
        let started = self.bands.partition_point(|&(first, _)| first < index);
        match started.checked_sub(1) {
            None => index,
            Some(i) => {
                let (first, last) = self.bands[i];
                index - (self.before[i] + index.min(last.saturating_add(1)) - first)
            }
        }
    }
}

/// What the plan deletes in one sheet.
#[derive(Debug, Default, Clone)]
struct SheetCuts {
    deleted: bool,
    rows: Cuts,
    columns: Cuts,
}

/// The functions that give the error of the first argument that has one: an error that goes in
/// is the result. A function that is not here is left to the engine (it may trap the error, as
/// `IFERROR` does, or choose by it, as `IF` does).
const PASS_THE_ERROR_ON: &[&str] = &[
    "ABS",
    "AVERAGE",
    "CONCAT",
    "CONCATENATE",
    "DATE",
    "DAY",
    "EXP",
    "FV",
    "INT",
    "IPMT",
    "LEFT",
    "LEN",
    "LN",
    "LOG",
    "LOG10",
    "LOWER",
    "MAX",
    "MID",
    "MIN",
    "MOD",
    "MONTH",
    "NPER",
    "PMT",
    "POWER",
    "PPMT",
    "PRODUCT",
    "PV",
    "RATE",
    "RIGHT",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "SIGN",
    "SQRT",
    "SUM",
    "TEXT",
    "TRIM",
    "UPPER",
    "VALUE",
    "YEAR",
];

/// The functions whose result depends on where things are, and not only on what they read: a
/// deletion changes it whatever the formula points at.
const READ_POSITIONS: &[&str] = &[
    "ADDRESS",
    "AGGREGATE",
    "AREAS",
    "CELL",
    "COLUMN",
    "COLUMNS",
    "FORMULA",
    "INFO",
    "ISFORMULA",
    "ROW",
    "ROWS",
    "SHEET",
    "SHEETS",
    "SUBTOTAL",
];

/// What an expression gives, as far as the analysis can tell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Gives {
    /// Always the error `#REF!`.
    Ref,
    /// A value that is never an error.
    Value,
    /// Anything else.
    Unknown,
}

/// The formulas of the workbook the plan leaves stale, and who calculates them.
///
/// `book` is the workbook the pass runs on (the input of the pass that applies `changes`).
pub(crate) fn analyze(book: &Book, changes: &Changes) -> Verdict {
    // The cached values that are not numbers, booleans or empty are taken from the model the program
    // is built from, so the workbook is read once.
    let Ok((program, _, unsure)) =
        compile_keeping(book, &Options::default(), Mode::Recalculate, unsure_cells)
    else {
        return Verdict::Unknown;
    };
    if !program.cycles.is_empty() {
        return Verdict::Unknown;
    }
    let formulas = program
        .cells
        .iter()
        .filter(|cell| matches!(cell.kind, CellKind::Formula(_)))
        .count() as u64;
    let listed: u64 = program
        .formula_cells
        .iter()
        .map(|run| u64::from(run.last - run.first) + 1)
        .sum();
    if formulas != listed {
        // Formulas the compiler keeps as the values the file cached: what they read is not known.
        return Verdict::Unknown;
    }
    // A defined name that points into what the plan changes is read by the formulas that use it,
    // and the program has no trace of which they are: the name is not told apart there. A name that
    // points to cells the plan neither deletes nor shrinks around (`Data!$I$1` with rows
    // `Data!10:20` deleted) is read as it was.
    match book.defined_names() {
        Ok((names, _)) if names.values().all(|text| name_is_left_alone(text, changes)) => {}
        _ => return Verdict::Unknown,
    }
    let mut cuts: Vec<SheetCuts> = vec![SheetCuts::default(); program.sheets.len()];
    for (index, name) in program.sheets.iter().enumerate() {
        let cut = &mut cuts[index];
        cut.deleted = changes.sheets.iter().any(|s| s == name);
        cut.rows = Cuts::new(
            changes
                .rows
                .iter()
                .filter(|r| &r.sheet == name)
                .map(|r| (r.first, r.last))
                .collect(),
        );
        cut.columns = Cuts::new(
            changes
                .columns
                .iter()
                .filter(|r| &r.sheet == name)
                .map(|r| (r.first, r.last))
                .collect(),
        );
    }
    let converted = |at: CellRef| {
        let name = &program.sheets[usize::from(at.sheet)];
        changes
            .converted
            .iter()
            .any(|scope| &scope.sheet == name && scope.contains(at.row, at.col))
    };
    let mut stale = Stale {
        cuts: &cuts,
        assigned: HashMap::new(),
        index: BTreeSet::new(),
        unsure,
    };
    let mut known = Known::default();
    for cell in &program.cells {
        let CellKind::Formula(expr) = &cell.kind else {
            continue;
        };
        let at = cell.at;
        let own = &cuts[usize::from(at.sheet)];
        // A formula that goes with its row, column or sheet, or that becomes a value, is not a
        // formula of the output.
        if own.deleted || own.rows.contains(at.row) || own.columns.contains(at.col) || converted(at)
        {
            continue;
        }
        if !stale.is_stale(expr) {
            continue;
        }
        let certain = stale.result(expr) == Gives::Ref;
        stale.add(at, certain);
        let name = program.sheets[usize::from(at.sheet)].clone();
        let moved = (name, own.rows.moved(at.row), own.columns.moved(at.col));
        if certain {
            known.assigned.push(moved);
        } else {
            known.engine.push(moved);
        }
    }
    Verdict::Known(known)
}

/// The cells found stale so far (the formulas read before the one that is looked at).
struct Stale<'a> {
    cuts: &'a [SheetCuts],
    /// Whether each stale cell is given `#REF!` (`true`) or calculated by the engine.
    assigned: HashMap<CellRef, bool>,
    /// The same cells, ordered, to find the ones inside a range.
    index: BTreeSet<(u16, u32, u32)>,
    /// The cells of the input whose cached value is a text, an error or an array formula.
    unsure: BTreeSet<(u16, u32, u32)>,
}

impl Stale<'_> {
    fn add(&mut self, at: CellRef, assigned: bool) {
        self.assigned.insert(at, assigned);
        self.index.insert((at.sheet, at.row, at.col));
    }

    fn deleted(&self, at: CellRef) -> bool {
        let cut = &self.cuts[usize::from(at.sheet)];
        cut.deleted || cut.rows.contains(at.row) || cut.columns.contains(at.col)
    }

    /// Whether every cell of the range goes (a sheet, or one band of rows or columns holds it all).
    fn range_gone(&self, r: &Rect) -> bool {
        let cut = &self.cuts[usize::from(r.sheet)];
        cut.deleted || cut.rows.covers(r.r0, r.r1) || cut.columns.covers(r.c0, r.c1)
    }

    /// Whether the range loses some cells without going.
    fn range_shrinks(&self, r: &Rect) -> bool {
        let cut = &self.cuts[usize::from(r.sheet)];
        !self.range_gone(r) && (cut.rows.touches(r.r0, r.r1) || cut.columns.touches(r.c0, r.c1))
    }

    /// Whether some stale cell is inside the range.
    fn range_reads_stale(&self, r: &Rect) -> bool {
        self.index
            .range((r.sheet, r.r0, 0)..=(r.sheet, r.r1, u32::MAX))
            .any(|&(_, _, col)| (r.c0..=r.c1).contains(&col))
    }

    /// Whether the result of the formula is out of date after the changes: it reads what they
    /// change, or something stale, or what depends on where it is.
    fn is_stale(&self, expr: &Expr) -> bool {
        let mut stale = false;
        expr.walk(&mut |e| match e {
            Expr::Cell(at) => {
                stale |= self.deleted(*at) || self.assigned.contains_key(at);
            }
            Expr::Range(r) => {
                stale |= self.range_gone(r) || self.range_shrinks(r) || self.range_reads_stale(r);
            }
            Expr::Call(id, _) => {
                let def = id.def();
                stale |= def.dynamic || READ_POSITIONS.contains(&def.name);
            }
            Expr::Unknown(_) => stale = true,
            _ => {}
        });
        stale
    }

    /// What the expression gives once the cells deleted are gone, as far as is certain.
    fn result(&self, expr: &Expr) -> Gives {
        match expr {
            Expr::Const(Value::Error(_) | Value::Text(_)) => Gives::Unknown,
            Expr::Const(_) | Expr::Missing => Gives::Value,
            Expr::Cell(at) => {
                if self.deleted(*at) || self.assigned.get(at) == Some(&true) {
                    Gives::Ref
                } else if self.assigned.contains_key(at)
                    || self.unsure.contains(&(at.sheet, at.row, at.col))
                {
                    Gives::Unknown
                } else {
                    // The cached result of a cell the plan leaves alone is valid: a number, a
                    // boolean or empty.
                    Gives::Value
                }
            }
            Expr::Range(r) => {
                // A range as an operand may be read through an implicit intersection that fails:
                // only a range that is gone is certain.
                if self.range_gone(r) {
                    Gives::Ref
                } else {
                    Gives::Unknown
                }
            }
            Expr::Unary(_, a) => self.result(a),
            Expr::Binary(op, a, b) => {
                // Operators evaluate the left operand first: its error is the result; the right
                // operand's error is the result only when the left one is no error.
                let _: &BinOp = op;
                self.first_error(&[a, b])
            }
            Expr::Call(id, args) if PASS_THE_ERROR_ON.contains(&id.def().name) => {
                let args: Vec<&Expr> = args.iter().collect();
                self.first_error(&args)
            }
            Expr::Call(..) | Expr::Unknown(_) => Gives::Unknown,
        }
    }

    /// The result of something that gives the error of the first operand that has one.
    fn first_error(&self, operands: &[&Expr]) -> Gives {
        for operand in operands {
            match self.result(operand) {
                Gives::Ref => return Gives::Ref,
                Gives::Value => {}
                Gives::Unknown => return Gives::Unknown,
            }
        }
        Gives::Unknown
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn deleting_rows(sheet: &str, first: u32, last: u32) -> Changes {
        Changes {
            rows: vec![RowRange {
                sheet: sheet.to_string(),
                first,
                last,
            }],
            ..Changes::default()
        }
    }

    #[test]
    fn a_name_that_points_above_the_deleted_rows_of_its_sheet_is_left_alone() {
        let changes = deleting_rows("Data", 9, 19);
        assert!(name_is_left_alone("Data!$I$1", &changes));
        assert!(name_is_left_alone("'Data'!$I$30:$I$31", &changes));
    }

    #[test]
    fn a_name_that_points_into_the_deleted_rows_or_to_a_range_that_loses_some_is_not() {
        let changes = deleting_rows("Data", 9, 19);
        assert!(!name_is_left_alone("Data!$I$10", &changes));
        assert!(!name_is_left_alone("Data!$I$5:$I$12", &changes));
        assert!(!name_is_left_alone("Data!$I:$I", &changes));
    }

    #[test]
    fn a_name_on_another_sheet_and_a_name_that_is_not_a_plain_reference() {
        let changes = deleting_rows("Data", 9, 19);
        assert!(name_is_left_alone("Other!$A$10", &changes));
        assert!(!name_is_left_alone("OFFSET(Data!$A$1,2,0)", &changes));
    }
}
