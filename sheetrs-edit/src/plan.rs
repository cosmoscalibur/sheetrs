//! A plan: a sequence of operations on one workbook, and the passes over the file that run it.
//!
//! The operations are written as a sequence, and each one addresses the workbook **as the previous
//! ones left it**: after `delete_rows("S!3")`, row 3 is what was row 4. That is what running them
//! one at a time gives, and it is the meaning of a plan. Running them one at a time costs a read
//! and a write of the file each. [`EditPlan::passes`] resolves the sequence into the fewest passes
//! that give the same workbook, and [`EditPlan::execute`] runs those. Nothing is read until
//! `execute`: a plan is built, inspected ([`EditPlan::explain`]) and run, once or several times.
//!
//! One pass applies what its operations ask in the order the writer does, all against the input
//! of that pass: formulas converted to values first, then, in each sheet, the rows and then the
//! columns deleted from the last to the first, then the named ranges removed, then the sheets
//! removed. Operations can share a pass when that order gives what their sequence gives:
//!
//! - Deletions of one sheet and axis share it when they do not touch: the positions of a later
//!   one are translated to the input of the pass, through the deletions before it. Two that touch
//!   (`delete_rows("S!3")` twice is the rows 3 and 4 of the input) are not one band to the writer,
//!   which refuses a band that removes the anchor of a shape together with the rows before it,
//!   so the second starts a new pass.
//! - Deletions in different sheets, or on rows and on columns, do not depend on each other: a
//!   reference moves only with the rows and columns of the sheet it names (a reference that spans
//!   sheets, `Sheet1:Sheet3!A1`, would tie them, and is not supported: issue #75).
//! - Removing names and sheets commutes with the deletions.
//! - A conversion reads the values that every other change would have dropped or rewritten, so
//!   it cannot move before one: a conversion after any other change starts a new pass.
//!
//! Whether the output is marked to be recalculated when it is opened is the plan's to decide
//! ([`Recalculation`]), not something to remember to ask for. A deletion or a removal leaves the
//! cached results of formulas that read what went (their text may not change, and so their value
//! is stale), so a plan with one marks the output, as the last thing it does before the file is
//! written: it changes how the file is opened, not what it holds, so it comes after every
//! operation that reads and rewrites the content. (An ODS applies it by dropping every cached
//! result; in the middle of a sequence it would change what a later conversion finds.) Whoever
//! knows better can ask for the mark always or never.
//!
//! A plan that addresses a sheet an earlier operation removed is an error when it is planned.
//!
//! The plan is a description, not a model of the workbook: it holds the operations (a few bands
//! and names), and each pass streams the file once, rewriting each cell as it goes by. The memory
//! of a pass does not depend on the size of the workbook.

use crate::writer::structure::Bands;
use crate::writer::{self, ColumnRange, RowRange, ValueScope, ValueTarget, WorkbookModifications};
use anyhow::{Context, Result, bail};
use sheetrs_core::reader::Source;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::path::Path;

/// One operation on a workbook, on the workbook as the operations before it leave it. The
/// ranges of one operation are simultaneous: `DeleteRows(["S!3", "S!4"])` deletes two rows.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Operation {
    /// Remove the sheets of these names.
    RemoveSheets(Vec<String>),
    /// Remove the defined names (or database ranges).
    RemoveNamedRanges(Vec<String>),
    /// Replace the formulas of the scopes (`Sheet`, `Sheet!B:D`, `Sheet!3:10`) by their cached
    /// values.
    ConvertToValues(Vec<ValueScope>),
    /// Delete the rows of the ranges; references to them become `#REF!`.
    DeleteRows(Vec<RowRange>),
    /// Delete the columns of the ranges.
    DeleteColumns(Vec<ColumnRange>),
}

impl fmt::Display for Operation {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fn list<T: fmt::Display>(items: &[T]) -> String {
            items
                .iter()
                .map(ToString::to_string)
                .collect::<Vec<_>>()
                .join(", ")
        }
        match self {
            Self::RemoveSheets(names) => write!(f, "remove sheets {}", list(names)),
            Self::RemoveNamedRanges(names) => write!(f, "remove named ranges {}", list(names)),
            Self::ConvertToValues(scopes) => {
                write!(f, "convert formulas to values in {}", list(scopes))
            }
            Self::DeleteRows(rows) => write!(f, "delete rows {}", list(rows)),
            Self::DeleteColumns(columns) => write!(f, "delete columns {}", list(columns)),
        }
    }
}

impl Operation {
    fn is_empty(&self) -> bool {
        match self {
            Self::RemoveSheets(items) | Self::RemoveNamedRanges(items) => items.is_empty(),
            Self::ConvertToValues(items) => items.is_empty(),
            Self::DeleteRows(items) => items.is_empty(),
            Self::DeleteColumns(items) => items.is_empty(),
        }
    }
}

/// Whether the output is marked so that the application that opens it recalculates every formula.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum Recalculation {
    /// The plan decides: it marks the output when an operation leaves cached results stale (rows,
    /// columns, sheets or names removed), and leaves it alone otherwise.
    #[default]
    Auto,
    /// Always mark it.
    Always,
    /// Never mark it, whatever the operations: the cached results are left as they are.
    Never,
}

/// A sequence of operations on one workbook, and what to do about recalculation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EditPlan {
    operations: Vec<Operation>,
    recalculation: Recalculation,
}

impl EditPlan {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds an operation after the ones there are. An operation with nothing in it is not added.
    pub fn push(&mut self, operation: Operation) -> &mut Self {
        if !operation.is_empty() {
            self.operations.push(operation);
        }
        self
    }

    pub fn remove_sheets(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        self.push(Operation::RemoveSheets(
            names.into_iter().map(Into::into).collect(),
        ))
    }

    pub fn remove_named_ranges(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        self.push(Operation::RemoveNamedRanges(
            names.into_iter().map(Into::into).collect(),
        ))
    }

    pub fn convert_to_values(&mut self, scopes: impl IntoIterator<Item = ValueScope>) -> &mut Self {
        self.push(Operation::ConvertToValues(scopes.into_iter().collect()))
    }

    pub fn delete_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self {
        self.push(Operation::DeleteRows(rows.into_iter().collect()))
    }

    pub fn delete_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self {
        self.push(Operation::DeleteColumns(columns.into_iter().collect()))
    }

    /// Overrides what the plan decides about marking the output to be recalculated when opened
    /// (see [`Recalculation`]).
    pub fn recalculation(&mut self, recalculation: Recalculation) -> &mut Self {
        self.recalculation = recalculation;
        self
    }

    /// Whether the output will be marked to be recalculated when it is opened: what was asked for
    /// or, by default, what the operations of the plan require.
    #[must_use]
    pub fn marks_recalculation(&self) -> bool {
        match self.recalculation {
            Recalculation::Always => true,
            Recalculation::Never => false,
            Recalculation::Auto => self
                .operations
                .iter()
                .any(|operation| !matches!(operation, Operation::ConvertToValues(_))),
        }
    }

    /// The plan for operations that are all written against the **input** workbook, in any order,
    /// as a command line gives them: they are grouped by kind into the order of a pass, so the
    /// plan has one pass whatever the order they came in, and the same operations give the same
    /// plan. This is the declarative counterpart of the sequence: the same operations given as a
    /// sequence would need their positions translated by whoever wrote them.
    pub fn against_input(operations: impl IntoIterator<Item = Operation>) -> Self {
        fn unique<T: PartialEq>(into: &mut Vec<T>, from: Vec<T>) {
            for item in from {
                if !into.contains(&item) {
                    into.push(item);
                }
            }
        }
        let (mut converts, mut rows, mut columns) = (Vec::new(), Vec::new(), Vec::new());
        let (mut names, mut sheets) = (Vec::new(), Vec::new());
        for operation in operations {
            match operation {
                Operation::ConvertToValues(scopes) => unique(&mut converts, scopes),
                Operation::DeleteRows(ranges) => unique(&mut rows, ranges),
                Operation::DeleteColumns(ranges) => unique(&mut columns, ranges),
                Operation::RemoveNamedRanges(n) => names.extend(n),
                Operation::RemoveSheets(n) => sheets.extend(n),
            }
        }
        names.sort();
        names.dedup();
        sheets.sort();
        sheets.dedup();
        converts.sort_by_key(|scope| (scope.sheet.clone(), target_key(scope.target)));
        rows.sort();
        columns.sort();
        let mut plan = Self::new();
        plan.convert_to_values(converts)
            .delete_rows(rows)
            .delete_columns(columns)
            .remove_named_ranges(names)
            .remove_sheets(sheets);
        plan
    }

    /// The operations as they were written.
    #[must_use]
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    /// Whether there is nothing to do: no operation, and no mark to recalculate asked for.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.operations.is_empty() && self.recalculation != Recalculation::Always
    }

    /// The passes over the file that run the plan: the operations of each, normalized and in the
    /// order of the pass (see the module documentation), against the input of that pass. Their
    /// number is what the plan costs in reads and writes of the file; `against_input` gives one.
    /// The mark to recalculate is not in them: it goes with the last (`marks_recalculation`).
    ///
    /// # Errors
    ///
    /// When an operation addresses a sheet that an earlier one removed.
    pub fn passes(&self) -> Result<Vec<Vec<Operation>>> {
        let mut passes = Vec::new();
        let mut current = Pass::default();
        let mut removed: BTreeSet<String> = BTreeSet::new();
        for operation in &self.operations {
            for sheet in sheets_of(operation) {
                if removed.contains(sheet) {
                    bail!(
                        "'{operation}' addresses the sheet '{sheet}', which an earlier operation removed"
                    );
                }
            }
            if !current.admits(operation) {
                passes.push(std::mem::take(&mut current).into_operations());
            }
            current.add(operation);
            if let Operation::RemoveSheets(names) = operation {
                removed.extend(names.iter().cloned());
            }
        }
        if !current.is_empty() {
            passes.push(current.into_operations());
        }
        Ok(passes)
    }

    /// The passes as text, one operation per line, for a person to read.
    ///
    /// # Errors
    ///
    /// As [`passes`](Self::passes).
    pub fn explain(&self) -> Result<String> {
        use std::fmt::Write as _;
        let passes = self.passes()?;
        let mut out = String::new();
        for (i, pass) in passes.iter().enumerate() {
            let _ = writeln!(out, "pass {} of {}:", i + 1, passes.len());
            for operation in pass {
                let _ = writeln!(out, "  {operation}");
            }
        }
        if self.marks_recalculation() {
            out.push_str("last, on the output: recalculate when opened\n");
        }
        Ok(out)
    }

    /// Runs the plan on `input` and writes `output`, which must not be the input file. Each pass
    /// reads and writes the file once, as a stream; a plan of several passes keeps the file between
    /// them in a temporary folder, and so does a source of bytes (the writer works on files). A
    /// failed run leaves no `output` behind. An empty plan is an error: there is nothing to do,
    /// and a copy of the file is not what was asked for.
    ///
    /// # Errors
    ///
    /// When the plan is empty or cannot be planned ([`passes`](Self::passes)), when a file cannot
    /// be read or written, or when a pass fails (a sheet that does not exist, a formula with a 3D
    /// reference, a change the writer cannot make faithfully).
    pub fn execute(&self, input: &Source, output: impl AsRef<Path>) -> Result<()> {
        if self.is_empty() {
            bail!("No operations to apply");
        }
        let output = output.as_ref();
        let mut passes = self.passes()?;
        if passes.is_empty() {
            // Only the mark to recalculate was asked for: one pass that does nothing else.
            passes.push(Vec::new());
        }
        let extension = input.format().extension();
        let folder = (passes.len() > 1 || input.path().is_none())
            .then(tempfile::tempdir)
            .transpose()
            .context("Cannot create a temporary folder")?;
        let mut current = match (input, &folder) {
            (Source::File { path, .. }, _) => path.clone(),
            (Source::Bytes { data, .. }, Some(folder)) => {
                let path = folder.path().join(format!("input.{extension}"));
                std::fs::write(&path, data).context("Cannot write the workbook to edit")?;
                path
            }
            (Source::Bytes { .. }, None) => unreachable!("a source of bytes has a folder"),
        };
        for (i, pass) in passes.iter().enumerate() {
            let last = i + 1 == passes.len();
            let target = match (&folder, last) {
                (Some(folder), false) => folder.path().join(format!("pass{i}.{extension}")),
                _ => output.to_path_buf(),
            };
            let mut mods = modifications_of(pass);
            mods.recalculate_on_open = last && self.marks_recalculation();
            writer::modify_workbook(&current, &target, &mods)
                .with_context(|| format!("Failed in pass {} of {}", i + 1, passes.len()))?;
            current = target;
        }
        Ok(())
    }
}

impl Extend<Operation> for EditPlan {
    fn extend<T: IntoIterator<Item = Operation>>(&mut self, operations: T) {
        for operation in operations {
            self.push(operation);
        }
    }
}

impl FromIterator<Operation> for EditPlan {
    fn from_iter<T: IntoIterator<Item = Operation>>(operations: T) -> Self {
        let mut plan = Self::new();
        plan.extend(operations);
        plan
    }
}

impl IntoIterator for EditPlan {
    type Item = Operation;
    type IntoIter = std::vec::IntoIter<Operation>;

    fn into_iter(self) -> Self::IntoIter {
        self.operations.into_iter()
    }
}

impl<'a> IntoIterator for &'a EditPlan {
    type Item = &'a Operation;
    type IntoIter = std::slice::Iter<'a, Operation>;

    fn into_iter(self) -> Self::IntoIter {
        self.operations.iter()
    }
}

fn target_key(target: ValueTarget) -> (u8, u32, u32) {
    match target {
        ValueTarget::Sheet => (0, 0, 0),
        ValueTarget::Columns(first, last) => (1, first, last),
        ValueTarget::Rows(first, last) => (2, first, last),
    }
}

/// The sheets an operation addresses, which must exist when it runs. Removing a sheet the workbook
/// does not have is not an error of the writer (it removes nothing), so it is not one here.
fn sheets_of(operation: &Operation) -> Vec<&str> {
    match operation {
        Operation::ConvertToValues(scopes) => scopes.iter().map(|s| s.sheet.as_str()).collect(),
        Operation::DeleteRows(rows) => rows.iter().map(|r| r.sheet.as_str()).collect(),
        Operation::DeleteColumns(columns) => columns.iter().map(|c| c.sheet.as_str()).collect(),
        Operation::RemoveSheets(_) | Operation::RemoveNamedRanges(_) => Vec::new(),
    }
}

/// The operations of one pass as the writer takes them.
fn modifications_of(pass: &[Operation]) -> WorkbookModifications {
    let mut mods = WorkbookModifications::default();
    for operation in pass {
        match operation {
            Operation::RemoveSheets(names) => mods
                .remove_sheets
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::RemoveNamedRanges(names) => mods
                .remove_named_ranges
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::ConvertToValues(scopes) => {
                mods.convert_to_values.extend(scopes.iter().cloned());
            }
            Operation::DeleteRows(rows) => {
                mods.delete_rows
                    .extend(rows.iter().cloned().map(ValueScope::from));
            }
            Operation::DeleteColumns(columns) => {
                mods.delete_columns
                    .extend(columns.iter().cloned().map(ValueScope::from));
            }
        }
    }
    mods
}

/// The operations of one pass while it is built: deletions by sheet as the bands of the input of
/// the pass.
#[derive(Default)]
struct Pass {
    converts: Vec<ValueScope>,
    rows: BTreeMap<String, Vec<(u32, u32)>>,
    columns: BTreeMap<String, Vec<(u32, u32)>>,
    names: BTreeSet<String>,
    sheets: BTreeSet<String>,
    /// Whether anything but a conversion is in the pass.
    changes: bool,
}

type Bounds<'a> = Vec<(&'a str, u32, u32)>;

fn row_bounds(rows: &[RowRange]) -> Bounds<'_> {
    rows.iter()
        .map(|r| (r.sheet.as_str(), r.first, r.last))
        .collect()
}

fn column_bounds(columns: &[ColumnRange]) -> Bounds<'_> {
    columns
        .iter()
        .map(|c| (c.sheet.as_str(), c.first, c.last))
        .collect()
}

impl Pass {
    fn is_empty(&self) -> bool {
        !self.changes && self.converts.is_empty()
    }

    /// Whether `operation` can join the pass: a conversion cannot follow another kind of change,
    /// and a deletion that touches one already there cannot either (see the module
    /// documentation).
    fn admits(&self, operation: &Operation) -> bool {
        match operation {
            Operation::ConvertToValues(_) => !self.changes,
            Operation::DeleteRows(rows) => !Self::touches(&self.rows, &row_bounds(rows)),
            Operation::DeleteColumns(columns) => {
                !Self::touches(&self.columns, &column_bounds(columns))
            }
            _ => true,
        }
    }

    fn touches(by_sheet: &BTreeMap<String, Vec<(u32, u32)>>, bounds: &Bounds<'_>) -> bool {
        bounds.iter().any(|&(sheet, first, last)| {
            let Some(ranges) = by_sheet.get(sheet) else {
                return false;
            };
            let bands = Bands::new(ranges.clone());
            let (first, last) = (bands.input_index(first), bands.input_index(last));
            ranges
                .iter()
                .any(|&(x, y)| first <= y.saturating_add(1) && x <= last.saturating_add(1))
        })
    }

    fn add(&mut self, operation: &Operation) {
        match operation {
            Operation::ConvertToValues(scopes) => {
                for scope in scopes {
                    if !self.converts.contains(scope) {
                        self.converts.push(scope.clone());
                    }
                }
            }
            Operation::DeleteRows(rows) => Self::delete(&mut self.rows, &row_bounds(rows)),
            Operation::DeleteColumns(columns) => {
                Self::delete(&mut self.columns, &column_bounds(columns));
            }
            Operation::RemoveNamedRanges(names) => self.names.extend(names.iter().cloned()),
            Operation::RemoveSheets(names) => self.sheets.extend(names.iter().cloned()),
        }
        if !matches!(operation, Operation::ConvertToValues(_)) {
            self.changes = true;
        }
    }

    /// Adds deletions given against the workbook as the pass has left it so far: each position is
    /// translated to the input of the pass through the deletions already in it. The ranges of one
    /// operation are simultaneous, so they are all translated before any of them is added.
    fn delete(by_sheet: &mut BTreeMap<String, Vec<(u32, u32)>>, bounds: &Bounds<'_>) {
        let translated: Vec<(String, u32, u32)> = bounds
            .iter()
            .map(|&(sheet, first, last)| {
                let bands = Bands::new(by_sheet.get(sheet).cloned().unwrap_or_default());
                (
                    sheet.to_string(),
                    bands.input_index(first),
                    bands.input_index(last),
                )
            })
            .collect();
        for (sheet, first, last) in translated {
            by_sheet.entry(sheet).or_default().push((first, last));
        }
    }

    fn into_operations(self) -> Vec<Operation> {
        let bands = |by_sheet: BTreeMap<String, Vec<(u32, u32)>>| -> Vec<(String, u32, u32)> {
            by_sheet
                .into_iter()
                .flat_map(|(sheet, ranges)| {
                    Bands::new(ranges)
                        .ranges()
                        .map(|(first, last)| (sheet.clone(), first, last))
                        .collect::<Vec<_>>()
                })
                .collect()
        };
        let mut plan = EditPlan::new();
        plan.convert_to_values(self.converts)
            .delete_rows(
                bands(self.rows)
                    .into_iter()
                    .map(|(sheet, first, last)| RowRange { sheet, first, last }),
            )
            .delete_columns(
                bands(self.columns)
                    .into_iter()
                    .map(|(sheet, first, last)| ColumnRange { sheet, first, last }),
            )
            .remove_named_ranges(self.names)
            .remove_sheets(self.sheets);
        plan.operations
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
        let mut plan = EditPlan::new();
        plan.remove_sheets(Vec::<String>::new())
            .convert_to_values([])
            .delete_rows([])
            .delete_columns([]);
        assert!(plan.is_empty());
    }

    #[test]
    fn a_plan_is_collected_and_iterated_like_a_collection() {
        let plan: EditPlan = [
            Operation::RemoveSheets(vec!["A".into()]),
            Operation::DeleteRows(rows(&["T!3"])),
        ]
        .into_iter()
        .collect();
        let mut extended = EditPlan::default();
        extended.extend(plan.clone());
        assert_eq!(extended, plan);
        assert_eq!((&plan).into_iter().count(), 2);
        assert_eq!(plan.into_iter().count(), 2);
    }

    #[test]
    fn later_positions_are_those_the_earlier_deletions_left() {
        // After row 3 goes, row 4 of the input is the new row 3, and 6 is the new 5.
        let mut plan = EditPlan::new();
        plan.delete_rows(rows(&["T!3"]))
            .delete_rows(rows(&["T!4"]))
            .delete_rows(rows(&["T!10"]));
        assert_eq!(
            described(&plan.passes().unwrap()),
            [["delete rows T!3:3, T!5:5, T!12:12"]]
        );
    }

    #[test]
    fn deletions_that_touch_are_not_one_band_to_the_writer_so_they_do_not_share_a_pass() {
        // Row 3 twice are the input's rows 3 and 4, which touch: one operation at a time.
        let mut plan = EditPlan::new();
        plan.delete_rows(rows(&["T!3"])).delete_rows(rows(&["T!3"]));
        assert_eq!(
            described(&plan.passes().unwrap()),
            [["delete rows T!3:3"], ["delete rows T!3:3"]]
        );
        // Written together they are one deletion.
        let mut plan = EditPlan::new();
        plan.delete_rows(rows(&["T!3", "T!4"]));
        assert_eq!(described(&plan.passes().unwrap()), [["delete rows T!3:4"]]);
    }

    #[test]
    fn deletions_names_and_sheets_share_a_pass_and_a_conversion_before_them_too() {
        let mut plan = EditPlan::new();
        plan.convert_to_values(scopes(&["F"]))
            .delete_columns(columns(&["T!D"]))
            .delete_rows(rows(&["U!2:3"]))
            .remove_named_ranges(["N"])
            .remove_sheets(["V"]);
        assert_eq!(
            described(&plan.passes().unwrap()),
            [[
                "convert formulas to values in F",
                "delete rows U!2:3",
                "delete columns T!D:D",
                "remove named ranges N",
                "remove sheets V",
            ]]
        );
    }

    #[test]
    fn a_conversion_after_a_change_starts_a_new_pass() {
        let mut plan = EditPlan::new();
        plan.delete_rows(rows(&["T!3"]))
            .convert_to_values(scopes(&["F"]))
            .delete_rows(rows(&["T!1"]));
        assert_eq!(
            described(&plan.passes().unwrap()),
            [
                vec!["delete rows T!3:3"],
                vec!["convert formulas to values in F", "delete rows T!1:1"]
            ]
        );
        assert!(plan.explain().unwrap().starts_with("pass 1 of 2:"));
    }

    #[test]
    fn the_plan_marks_the_output_for_recalculation_when_an_operation_leaves_stale_results() {
        let marks = |build: &dyn Fn(&mut EditPlan)| {
            let mut plan = EditPlan::new();
            build(&mut plan);
            plan.marks_recalculation()
        };
        // Deleting or removing leaves formulas whose text did not change with a stale result.
        assert!(marks(&|p| {
            p.delete_rows(rows(&["T!3"]));
        }));
        assert!(marks(&|p| {
            p.delete_columns(columns(&["T!B"]));
        }));
        assert!(marks(&|p| {
            p.remove_sheets(["U"]);
        }));
        assert!(marks(&|p| {
            p.remove_named_ranges(["N"]);
        }));
        // Converting formulas to the values they have changes no other cell.
        assert!(!marks(&|p| {
            p.convert_to_values(scopes(&["F"]));
        }));
        assert!(!marks(&|_| {}));
    }

    #[test]
    fn the_mark_can_be_forced_or_avoided_and_is_never_an_operation_of_the_sequence() {
        let mut plan = EditPlan::new();
        plan.convert_to_values(scopes(&["F"]))
            .delete_rows(rows(&["T!3"]));
        // It splits nothing and is in no pass; it is said apart, as the last thing.
        assert_eq!(plan.passes().unwrap().len(), 1);
        assert!(plan.marks_recalculation());
        assert!(
            plan.explain()
                .unwrap()
                .ends_with("recalculate when opened\n")
        );
        plan.recalculation(Recalculation::Never);
        assert!(!plan.marks_recalculation());
        assert!(!plan.explain().unwrap().contains("recalculate"));
        // Forced, it makes a plan of its own: the file is marked and nothing else changes.
        let mut only_the_mark = EditPlan::new();
        assert!(only_the_mark.is_empty());
        only_the_mark.recalculation(Recalculation::Always);
        assert!(!only_the_mark.is_empty());
        assert!(only_the_mark.passes().unwrap().is_empty());
    }

    #[test]
    fn an_operation_on_a_removed_sheet_is_an_error_when_planned() {
        let mut plan = EditPlan::new();
        plan.remove_sheets(["U"]).delete_rows(rows(&["U!2"]));
        let error = plan.passes().unwrap_err().to_string();
        assert!(
            error.contains("'U'") && error.contains("removed"),
            "{error}"
        );
    }

    #[test]
    fn against_input_gives_one_pass_whatever_the_order() {
        let build = |order: &[usize]| {
            let ops = [
                Operation::RemoveSheets(vec!["U".into()]),
                Operation::DeleteRows(rows(&["T!8:9"])),
                Operation::ConvertToValues(scopes(&["F"])),
                Operation::DeleteRows(rows(&["T!2", "V!6"])),
                Operation::DeleteColumns(columns(&["T!D"])),
                Operation::RemoveNamedRanges(vec!["N".into()]),
            ];
            EditPlan::against_input(order.iter().map(|&i| ops[i].clone()))
        };
        let one = build(&[0, 1, 2, 3, 4, 5]);
        let other = build(&[5, 4, 3, 2, 1, 0]);
        assert_eq!(one, other);
        let passes = one.passes().unwrap();
        assert_eq!(passes.len(), 1);
        assert_eq!(
            described(&passes),
            [[
                "convert formulas to values in F",
                "delete rows T!2:2, T!8:9, V!6:6",
                "delete columns T!D:D",
                "remove named ranges N",
                "remove sheets U",
            ]]
        );
    }

    #[test]
    fn an_empty_plan_has_no_passes_and_does_not_execute() {
        let plan = EditPlan::new();
        assert!(plan.passes().unwrap().is_empty());
        let input = Source::from_file("in.xlsx").unwrap();
        let error = plan.execute(&input, "out.xlsx").unwrap_err();
        assert!(error.to_string().contains("No operations"));
    }
}
