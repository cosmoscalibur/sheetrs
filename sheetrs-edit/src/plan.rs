//! A plan: a sequence of operations on one workbook, and the passes over the file that run it.
//!
//! The operations are written as a sequence, and each one addresses the workbook **as the previous
//! ones left it**: after `delete_rows("S!3")`, row 3 is what was row 4. That is what running them
//! one at a time gives, and it is the meaning of a plan. Running them one at a time costs a read
//! and a write of the file each. [`EditPlan::passes`] resolves the sequence into the fewest passes
//! that give the same workbook, and [`EditPlan::execute`] runs those, so an eager caller (one
//! operation at a time) and a lazy one (the whole plan at once) get the same file.
//!
//! One pass applies what its operations ask in the order the writer does, all against the input
//! of that pass: formulas converted to values first, then, in each sheet, the rows and then the
//! columns deleted from the last to the first, then the named ranges removed, then the sheets
//! removed. Operations can share a pass when that order gives what their sequence gives:
//!
//! - Deletions of one sheet and axis share it: the positions of a later one are translated to
//!   the input of the pass, through the deletions before it (`delete_rows("S!3")` twice is the
//!   rows 3 and 4 of the input).
//! - Deletions in different sheets, or on rows and on columns, do not depend on each other: a
//!   reference moves only with the rows and columns of the sheet it names (a reference that spans
//!   sheets, `Sheet1:Sheet3!A1`, would tie them, and is not supported: issue #75).
//! - Removing names and sheets, and the mark to recalculate, commute with the deletions.
//! - A conversion reads the values that every other change would have dropped or rewritten, so
//!   it cannot move before one: a conversion after any other change starts a new pass.
//!
//! A plan that addresses a sheet an earlier operation removed is an error when it is planned.
//!
//! The plan is a description, not a model of the workbook: it holds the operations (a few bands
//! and names), and each pass streams the file once, rewriting each cell as it goes by. The memory
//! of a pass does not depend on the size of the workbook.

use crate::writer::structure::Bands;
use crate::writer::{self, ValueScope, ValueTarget, WorkbookModifications};
use anyhow::{Context, Result, bail};
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::path::Path;

/// One operation on a workbook, on the workbook as the operations before it leave it. The
/// scopes of one operation are simultaneous: `DeleteRows(["S!3", "S!4"])` deletes two rows.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Operation {
    /// Remove the sheets of these names.
    RemoveSheets(Vec<String>),
    /// Remove the defined names (or database ranges).
    RemoveNamedRanges(Vec<String>),
    /// Replace the formulas of the scopes by their cached values.
    ConvertToValues(Vec<ValueScope>),
    /// Delete the rows of the scopes (`Sheet!3:10`); references to them become `#REF!`.
    DeleteRows(Vec<ValueScope>),
    /// Delete the columns of the scopes (`Sheet!B:D`).
    DeleteColumns(Vec<ValueScope>),
    /// Mark the output so the application that opens it recalculates every formula.
    RecalculateOnOpen,
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
            Self::DeleteRows(scopes) => write!(f, "delete rows {}", list(scopes)),
            Self::DeleteColumns(scopes) => write!(f, "delete columns {}", list(scopes)),
            Self::RecalculateOnOpen => write!(f, "recalculate when opened"),
        }
    }
}

/// A sequence of operations on one workbook.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EditPlan {
    operations: Vec<Operation>,
}

impl EditPlan {
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds an operation after the ones there are. Rows are deleted with row ranges
    /// (`Sheet!3:10`) and columns with column ranges (`Sheet!B:D`): any other scope is refused
    /// here, not when the plan runs. An operation with nothing in it is not added.
    pub fn push(&mut self, operation: Operation) -> Result<&mut Self> {
        let empty = match &operation {
            Operation::RemoveSheets(names) | Operation::RemoveNamedRanges(names) => {
                names.is_empty()
            }
            Operation::ConvertToValues(scopes) => scopes.is_empty(),
            Operation::DeleteRows(scopes) => {
                if let Some(scope) = scopes
                    .iter()
                    .find(|scope| !matches!(scope.target, ValueTarget::Rows(..)))
                {
                    bail!("'{scope}' is not a row range such as 'Sheet!3:10'");
                }
                scopes.is_empty()
            }
            Operation::DeleteColumns(scopes) => {
                if let Some(scope) = scopes
                    .iter()
                    .find(|scope| !matches!(scope.target, ValueTarget::Columns(..)))
                {
                    bail!("'{scope}' is not a column range such as 'Sheet!B:D'");
                }
                scopes.is_empty()
            }
            Operation::RecalculateOnOpen => false,
        };
        if !empty {
            self.operations.push(operation);
        }
        Ok(self)
    }

    pub fn remove_sheets(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        let names = names.into_iter().map(Into::into).collect();
        self.push(Operation::RemoveSheets(names))
            .expect("a sheet name is always valid")
    }

    pub fn remove_named_ranges(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        let names = names.into_iter().map(Into::into).collect();
        self.push(Operation::RemoveNamedRanges(names))
            .expect("a name is always valid")
    }

    pub fn convert_to_values(&mut self, scopes: impl IntoIterator<Item = ValueScope>) -> &mut Self {
        self.push(Operation::ConvertToValues(scopes.into_iter().collect()))
            .expect("every scope can be converted")
    }

    pub fn delete_rows(
        &mut self,
        scopes: impl IntoIterator<Item = ValueScope>,
    ) -> Result<&mut Self> {
        self.push(Operation::DeleteRows(scopes.into_iter().collect()))
    }

    pub fn delete_columns(
        &mut self,
        scopes: impl IntoIterator<Item = ValueScope>,
    ) -> Result<&mut Self> {
        self.push(Operation::DeleteColumns(scopes.into_iter().collect()))
    }

    pub fn recalculate_on_open(&mut self) -> &mut Self {
        self.push(Operation::RecalculateOnOpen)
            .expect("the mark is always valid")
    }

    /// The plan for operations that are all written against the **input** workbook, in any order,
    /// as a command line gives them: its operations are grouped by kind into the order of a
    /// pass, so the plan has one pass whatever the order they came in. This is the declarative
    /// counterpart of the sequence: the same operations given as a sequence would need their
    /// positions translated by whoever wrote them.
    pub fn against_input(operations: impl IntoIterator<Item = Operation>) -> Result<Self> {
        let (mut converts, mut rows, mut cols) = (Vec::new(), Vec::new(), Vec::new());
        let (mut names, mut sheets, mut recalc) = (Vec::new(), Vec::new(), false);
        let unique = |into: &mut Vec<ValueScope>, from: Vec<ValueScope>| {
            for scope in from {
                if !into.contains(&scope) {
                    into.push(scope);
                }
            }
        };
        let mut checker = Self::new();
        for operation in operations {
            checker.push(operation.clone())?;
            match operation {
                Operation::ConvertToValues(scopes) => unique(&mut converts, scopes),
                Operation::DeleteRows(scopes) => unique(&mut rows, scopes),
                Operation::DeleteColumns(scopes) => unique(&mut cols, scopes),
                Operation::RemoveNamedRanges(n) => names.extend(n),
                Operation::RemoveSheets(n) => sheets.extend(n),
                Operation::RecalculateOnOpen => recalc = true,
            }
        }
        names.sort();
        names.dedup();
        sheets.sort();
        sheets.dedup();
        // The same operations give the same plan, whatever the order they came in.
        for scopes in [&mut converts, &mut rows, &mut cols] {
            scopes.sort_by_key(|scope| (scope.sheet.clone(), target_key(scope.target)));
        }
        let mut plan = Self::new();
        plan.convert_to_values(converts);
        plan.delete_rows(rows)?;
        plan.delete_columns(cols)?;
        plan.remove_named_ranges(names);
        plan.remove_sheets(sheets);
        if recalc {
            plan.recalculate_on_open();
        }
        Ok(plan)
    }

    /// The operations as they were written.
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// The passes over the file that run the plan: the operations of each, normalized and in the
    /// order of the pass (see the module documentation), against the input of that pass. Their
    /// number is what the plan costs in reads and writes of the file; `against_input` gives one.
    ///
    /// Fails for an operation on a sheet that an earlier one removed.
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

    /// The passes as text, one operation per line, for a person to read (`explain` in a lazy
    /// frame of other libraries).
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
        Ok(out)
    }

    /// Runs the plan on `input` and writes `output`, which must not be the input. Each pass reads
    /// and writes the file once, as a stream; a plan of several passes keeps the file between
    /// them in a temporary folder. A failed run leaves no `output` behind. An empty plan is an
    /// error: there is nothing to do, and a copy of the file is not what was asked for.
    pub fn execute(&self, input: impl AsRef<Path>, output: impl AsRef<Path>) -> Result<()> {
        if self.is_empty() {
            bail!("No operations to apply");
        }
        let (input, output) = (input.as_ref(), output.as_ref());
        let passes = self.passes()?;
        let extension = input
            .extension()
            .and_then(|e| e.to_str())
            .context("The input file has no extension")?;
        let folder = (passes.len() > 1)
            .then(tempfile::tempdir)
            .transpose()
            .context("Cannot create a temporary folder")?;
        let mut current = input.to_path_buf();
        for (i, pass) in passes.iter().enumerate() {
            let last = i + 1 == passes.len();
            let target = match (&folder, last) {
                (Some(folder), false) => folder.path().join(format!("pass{i}.{extension}")),
                _ => output.to_path_buf(),
            };
            writer::modify_workbook(&current, &target, &modifications_of(pass))
                .with_context(|| format!("Failed in pass {} of {}", i + 1, passes.len()))?;
            current = target;
        }
        Ok(())
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
        Operation::ConvertToValues(scopes)
        | Operation::DeleteRows(scopes)
        | Operation::DeleteColumns(scopes) => scopes.iter().map(|s| s.sheet.as_str()).collect(),
        Operation::RemoveSheets(_)
        | Operation::RemoveNamedRanges(_)
        | Operation::RecalculateOnOpen => Vec::new(),
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
                mods.convert_to_values.extend(scopes.iter().cloned())
            }
            Operation::DeleteRows(scopes) => mods.delete_rows.extend(scopes.iter().cloned()),
            Operation::DeleteColumns(scopes) => mods.delete_columns.extend(scopes.iter().cloned()),
            Operation::RecalculateOnOpen => mods.recalculate_on_open = true,
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
    cols: BTreeMap<String, Vec<(u32, u32)>>,
    names: BTreeSet<String>,
    sheets: BTreeSet<String>,
    recalculate: bool,
    /// Whether anything but a conversion or the mark is in the pass.
    changes: bool,
}

impl Pass {
    fn is_empty(&self) -> bool {
        !self.changes && self.converts.is_empty() && !self.recalculate
    }

    /// Whether `operation` can join the pass: a conversion cannot follow another kind of change,
    /// and a deletion that touches one already there cannot either. Two deletions in a row that
    /// touch are not one deletion to the writer: it refuses a band that removes the anchor of a
    /// shape together with the rows before it, and the same rows deleted one operation at a time
    /// are not that band.
    fn admits(&self, operation: &Operation) -> bool {
        match operation {
            Operation::ConvertToValues(_) => !self.changes,
            Operation::DeleteRows(scopes) => !Self::touches(&self.rows, scopes),
            Operation::DeleteColumns(scopes) => !Self::touches(&self.cols, scopes),
            _ => true,
        }
    }

    fn touches(by_sheet: &BTreeMap<String, Vec<(u32, u32)>>, scopes: &[ValueScope]) -> bool {
        scopes.iter().any(|scope| {
            let (ValueTarget::Rows(first, last) | ValueTarget::Columns(first, last)) = scope.target
            else {
                return false;
            };
            let Some(ranges) = by_sheet.get(&scope.sheet) else {
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
            Operation::DeleteRows(scopes) => {
                Self::delete(&mut self.rows, scopes, &mut self.changes)
            }
            Operation::DeleteColumns(scopes) => {
                Self::delete(&mut self.cols, scopes, &mut self.changes)
            }
            Operation::RemoveNamedRanges(names) => {
                self.names.extend(names.iter().cloned());
                self.changes = true;
            }
            Operation::RemoveSheets(names) => {
                self.sheets.extend(names.iter().cloned());
                self.changes = true;
            }
            Operation::RecalculateOnOpen => self.recalculate = true,
        }
    }

    /// Adds deletions given against the workbook as the pass has left it so far: each position is
    /// translated to the input of the pass through the deletions already in it. The scopes of one
    /// operation are simultaneous, so they are all translated before any of them is added.
    fn delete(
        by_sheet: &mut BTreeMap<String, Vec<(u32, u32)>>,
        scopes: &[ValueScope],
        changes: &mut bool,
    ) {
        let mut translated = Vec::new();
        for scope in scopes {
            let (ValueTarget::Rows(first, last) | ValueTarget::Columns(first, last)) = scope.target
            else {
                continue;
            };
            let bands = Bands::new(by_sheet.get(&scope.sheet).cloned().unwrap_or_default());
            translated.push((
                scope.sheet.clone(),
                bands.input_index(first),
                bands.input_index(last),
            ));
        }
        for (sheet, first, last) in translated {
            by_sheet.entry(sheet).or_default().push((first, last));
        }
        *changes = true;
    }

    fn into_operations(self) -> Vec<Operation> {
        let bands = |by_sheet: BTreeMap<String, Vec<(u32, u32)>>, rows: bool| {
            let scopes: Vec<ValueScope> = by_sheet
                .into_iter()
                .flat_map(|(sheet, ranges)| {
                    Bands::new(ranges)
                        .ranges()
                        .map(|(first, last)| ValueScope {
                            sheet: sheet.clone(),
                            target: if rows {
                                ValueTarget::Rows(first, last)
                            } else {
                                ValueTarget::Columns(first, last)
                            },
                        })
                        .collect::<Vec<_>>()
                })
                .collect();
            scopes
        };
        let mut operations = Vec::new();
        let mut push = |operation: Operation, empty: bool| {
            if !empty {
                operations.push(operation);
            }
        };
        push(
            Operation::ConvertToValues(self.converts.clone()),
            self.converts.is_empty(),
        );
        let rows = bands(self.rows, true);
        push(Operation::DeleteRows(rows.clone()), rows.is_empty());
        let cols = bands(self.cols, false);
        push(Operation::DeleteColumns(cols.clone()), cols.is_empty());
        push(
            Operation::RemoveNamedRanges(self.names.iter().cloned().collect()),
            self.names.is_empty(),
        );
        push(
            Operation::RemoveSheets(self.sheets.iter().cloned().collect()),
            self.sheets.is_empty(),
        );
        push(Operation::RecalculateOnOpen, !self.recalculate);
        operations
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scope(spec: &str) -> ValueScope {
        spec.parse().unwrap()
    }

    fn scopes(specs: &[&str]) -> Vec<ValueScope> {
        specs.iter().map(|s| scope(s)).collect()
    }

    fn described(passes: &[Vec<Operation>]) -> Vec<Vec<String>> {
        passes
            .iter()
            .map(|pass| pass.iter().map(ToString::to_string).collect())
            .collect()
    }

    #[test]
    fn a_deletion_needs_the_kind_of_range_it_deletes() {
        let mut plan = EditPlan::new();
        assert!(plan.delete_rows(scopes(&["S!B:D"])).is_err());
        assert!(plan.delete_rows(scopes(&["S"])).is_err());
        assert!(plan.delete_columns(scopes(&["S!3:10"])).is_err());
        assert!(plan.is_empty());
    }

    #[test]
    fn an_operation_with_nothing_in_it_is_not_added() {
        let mut plan = EditPlan::new();
        plan.remove_sheets(Vec::<String>::new())
            .convert_to_values([]);
        plan.delete_rows([]).unwrap();
        assert!(plan.is_empty());
    }

    #[test]
    fn later_positions_are_those_the_earlier_deletions_left() {
        // After row 3 goes, row 4 of the input is the new row 3, and 6 is the new 5.
        let mut plan = EditPlan::new();
        plan.delete_rows(scopes(&["T!3"])).unwrap();
        plan.delete_rows(scopes(&["T!4"])).unwrap();
        plan.delete_rows(scopes(&["T!10"])).unwrap();
        assert_eq!(
            described(&plan.passes().unwrap()),
            [["delete rows T!3:3, T!5:5, T!12:12"]]
        );
    }

    #[test]
    fn deletions_that_touch_are_not_one_band_to_the_writer_so_they_do_not_share_a_pass() {
        // Row 3 twice are the input's rows 3 and 4, which touch: one operation at a time.
        let mut plan = EditPlan::new();
        plan.delete_rows(scopes(&["T!3"])).unwrap();
        plan.delete_rows(scopes(&["T!3"])).unwrap();
        assert_eq!(
            described(&plan.passes().unwrap()),
            [["delete rows T!3:3"], ["delete rows T!3:3"]]
        );
        // Written together they are one deletion.
        let mut plan = EditPlan::new();
        plan.delete_rows(scopes(&["T!3", "T!4"])).unwrap();
        assert_eq!(described(&plan.passes().unwrap()), [["delete rows T!3:4"]]);
    }

    #[test]
    fn the_scopes_of_one_operation_are_simultaneous() {
        let mut plan = EditPlan::new();
        plan.delete_rows(scopes(&["T!3", "T!4"])).unwrap();
        assert_eq!(described(&plan.passes().unwrap()), [["delete rows T!3:4"]]);
    }

    #[test]
    fn deletions_names_and_sheets_share_a_pass_and_a_conversion_before_them_too() {
        let mut plan = EditPlan::new();
        plan.convert_to_values(scopes(&["F"]));
        plan.delete_columns(scopes(&["T!D"])).unwrap();
        plan.delete_rows(scopes(&["U!2:3"])).unwrap();
        plan.remove_named_ranges(["N"]).remove_sheets(["V"]);
        plan.recalculate_on_open();
        assert_eq!(
            described(&plan.passes().unwrap()),
            [[
                "convert formulas to values in F",
                "delete rows U!2:3",
                "delete columns T!D:D",
                "remove named ranges N",
                "remove sheets V",
                "recalculate when opened",
            ]]
        );
    }

    #[test]
    fn a_conversion_after_a_change_starts_a_new_pass() {
        let mut plan = EditPlan::new();
        plan.delete_rows(scopes(&["T!3"])).unwrap();
        plan.convert_to_values(scopes(&["F"]));
        plan.delete_rows(scopes(&["T!1"])).unwrap();
        let passes = described(&plan.passes().unwrap());
        assert_eq!(
            passes,
            [
                vec!["delete rows T!3:3"],
                vec!["convert formulas to values in F", "delete rows T!1:1"]
            ]
        );
        assert!(plan.explain().unwrap().starts_with("pass 1 of 2:"));
    }

    #[test]
    fn an_operation_on_a_removed_sheet_is_an_error_when_planned() {
        let mut plan = EditPlan::new();
        plan.remove_sheets(["U"]);
        plan.delete_rows(scopes(&["U!2"])).unwrap();
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
                Operation::DeleteRows(scopes(&["T!8:9"])),
                Operation::ConvertToValues(scopes(&["F"])),
                Operation::DeleteRows(scopes(&["T!2", "V!6"])),
                Operation::RecalculateOnOpen,
                Operation::DeleteColumns(scopes(&["T!D"])),
                Operation::RemoveNamedRanges(vec!["N".into()]),
            ];
            EditPlan::against_input(order.iter().map(|&i| ops[i].clone())).unwrap()
        };
        let one = build(&[0, 1, 2, 3, 4, 5, 6]);
        let other = build(&[6, 5, 4, 3, 2, 1, 0]);
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
                "recalculate when opened",
            ]]
        );
    }

    #[test]
    fn an_empty_plan_has_no_passes_and_does_not_execute() {
        let plan = EditPlan::new();
        assert!(plan.passes().unwrap().is_empty());
        let error = plan.execute("in.xlsx", "out.xlsx").unwrap_err();
        assert!(error.to_string().contains("No operations"));
    }
}
