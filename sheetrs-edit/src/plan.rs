//! A plan: the operations on one workbook, resolved into the order in which they give the right
//! result, and run together.
//!
//! Operations are written against the **input** workbook: `delete_rows("S!3")` and
//! `convert_to_values("S!A:B")` address the rows and columns the input file has, whatever the
//! order they were added in. Applied one after another as written, they would not mean that: a
//! deletion shifts every position below it, so the next operation would address other cells, a
//! formula converted after the cells it reads were deleted would have lost its result, and a
//! reference across sheets would shrink by the deletions of one sheet and not by those of another
//! depending on which came first.
//!
//! The plan resolves this once. [`EditPlan::schedule`] is the plan in its canonical order, the
//! only one the result is defined by, and the writer applies exactly that order in a single pass
//! over the file:
//!
//! 1. formulas converted to values, while every cached result is still there (a scope on a sheet
//!    that goes is dropped with it);
//! 2. rows deleted, in every sheet, from the last to the first, with overlapping or touching
//!    ranges merged; each deletion shifts only what is below it, so the positions of the next one
//!    are still the input's;
//! 3. columns deleted the same way;
//! 4. defined names removed;
//! 5. sheets removed, last: a deletion in a sheet that goes still shapes the references that span
//!    it, as if it had been done before the sheet went;
//! 6. the mark to recalculate when opened.
//!
//! So the result is the same for any order of the same operations, equals what running the
//! schedule one operation per pass gives (the plan is the optimization, not another meaning), and
//! costs one read and one write of the file instead of one per operation.

use crate::writer::structure::StructureEdits;
use crate::writer::{self, ValueScope, ValueTarget, WorkbookModifications};
use anyhow::{Result, bail};
use sheetrs_formula::edit::Edit;
use std::collections::{BTreeSet, HashSet};
use std::fmt;
use std::path::Path;

/// One operation on a workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Operation {
    /// Remove the sheet of this name.
    RemoveSheet(String),
    /// Remove the defined name (or database range).
    RemoveNamedRange(String),
    /// Replace the formulas of the scope by their cached values.
    ConvertToValues(ValueScope),
    /// Delete the rows of the scope (`Sheet!3:10`); references to them become `#REF!`.
    DeleteRows(ValueScope),
    /// Delete the columns of the scope (`Sheet!B:D`).
    DeleteColumns(ValueScope),
    /// Mark the output so the application that opens it recalculates every formula.
    RecalculateOnOpen,
}

impl fmt::Display for Operation {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::RemoveSheet(name) => write!(f, "remove sheet {name}"),
            Self::RemoveNamedRange(name) => write!(f, "remove named range {name}"),
            Self::ConvertToValues(scope) => write!(f, "convert formulas to values in {scope}"),
            Self::DeleteRows(scope) => write!(f, "delete rows {scope}"),
            Self::DeleteColumns(scope) => write!(f, "delete columns {scope}"),
            Self::RecalculateOnOpen => write!(f, "recalculate when opened"),
        }
    }
}

/// Operations to apply to a workbook in a single pass.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EditPlan {
    operations: Vec<Operation>,
}

impl EditPlan {
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds an operation. Repeating one that is already in the plan changes nothing. Rows are
    /// deleted with a row range (`Sheet!3:10`) and columns with a column range (`Sheet!B:D`): any
    /// other scope is refused here, not when the plan runs.
    pub fn push(&mut self, operation: Operation) -> Result<&mut Self> {
        match &operation {
            Operation::DeleteRows(scope) if !matches!(scope.target, ValueTarget::Rows(..)) => {
                bail!("'{scope}' is not a row range such as 'Sheet!3:10'")
            }
            Operation::DeleteColumns(scope)
                if !matches!(scope.target, ValueTarget::Columns(..)) =>
            {
                bail!("'{scope}' is not a column range such as 'Sheet!B:D'")
            }
            _ => {}
        }
        if !self.operations.contains(&operation) {
            self.operations.push(operation);
        }
        Ok(self)
    }

    pub fn remove_sheet(&mut self, name: impl Into<String>) -> &mut Self {
        self.push(Operation::RemoveSheet(name.into()))
            .expect("a sheet name is always valid")
    }

    pub fn remove_named_range(&mut self, name: impl Into<String>) -> &mut Self {
        self.push(Operation::RemoveNamedRange(name.into()))
            .expect("a name is always valid")
    }

    pub fn convert_to_values(&mut self, scope: ValueScope) -> &mut Self {
        self.push(Operation::ConvertToValues(scope))
            .expect("every scope can be converted")
    }

    pub fn delete_rows(&mut self, scope: ValueScope) -> Result<&mut Self> {
        self.push(Operation::DeleteRows(scope))
    }

    pub fn delete_columns(&mut self, scope: ValueScope) -> Result<&mut Self> {
        self.push(Operation::DeleteColumns(scope))
    }

    pub fn recalculate_on_open(&mut self) -> &mut Self {
        self.push(Operation::RecalculateOnOpen)
            .expect("the mark is always valid")
    }

    /// The operations as they were added.
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// The plan as the writer takes it. The order of the operations plays no part in it: the
    /// writer applies the canonical order (see the module documentation).
    pub fn to_modifications(&self) -> WorkbookModifications {
        let mut mods = WorkbookModifications::default();
        for operation in &self.operations {
            match operation {
                Operation::RemoveSheet(name) => {
                    mods.remove_sheets
                        .get_or_insert_with(Default::default)
                        .insert(name.clone());
                }
                Operation::RemoveNamedRange(name) => {
                    mods.remove_named_ranges
                        .get_or_insert_with(Default::default)
                        .insert(name.clone());
                }
                Operation::ConvertToValues(scope) => mods.convert_to_values.push(scope.clone()),
                Operation::DeleteRows(scope) => mods.delete_rows.push(scope.clone()),
                Operation::DeleteColumns(scope) => mods.delete_columns.push(scope.clone()),
                Operation::RecalculateOnOpen => mods.recalculate_on_open = true,
            }
        }
        mods
    }

    /// The plan in its canonical order: the operations normalized (ranges merged, repeats and
    /// scopes on removed sheets dropped) and sorted as the module documentation lists, which is
    /// the order the writer applies. Two plans with the same schedule give the same workbook.
    ///
    /// Every position in it is the input's: a step does not need the ones before it to be read.
    /// Running the steps one per pass, each on the output of the previous one, gives the same
    /// result as running the plan.
    pub fn schedule(&self) -> Vec<Operation> {
        let mods = self.to_modifications();
        let removed: HashSet<String> = mods.remove_sheets.clone().unwrap_or_default();
        let mut steps = Vec::new();

        let converts: BTreeSet<(String, (u8, u32, u32))> = mods
            .convert_to_values
            .iter()
            .filter(|scope| !removed.contains(&scope.sheet))
            .map(|scope| (scope.sheet.clone(), target_key(scope.target)))
            .collect();
        steps.extend(converts.into_iter().map(|(sheet, key)| {
            Operation::ConvertToValues(ValueScope {
                sheet,
                target: key_target(key),
            })
        }));

        // The same resolution the writer makes of the deletions, so the two cannot differ.
        let structure =
            StructureEdits::new(None, &removed, &mods.delete_rows, &mods.delete_columns)
                .expect("the kind of every deletion was checked when it was added");
        let mut sheets = Vec::new();
        for edit in structure.edits() {
            match edit {
                Edit::DeleteRows { sheet, at, count } => {
                    steps.push(Operation::DeleteRows(ValueScope {
                        sheet: sheet.to_string(),
                        target: ValueTarget::Rows(at, at + count - 1),
                    }));
                }
                Edit::DeleteCols { sheet, at, count } => {
                    steps.push(Operation::DeleteColumns(ValueScope {
                        sheet: sheet.to_string(),
                        target: ValueTarget::Columns(at, at + count - 1),
                    }));
                }
                Edit::DeleteSheet { sheet } => sheets.push(Operation::RemoveSheet(sheet.into())),
                _ => unreachable!("the writer builds no other edit"),
            }
        }

        let names: BTreeSet<String> = mods
            .remove_named_ranges
            .unwrap_or_default()
            .into_iter()
            .collect();
        steps.extend(names.into_iter().map(Operation::RemoveNamedRange));
        steps.extend(sheets);
        if mods.recalculate_on_open {
            steps.push(Operation::RecalculateOnOpen);
        }
        steps
    }

    /// Applies every operation to `input` in one pass and writes `output`, which must not be the
    /// input. A failed run leaves no `output` behind. An empty plan is an error: there is nothing
    /// to do and a copy of the file is not what was asked for.
    pub fn execute(&self, input: impl AsRef<Path>, output: impl AsRef<Path>) -> Result<()> {
        if self.is_empty() {
            bail!("No operations to apply");
        }
        writer::modify_workbook(input.as_ref(), output.as_ref(), &self.to_modifications())
    }
}

fn target_key(target: ValueTarget) -> (u8, u32, u32) {
    match target {
        ValueTarget::Sheet => (0, 0, 0),
        ValueTarget::Columns(first, last) => (1, first, last),
        ValueTarget::Rows(first, last) => (2, first, last),
    }
}

fn key_target((kind, first, last): (u8, u32, u32)) -> ValueTarget {
    match kind {
        0 => ValueTarget::Sheet,
        1 => ValueTarget::Columns(first, last),
        _ => ValueTarget::Rows(first, last),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scope(spec: &str) -> ValueScope {
        spec.parse().unwrap()
    }

    fn described(steps: &[Operation]) -> Vec<String> {
        steps.iter().map(ToString::to_string).collect()
    }

    #[test]
    fn operations_keep_their_order_and_ignore_repeats() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("B")
            .recalculate_on_open()
            .remove_sheet("A")
            .remove_sheet("B");
        assert_eq!(
            described(plan.operations()),
            [
                "remove sheet B",
                "recalculate when opened",
                "remove sheet A"
            ]
        );
    }

    #[test]
    fn a_plan_folds_into_the_modifications_of_the_writer() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("A")
            .remove_named_range("N")
            .convert_to_values(scope("S!B:D"))
            .recalculate_on_open();
        plan.delete_rows(scope("S!3:10")).unwrap();
        plan.delete_columns(scope("S!C")).unwrap();
        let mods = plan.to_modifications();

        assert_eq!(mods.remove_sheets.unwrap().len(), 1);
        assert_eq!(mods.remove_named_ranges.unwrap().len(), 1);
        assert_eq!(mods.convert_to_values, [scope("S!B:D")]);
        assert_eq!(mods.delete_rows, [scope("S!3:10")]);
        assert_eq!(mods.delete_columns, [scope("S!C")]);
        assert!(mods.recalculate_on_open);
    }

    #[test]
    fn a_deletion_needs_the_kind_of_range_it_deletes() {
        let mut plan = EditPlan::new();
        assert!(plan.delete_rows(scope("S!B:D")).is_err());
        assert!(plan.delete_rows(scope("S")).is_err());
        assert!(plan.delete_columns(scope("S!3:10")).is_err());
        assert!(plan.is_empty());
    }

    #[test]
    fn the_schedule_is_the_canonical_order_whatever_the_order_written() {
        let mut one = EditPlan::new();
        one.recalculate_on_open()
            .remove_sheet("U")
            .remove_named_range("N");
        one.delete_columns(scope("T!D")).unwrap();
        one.delete_rows(scope("T!2")).unwrap();
        one.delete_rows(scope("V!6")).unwrap();
        one.delete_rows(scope("U!4:5")).unwrap();
        one.convert_to_values(scope("F"));

        let mut other = EditPlan::new();
        other.convert_to_values(scope("F"));
        other.delete_rows(scope("U!4:5")).unwrap();
        other.delete_rows(scope("V!6")).unwrap();
        other.delete_rows(scope("T!2")).unwrap();
        other.delete_columns(scope("T!D")).unwrap();
        other
            .remove_named_range("N")
            .remove_sheet("U")
            .recalculate_on_open();

        assert_eq!(one.schedule(), other.schedule());
        assert_eq!(
            described(&one.schedule()),
            [
                "convert formulas to values in F",
                // Rows from the last to the first across sheets, then columns.
                "delete rows V!6:6",
                "delete rows U!4:5",
                "delete rows T!2:2",
                "delete columns T!D:D",
                "remove named range N",
                "remove sheet U",
                "recalculate when opened",
            ]
        );
    }

    #[test]
    fn the_schedule_merges_the_ranges_of_a_sheet() {
        let mut plan = EditPlan::new();
        for spec in ["T!5", "T!3:6", "T!7", "T!20"] {
            plan.delete_rows(scope(spec)).unwrap();
        }
        assert_eq!(
            described(&plan.schedule()),
            ["delete rows T!20:20", "delete rows T!3:7"]
        );
    }

    #[test]
    fn a_conversion_goes_with_its_sheet_and_a_deletion_does_not() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("U").convert_to_values(scope("U!A:B"));
        plan.delete_rows(scope("U!2")).unwrap();
        assert_eq!(
            described(&plan.schedule()),
            ["delete rows U!2:2", "remove sheet U"]
        );
    }

    #[test]
    fn an_empty_plan_has_no_schedule_and_does_not_execute() {
        let plan = EditPlan::new();
        assert!(plan.is_empty());
        assert!(plan.schedule().is_empty());
        assert!(plan.to_modifications().remove_sheets.is_none());
        let error = plan.execute("in.xlsx", "out.xlsx").unwrap_err();
        assert!(error.to_string().contains("No operations"));
    }
}
