//! A plan: the operations on one workbook, resolved into the order in which they give the right
//! result, and run together.
//!
//! Operations are written against the **input** workbook: `delete_rows("S!3")` and
//! `convert_to_values("S!A:B")` address the rows and columns the input file has, whatever the
//! order they were added in. Applied one after another as written they would not mean that: a
//! deletion shifts every position below it, so the next operation would address other cells, and a
//! formula converted after the cells it reads were deleted has lost its result.
//!
//! So [`EditPlan::operations`] is not the list of what was added: it is the plan resolved, the
//! only order the result is defined by, and the one the writer applies in a single pass:
//!
//! 1. formulas converted to values, while every cached result is still there (a scope on a sheet
//!    that goes is dropped with it);
//! 2. in each sheet, rows deleted from the last to the first, overlapping or touching ranges
//!    merged, so each deletion shifts only what is below it and the positions of the next one
//!    are still the input's;
//! 3. in each sheet, columns deleted the same way;
//! 4. defined names removed;
//! 5. sheets removed;
//! 6. the mark to recalculate when opened.
//!
//! The only dependencies are the ones that order lists: a conversion reads values that a deletion
//! or a removal would lose, and the deletions of one sheet shift one another. Operations on
//! different sheets are independent, since a reference moves only with the rows and columns of the
//! sheet it names; a reference that spans sheets (`Sheet1:Sheet3!A1`) would tie them, and is not
//! supported (issue #75): a plan that meets one fails instead of rewriting it wrongly. So no
//! ordering between sheets is needed, and they go by name only so the result never depends on a
//! hash order.
//!
//! The plan is a description, not a model of the workbook: it holds the operations (a few bands
//! and names), and the writer streams the file once, rewriting each cell as it goes by.

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

/// Operations to apply to a workbook in a single pass, written against the input workbook.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EditPlan {
    added: Vec<Operation>,
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
        if !self.added.contains(&operation) {
            self.added.push(operation);
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

    pub fn is_empty(&self) -> bool {
        self.added.is_empty()
    }

    /// The plan as the writer takes it. The order the operations were added in plays no part in
    /// it: the writer applies the order of [`operations`](Self::operations).
    pub fn to_modifications(&self) -> WorkbookModifications {
        let mut mods = WorkbookModifications::default();
        for operation in &self.added {
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

    /// The operations of the plan, resolved: normalized (ranges merged, repeats and scopes on
    /// removed sheets dropped) and in the order the module documentation lists, which is the
    /// order the writer applies. Two plans with the same operations give the same workbook,
    /// whatever the order they were written in.
    ///
    /// Every position in it is the input's: an operation does not need the ones before it to be
    /// read. Running them one per pass, each on the output of the previous one, gives the same
    /// result as running the plan, which reads and writes the file once.
    pub fn operations(&self) -> Vec<Operation> {
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
    fn repeats_count_once_and_the_order_added_is_not_kept() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("B")
            .recalculate_on_open()
            .remove_sheet("A")
            .remove_sheet("B");
        assert_eq!(
            described(&plan.operations()),
            [
                "remove sheet A",
                "remove sheet B",
                "recalculate when opened"
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
    fn the_operations_are_in_the_canonical_order_whatever_the_order_written() {
        let mut one = EditPlan::new();
        one.recalculate_on_open()
            .remove_sheet("U")
            .remove_named_range("N");
        one.delete_columns(scope("T!D")).unwrap();
        one.delete_rows(scope("T!2")).unwrap();
        one.delete_rows(scope("V!6")).unwrap();
        one.delete_rows(scope("T!8:9")).unwrap();
        one.convert_to_values(scope("F"));

        let mut other = EditPlan::new();
        other.convert_to_values(scope("F"));
        other.delete_rows(scope("T!8:9")).unwrap();
        other.delete_rows(scope("V!6")).unwrap();
        other.delete_rows(scope("T!2")).unwrap();
        other.delete_columns(scope("T!D")).unwrap();
        other
            .remove_named_range("N")
            .remove_sheet("U")
            .recalculate_on_open();

        assert_eq!(one.operations(), other.operations());
        assert_eq!(
            described(&one.operations()),
            [
                "convert formulas to values in F",
                // Each sheet's rows from the last band to the first; sheets by name.
                "delete rows T!8:9",
                "delete rows T!2:2",
                "delete rows V!6:6",
                "delete columns T!D:D",
                "remove named range N",
                "remove sheet U",
                "recalculate when opened",
            ]
        );
    }

    #[test]
    fn the_operations_merge_the_ranges_of_a_sheet() {
        let mut plan = EditPlan::new();
        for spec in ["T!5", "T!3:6", "T!7", "T!20"] {
            plan.delete_rows(scope(spec)).unwrap();
        }
        assert_eq!(
            described(&plan.operations()),
            ["delete rows T!20:20", "delete rows T!3:7"]
        );
    }

    #[test]
    fn the_scopes_of_a_removed_sheet_go_with_it() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("U").convert_to_values(scope("U!A:B"));
        plan.delete_rows(scope("U!2")).unwrap();
        plan.delete_columns(scope("T!B")).unwrap();
        assert_eq!(
            described(&plan.operations()),
            ["delete columns T!B:B", "remove sheet U"]
        );
    }

    #[test]
    fn an_empty_plan_has_no_operations_and_does_not_execute() {
        let plan = EditPlan::new();
        assert!(plan.is_empty());
        assert!(plan.operations().is_empty());
        assert!(plan.to_modifications().remove_sheets.is_none());
        let error = plan.execute("in.xlsx", "out.xlsx").unwrap_err();
        assert!(error.to_string().contains("No operations"));
    }
}
