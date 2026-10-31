//! A set of operations on one workbook, run together.
//!
//! The writer applies every operation of a [`WorkbookModifications`] while it streams the archive
//! once, so the cost of an edit is that of reading and writing the file, not of each operation.
//! [`EditPlan`] is the way to describe such a set: operations are added one at a time (the shape a
//! command line, a script or a binding has them) and nothing runs until [`EditPlan::execute`].
//!
//! All the operations of a plan refer to the **input** workbook, whatever order they are added in:
//! `delete_rows("S!3")` and `convert_to_values("S!A:B")` address the rows and columns the input
//! file has. Running operations one after another, each on the output of the previous one, gives
//! each the coordinates of the result of the earlier ones; that costs one pass per operation.

use crate::writer::{self, ValueScope, WorkbookModifications};
use anyhow::{Result, bail};
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
    /// Delete the rows of the scope; references to them become `#REF!`.
    DeleteRows(ValueScope),
    /// Delete the columns of the scope.
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

    /// Adds an operation. Repeating one that is already in the plan changes nothing.
    pub fn push(&mut self, operation: Operation) -> &mut Self {
        if !self.operations.contains(&operation) {
            self.operations.push(operation);
        }
        self
    }

    pub fn remove_sheet(&mut self, name: impl Into<String>) -> &mut Self {
        self.push(Operation::RemoveSheet(name.into()))
    }

    pub fn remove_named_range(&mut self, name: impl Into<String>) -> &mut Self {
        self.push(Operation::RemoveNamedRange(name.into()))
    }

    pub fn convert_to_values(&mut self, scope: ValueScope) -> &mut Self {
        self.push(Operation::ConvertToValues(scope))
    }

    pub fn delete_rows(&mut self, scope: ValueScope) -> &mut Self {
        self.push(Operation::DeleteRows(scope))
    }

    pub fn delete_columns(&mut self, scope: ValueScope) -> &mut Self {
        self.push(Operation::DeleteColumns(scope))
    }

    pub fn recalculate_on_open(&mut self) -> &mut Self {
        self.push(Operation::RecalculateOnOpen)
    }

    /// The operations, in the order they were added.
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// The plan as the writer takes it.
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

#[cfg(test)]
mod tests {
    use super::*;

    fn scope(spec: &str) -> ValueScope {
        spec.parse().unwrap()
    }

    #[test]
    fn operations_keep_their_order_and_ignore_repeats() {
        let mut plan = EditPlan::new();
        plan.remove_sheet("B")
            .recalculate_on_open()
            .remove_sheet("A")
            .remove_sheet("B");
        let described: Vec<String> = plan.operations().iter().map(ToString::to_string).collect();
        assert_eq!(
            described,
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
            .delete_rows(scope("S!3:10"))
            .delete_columns(scope("S!C"))
            .recalculate_on_open();
        let mods = plan.to_modifications();

        assert_eq!(mods.remove_sheets.unwrap().len(), 1);
        assert_eq!(mods.remove_named_ranges.unwrap().len(), 1);
        assert_eq!(mods.convert_to_values, [scope("S!B:D")]);
        assert_eq!(mods.delete_rows, [scope("S!3:10")]);
        assert_eq!(mods.delete_columns, [scope("S!C")]);
        assert!(mods.recalculate_on_open);
    }

    #[test]
    fn an_empty_plan_has_no_modifications_and_does_not_execute() {
        let plan = EditPlan::new();
        assert!(plan.is_empty());
        assert!(plan.to_modifications().remove_sheets.is_none());
        let error = plan.execute("in.xlsx", "out.xlsx").unwrap_err();
        assert!(error.to_string().contains("No operations"));
    }
}
