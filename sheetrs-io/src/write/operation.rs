//! The operations a workbook is asked to undergo, as data: what a [`Book`](crate::Book) holds
//! while they are pending and what the planner of `sheetrs-edit` turns into passes of the writer.

use super::{CellWrites, ColumnRange, RowRange, ValueScope, VisibilityChange};
use serde_json::{Map, Value as Json};
use sheetrs_protocol::{MajorDimension, Replace, Scope};
use std::fmt;

/// A replacement and where it applies.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReplaceSpec {
    /// Empty: the whole workbook.
    pub scopes: Vec<Scope>,
    /// What is found and what takes its place.
    pub replace: Replace,
}

/// What a `calculate` step asks of the engine.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Calculation {
    /// The values written into the cells they name (`"Data!B2"`, a defined name, a range of
    /// constants) before the formulas are calculated; empty when the workbook is calculated as it
    /// is.
    pub inputs: Map<String, Json>,
    /// The ranges whose values the answer holds; `None` is the outputs the workbook declares or,
    /// with none declared, its formula cells.
    pub ranges: Option<Vec<String>>,
    /// Whether the answer lists its arrays as rows or as columns.
    pub major: MajorDimension,
    /// Whether the answer (the value ranges) is wanted besides the workbook.
    pub answer: bool,
    /// The serial number that `NOW()` and `TODAY()` give in this calculation (the `now` of `sheetc`);
    /// `None` is the clock. It also fixes the recalculation that the same plan runs after this step.
    pub now: Option<f64>,
    /// The seed of `RAND()` and its relatives in this calculation (the `seed` of `sheetc`); `None` is
    /// seeded from the clock. It also fixes the recalculation that the same plan runs after this step.
    pub seed: Option<u64>,
}

/// One operation on a workbook, on the workbook as the operations before it leave it. The
/// ranges of one operation are simultaneous: `DeleteRows(["S!3", "S!4"])` deletes two rows.
#[derive(Debug, Clone, PartialEq)]
pub enum Operation {
    /// Delete the sheets of these names.
    DeleteSheets(Vec<String>),
    /// Delete the defined names (or database ranges).
    DeleteNames(Vec<String>),
    /// Replace the formulas of the scopes (`Sheet`, `Sheet!B:D`, `Sheet!3:10`) by their cached
    /// values.
    FormulaToValue(Vec<ValueScope>),
    /// Delete the rows of the ranges; references to them become `#REF!`.
    DeleteRows(Vec<RowRange>),
    /// Delete the columns of the ranges.
    DeleteColumns(Vec<ColumnRange>),
    /// Find and replace text in the cells of the scopes, as `LibreOffice` does. It reads the
    /// content of the cells, so it is a pass of its own: what it decides is decided against the
    /// workbook as the operations before it leave it.
    Replace(ReplaceSpec),
    /// Hide or show rows, columns and sheets, in the coordinates of the workbook as it is. The
    /// changes keep the order they were pushed in (a later one wins over an earlier one) and are
    /// applied before any deletion.
    SetVisibility(Vec<VisibilityChange>),
    /// Write values into cells, decided by the caller (see [`CellWrites`]).
    WriteCells(CellWrites),
    /// Calculate every formula again. The plan puts it where a conversion would read results that
    /// the changes before it left stale; it is never written by hand.
    Recalculate,
    /// Write the inputs and calculate every formula, with the engine (see
    /// [`Calculation`]): the workbook as the operations before it leave it, calculated. It is a pass
    /// of its own. Where the engine cannot, the workbook is marked to be recalculated when opened.
    Calculate(Calculation),
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
            Self::DeleteSheets(names) => write!(f, "delete sheets {}", list(names)),
            Self::DeleteNames(names) => write!(f, "delete names {}", list(names)),
            Self::FormulaToValue(scopes) => {
                write!(f, "formula to value in {}", list(scopes))
            }
            Self::SetVisibility(changes) => write!(f, "{}", list(changes)),
            Self::DeleteRows(rows) => write!(f, "delete rows {}", list(rows)),
            Self::DeleteColumns(columns) => write!(f, "delete columns {}", list(columns)),
            Self::Recalculate => f.write_str("recalculate every formula (the results are stale)"),
            Self::Calculate(calculation) => {
                let inputs = calculation.inputs.len();
                write!(
                    f,
                    "calculate every formula, with {inputs} input{}",
                    if inputs == 1 { "" } else { "s" }
                )?;
                if let Some(ranges) = &calculation.ranges {
                    write!(f, " and the ranges {}", list(ranges))?;
                }
                if calculation.major == MajorDimension::Columns {
                    f.write_str(", as columns")?;
                }
                Ok(())
            }
            Self::WriteCells(writes) => {
                let mut sheets: Vec<&str> = writes.sheets().keys().map(String::as_str).collect();
                sheets.sort_unstable();
                write!(f, "write {} cells in {}", writes.len(), sheets.join(", "))
            }
            Self::Replace(spec) => {
                let find = &spec.replace.find;
                write!(f, "replace {:?} with {:?}", find.text, spec.replace.with)?;
                let mut how = vec![format!("{:?}", find.mode).to_lowercase()];
                if find.match_case {
                    how.push("match case".into());
                }
                if find.entire_cells {
                    how.push("entire cells".into());
                }
                how.push(format!("in {:?}", find.search_in).to_lowercase());
                write!(f, " ({})", how.join(", "))?;
                if spec.scopes.is_empty() {
                    f.write_str(" in the workbook")
                } else {
                    let scopes: Vec<String> = spec
                        .scopes
                        .iter()
                        .map(|s| match s {
                            Scope::Workbook => "the workbook".to_string(),
                            Scope::Sheet(name) => format!("sheet {name}"),
                            Scope::Range(r) => r.as_str().to_string(),
                        })
                        .collect();
                    write!(f, " in {}", scopes.join(", "))
                }
            }
        }
    }
}

impl Operation {
    /// Whether there is nothing in it to do (an empty list of names, of ranges or of cells).
    #[must_use]
    pub fn is_empty(&self) -> bool {
        match self {
            Self::DeleteSheets(items) | Self::DeleteNames(items) => items.is_empty(),
            Self::FormulaToValue(items) => items.is_empty(),
            Self::DeleteRows(items) => items.is_empty(),
            Self::DeleteColumns(items) => items.is_empty(),
            Self::SetVisibility(items) => items.is_empty(),
            Self::Replace(_) | Self::Recalculate | Self::Calculate(_) => false,
            Self::WriteCells(writes) => writes.is_empty(),
        }
    }
}
