//! The operations a `Book` accumulates and how `save` and `to_dict` run them.
//!
//! The operations go to one [`Book`] (a copy of the one the Python object refers to), whose planner
//! runs all of them in the fewest passes over the file, in an order of its own (every operation is
//! written against the workbook as it is: see `sheetrs_edit::BookEdit`), and writes the result. The
//! Python layer adds nothing to the plan.

use serde_json::{Map, Value};
use sheetrs_edit::{
    BookEdit, Calculation, ColumnRange, Report, RowRange, ValueScope, VisibilityChange,
};
use sheetrs_io::Book;
use sheetrs_io::write::Operation;
use sheetrs_protocol::{MajorDimension, Replace, Scope};
use std::io::Write;
use std::path::Path;

/// What can go wrong running the operations.
#[derive(Debug, thiserror::Error)]
pub(crate) enum PlanError {
    #[error(transparent)]
    Edit(#[from] sheetrs_edit::Error),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error("{0}")]
    Message(String),
}

/// One operation of a `Book`, as it was asked.
#[derive(Clone)]
pub(crate) enum Step {
    DeleteSheets(Vec<String>),
    DeleteNames(Vec<String>),
    FormulaToValue(Vec<ValueScope>),
    DeleteRows(Vec<RowRange>),
    DeleteColumns(Vec<ColumnRange>),
    /// Rows, columns or sheets hidden or shown.
    SetVisibility(Vec<VisibilityChange>),
    Replace {
        scopes: Vec<Scope>,
        replace: Replace,
    },
    /// The engine writes the inputs and calculates every formula. `ranges` are the ranges the
    /// answer holds (`None`: the outputs the workbook declares, or its formula cells), listed as
    /// `major` says.
    Calculate {
        inputs: Map<String, Value>,
        ranges: Option<Vec<String>>,
        major: MajorDimension,
        /// The serial number of `NOW()` and `TODAY()`; `None` is the clock.
        now: Option<f64>,
        /// The seed of `RAND()`; `None` is seeded from the clock.
        seed: Option<u64>,
    },
}

impl Step {
    /// The ranges a `calculate` step was asked to answer.
    fn ranges(&self) -> Option<&Vec<String>> {
        match self {
            Self::Calculate { ranges, .. } => ranges.as_ref(),
            _ => None,
        }
    }
}

/// The ranges the last `calculate` step asked for, by name (`None`: those of the workbook).
pub(crate) fn asked_ranges(steps: &[Step]) -> Option<Vec<String>> {
    steps.iter().rev().find_map(|step| match step {
        Step::Calculate { .. } => Some(step.ranges().cloned()),
        _ => None,
    })?
}

/// How the last `calculate` step lists its answer.
pub(crate) fn asked_major(steps: &[Step]) -> MajorDimension {
    steps
        .iter()
        .rev()
        .find_map(|step| match step {
            Step::Calculate { major, .. } => Some(*major),
            _ => None,
        })
        .unwrap_or_default()
}

/// Whether there is a `calculate` step.
pub(crate) fn has_calculate(steps: &[Step]) -> bool {
    steps.iter().any(|s| matches!(s, Step::Calculate { .. }))
}

fn editor(source: &Book, steps: &[Step], keep_answer: bool) -> Book {
    let mut editor = source.clone();
    // The answer kept is that of the last calculation.
    let last_calculation = steps
        .iter()
        .rposition(|step| matches!(step, Step::Calculate { .. }));
    for (index, step) in steps.iter().enumerate() {
        match step {
            Step::DeleteSheets(names) => {
                editor.delete_sheets(names.clone());
            }
            Step::DeleteNames(names) => {
                editor.delete_names(names.clone());
            }
            Step::FormulaToValue(scopes) => {
                editor.formula_to_value(scopes.clone());
            }
            Step::DeleteRows(rows) => {
                editor.delete_rows(rows.clone());
            }
            Step::DeleteColumns(columns) => {
                editor.delete_columns(columns.clone());
            }
            Step::SetVisibility(changes) => {
                editor.push(Operation::SetVisibility(changes.clone()));
            }
            Step::Replace { scopes, replace } => {
                editor.replace(scopes.clone(), [replace.clone()]);
            }
            Step::Calculate {
                inputs,
                ranges,
                major,
                now,
                seed,
            } => {
                editor.calculate(Calculation {
                    inputs: inputs.clone(),
                    ranges: ranges.clone(),
                    major: *major,
                    answer: keep_answer && Some(index) == last_calculation,
                    now: *now,
                    seed: *seed,
                });
            }
        }
    }
    editor
}

/// What `save` will do, as text.
pub(crate) fn explain(source: &Book, steps: &[Step]) -> String {
    if steps.is_empty() {
        return "nothing to do: no operations".to_owned();
    }
    editor(source, steps, false).explain()
}

/// Where a run writes the workbook.
pub(crate) enum Target<'a> {
    Path(&'a Path),
    InPlace,
    /// Nowhere: the run is for what it answers (`to_dict`).
    Nowhere,
    /// Memory: the workbook comes back as bytes in [`Done::bytes`] (`to_bytes`); no file is made.
    Bytes,
}

/// What running the operations did.
pub(crate) struct Done {
    /// What the run reports: the cells replaced and written, the value ranges, the diagnostics.
    pub report: Report,
    /// The bytes of the workbook, for [`Target::Bytes`].
    pub bytes: Option<Vec<u8>>,
}

/// Writes `bytes` to `path` through a temporary file in the same folder and a rename, so that a
/// failure leaves the file as it was and no half-written one behind.
pub(crate) fn write_atomic(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let folder = path
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = tempfile::Builder::new()
        .prefix(".sheetrs-")
        .tempfile_in(folder)?;
    temporary.write_all(bytes)?;
    if let Ok(meta) = std::fs::metadata(path) {
        let _ = std::fs::set_permissions(temporary.path(), meta.permissions());
    }
    temporary.persist(path).map_err(|e| e.error)?;
    Ok(())
}

/// Runs the operations and writes the workbook to `target`; `keep_answer` keeps the answer of the
/// calculation in the [`Done`].
pub(crate) fn run(
    source: &Book,
    steps: &[Step],
    target: &Target<'_>,
    keep_answer: bool,
) -> Result<Done, PlanError> {
    if steps.is_empty() {
        return Err(PlanError::Message(
            "there are no operations to save: a copy of the file is not what was asked for".into(),
        ));
    }
    let editor = editor(source, steps, keep_answer);
    let mut bytes = None;
    let report = match target {
        Target::Path(path) => editor.write(path)?,
        Target::InPlace => editor.write_in_place()?,
        Target::Nowhere => editor.write_to(std::io::sink())?,
        Target::Bytes => {
            let mut memory = Vec::new();
            let report = editor.write_to(&mut memory)?;
            bytes = Some(memory);
            report
        }
    };
    Ok(Done { report, bytes })
}
