//! INT402: Formula sequence interrupted by an empty cell
//!
//! A hole in a column or row of otherwise identical formulas (`=B1+C1`, `=B2+C2`, nothing, `=B4+C4`)
//! is usually a formula that was deleted or never copied, so one line of a table is missing from
//! its totals. INT402 reports the interruption where the cell is empty.
//!
//! This rule is the `InterruptionKind::Empty` case of [`InterruptionRule`], the implementation shared
//! with INT401 and INT403; the full algorithm is described in
//! [`int401_interrupted_by_data`](super::int401_interrupted_by_data).
//!
//! # Reports
//!
//! One violation per empty cell (unit: the cell) inside a run of formulas with the same normalized
//! pattern in a row or a column, with at least `int_min_formula_sequence` formulas of the pattern
//! directly on one side of it. Empty means a cell with neither value nor formula: either stored
//! without content (for example only formatted), or a single absent position between two cells of the
//! run. A cell can be reported from its row and from its column.
//!
//! # Does not report
//!
//! - A gap of two or more absent cells: it ends the run and is read as a deliberate separation.
//! - Empty cells after the last formula of a run.
//! - Runs too short for the minimum.
//! - A value in place of a formula (INT401) or a different formula (INT403).
//!
//! # Parameters
//!
//! - `int_min_formula_sequence` (integer, default `3`), global only; see INT401.
//!
//! # Severity
//!
//! Always `Warning`.

pub use super::int401_interrupted_by_data::{InterruptionKind, InterruptionRule};
