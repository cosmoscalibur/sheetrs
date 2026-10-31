//! INT403: Formula sequence interrupted by a different formula
//!
//! In a column or row of formulas that follow one pattern (`=B1+C1`, `=B2+C2`, `=SUM(B3:C3)`,
//! `=B4+C4`), a formula written differently is often a leftover of manual editing and gives a result
//! computed by another rule than its neighbours. INT403 reports the interruption where the cell is
//! itself a formula, but a different one.
//!
//! This rule is the `InterruptionKind::Other` case of [`InterruptionRule`], the implementation shared
//! with INT401 and INT402; the full algorithm is described in
//! [`int401_interrupted_by_data`](super::int401_interrupted_by_data).
//!
//! # Reports
//!
//! One violation per formula cell (unit: the cell) in a run of formulas whose normalized pattern
//! differs from the run's first formula, with at least `int_min_formula_sequence` formulas of the
//! run's pattern directly on one side of it. Patterns are compared after rewriting references as
//! offsets and upper-casing, so only the structure matters: `=B2+C2` below `=B1+C1` is the same
//! pattern, `=SUM(B2:C2)` is not. A cell can be reported from its row and from its column.
//!
//! # Does not report
//!
//! - Formulas with the same pattern: copied formulas that differ only in the cells they point to.
//! - A value (INT401) or an empty cell (INT402) in place of a formula.
//! - Runs, or sides of an interruption, shorter than the minimum.
//! - Two different patterns that each last fewer formulas than the minimum, such as alternating
//!   `A B A B` formulas.
//! - An odd formula at the start of a run: the run's pattern is its first formula's, so the formulas
//!   that follow differ from it and none has enough matching formulas before it; nothing is reported.
//!
//! # Parameters
//!
//! - `int_min_formula_sequence` (integer, default `3`), global only; see INT401.
//!
//! # Severity
//!
//! Always `Warning`.

pub use super::int401_interrupted_by_data::{InterruptionKind, InterruptionRule};
