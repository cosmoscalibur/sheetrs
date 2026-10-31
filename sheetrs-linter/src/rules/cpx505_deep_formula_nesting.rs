//! CPX505: Deep formula nesting
//!
//! Formulas with many levels of parentheses are hard to read, understand and debug; intermediate
//! helper cells or `LET` make each step visible.
//!
//! This rule is emitted by [`cpx504_deep_if_nesting`](super::cpx504_deep_if_nesting), which evaluates
//! the whole CPX5xx chain for each cell. The walker rule in this module does nothing and is
//! registered only so that CPX505 can be enabled or disabled.
//!
//! # Reports
//! One violation per formula cell (scope: cell) whose parenthesis nesting depth is greater than
//! `max_formula_nesting`. The depth is the largest number of `(` open at the same time in the
//! formula text, for every kind of function call or grouping. Parentheses inside string literals and
//! quoted sheet names are not counted. A cell is only reported here if CPX504 did not fire for it; if this rule fires,
//! CPX506, CPX508 and CPX509 are not evaluated for the cell.
//!
//! # Does not report
//! Deep IF nesting, which CPX504 reports first and more specifically. Formulas that are long or have
//! many operators but are shallow ([`cpx506_many_operations`](super::cpx506_many_operations),
//! [`cpx509_long_formula`](super::cpx509_long_formula)).
//!
//! # Parameters
//! - `max_formula_nesting` (integer, default `5`): the deepest nesting still accepted.
//!
//! # Severity
//! Always `Warning`.

use super::{RuleCategory, WalkerRule};
use crate::violation::RuleId;

/// Rule that detects excessively nested formulas.
///
/// This walker rule is a no-op. Actual violations are emitted by CPX504's
/// `on_cell` to enforce the dependency cascade.
pub struct DeepFormulaNestingRule;

impl WalkerRule for DeepFormulaNestingRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx505
    }

    fn name(&self) -> &'static str {
        "Many Nested Functions"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    // All hooks are no-op: violations are emitted by CPX504's on_cell.
}
