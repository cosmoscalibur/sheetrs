//! CPX509: Long formula
//!
//! A very long formula is difficult to read, review and change, whatever it contains.
//!
//! This rule is emitted by [`cpx504_deep_if_nesting`](super::cpx504_deep_if_nesting), which evaluates
//! the whole CPX5xx chain for each cell. The walker rule in this module does nothing and is
//! registered only so that CPX509 can be enabled or disabled.
//!
//! # Reports
//! One violation per formula cell (scope: cell) whose text has more than `max_formula_length`
//! characters. Characters are counted, not bytes, and the leading `=` is part of the length. The
//! namespace prefix of a function (`_xlfn.`, `COM.MICROSOFT.`) is not: both spellings of
//! `CONCAT` count as `CONCAT`, as the application shows it. The
//! message gives the length. The cell is checked only if CPX504 and CPX505 did not fire; CPX506 does
//! not suppress it, so a cell can be reported both here and by CPX508.
//!
//! # Does not report
//! A cell for which CPX504 or CPX505 already fired: that finding is reported in this rule's place.
//! Non-formula cells, so a long text constant is not reported.
//!
//! # Parameters
//! - `max_formula_length` (integer, default `255`): the longest formula text still accepted.
//!
//! # Severity
//! Always `Warning`.

use super::{RuleCategory, WalkerRule};
use crate::violation::RuleId;

/// Rule that detects excessively long formulas.
///
/// This walker rule is a no-op. Actual violations are emitted by CPX504's
/// `on_cell` to enforce the dependency cascade.
pub struct LongFormulaRule;

impl WalkerRule for LongFormulaRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx509
    }

    fn name(&self) -> &'static str {
        "Long Formula"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    // All hooks are no-op: violations are emitted by CPX504's on_cell.
}
