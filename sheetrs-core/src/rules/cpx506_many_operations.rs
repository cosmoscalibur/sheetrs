//! CPX506: Many operations
//!
//! A formula that chains many arithmetic or comparison operators packs a calculation into one
//! expression that cannot be checked at a glance; splitting it over cells or names makes each step
//! reviewable.
//!
//! This rule is emitted by [`cpx504_deep_if_nesting`](super::cpx504_deep_if_nesting), which evaluates
//! the whole CPX5xx chain for each cell. The walker rule in this module does nothing and is
//! registered only so that CPX506 can be enabled or disabled.
//!
//! # Reports
//! One violation per formula cell (scope: cell) with more than `max_operations` operators. Counted
//! outside string literals: `+ - * / ^ &` and the comparison operators `= < >`, where `<=`, `>=` and
//! `<>` count once each. The leading `=` of the formula is not counted. Unary signs count like
//! binary operators. The message gives the count. The cell is reported only if CPX504 and CPX505
//! did not fire for it; when this rule fires, CPX508 is not evaluated for the cell.
//!
//! # Does not report
//! Function calls: a formula made of many nested functions but few operators is covered by CPX505
//! or CPX509. Operators inside quoted text.
//!
//! # Parameters
//! - `max_operations` (integer, default `8`): the highest operator count still accepted.
//!
//! # Severity
//! Always `Warning`.

use super::{RuleCategory, WalkerRule};
use crate::violation::RuleId;

/// Rule that identifies formulas with too many operations.
///
/// This walker rule is a no-op. Actual violations are emitted by CPX504's
/// `on_cell` to enforce the dependency cascade.
pub struct ManyOperationsRule;

impl WalkerRule for ManyOperationsRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx506
    }

    fn name(&self) -> &'static str {
        "Many Operations"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    // All hooks are no-op: violations are emitted by CPX504's on_cell.
}
