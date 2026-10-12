//! CPX509: Long formula detection
//!
//! Description: Excessively long formulas are hard to read, maintain, and debug.
//!
//! Violations for this rule are emitted by [`cpx504_deep_if_nesting::DeepIfNestingRule::on_cell`]
//! as part of the CPX5xx dependency cascade. This struct is registered as a walker
//! rule only for configuration enable/disable.

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
