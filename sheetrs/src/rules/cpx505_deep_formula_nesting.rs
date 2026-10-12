//! CPX505: Deep formula nesting detection
//!
//! Description: Excessively nested formulas are hard to read, understand, and debug.
//!
//! Violations for this rule are emitted by [`cpx504_deep_if_nesting::DeepIfNestingRule::on_cell`]
//! as part of the CPX5xx dependency cascade. This struct is registered as a walker
//! rule only for configuration enable/disable.

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
