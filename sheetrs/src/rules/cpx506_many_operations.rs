//! CPX506: Many Operations detection
//!
//! Description: Detects formulas with an excessive number of mathematical or logical operations.
//!
//! Violations for this rule are emitted by [`super::cpx504_deep_if_nesting::DeepIfNestingRule::on_cell`]
//! as part of the CPX5xx dependency cascade. This struct is registered as a walker
//! rule only for configuration enable/disable.

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
