//! CPX508: Many References detection
//!
//! Description: Detects formulas with an excessive number of individual cell or range references.
//!
//! Violations for this rule are emitted by [`super::cpx504_deep_if_nesting::DeepIfNestingRule::on_cell`]
//! as part of the CPX5xx dependency cascade. This struct is registered as a walker
//! rule only for configuration enable/disable.

use super::{RuleCategory, WalkerRule};
use crate::violation::RuleId;

/// Rule that identifies formulas with many references.
///
/// This walker rule is a no-op. Actual violations are emitted by CPX504's
/// `on_cell` to enforce the dependency cascade.
pub struct ManyReferencesRule;

impl WalkerRule for ManyReferencesRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx508
    }

    fn name(&self) -> &'static str {
        "Many References"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    // All hooks are no-op: violations are emitted by CPX504's on_cell.
}
