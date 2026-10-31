//! CPX508: Many references
//!
//! A formula that reads many individual cells or ranges is hard to audit: each reference has to be
//! traced, and a mistake in one of them is difficult to see.
//!
//! This rule is emitted by [`cpx504_deep_if_nesting`](super::cpx504_deep_if_nesting), which evaluates
//! the whole CPX5xx chain for each cell. The walker rule in this module does nothing and is
//! registered only so that CPX508 can be enabled or disabled.
//!
//! # Reports
//! One violation per formula cell (scope: cell) with more than `max_references` references. Counted
//! outside string literals: each token of the form `A1`, `$B$2` or `AB12` (one to three letters,
//! then digits). Both ends of a range count, so `A1:B2` is two. The count is textual, so a
//! sheet-name prefix that itself looks like a reference can add to it. The message gives the
//! count. A cell is checked only if CPX504, CPX505 and CPX506 did not fire: an expression that
//! already has too many operators is reported there and not a second time here. CPX509 is evaluated
//! independently of this rule.
//!
//! # Does not report
//! Whole-column or whole-row references (`A:A`, `1:1`), named ranges and function names. References
//! to many different sheets are CPX507.
//!
//! # Parameters
//! - `max_references` (integer, default `10`): the highest reference count still accepted.
//!
//! # Severity
//! Always `Warning`.

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
