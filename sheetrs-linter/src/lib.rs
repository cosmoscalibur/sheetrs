//! sheetrs-linter: a fast, extensible linter for XLSX and ODS workbooks with hierarchical
//! violation reporting.
//!
//! The library behind the `sheetlint` executable and the language bindings. The rules exist only
//! to be run by the [`Linter`], which reads nothing itself: it analyzes a
//! [`sheetrs_io::reader::Workbook`] the caller obtained from the readers of `sheetrs-io`.
//!
//! - [`config`]: the `sheetlint.toml` model.
//! - [`rules`]: the rules, their registry and the catalog ([`rules::registry`]).
//! - [`violation`]: what a rule reports.
//! - [`report`]: the findings grouped by sheet and cell, and their flat serializable form, for
//!   every renderer (the executable, the bindings).

pub mod config;
mod error;
pub mod report;
pub mod rules;
pub mod violation;

use sheetrs_io::reader::{self, Source};
use std::collections::HashSet;

pub use config::LinterConfig;
pub use error::{Error, Result};
pub use violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

use rules::walker::WorkbookWalker;

/// Main linter interface
#[derive(Debug)]
pub struct Linter {
    config: LinterConfig,
    /// Enabled rules and their providers, in execution order; instantiated
    /// afresh for every lint run.
    walker_rule_ids: Vec<RuleId>,
    /// Providers that run only for enabled dependents: their violations are
    /// dropped. Empty unless an enabled rule has a provider that is not enabled.
    unreported_rule_ids: HashSet<RuleId>,
}

impl Linter {
    /// Create a new linter with default configuration
    #[must_use]
    pub fn new() -> Self {
        Self::with_config(LinterConfig::default())
    }

    /// Create a new linter with custom configuration (a [`LinterConfig`] is valid by construction)
    #[must_use]
    pub fn with_config(config: LinterConfig) -> Self {
        let all_rule_ids: Vec<RuleId> = rules::registry::create_all_walker_rules(&config)
            .iter()
            .map(|rule| rule.id())
            .collect();
        let reported_rule_ids: HashSet<RuleId> = all_rule_ids
            .iter()
            .copied()
            .filter(|&id| rules::registry::is_rule_active(id, &config))
            .collect();
        let walker_rule_ids = rules::registry::run_rule_ids(&all_rule_ids, &reported_rule_ids);
        let unreported_rule_ids = walker_rule_ids
            .iter()
            .copied()
            .filter(|id| !reported_rule_ids.contains(id))
            .collect();
        Self {
            config,
            walker_rule_ids,
            unreported_rule_ids,
        }
    }

    /// Lints the workbook `source` refers to.
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read.
    pub fn lint(&self, source: &Source) -> sheetrs_io::Result<Vec<Violation>> {
        Ok(self.lint_workbook(&source.read_workbook()?))
    }

    /// Lints a workbook already read, for a caller that needs the model too (to name the sheets of
    /// the violations, for example) or that read it from somewhere a [`Source`] cannot describe.
    #[must_use]
    pub fn lint_workbook(&self, workbook: &reader::Workbook) -> Vec<Violation> {
        let mut violations = Vec::new();

        let walker_rules =
            rules::registry::create_walker_rules(&self.walker_rule_ids, &self.config);
        let walker = WorkbookWalker::new(workbook, walker_rules);
        let walker_violations = walker.walk();

        // Drop violations of providers that are not enabled, then filter by
        // sheet config
        for violation in walker_violations {
            if self.unreported_rule_ids.contains(&violation.rule_id) {
                continue;
            }
            let enabled = if let Some(sheet_idx) = violation.scope.sheet_index() {
                if let Some(sheet_name) = workbook.sheet_name_by_index(sheet_idx) {
                    self.config
                        .is_rule_enabled_for_sheet(violation.rule_id.as_str(), sheet_name)
                } else {
                    true
                }
            } else {
                true
            };
            if enabled {
                violations.push(violation);
            }
        }

        // Hierarchical order by scope, then rule. Stable sort: ties keep each
        // rule's emission order.
        violations.sort();

        violations
    }

    /// Get a reference to the linter configuration
    #[must_use]
    pub fn config(&self) -> &LinterConfig {
        &self.config
    }
}

impl Default for Linter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_io::reader::Format;

    const XLSX: &[u8] = include_bytes!("../../tests/minimal_test.xlsx");

    #[test]
    fn lint_takes_a_source_and_gives_what_linting_its_workbook_gives() {
        let source = Source::from_bytes(XLSX, Format::Xlsx);
        let linter = Linter::new();
        let workbook = source.read_workbook().unwrap();
        assert_eq!(
            linter.lint(&source).unwrap(),
            linter.lint_workbook(&workbook)
        );
    }

    #[test]
    fn lint_reports_a_workbook_that_cannot_be_read_with_the_core_error() {
        let missing = Source::from_file("does-not-exist.xlsx").unwrap();
        assert!(matches!(
            Linter::new().lint(&missing),
            Err(sheetrs_io::Error::Open { .. })
        ));
    }
}
