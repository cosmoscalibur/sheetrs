//! sheetrs-linter: a fast, extensible linter for XLSX and ODS workbooks with hierarchical
//! violation reporting.
//!
//! The library behind the `sheetlint` executable and the language bindings. The rules exist only
//! to be run by the [`Linter`], which reads nothing itself: it analyzes a
//! [`sheetrs_core::reader::Workbook`] the caller obtained from the readers of `sheetrs-core`.
//!
//! - [`config`]: the `sheetlint.toml` model.
//! - [`rules`]: the rules, their registry and the catalog ([`rules::registry`]).
//! - [`violation`]: what a rule reports.
//! - [`report`]: the findings grouped by sheet and cell, and their flat serializable form, for
//!   every renderer (the executable, the bindings).

pub mod config;
pub mod report;
pub mod rules;
pub mod violation;

use anyhow::Result;
use sheetrs_core::reader;
use std::collections::HashSet;
use std::path::Path;

pub use config::LinterConfig;
pub use violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

use rules::walker::WorkbookWalker;

/// Main linter interface
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
    pub fn new() -> Self {
        Self::with_config(LinterConfig::default())
    }

    /// Create a linter, rejecting a configuration that names rules, categories or parameters the
    /// linter does not know: a misspelled rule would otherwise be ignored without a word.
    pub fn try_with_config(config: LinterConfig) -> Result<Self> {
        config.validate_rules(&rules::registry::get_all_valid_tokens())?;
        Ok(Self::with_config(config))
    }

    /// Create a new linter with custom configuration, as given (see
    /// [`try_with_config`](Self::try_with_config) to validate it)
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

    /// Lint a spreadsheet file and return violations
    pub fn lint_file<P: AsRef<Path>>(&self, path: P) -> Result<Vec<Violation>> {
        let workbook = reader::read_workbook(path)?;
        self.lint_workbook(&workbook)
    }

    /// Lint a spreadsheet file and return both violations and workbook
    ///
    /// This is useful when the consumer needs to resolve sheet indices
    /// to sheet names for display purposes.
    pub fn lint_file_with_workbook<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> Result<(Vec<Violation>, reader::Workbook)> {
        let workbook = reader::read_workbook(path)?;
        let violations = self.lint_workbook(&workbook)?;
        Ok((violations, workbook))
    }

    /// Lint a workbook and return violations
    pub fn lint_workbook(&self, workbook: &reader::Workbook) -> Result<Vec<Violation>> {
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

        Ok(violations)
    }

    /// Get a reference to the linter configuration
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

    #[test]
    fn a_misspelled_rule_is_rejected_when_the_linter_is_built() {
        let config = LinterConfig::from_toml("[global]\ndisabled_rules = [\"ERR1O2\"]\n").unwrap();
        let error = Linter::try_with_config(config.clone())
            .err()
            .expect("rejected");
        assert!(error.to_string().contains("ERR1O2"), "{error}");
        // `with_config` takes the configuration as given.
        let _ = Linter::with_config(config);
    }

    #[test]
    fn known_rules_and_categories_are_accepted() {
        let config =
            LinterConfig::from_toml("[global]\ndisabled_rules = [\"ERR102\", \"CPX\"]\n").unwrap();
        assert!(Linter::try_with_config(config).is_ok());
    }
}
