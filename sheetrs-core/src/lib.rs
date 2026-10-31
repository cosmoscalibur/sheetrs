//! sheetrs-core: Core library for Excel/ODS linting
//!
//! This library provides a fast, extensible linting framework for spreadsheet files
//! with hierarchical violation reporting.

pub mod config;
pub mod reader;
pub mod rules;
pub mod violation;
pub mod writer;

use anyhow::Result;
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

    /// Create a new linter with custom configuration
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
