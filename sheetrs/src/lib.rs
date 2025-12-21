//! sheetrs: Core library for Excel/ODS linting
//!
//! This library provides a fast, extensible linting framework for spreadsheet files
//! with hierarchical violation reporting.

pub mod config;
pub mod reader;
pub mod rules;
pub mod violation;
pub mod writer;

use anyhow::Result;
use std::path::Path;

pub use config::LinterConfig;
pub use rules::LinterRule;
pub use violation::{Severity, Violation, ViolationScope};

/// Main linter interface
pub struct Linter {
    config: LinterConfig,
    rules: Vec<Box<dyn LinterRule>>,
}

impl Linter {
    /// Create a new linter with default configuration
    pub fn new() -> Self {
        Self::with_config(LinterConfig::default())
    }

    /// Create a new linter with custom configuration
    pub fn with_config(config: LinterConfig) -> Self {
        let rules = rules::registry::create_enabled_rules(&config);
        Self { config, rules }
    }

    /// Lint a spreadsheet file and return violations
    pub fn lint_file<P: AsRef<Path>>(&self, path: P) -> Result<Vec<Violation>> {
        let workbook = reader::read_workbook(path)?;
        let mut violations = Vec::new();

        for rule in &self.rules {
            let rule_violations = rule.check(&workbook)?;

            // Filter violations based on sheet configuration
            for violation in rule_violations {
                let enabled = if let Some(sheet_name) = violation.scope.sheet_name() {
                    self.config
                        .is_rule_enabled_for_sheet(&violation.rule_id, sheet_name)
                } else {
                    // Book-level violations are enabled if the rule itself is enabled (config logic handles this)
                    // But wait, the rules vector already contains only globally enabled rules.
                    // However, we should double check if there's any reason a book-level rule would be disabled?
                    // Usually book-level rules aren't sheet-specific, so default to true here.
                    true
                };

                if enabled {
                    violations.push(violation);
                }
            }
        }

        // Sort violations by scope for hierarchical reporting
        violations.sort_by(|a, b| a.scope.cmp(&b.scope));

        Ok(violations)
    }
}

impl Default for Linter {
    fn default() -> Self {
        Self::new()
    }
}
