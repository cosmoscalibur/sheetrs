//! sheetrs-linter: a fast, extensible linter for XLSX and ODS workbooks with hierarchical
//! violation reporting.
//!
//! The library behind `sheetrs lint` and the language bindings. The rules exist only
//! to be run by the [`Linter`], which reads nothing itself: it analyzes a
//! [`sheetrs_io::reader::Workbook`] the caller obtained from the readers of `sheetrs-io`.
//!
//! - [`config`]: the `sheetrs.toml` model.
//! - [`rules`]: the rules, their registry and the catalog ([`rules::registry`]).
//! - [`violation`]: what a rule reports.
//! - `human`, `json` and [`sarif`]: the findings of a run for a person, for an interface of
//!   your own and as a SARIF 2.1.0 log; [`output`] is the one entry point to the three.
//! - [`report`]: the findings grouped by sheet and cell, and their flat serializable form, for
//!   every renderer (the command, the bindings).

#[cfg(feature = "cli")]
pub mod cli;
pub mod config;
mod error;
mod human;
mod json;
pub mod output;
pub mod report;
pub mod rules;
pub mod sarif;
pub mod violation;

use sheetrs_io::Book;
use sheetrs_io::reader::{self};
use std::collections::HashSet;

pub use config::LinterConfig;
pub use error::{Error, Result};
pub use output::OutputFormat;
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

    /// Lints a workbook already read, for a caller that needs the model too (to name the sheets of
    /// the violations, for example) or that read it from somewhere a [`Book`] cannot describe.
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

/// The result of linting a workbook: its [`violations`](Self::violations), ordered by scope (book, sheet,
/// cell) and rule, and what is needed to write them in each of the three forms of `sheetrs lint`: the
/// report for a person ([`to_text`](Self::to_text)), the JSON document ([`to_json`](Self::to_json)) and the
/// SARIF 2.1.0 log ([`to_sarif`](Self::to_sarif)), or into a writer ([`write`](Self::write)). `file` is the
/// name the documents give the workbook (the path of the book by default, `workbook.xlsx` for bytes).
#[derive(Debug, Clone)]
pub struct LintReport {
    violations: Vec<Violation>,
    workbook: sheetrs_io::reader::Workbook,
    config: LinterConfig,
    file: String,
}

impl LintReport {
    /// The violations, in the order of the report.
    #[must_use]
    pub fn violations(&self) -> &[Violation] {
        &self.violations
    }

    /// The violations with everything resolved (rule and sheet names, the cell in A1 notation, the
    /// message), in the same order.
    #[must_use]
    pub fn records(&self) -> Vec<report::ViolationRecord> {
        report::violation_records(&self.violations, &self.workbook, &self.config)
    }

    /// The name the documents give the workbook when none is asked for.
    #[must_use]
    pub fn file(&self) -> &str {
        &self.file
    }

    /// Writes the report in `format` to `out`; `file` names the workbook in the document (see
    /// [`file`](Self::file)).
    ///
    /// # Errors
    ///
    /// A write error of `out`.
    pub fn write(
        &self,
        out: &mut impl std::io::Write,
        format: OutputFormat,
        file: Option<&str>,
    ) -> std::io::Result<()> {
        output::write_report(
            out,
            format,
            file.unwrap_or(&self.file),
            &self.violations,
            &self.workbook,
            &self.config,
            false,
        )
    }

    fn text(&self, format: OutputFormat, file: Option<&str>) -> String {
        let mut out = Vec::new();
        // A `Vec` does not fail and the writers produce UTF-8.
        let _ = self.write(&mut out, format, file);
        String::from_utf8_lossy(&out).into_owned()
    }

    /// The report for a person, as plain text.
    #[must_use]
    pub fn to_text(&self, file: Option<&str>) -> String {
        self.text(OutputFormat::Human, file)
    }

    /// The JSON document of `sheetrs lint --format json` (`file`, `summary`, `violations`).
    #[must_use]
    pub fn to_json(&self, file: Option<&str>) -> String {
        self.text(OutputFormat::Json, file)
    }

    /// The SARIF 2.1.0 log of `sheetrs lint --format sarif`.
    #[must_use]
    pub fn to_sarif(&self, file: Option<&str>) -> String {
        self.text(OutputFormat::Sarif, file)
    }
}

impl PartialEq for LintReport {
    fn eq(&self, other: &Self) -> bool {
        self.violations == other.violations && self.file == other.file
    }
}

/// What a [`Book`] is linted with.
pub trait BookLint {
    /// Lints the workbook with `linter`: the [`LintReport`] of its rules.
    ///
    /// ```no_run
    /// use sheetrs_io::Book;
    /// use sheetrs_linter::{BookLint, Linter};
    ///
    /// # fn main() -> sheetrs_io::Result<()> {
    /// let book = Book::from_file("book.xlsx")?;
    /// let report = book.lint(&Linter::new())?;
    /// println!("{} violations", report.violations().len());
    /// std::fs::write("book.sarif", report.to_sarif(None)).expect("a writable folder");
    /// # Ok(()) }
    /// ```
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read.
    fn lint(&self, linter: &Linter) -> sheetrs_io::Result<LintReport>;
}

impl BookLint for Book {
    fn lint(&self, linter: &Linter) -> sheetrs_io::Result<LintReport> {
        let workbook = self.workbook()?;
        let violations = linter.lint_workbook(&workbook);
        let file = self.path().map_or_else(
            || format!("workbook.{}", self.format().extension()),
            |path| path.display().to_string(),
        );
        Ok(LintReport {
            violations,
            workbook,
            config: linter.config().clone(),
            file,
        })
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

    const XLSX: &[u8] = include_bytes!("../../tests/lint_book.xlsx");

    #[test]
    fn lint_takes_a_book_and_gives_what_linting_its_workbook_gives() {
        let source = Book::from_bytes(XLSX).unwrap();
        let linter = Linter::new();
        let workbook = source.workbook().unwrap();
        assert_eq!(
            source.lint(&linter).unwrap().violations(),
            linter.lint_workbook(&workbook)
        );
    }

    #[test]
    fn lint_reports_a_workbook_that_cannot_be_read_with_the_core_error() {
        // A file that is not there is known when the source is made.
        assert!(matches!(
            Book::from_file("does-not-exist.xlsx"),
            Err(sheetrs_io::Error::Open { .. })
        ));
    }
}
