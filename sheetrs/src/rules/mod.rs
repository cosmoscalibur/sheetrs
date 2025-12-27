//! Linter rule system

pub mod registry;

// Rule implementations
pub mod err001_error_cells;
pub mod err002_broken_named_ranges;
pub mod err003_circular_references;
pub mod form001_long_formula;
pub mod form002_volatile_functions;
pub mod form003_duplicate_formulas;
pub mod form004_whole_column_row_refs;
pub mod form005_empty_string_test;
pub mod form006_deep_formula_nesting;
pub mod form007_deep_if_nesting;
pub mod form008_hardcoded_values_in_formulas;
pub mod form009_vlookup_hlookup_usage;
pub mod perf001_unused_named_ranges;
pub mod perf002_unused_sheets;
pub mod perf003_large_used_range;
pub mod perf004_excessive_conditional_formatting;
pub mod perf005_empty_sheets;
pub mod sec001_external_workbooks;
pub mod sec002_hidden_sheets;
pub mod sec003_hidden_columns_rows;
pub mod sec004_has_macros;
pub mod sec005_web_urls;

pub mod sm001_excessive_sheet_counts;
pub mod sm002_duplicate_sheet_names;
pub mod sm003_long_text_cell;
pub mod sm004_merged_cells;
pub mod sm005_non_descriptive_sheet_name;
pub mod ux001_inconsistent_number_format;
pub mod ux002_inconsistent_date_format;
pub mod ux003_blank_rows_columns;

use crate::reader::Workbook;
use crate::violation::Violation;
use anyhow::Result;

/// Trait that all linter rules must implement
pub trait LinterRule: Send + Sync {
    /// Unique rule identifier (e.g., "ERR001")
    fn id(&self) -> &str;

    /// Human-readable rule name
    fn name(&self) -> &str;

    /// Rule category
    fn category(&self) -> RuleCategory;

    /// Check the workbook for violations
    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>>;
}

/// Rule categories
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleCategory {
    UnresolvedErrors,
    SecurityAndPrivacy,
    FormattingAndUsability,
    StructuralAndMaintainability,
    Performance,
    Formula,
}

impl RuleCategory {
    pub fn as_str(&self) -> &str {
        match self {
            RuleCategory::UnresolvedErrors => "Unresolved Errors",
            RuleCategory::SecurityAndPrivacy => "Security and Privacy",
            RuleCategory::FormattingAndUsability => "Formatting and Usability",
            RuleCategory::StructuralAndMaintainability => "Structural and Maintainability",
            RuleCategory::Performance => "Performance",
            RuleCategory::Formula => "Formula",
        }
    }
}
