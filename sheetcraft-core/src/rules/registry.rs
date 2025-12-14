//! Rule registry for managing and creating rule instances

use super::*;
use crate::config::LinterConfig;
use std::collections::HashSet;

/// Get all valid configuration tokens (Rule IDs, Category Prefixes, "ALL")
pub fn get_all_valid_tokens() -> HashSet<String> {
    let mut tokens = HashSet::new();
    tokens.insert("ALL".to_string());

    // Category prefixes
    let prefixes = ["ERR", "SEC", "UX", "SM", "PERF", "FORM"];
    for prefix in prefixes {
        tokens.insert(prefix.to_string());
    }

    // Rule IDs
    let config = LinterConfig::default();
    let rules = create_all_rules(&config);
    for rule in rules {
        tokens.insert(rule.id().to_string());
    }

    tokens
}

/// Create all enabled rules based on configuration
pub fn create_enabled_rules(config: &LinterConfig) -> Vec<Box<dyn LinterRule>> {
    let all_rules = create_all_rules(config);

    all_rules
        .into_iter()
        .filter(|rule| {
            // Use default active status if not explicitly configured
            let default_enabled = rule.default_active();
            config.is_rule_enabled(rule.id())
                && (default_enabled || config.global.enabled_rules.contains(rule.id()))
        })
        .collect()
}

/// Create instances of all available rules
fn create_all_rules(config: &LinterConfig) -> Vec<Box<dyn LinterRule>> {
    vec![
        Box::new(err001_error_cells::ErrorCellsRule),
        Box::new(err002_broken_named_ranges::BrokenNamedRangesRule),
        Box::new(sec001_external_links::ExternalLinksRule::new(config)),
        Box::new(sec002_hidden_sheets::HiddenSheetsRule),
        Box::new(sec003_hidden_columns_rows::HiddenColumnsRowsRule),
        Box::new(sec004_macros_vba::MacrosVbaRule),
        Box::new(sec005_possible_corruption::PossibleCorruptionRule::new(
            config,
        )),
        Box::new(ux001_inconsistent_number_format::InconsistentNumberFormatRule),
        Box::new(ux003_blank_rows_columns::BlankRowsColumnsRule::new(config)),
        Box::new(perf001_unused_named_ranges::UnusedNamedRangesRule),
        Box::new(perf002_unused_sheets::UnusedSheetsRule),
        Box::new(perf003_large_used_range::LargeUsedRangeRule::new(config)),
        Box::new(form002_volatile_functions::VolatileFunctionsRule::new(
            config,
        )),
        Box::new(form003_duplicate_formulas::DuplicateFormulasRule),
        Box::new(
            perf004_excessive_conditional_formatting::ExcessiveConditionalFormattingRule::new(
                config,
            ),
        ),
        Box::new(form004_whole_column_row_refs::WholeColumnRowRefsRule::new()),
        Box::new(sm001_excessive_sheet_counts::ExcessiveSheetCountsRule::new(
            config,
        )),
        Box::new(sm002_duplicate_sheet_names::DuplicateSheetNamesRule),
        Box::new(form001_long_formula::LongFormulaRule::new(config)),
        Box::new(sm003_long_text_cell::LongTextCellRule::new(config)),
        Box::new(sm004_merged_cells::MergedCellsRule),
        Box::new(sm005_non_descriptive_sheet_name::NonDescriptiveSheetNameRule::new(config)),
        Box::new(form005_empty_string_test::EmptyStringTestRule::new()),
        Box::new(form006_deep_formula_nesting::DeepFormulaNestingRule::new(
            config,
        )),
        Box::new(form007_deep_if_nesting::DeepIfNestingRule::new(config)),
        Box::new(form008_hardcoded_values_in_formulas::HardcodedValuesInFormulasRule::new(config)),
        Box::new(form009_vlookup_hlookup_usage::VLookupHLookupUsageRule::new(
            config,
        )),
        Box::new(ux002_inconsistent_date_format::InconsistentDateFormatRule::new(config)),
        Box::new(err003_circular_references::CircularReferenceRule::new()),
    ]
}

/// Get all rule IDs that are active by default
pub fn default_active_rule_ids() -> Vec<&'static str> {
    vec![
        "ERR001", "ERR002", "ERR003", "SEC001", "UX001", "PERF001", "PERF002", "PERF003", "SM001",
        "SM002", "SM005", "FORM002", "FORM003", "FORM004", "FORM005", "FORM008", "FORM009",
    ]
}
