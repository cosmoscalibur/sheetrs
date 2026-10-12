//! Rule registry for managing and creating rule instances

use super::*;
use crate::config::LinterConfig;
use crate::violation::RuleId;
use int401_interrupted_by_data::{InterruptionKind, InterruptionRule};
use std::collections::HashSet;

/// List of rule IDs that are active by default
pub const DEFAULT_ACTIVE_RULES: &[RuleId] = &[
    RuleId::Err101,
    RuleId::Err102,
    RuleId::Err103,
    RuleId::Calc201,
    RuleId::Calc202,
    RuleId::Calc203,
    RuleId::Calc204,
    RuleId::Calc205,
    RuleId::Ref301,
    RuleId::Ref302,
    RuleId::Ref303,
    RuleId::Ref304,
    RuleId::Ref305,
    RuleId::Ref306,
    RuleId::Ref307,
    RuleId::Ref309,
    RuleId::Ref310,
    RuleId::Int401,
    RuleId::Int402,
    RuleId::Int403,
    RuleId::Cpx501,
    RuleId::Cpx502,
    RuleId::Cpx503,
    RuleId::Cpx504,
    RuleId::Cpx505,
    RuleId::Cpx506,
    RuleId::Cpx507,
    RuleId::Cpx508,
    RuleId::Cpx509,
    RuleId::Vul601,
    RuleId::Vul602,
    RuleId::Vul603,
    RuleId::Vul604,
    RuleId::Vul605,
    RuleId::Vul606,
    RuleId::Data701,
    RuleId::Data702,
    RuleId::Data703,
    RuleId::Data704,
    RuleId::Data705,
    RuleId::Data706,
    RuleId::Ext801,
    RuleId::Ext802,
    RuleId::Hid901,
    RuleId::Hid902,
    RuleId::File1001,
    RuleId::File1002,
    RuleId::File1003,
    RuleId::Vba1101,
];

/// Get all valid configuration tokens (Rule IDs, Category Prefixes, "ALL")
pub fn get_all_valid_tokens() -> HashSet<String> {
    let mut tokens = HashSet::new();
    tokens.insert("ALL".to_string());

    // Category prefixes matching hierarchical system
    let prefixes = [
        "ERR", "CALC", "REF", "INT", "CPX", "VUL", "DATA", "EXT", "HID", "FILE", "VBA",
    ];
    for prefix in prefixes {
        tokens.insert(prefix.to_string());
    }

    // Rule IDs from walker rules
    let config = LinterConfig::default();
    let walker_rules = create_all_walker_rules(&config);
    for rule in walker_rules {
        tokens.insert(rule.id().to_string());
    }

    tokens
}

/// Metadata for a single linter rule (for API/UI consumers).
pub struct RuleMetadata {
    /// Unique rule identifier.
    pub id: RuleId,
    /// Human-readable rule name.
    pub name: String,
    /// Rule category.
    pub category: RuleCategory,
    /// Whether the rule is active by default.
    pub is_default: bool,
}

/// Collect metadata from all rules.
pub fn get_all_rule_metadata(config: &LinterConfig) -> Vec<RuleMetadata> {
    let mut metadata: Vec<RuleMetadata> = create_all_walker_rules(config)
        .into_iter()
        .map(|rule| RuleMetadata {
            id: rule.id(),
            name: rule.name().to_string(),
            category: rule.category(),
            is_default: DEFAULT_ACTIVE_RULES.contains(&rule.id()),
        })
        .collect();

    // Sort by rule ID string for stable ordering
    metadata.sort_by(|a, b| a.id.as_str().cmp(b.id.as_str()));
    metadata
}

/// Create instances of all available walker rules
pub fn create_all_walker_rules(config: &LinterConfig) -> Vec<Box<dyn WalkerRule>> {
    let mut rules: Vec<Box<dyn WalkerRule>> = vec![
        // Excel Errors (1xx)
        Box::new(err101_broken_named_ranges::BrokenNamedRangesRule),
        // Unreliable Calculations (2xx)
        Box::new(calc201_hardcoded_values::HardcodedValuesInFormulasRule::new(config)),
        Box::new(calc202_circular_references::CircularReferenceRule::new()),
        // ERR102/ERR103 after CALC202 so on_workbook_end reads circular_cells
        Box::new(err102_error_cells::ErrorCellsRule),
        Box::new(err103_ref_to_error::RefToErrorRule),
        Box::new(calc203_double_operator::DoubleOperatorRule::new(config)),
        Box::new(calc204_approximate_lookup::ApproximateLookupRule),
        Box::new(calc205_double_count::DoubleCountRule),
        // Reference Issues (3xx)
        Box::new(ref301_unused_named_ranges::UnusedNamedRangesRule::new()),
        Box::new(ref302_duplicate_names::DuplicateSheetNamesRule),
        Box::new(ref303_empty_sheets::EmptySheetsRule),
        Box::new(ref304_large_used_range::LargeUsedRangeRule::new()),
        Box::new(ref305_blank_rows_columns::BlankRowsColumnsRule::new()),
        Box::new(ref306_unused_sheets::UnusedSheetsRule),
        Box::new(ref307_whole_column_row_refs::WholeColumnRowRefsRule::new()),
        Box::new(ref309_ref_to_empty_cell::RefToEmptyCellRule),
        Box::new(ref310_longer_ref_expected::LongerRefExpectedRule::new(
            config,
        )),
    ];

    // Formula Interruptions (4xx)
    for kind in [
        InterruptionKind::Data,
        InterruptionKind::Empty,
        InterruptionKind::Other,
    ] {
        rules.push(Box::new(InterruptionRule::new(vec![kind], config)));
    }

    // Remaining rules
    rules.extend(vec![
        // Complexity (5xx)
        Box::new(cpx501_sheet_counts::ExcessiveSheetCountsRule::new(config)) as Box<dyn WalkerRule>,
        Box::new(cpx502_merged_cells::MergedCellsRule),
        Box::new(
            cpx503_excessive_conditional_formatting::ExcessiveConditionalFormattingRule::new(
                config,
            ),
        ),
        Box::new(cpx504_deep_if_nesting::DeepIfNestingRule::new(config)),
        Box::new(cpx505_deep_formula_nesting::DeepFormulaNestingRule),
        Box::new(cpx506_many_operations::ManyOperationsRule),
        Box::new(cpx507_multiple_sheet_ref::MultipleSheetRefRule::new(config)),
        Box::new(cpx508_many_references::ManyReferencesRule),
        Box::new(cpx509_long_formula::LongFormulaRule),
        // Vulnerable Formulas (6xx)
        Box::new(vul601_duplicate_formulas::DuplicateFormulasRule::new()),
        Box::new(vul602_volatile_functions::VolatileFunctionsRule::new()),
        Box::new(vul603_empty_string_test::EmptyStringTestRule::new()),
        Box::new(vul604_error_prone_functions::ErrorProneFunctionsRule::new(
            config,
        )),
        Box::new(vul605_legacy_array::LegacyArrayRule::new()),
        Box::new(vul606_deprecated_func::DeprecatedFuncRule::new()),
        // Data Issues (7xx)
        Box::new(dat701_sheet_names::NonDescriptiveSheetNameRule::new(config)),
        Box::new(dat702_numeric_formats::InconsistentNumberFormatRule::new()),
        Box::new(dat703_date_formats::InconsistentDateFormatRule::new(config)),
        Box::new(dat704_long_text::LongTextCellRule::new(config)),
        Box::new(dat705_unnecessary_space::UnnecessarySpaceRule),
        Box::new(dat706_numeric_text_calc::NumericTextCalcRule),
        // External References (8xx)
        Box::new(ext801_external_workbook::ExternalWorkbooksRule::new()),
        Box::new(ext802_web_urls::WebUrlsRule::new(config)),
        // Hidden Information (9xx)
        Box::new(hid901_hidden_worksheet::HiddenWorksheetRule),
        Box::new(hid902_hidden_columns_rows::HiddenColumnsRowsRule),
        // Files & Settings (10xx)
        Box::new(file1001_large_file_size::LargeFileSizeRule::new(config)),
        Box::new(file1002_old_spreadsheet::OldSpreadsheetRule::new(config)),
        Box::new(file1003_date_system_1904::DateSystem1904Rule::new()),
        // VBA Issues (11xx)
        Box::new(vba1101_has_macros::HasMacrosRule),
    ]);

    rules
}

/// Create enabled walker rules based on configuration
pub fn create_enabled_walker_rules(config: &LinterConfig) -> Vec<Box<dyn WalkerRule>> {
    let all_rules = create_all_walker_rules(config);

    all_rules
        .into_iter()
        .filter(|rule| {
            let is_enabled_in_config = config.is_rule_enabled(rule.id().as_str());

            if config.global.enabled_rules.is_empty() {
                DEFAULT_ACTIVE_RULES.contains(&rule.id()) && is_enabled_in_config
            } else {
                is_enabled_in_config
            }
        })
        .collect()
}

/// Creates fresh instances of the given rules, in order, for one lint run.
///
/// INT401–403 share one [`InterruptionRule`] reporting every requested kind,
/// placed at the position of the first of them.
pub fn create_walker_rules(ids: &[RuleId], config: &LinterConfig) -> Vec<Box<dyn WalkerRule>> {
    let interruption_kinds: Vec<InterruptionKind> = ids
        .iter()
        .filter_map(|&id| InterruptionKind::from_rule_id(id))
        .collect();
    let mut interruption_kinds = Some(interruption_kinds);
    let mut rules: Vec<Box<dyn WalkerRule>> = Vec::with_capacity(ids.len());
    for &id in ids {
        if InterruptionKind::from_rule_id(id).is_none() {
            rules.push(create_walker_rule(id, config));
        } else if let Some(kinds) = interruption_kinds.take() {
            rules.push(Box::new(InterruptionRule::new(kinds, config)));
        }
    }
    rules
}

/// Creates a fresh instance of a single rule.
fn create_walker_rule(id: RuleId, config: &LinterConfig) -> Box<dyn WalkerRule> {
    match id {
        RuleId::Err101 => Box::new(err101_broken_named_ranges::BrokenNamedRangesRule),
        RuleId::Err102 => Box::new(err102_error_cells::ErrorCellsRule),
        RuleId::Err103 => Box::new(err103_ref_to_error::RefToErrorRule),
        RuleId::Calc201 => {
            Box::new(calc201_hardcoded_values::HardcodedValuesInFormulasRule::new(config))
        }
        RuleId::Calc202 => Box::new(calc202_circular_references::CircularReferenceRule::new()),
        RuleId::Calc203 => Box::new(calc203_double_operator::DoubleOperatorRule::new(config)),
        RuleId::Calc204 => Box::new(calc204_approximate_lookup::ApproximateLookupRule),
        RuleId::Calc205 => Box::new(calc205_double_count::DoubleCountRule),
        RuleId::Ref301 => Box::new(ref301_unused_named_ranges::UnusedNamedRangesRule::new()),
        RuleId::Ref302 => Box::new(ref302_duplicate_names::DuplicateSheetNamesRule),
        RuleId::Ref303 => Box::new(ref303_empty_sheets::EmptySheetsRule),
        RuleId::Ref304 => Box::new(ref304_large_used_range::LargeUsedRangeRule::new()),
        RuleId::Ref305 => Box::new(ref305_blank_rows_columns::BlankRowsColumnsRule::new()),
        RuleId::Ref306 => Box::new(ref306_unused_sheets::UnusedSheetsRule),
        RuleId::Ref307 => Box::new(ref307_whole_column_row_refs::WholeColumnRowRefsRule::new()),
        RuleId::Ref309 => Box::new(ref309_ref_to_empty_cell::RefToEmptyCellRule),
        RuleId::Ref310 => Box::new(ref310_longer_ref_expected::LongerRefExpectedRule::new(
            config,
        )),
        RuleId::Int401 | RuleId::Int402 | RuleId::Int403 => Box::new(InterruptionRule::new(
            InterruptionKind::from_rule_id(id).into_iter().collect(),
            config,
        )),
        RuleId::Cpx501 => Box::new(cpx501_sheet_counts::ExcessiveSheetCountsRule::new(config)),
        RuleId::Cpx502 => Box::new(cpx502_merged_cells::MergedCellsRule),
        RuleId::Cpx503 => Box::new(
            cpx503_excessive_conditional_formatting::ExcessiveConditionalFormattingRule::new(
                config,
            ),
        ),
        RuleId::Cpx504 => Box::new(cpx504_deep_if_nesting::DeepIfNestingRule::new(config)),
        RuleId::Cpx505 => Box::new(cpx505_deep_formula_nesting::DeepFormulaNestingRule),
        RuleId::Cpx506 => Box::new(cpx506_many_operations::ManyOperationsRule),
        RuleId::Cpx507 => Box::new(cpx507_multiple_sheet_ref::MultipleSheetRefRule::new(config)),
        RuleId::Cpx508 => Box::new(cpx508_many_references::ManyReferencesRule),
        RuleId::Cpx509 => Box::new(cpx509_long_formula::LongFormulaRule),
        RuleId::Vul601 => Box::new(vul601_duplicate_formulas::DuplicateFormulasRule::new()),
        RuleId::Vul602 => Box::new(vul602_volatile_functions::VolatileFunctionsRule::new()),
        RuleId::Vul603 => Box::new(vul603_empty_string_test::EmptyStringTestRule::new()),
        RuleId::Vul604 => Box::new(vul604_error_prone_functions::ErrorProneFunctionsRule::new(
            config,
        )),
        RuleId::Vul605 => Box::new(vul605_legacy_array::LegacyArrayRule::new()),
        RuleId::Vul606 => Box::new(vul606_deprecated_func::DeprecatedFuncRule::new()),
        RuleId::Data701 => Box::new(dat701_sheet_names::NonDescriptiveSheetNameRule::new(config)),
        RuleId::Data702 => Box::new(dat702_numeric_formats::InconsistentNumberFormatRule::new()),
        RuleId::Data703 => Box::new(dat703_date_formats::InconsistentDateFormatRule::new(config)),
        RuleId::Data704 => Box::new(dat704_long_text::LongTextCellRule::new(config)),
        RuleId::Data705 => Box::new(dat705_unnecessary_space::UnnecessarySpaceRule),
        RuleId::Data706 => Box::new(dat706_numeric_text_calc::NumericTextCalcRule),
        RuleId::Ext801 => Box::new(ext801_external_workbook::ExternalWorkbooksRule::new()),
        RuleId::Ext802 => Box::new(ext802_web_urls::WebUrlsRule::new(config)),
        RuleId::Hid901 => Box::new(hid901_hidden_worksheet::HiddenWorksheetRule),
        RuleId::Hid902 => Box::new(hid902_hidden_columns_rows::HiddenColumnsRowsRule),
        RuleId::File1001 => Box::new(file1001_large_file_size::LargeFileSizeRule::new(config)),
        RuleId::File1002 => Box::new(file1002_old_spreadsheet::OldSpreadsheetRule::new(config)),
        RuleId::File1003 => Box::new(file1003_date_system_1904::DateSystem1904Rule::new()),
        RuleId::Vba1101 => Box::new(vba1101_has_macros::HasMacrosRule),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_prefix_activation() {
        let mut config = LinterConfig::default();
        config.global.enabled_rules.insert("ERR".to_string());
        let enabled = create_enabled_walker_rules(&config);
        assert!(enabled.iter().any(|r| r.id() == RuleId::Err102));
    }

    #[test]
    fn test_create_walker_rules_merges_interruption_rules() {
        let config = LinterConfig::default();
        let ids = [
            RuleId::Calc202,
            RuleId::Int402,
            RuleId::Err102,
            RuleId::Int401,
        ];
        let rules = create_walker_rules(&ids, &config);
        let created: Vec<RuleId> = rules.iter().map(|r| r.id()).collect();
        assert_eq!(created, [RuleId::Calc202, RuleId::Int402, RuleId::Err102]);
    }

    #[test]
    fn test_interruption_rules_listed_individually() {
        let config = LinterConfig::default();
        let metadata_ids: Vec<RuleId> = get_all_rule_metadata(&config)
            .iter()
            .map(|m| m.id)
            .collect();
        let tokens = get_all_valid_tokens();
        for id in [RuleId::Int401, RuleId::Int402, RuleId::Int403] {
            assert!(metadata_ids.contains(&id));
            assert!(tokens.contains(id.as_str()));
        }
    }

    #[test]
    fn test_default_activation() {
        let config = LinterConfig::default();
        let walker_enabled = create_enabled_walker_rules(&config);
        assert!(walker_enabled.iter().any(|r| r.id() == RuleId::Err101));
        assert!(walker_enabled.iter().any(|r| r.id() == RuleId::Err102));
        assert!(walker_enabled.iter().any(|r| r.id() == RuleId::Err103));
    }
}
