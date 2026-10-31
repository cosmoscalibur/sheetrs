//! REF302: Duplicate or confusingly similar sheet names
//!
//! Sheets named `Data`, `data` and `Data-` are easy to confuse when typing a reference or choosing
//! a tab, and a formula can silently point at the wrong one. Excel itself rejects two sheet names that
//! differ only in case, so the cases that occur in practice are names that look alike.
//!
//! # Reports
//!
//! One violation per group (unit: the group, reported at book level) of two or more sheets whose names
//! are equal after normalization: lowercased with every character that is not a letter or digit
//! removed. `Sheet1`, `sheet 1` and `Sheet-1` form one group, `Data_2024` and `data2024` another.
//! The message lists every sheet name of the group. Groups come in order of their first sheet.
//!
//! # Does not report
//!
//! - Sheets whose normalized names differ, however close (`Sales` and `Sales2`).
//! - Sheets of one workbook that all have generic names but different ones: that is DATA701.
//! - A single sheet: there is nothing to compare, and a sheet with no letters or digits in its name
//!   matches only other such sheets.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use std::collections::HashMap;

/// Rule that checks for duplicate sheet names (case-insensitive)
pub struct DuplicateSheetNamesRule;

/// Incident data for REF302.
#[derive(Debug)]
pub struct DuplicateSheetNameData {
    /// 0-based sheet indices that share a normalized name.
    pub sheet_indices: Vec<u16>,
}

impl ViolationData for DuplicateSheetNameData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let names: Vec<&str> = self
            .sheet_indices
            .iter()
            .map(|idx| ctx.workbook.sheet_name_by_index(*idx).unwrap_or("Unknown"))
            .collect();
        format!("Confusingly similar sheet names: {}", names.join(", "))
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Normalize sheet name for comparison by converting to lowercase and removing non-alphanumeric characters
fn normalize_sheet_name(name: &str) -> String {
    name.to_lowercase()
        .chars()
        .filter(|c| c.is_alphanumeric())
        .collect()
}

impl WalkerRule for DuplicateSheetNamesRule {
    fn id(&self) -> RuleId {
        RuleId::Ref302
    }

    fn name(&self) -> &'static str {
        "Duplicate Sheet Name"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let mut name_map: HashMap<String, Vec<u16>> = HashMap::new();

        // Group sheet indices by their normalized name
        for sheet in &workbook.sheets {
            let normalized = normalize_sheet_name(&sheet.name);
            name_map
                .entry(normalized)
                .or_default()
                .push(sheet.sheet_index);
        }

        let mut duplicates: Vec<_> = name_map
            .into_values()
            .filter(|indices| indices.len() > 1)
            .collect();
        duplicates.sort_unstable();

        out.extend(duplicates.into_iter().map(|indices| {
            Violation::with_data(
                RuleId::Ref302,
                ViolationScope::Book,
                DuplicateSheetNameData {
                    sheet_indices: indices,
                },
                Severity::Warning,
            )
        }));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::path::PathBuf;

    #[test]
    fn test_duplicate_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Data".to_string(),
                sheet_index: 0,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data".to_string(),
                sheet_index: 1,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Summary".to_string(),
                sheet_index: 2,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
        ];

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            ..Default::default()
        };

        let mut rule = DuplicateSheetNamesRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref302);
    }

    #[test]
    fn test_confusingly_similar_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Sheet1".to_string(),
                sheet_index: 0,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "sheet 1".to_string(),
                sheet_index: 1,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Sheet-1".to_string(),
                sheet_index: 2,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Data_2024".to_string(),
                sheet_index: 3,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data2024".to_string(),
                sheet_index: 4,
                cells: Default::default(),
                used_range: None,
                ..Default::default()
            },
        ];

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            ..Default::default()
        };

        let mut rule = DuplicateSheetNamesRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

        // Should detect 2 groups of duplicates:
        // 1. "Sheet1", "sheet 1", "Sheet-1" (all normalize to "sheet1")
        // 2. "Data_2024", "data2024" (both normalize to "data2024")
        assert_eq!(violations.len(), 2);
    }

    #[test]
    fn test_violations_ordered_by_sheet() {
        let sheets = (0..32)
            .map(|i| Sheet {
                name: format!("{}name{}", if i % 2 == 0 { "" } else { " " }, i / 2),
                sheet_index: i,
                ..Default::default()
            })
            .collect();
        let workbook = Workbook {
            sheets,
            ..Default::default()
        };

        let violations = collect_violations(|out| {
            DuplicateSheetNamesRule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        let groups: Vec<_> = violations
            .iter()
            .map(|v| {
                v.data::<DuplicateSheetNameData>()
                    .unwrap()
                    .sheet_indices
                    .clone()
            })
            .collect();
        let mut sorted = groups.clone();
        sorted.sort();
        assert_eq!(groups.len(), 16);
        assert_eq!(groups, sorted);
    }
}
