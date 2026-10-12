//! REF302: Duplicate sheet names
//!
//! Description: Detects sheet names that are identical or confusingly similar (case-insensitive).

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
    ) -> Vec<Violation> {
        let mut violations = Vec::new();
        let mut name_map: HashMap<String, Vec<u16>> = HashMap::new();

        // Group sheet indices by their normalized name
        for sheet in &workbook.sheets {
            let normalized = normalize_sheet_name(&sheet.name);
            name_map
                .entry(normalized)
                .or_default()
                .push(sheet.sheet_index);
        }

        // Find duplicates
        for (_normalized, indices) in name_map {
            if indices.len() > 1 {
                violations.push(Violation::with_data(
                    RuleId::Ref302,
                    ViolationScope::Book,
                    DuplicateSheetNameData {
                        sheet_indices: indices,
                    },
                    Severity::Warning,
                ));
            }
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use crate::rules::LinterContext;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_duplicate_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Data".to_string(),
                sheet_index: 0,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data".to_string(),
                sheet_index: 1,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Summary".to_string(),
                sheet_index: 2,
                cells: HashMap::new(),
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
        let violations = rule.on_workbook_start(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref302);
    }

    #[test]
    fn test_confusingly_similar_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Sheet1".to_string(),
                sheet_index: 0,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "sheet 1".to_string(),
                sheet_index: 1,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Sheet-1".to_string(),
                sheet_index: 2,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Data_2024".to_string(),
                sheet_index: 3,
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data2024".to_string(),
                sheet_index: 4,
                cells: HashMap::new(),
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
        let violations = rule.on_workbook_start(&workbook, &mut ctx);

        // Should detect 2 groups of duplicates:
        // 1. "Sheet1", "sheet 1", "Sheet-1" (all normalize to "sheet1")
        // 2. "Data_2024", "data2024" (both normalize to "data2024")
        assert_eq!(violations.len(), 2);
    }
}
