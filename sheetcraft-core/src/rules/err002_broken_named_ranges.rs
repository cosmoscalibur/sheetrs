//! ERR002: Broken named ranges detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct BrokenNamedRangesRule;

impl LinterRule for BrokenNamedRangesRule {
    fn id(&self) -> &str {
        "ERR002"
    }

    fn name(&self) -> &str {
        "Broken named ranges"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::UnresolvedErrors
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // Check each defined name to see if it references a valid range
        for (name, reference) in &workbook.defined_names {
            if is_broken_reference(workbook, reference) {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!("Named range '{}' has broken reference: {}", name, reference),
                    Severity::Error,
                ));
            }
        }

        Ok(violations)
    }
}

/// Check if a reference is broken (points to non-existent sheet or invalid range)
fn is_broken_reference(workbook: &Workbook, reference: &str) -> bool {
    // Simple check: if reference contains a sheet name, verify it exists
    if let Some(sheet_name) = extract_sheet_name(reference) {
        return workbook.get_sheet(sheet_name).is_none();
    }
    
    // If we can't parse it, assume it might be broken
    // This is a simplified implementation - a full implementation would
    // need more sophisticated reference parsing
    false
}

/// Extract sheet name from a reference like "Sheet1!A1:B2"
fn extract_sheet_name(reference: &str) -> Option<&str> {
    reference.split('!').next()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_broken_named_ranges() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells: HashMap::new(),
            used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(), formula_parsing_error: None,
        };

        let mut defined_names = HashMap::new();
        defined_names.insert("ValidRange".to_string(), "Sheet1!A1:B2".to_string());
        defined_names.insert("BrokenRange".to_string(), "NonExistentSheet!A1:B2".to_string());

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names,
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = BrokenNamedRangesRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "ERR002");
        assert!(violations[0].message.contains("BrokenRange"));
    }
}
