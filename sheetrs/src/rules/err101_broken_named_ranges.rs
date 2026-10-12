//! ERR101: Broken named ranges detection
//!
//! Description: Detects named ranges pointing to invalid or deleted cell regions.

use super::{LinterContext, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that identifies broken named ranges (references to invalid/deleted locations)
///
/// Checks all defined names (named ranges) for reference strings containing "#REF!".
pub struct BrokenNamedRangesRule;

/// Incident data for ERR101.
#[derive(Debug)]
pub struct BrokenNamedRangeData {
    /// The named range name.
    pub name: String,
    /// The broken reference string.
    pub reference: String,
}

impl ViolationData for BrokenNamedRangeData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Named range '{}' has broken reference: {}",
            self.name, self.reference
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for BrokenNamedRangesRule {
    fn id(&self) -> RuleId {
        RuleId::Err101
    }

    fn name(&self) -> &'static str {
        "Broken Named Ranges"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::ExcelErrors
    }

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
    ) -> Vec<Violation> {
        let mut violations = Vec::new();

        for (name, reference) in &workbook.defined_names {
            if is_broken_reference(reference) {
                violations.push(Violation::with_data(
                    RuleId::Err101,
                    ViolationScope::Book,
                    BrokenNamedRangeData {
                        name: name.clone(),
                        reference: reference.clone(),
                    },
                    Severity::Error,
                ));
            }
        }

        violations
    }
}

/// Check if a reference is broken (contains #REF! error)
fn is_broken_reference(reference: &str) -> bool {
    // Validates REF error in range definition, ignoring sheet existence.
    // Example: "INGRESOS!#REF!"
    reference.contains("#REF!")
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
            sheet_index: 0,
            cells: HashMap::new(),
            used_range: None,
            ..Default::default()
        };

        let mut defined_names = HashMap::new();
        defined_names.insert("ValidRange".to_string(), "Sheet1!A1:B2".to_string());
        // This SHOULD be reported
        defined_names.insert("BrokenRange".to_string(), "Sheet1!#REF!".to_string());

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names,
            ..Default::default()
        };

        let mut rule = BrokenNamedRangesRule;
        let mut ctx = LinterContext::default();
        let violations = rule.on_workbook_start(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Err101);
        assert!(violations[0].message().contains("BrokenRange"));
    }
}
