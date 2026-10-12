//! VBA1101: Macros and scripts detection

use super::{LinterContext, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects the presence of macros (VBA/Scripts)
///
/// Macros can pose security risks or indicate legacy automation that may need review.
pub struct HasMacrosRule;

/// Incident data for VBA1101 — no variable fields.
#[derive(Debug)]
pub struct HasMacrosData;

impl ViolationData for HasMacrosData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        "Workbook contains macros or scripts. Review for security concerns.".to_string()
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for HasMacrosRule {
    fn id(&self) -> RuleId {
        RuleId::Vba1101
    }

    fn name(&self) -> &'static str {
        "Has Macros"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::VBA
    }

    fn on_workbook_start(&self, workbook: &Workbook, _ctx: &mut LinterContext) -> Vec<Violation> {
        let mut violations = Vec::new();

        if workbook.has_macros {
            violations.push(Violation::with_data(
                RuleId::Vba1101,
                ViolationScope::Book,
                HasMacrosData,
                Severity::Warning,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::path::PathBuf;

    #[test]
    fn test_workbook_with_macros() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsm"),
            has_macros: true,
            ..Default::default()
        };

        let rule = HasMacrosRule;
        let mut ctx = LinterContext::default();
        let violations = rule.on_workbook_start(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vba1101);
        assert!(violations[0].message().contains("macros"));
    }

    #[test]
    fn test_workbook_without_macros() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            ..Default::default()
        };

        let rule = HasMacrosRule;
        let mut ctx = LinterContext::default();
        let violations = rule.on_workbook_start(&workbook, &mut ctx);

        assert_eq!(violations.len(), 0);
    }
}
