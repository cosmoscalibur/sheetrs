//! REF311: Defined name defined more than once
//!
//! A name defined for the workbook and again for a sheet, or for two sheets, means a different
//! thing depending on the sheet a formula is in: inside a sheet its own definition wins over the
//! workbook's, and a formula elsewhere has to say which sheet's it means. Nothing in the formula
//! says that two things share the name, so a reader (or a tool that reads names without the sheet
//! of the formula, as this suite does: it keeps one definition per name) can take the wrong one.
//!
//! # Reports
//!
//! One violation per name (unit: the name, reported at book level) that the document defines in
//! more than one scope. The message lists the scopes: the workbook first, then the sheets in the
//! order of the workbook. Violations come in alphabetical order of the name.
//!
//! # Does not report
//!
//! - A name defined once, however local: a sheet's own name is fine.
//! - Names that start with `_xlnm.` (print area, print titles, filter database): reserved names
//!   the application writes once per sheet.
//! - Two names that differ only in case: that is a different matter and Excel refuses them in one
//!   scope.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use sheetrs_core::reader::{NameScope, Workbook};

/// Rule that detects defined names defined in more than one scope
#[derive(Debug)]
pub struct DuplicateDefinedNamesRule;

/// Incident data for REF311.
#[derive(Debug)]
pub struct DuplicateDefinedNameData {
    /// The name.
    pub name: String,
    /// The scopes that define it: two or more.
    pub scopes: Vec<NameScope>,
}

impl ViolationData for DuplicateDefinedNameData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let scopes: Vec<String> = self
            .scopes
            .iter()
            .map(|scope| match scope {
                NameScope::Workbook => "the workbook".to_string(),
                NameScope::Sheet(sheet) => format!("sheet '{sheet}'"),
            })
            .collect();
        format!(
            "Defined name '{}' is defined more than once: in {}",
            self.name,
            scopes.join(", ")
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for DuplicateDefinedNamesRule {
    fn id(&self) -> RuleId {
        RuleId::Ref311
    }

    fn name(&self) -> &'static str {
        "Duplicate Defined Name"
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
        out.extend(workbook.name_conflicts.iter().map(|conflict| {
            Violation::with_data(
                RuleId::Ref311,
                ViolationScope::Book,
                DuplicateDefinedNameData {
                    name: conflict.name.clone(),
                    scopes: conflict.scopes.clone(),
                },
                Severity::Warning,
            )
        }));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::collect_violations;
    use sheetrs_core::reader::NameConflict;

    #[test]
    fn a_name_defined_in_two_scopes_is_reported_with_its_scopes() {
        let workbook = Workbook {
            name_conflicts: vec![
                NameConflict {
                    name: "Rate".to_string(),
                    scopes: vec![NameScope::Workbook, NameScope::Sheet("Costs".to_string())],
                },
                NameConflict {
                    name: "Total".to_string(),
                    scopes: vec![
                        NameScope::Sheet("A".to_string()),
                        NameScope::Sheet("B".to_string()),
                    ],
                },
            ],
            ..Default::default()
        };
        let mut rule = DuplicateDefinedNamesRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

        assert_eq!(violations.len(), 2);
        assert!(violations.iter().all(|v| v.rule_id == RuleId::Ref311));
        assert!(violations.iter().all(|v| v.scope == ViolationScope::Book));
        assert_eq!(
            violations[0].message(),
            "Defined name 'Rate' is defined more than once: in the workbook, sheet 'Costs'"
        );
        assert_eq!(
            violations[1].message(),
            "Defined name 'Total' is defined more than once: in sheet 'A', sheet 'B'"
        );
    }

    #[test]
    fn a_workbook_without_conflicts_has_no_violations() {
        let workbook = Workbook::default();
        let mut rule = DuplicateDefinedNamesRule;
        let mut ctx = LinterContext::default();
        assert!(
            collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out)).is_empty()
        );
    }
}
