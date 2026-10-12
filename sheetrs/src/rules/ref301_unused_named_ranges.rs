//! REF301: Unused named ranges detection
//!
//! Description: Named ranges that are defined but never used in any formula.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet, Workbook};
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use std::collections::HashSet;
use std::sync::Mutex;

/// Rule that detects unused named ranges
pub struct UnusedNamedRangesRule {
    /// Named ranges found in formulas during the walk.
    used_names: Mutex<HashSet<String>>,
}

impl UnusedNamedRangesRule {
    /// Create a new instance of the rule.
    pub fn new() -> Self {
        Self {
            used_names: Mutex::new(HashSet::new()),
        }
    }
}

impl Default for UnusedNamedRangesRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for REF301.
#[derive(Debug)]
pub struct UnusedNamedRangeData {
    /// The unused named range name.
    pub name: String,
}

impl ViolationData for UnusedNamedRangeData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Named range '{}' is defined but never used", self.name)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for UnusedNamedRangesRule {
    fn id(&self) -> RuleId {
        RuleId::Ref301
    }

    fn name(&self) -> &'static str {
        "Unused Defined Name"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_cell(&self, _sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula) = cell.as_formula() {
            let mut used = self.used_names.lock().unwrap();
            // Record the formula text; matching is deferred to on_workbook_end
            used.insert(formula.to_string());
        }
        Vec::new()
    }

    fn on_workbook_end(&self, workbook: &Workbook, _ctx: &mut LinterContext) -> Vec<Violation> {
        let formulas = self.used_names.lock().unwrap();
        workbook
            .defined_names
            .keys()
            .filter(|name| !name.starts_with("_xlnm."))
            .filter(|name| !formulas.iter().any(|f| f.contains(name.as_str())))
            .map(|name| {
                Violation::with_data(
                    RuleId::Ref301,
                    ViolationScope::Book,
                    UnusedNamedRangeData { name: name.clone() },
                    Severity::Warning,
                )
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_unused_named_ranges() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=UsedRange")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut defined_names = HashMap::new();
        defined_names.insert("UsedRange".to_string(), "Sheet1!A1:B2".to_string());
        defined_names.insert("UnusedRange".to_string(), "Sheet1!C1:D2".to_string());

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names,
            ..Default::default()
        };

        let rule = UnusedNamedRangesRule::new();
        let mut ctx = LinterContext::default();

        // Walk cells
        for cell in workbook.sheets[0].cells.values() {
            rule.on_cell(&workbook.sheets[0], cell, &mut ctx);
        }

        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref301);
        assert!(violations[0].message().contains("UnusedRange"));
    }
}
