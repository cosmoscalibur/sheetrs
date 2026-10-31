//! ERR101: Broken Named Ranges
//!
//! A defined name whose reference has lost its target (the sheet or cells were deleted) now points
//! at `#REF!`. Every formula that uses the name evaluates to an error, and the name is a trap for
//! whoever picks it up next.
//!
//! # Reports
//!
//! One violation per defined name whose reference string contains `#REF!` (unit: the book, one
//! violation per name, ordered by name). This includes a name that still has a sheet part
//! (`Sheet1!#REF!`) and a `#REF!` inside a longer expression. The message gives the name and the
//! reference as stored.
//!
//! # Does not report
//!
//! - A name that points to an existing sheet and range, even when the range is empty or unused;
//!   unused names are REF301.
//! - A name whose sheet does not exist but whose reference is not `#REF!`: the check is on the
//!   reference text only, and sheet existence is not validated.
//! - Formulas that contain `#REF!` and cells that show it: those are ERR102 and ERR103.
//! - Other broken forms of a name (`#NAME?`, an unparsable definition); only `#REF!` is looked for.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Error`.

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
        out: &mut Vec<Violation>,
    ) {
        let mut defined_names: Vec<_> = workbook.defined_names.iter().collect();
        defined_names.sort_unstable();

        for (name, reference) in defined_names {
            if is_broken_reference(reference) {
                out.push(Violation::with_data(
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
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_broken_named_ranges() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: Default::default(),
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
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Err101);
        assert!(violations[0].message().contains("BrokenRange"));
    }

    #[test]
    fn test_violations_ordered_by_name() {
        let defined_names = (0..16)
            .map(|i| (format!("Broken{i:02}"), "Sheet1!#REF!".to_string()))
            .collect();
        let workbook = Workbook {
            defined_names,
            ..Default::default()
        };

        let violations = collect_violations(|out| {
            BrokenNamedRangesRule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        let names: Vec<_> = violations
            .iter()
            .map(|v| v.data::<BrokenNamedRangeData>().unwrap().name.clone())
            .collect();
        let mut sorted = names.clone();
        sorted.sort();
        assert_eq!(names.len(), 16);
        assert_eq!(names, sorted);
    }
}
