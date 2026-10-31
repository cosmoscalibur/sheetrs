//! VBA1101: Macros and Scripts
//!
//! Macros and scripts run code when the file is opened or used. They are a security risk, may
//! hide logic that cannot be reviewed as cells and formulas, and can stop the file from working
//! in other applications or online.
//!
//! # Reports
//!
//! One violation for the book when macros or scripts are found. In an XLSX/XLSM file that is a
//! `xl/vbaProject.bin` part or any part under `xl/macrosheets/` (Excel 4 macro sheets). In an ODS
//! file it is a `Basic/` or `Scripts/` library, found as an entry of the archive or as a path the
//! manifest (`META-INF/manifest.xml`) lists even when the archive has no entry for the folder, or a
//! macro bound to an event of the document, a form or a cell (a `vnd.sun.star.script:` target in
//! `content.xml`). The message is the same for every file and
//! says only that the workbook should be reviewed.
//!
//! # Does not report
//!
//! - What the macros do, how many there are or whether they are signed; the code is never read.
//! - The file extension: an `.xlsm` without a VBA project is not reported, and a macro in an
//!   `.xlsx` renamed from another type is.
//! - ODS scripts that are neither in a `Basic/` or `Scripts/` library nor bound to an event: there
//!   is no other place where a document keeps them.
//! - Formulas calling add-in or user-defined functions; those are cell-level matters.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

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

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if workbook.has_macros {
            out.push(Violation::with_data(
                RuleId::Vba1101,
                ViolationScope::Book,
                HasMacrosData,
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::collect_violations;

    use std::path::PathBuf;

    #[test]
    fn test_workbook_with_macros() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsm"),
            has_macros: true,
            ..Default::default()
        };

        let mut rule = HasMacrosRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

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

        let mut rule = HasMacrosRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_start(&workbook, &mut ctx, out));

        assert_eq!(violations.len(), 0);
    }
}
