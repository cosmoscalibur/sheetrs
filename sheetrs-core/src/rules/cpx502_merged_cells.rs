//! CPX502: Merged cells
//!
//! Merged cells break the regular grid that sorting, filtering, copy and paste, pivot tables and
//! formula fill-down rely on, and they make a sheet's structure depend on layout rather than data.
//!
//! # Reports
//! One violation per merged range recorded for a sheet (scope: sheet; the message gives the range,
//! for example `A1:C1`). Ranges are reported in the order the file lists them. Any merge counts,
//! whatever its size and whatever it contains: a title banner is reported like a merge inside a data
//! table.
//!
//! # Does not report
//! Visual alternatives to merging, such as "centre across selection". A sheet without merged ranges
//! is not reported. The rule has no exceptions for headers or for ranges outside the used range.
//!
//! # Parameters
//! None.
//!
//! # Severity
//! Always `Warning`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Sheet;
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};

/// Rule that detects merged cells
pub struct MergedCellsRule;

/// Incident data for CPX502.
#[derive(Debug)]
pub struct MergedCellsData {
    /// Range as (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for MergedCellsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (sr, sc, er, ec) = self.range;
        let start = CellReference::new(sr, sc);
        let end = CellReference::new(er, ec);
        format!("Merged cells in range: {start}:{end}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for MergedCellsRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx502
    }

    fn name(&self) -> &'static str {
        "Merged Cells"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        for &(start_row, start_col, end_row, end_col) in &sheet.merged_cells {
            out.push(Violation::with_data(
                RuleId::Cpx502,
                ViolationScope::Sheet(sheet.sheet_index),
                MergedCellsData {
                    range: (start_row, start_col, end_row, end_col),
                },
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::Workbook;
    use crate::reader::workbook::Sheet;
    use crate::rules::walker::WorkbookWalker;
    use std::path::PathBuf;

    #[test]
    fn test_merged_cells() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                merged_cells: vec![
                    (0, 0, 0, 2), // A1:C1
                    (2, 0, 4, 0), // A3:A5
                ],
                ..Default::default()
            }],
            ..Default::default()
        };

        let rule = MergedCellsRule;
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(rule)];
        let walker = WorkbookWalker::new(&workbook, rules);
        let violations = walker.walk();

        assert_eq!(violations.len(), 2);
        assert_eq!(violations[0].rule_id, RuleId::Cpx502);
        assert!(violations[0].message().contains("A1:C1"));
        assert!(violations[1].message().contains("A3:A5"));
    }

    #[test]
    fn test_no_merged_cells() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let rule = MergedCellsRule;
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(rule)];
        let walker = WorkbookWalker::new(&workbook, rules);
        let violations = walker.walk();

        assert_eq!(violations.len(), 0);
    }
}
