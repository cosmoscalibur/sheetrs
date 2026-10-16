//! REF301: Unused named ranges detection
//!
//! Description: Named ranges that are defined but never used in any formula.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet, Workbook};
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use aho_corasick::AhoCorasick;

/// Rule that detects unused named ranges
pub struct UnusedNamedRangesRule {
    /// Defined names to check, sorted, without reserved `_xlnm.` names.
    names: Vec<String>,
    /// Matcher over `names`: each pattern ID indexes `names` and `used`.
    matcher: Option<AhoCorasick>,
    /// Whether some formula contains the name at the same index.
    used: Vec<bool>,
    /// Number of `false` entries in `used`.
    unused_count: usize,
}

impl UnusedNamedRangesRule {
    /// Create a new instance of the rule.
    pub fn new() -> Self {
        Self {
            names: Vec::new(),
            matcher: None,
            used: Vec::new(),
            unused_count: 0,
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

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        self.names = workbook
            .defined_names
            .keys()
            .filter(|name| !name.starts_with("_xlnm."))
            .cloned()
            .collect();
        self.names.sort_unstable();
        self.used = vec![false; self.names.len()];
        self.unused_count = self.names.len();
        self.matcher = Some(
            AhoCorasick::new(&self.names).expect("defined names fit an Aho-Corasick automaton"),
        );
    }

    fn on_cell(
        &mut self,
        _sheet: &Sheet,
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if self.unused_count > 0
            && let Some(formula) = cell.as_formula()
        {
            let matcher = self
                .matcher
                .as_ref()
                .expect("on_workbook_start builds the matcher");
            // Overlapping matches also report a name that appears only inside
            // a longer one (`RATE` in `TAX_RATE`).
            for found in matcher.find_overlapping_iter(formula) {
                let used = &mut self.used[found.pattern()];
                if !*used {
                    *used = true;
                    self.unused_count -= 1;
                }
            }
        }
    }

    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        out.extend(
            self.names
                .iter()
                .zip(&self.used)
                .filter(|(_, used)| !**used)
                .map(|(name, _)| {
                    Violation::with_data(
                        RuleId::Ref301,
                        ViolationScope::Book,
                        UnusedNamedRangeData { name: name.clone() },
                        Severity::Warning,
                    )
                }),
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_workbook_from_reader;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use crate::rules::walker::WorkbookWalker;
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

        let mut rule = UnusedNamedRangesRule::new();
        let mut ctx = LinterContext::default();

        rule.on_workbook_start(&workbook, &mut ctx, &mut Vec::new());
        for cell in workbook.sheets[0].cells.values() {
            rule.on_cell(&workbook.sheets[0], cell, &mut ctx, &mut Vec::new());
        }

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref301);
        assert!(violations[0].message().contains("UnusedRange"));
    }

    #[test]
    fn test_violations_ordered_by_name() {
        let defined_names = (0..16)
            .map(|i| (format!("Unused{i:02}"), "Sheet1!A1".to_string()))
            .collect();
        let workbook = Workbook {
            defined_names,
            ..Default::default()
        };

        let mut rule = UnusedNamedRangesRule::new();
        let mut ctx = LinterContext::default();
        rule.on_workbook_start(&workbook, &mut ctx, &mut Vec::new());
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        let names: Vec<_> = violations
            .iter()
            .map(|v| v.data::<UnusedNamedRangeData>().unwrap().name.clone())
            .collect();
        let mut sorted = names.clone();
        sorted.sort();
        assert_eq!(names.len(), 16);
        assert_eq!(names, sorted);
    }

    /// Lints `formulas` in one sheet and returns the names REF301 reports.
    fn unused_names(names: &[&str], formulas: &[&str]) -> Vec<String> {
        let cells = formulas
            .iter()
            .enumerate()
            .map(|(row, formula)| {
                let row = u32::try_from(row).unwrap();
                let cell = Cell {
                    formula: Some((*formula).into()),
                    is_array: false,
                    num_fmt: None,
                    row,
                    col: 0,
                    value: CellValue::Empty,
                };
                ((row, 0), cell)
            })
            .collect();
        let workbook = Workbook {
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells,
                ..Default::default()
            }],
            defined_names: names
                .iter()
                .map(|name| ((*name).to_string(), "Sheet1!A1".to_string()))
                .collect(),
            ..Default::default()
        };
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(UnusedNamedRangesRule::new())];
        WorkbookWalker::new(&workbook, rules)
            .walk()
            .iter()
            .map(|v| v.data::<UnusedNamedRangeData>().unwrap().name.clone())
            .collect()
    }

    /// Lints a test asset with REF301 alone and returns the names it reports.
    fn unused_names_in_asset(bytes: &[u8], extension: &str) -> Vec<String> {
        let workbook =
            read_workbook_from_reader(std::io::Cursor::new(bytes), Some(extension)).unwrap();
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(UnusedNamedRangesRule::new())];
        WorkbookWalker::new(&workbook, rules)
            .walk()
            .iter()
            .map(|v| v.data::<UnusedNamedRangeData>().unwrap().name.clone())
            .collect()
    }

    #[test]
    fn test_parity_minimal_assets() {
        const TEST_XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");
        const TEST_ODS: &[u8] = include_bytes!("../../../tests/minimal_test.ods");
        let expected = ["text_not_used_range"];
        assert_eq!(unused_names_in_asset(TEST_XLSX, "xlsx"), expected);
        assert_eq!(unused_names_in_asset(TEST_ODS, "ods"), expected);
    }

    #[test]
    fn test_name_inside_longer_name_is_used() {
        assert!(unused_names(&["RATE", "TAX_RATE"], &["=TAX_RATE*2"]).is_empty());
    }

    #[test]
    fn test_name_inside_string_literal_is_used() {
        assert!(unused_names(&["Rate"], &["=\"Rate\"&A1"]).is_empty());
    }

    #[test]
    fn test_name_matching_is_case_sensitive() {
        assert_eq!(unused_names(&["TASA"], &["=Tasa*2"]), ["TASA"]);
    }

    #[test]
    fn test_reserved_names_are_skipped() {
        assert!(unused_names(&["_xlnm.Print_Area"], &[]).is_empty());
    }

    #[test]
    fn test_all_names_reported_without_formulas() {
        assert_eq!(unused_names(&["B", "A"], &[]), ["A", "B"]);
    }

    #[test]
    fn test_no_names_no_violations() {
        assert!(unused_names(&[], &["=A1+B1"]).is_empty());
    }

    #[test]
    fn test_names_used_across_formulas() {
        assert_eq!(
            unused_names(&["Alpha", "Beta", "Gamma"], &["=Alpha+1", "=SUM(Beta)"]),
            ["Gamma"]
        );
    }
}
