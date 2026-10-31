//! REF306: Unused sheets
//!
//! A sheet with no connection to the rest of the workbook: nothing reads it and it reads nothing.
//! It feeds no result and takes none from another sheet, so it may be a leftover, scratch work or
//! stale data that adds size and review effort without contributing.
//!
//! # Reports
//!
//! One violation per sheet (the unit is the sheet) that has no link with another sheet, in a
//! workbook of more than one sheet. Seen as a graph whose nodes are the sheets and whose edges
//! are "a formula of sheet A names a cell, a range, a whole column or a whole row of sheet B", a
//! sheet is unused when it is an endpoint of no edge: no other sheet refers to it **and** it
//! refers to none. The edges are collected from the parsed formulas by CALC202
//! ([`super::calc202_circular_references`], which provides this rule's data) in
//! [`LinterContext::sheet_links`]. A defined name used by a formula links that formula's sheet
//! with the sheets the name's definition refers to ([`LinterContext::name_uses`]).
//!
//! # Does not report
//!
//! - The only sheet of a workbook.
//! - A sheet that only reads from others (a summary) or is only read by others (a data sheet): each
//!   has an edge.
//! - A sheet whose only formulas refer to its own cells: a reference of a sheet to itself is not an
//!   edge, so such a sheet is reported when nothing else links it.
//! - Hidden or empty sheets get no special treatment (see HID901 and REF303 for those properties).
//!
//! A sheet used only in ways that are not formulas (charts, conditional formatting, data
//! validation, macros, external workbooks) has no edge and is reported.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always warning.

use super::helpers::formula_links;
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use std::collections::{HashMap, HashSet};

/// Rule that detects unused (standalone) sheets
pub struct UnusedSheetsRule;

/// Incident data for REF306.
#[derive(Debug)]
pub struct UnusedSheetData {
    /// 0-based sheet index of the unused sheet.
    pub sheet_index: u16,
}

impl ViolationData for UnusedSheetData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        format!(
            "Sheet '{name}' is not connected with other sheets: no other sheet refers to it and it refers to none"
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for UnusedSheetsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref306
    }

    fn name(&self) -> &'static str {
        "Unused Sheet"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_workbook_end(
        &mut self,
        workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        // The sheets that have a link with another one, in either direction.
        let mut linked: HashSet<u16> = HashSet::new();
        for &(from, to) in &ctx.sheet_links {
            linked.insert(from);
            linked.insert(to);
        }
        // A defined name used by a sheet links it to the other sheets its definition names.
        if !ctx.name_uses.is_empty() {
            let mut targets: HashMap<String, Vec<u16>> = HashMap::new();
            for (name, definition) in &workbook.defined_names {
                if let Some(links) = formula_links(definition, u16::MAX, &ctx.name_to_index) {
                    targets.insert(
                        name.to_uppercase(),
                        links
                            .sheets
                            .into_iter()
                            .filter(|s| *s != u16::MAX)
                            .collect(),
                    );
                }
            }
            for (from, name) in &ctx.name_uses {
                for &to in targets.get(name).into_iter().flatten() {
                    if to != *from {
                        linked.insert(*from);
                        linked.insert(to);
                    }
                }
            }
        }
        for sheet in &workbook.sheets {
            let is_only_sheet = workbook.sheets.len() == 1;
            if !is_only_sheet && !linked.contains(&sheet.sheet_index) {
                out.push(Violation::with_data(
                    RuleId::Ref306,
                    ViolationScope::Sheet(sheet.sheet_index),
                    UnusedSheetData {
                        sheet_index: sheet.sheet_index,
                    },
                    Severity::Warning,
                ));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_unreferenced_sheet_flagged() {
        let sheet1 = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: Default::default(),
            ..Default::default()
        };

        let mut cells2 = HashMap::new();
        cells2.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Number(42.0),
            },
        );

        let sheet2 = Sheet {
            name: "Sheet2".to_string(),
            sheet_index: 1,
            cells: cells2.into(),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet1, sheet2],
            ..Default::default()
        };

        let mut rule = UnusedSheetsRule;
        let mut ctx = LinterContext::default();

        // Simulate: a formula of Sheet1 (index 0) refers to a third sheet; Sheet2 (index 1) is
        // linked with nothing.
        ctx.sheet_links.insert((0, 2));

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

        // Sheet2 (index 1) should be flagged as unreferenced
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref306);
    }

    #[test]
    fn test_single_sheet_not_flagged() {
        let sheet = Sheet {
            name: "OnlySheet".to_string(),
            sheet_index: 0,
            cells: Default::default(),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let mut rule = UnusedSheetsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

        assert!(violations.is_empty());
    }

    fn workbook_of(names: &[&str]) -> Workbook {
        Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: names
                .iter()
                .enumerate()
                .map(|(i, n)| Sheet {
                    name: (*n).to_string(),
                    sheet_index: i as u16,
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        }
    }

    fn reported(workbook: &Workbook, ctx: &mut LinterContext) -> Vec<u16> {
        let mut rule = UnusedSheetsRule;
        collect_violations(|out| rule.on_workbook_end(workbook, ctx, out))
            .iter()
            .map(|v| match v.scope {
                ViolationScope::Sheet(i) => i,
                _ => panic!("not a sheet"),
            })
            .collect()
    }

    #[test]
    fn a_sheet_is_unused_only_if_it_has_no_link_in_either_direction() {
        let workbook = workbook_of(&["Data", "Summary", "Scratch"]);
        let mut ctx = LinterContext::default();
        // Summary reads Data; nothing refers to Summary: it is still linked (outgoing).
        ctx.sheet_links.insert((1, 0));
        assert_eq!(reported(&workbook, &mut ctx), vec![2]);
    }

    #[test]
    fn a_reference_of_a_sheet_to_itself_is_no_link() {
        // CALC202 does not record a link for it, so the sheet has none.
        let workbook = workbook_of(&["A", "B"]);
        let mut ctx = LinterContext::default();
        assert_eq!(reported(&workbook, &mut ctx), vec![0, 1]);
    }

    #[test]
    fn a_defined_name_links_the_sheet_that_uses_it_with_the_sheet_it_names() {
        let mut workbook = workbook_of(&["Calc", "Lists", "Scratch"]);
        workbook
            .defined_names
            .insert("Items".to_string(), "Lists!$A$1:$A$10".to_string());
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Calc".into(), 0);
        ctx.name_to_index.insert("Lists".into(), 1);
        ctx.name_to_index.insert("Scratch".into(), 2);
        ctx.name_uses.insert((0, "ITEMS".to_string()));
        assert_eq!(reported(&workbook, &mut ctx), vec![2]);
    }
}
