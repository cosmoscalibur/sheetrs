//! FORM005: Circular reference detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use regex::Regex;
use std::collections::{HashMap, HashSet};

pub struct CircularReferenceRule {
    cell_ref_pattern: Regex,
}

impl CircularReferenceRule {
    pub fn new() -> Self {
        // Regex to match cell references with optional sheet qualifier.
        // Captures:
        // 1. Optional sheet name with quotes: 'Sheet Name'!
        // 2. Optional sheet name without quotes: Sheet1!
        // 3. Cell reference: $A$1
        // Note: This regex is an approximation. Excel formula parsing is complex.
        // It handles: 'Sheet'!A1, Sheet!A1, A1.
        // It does NOT handle external workbooks like [Book.xlsx]Sheet!A1 correctly (might capture Sheet!A1 part).
        let cell_ref_pattern =
            Regex::new(r"(?:('[^']+'|[A-Za-z0-9_\.\-]+)!)?(\$?[A-Z]+\$?[0-9]+)").unwrap();

        Self { cell_ref_pattern }
    }
}

impl Default for CircularReferenceRule {
    fn default() -> Self {
        Self::new()
    }
}

impl LinterRule for CircularReferenceRule {
    fn id(&self) -> &str {
        "ERR003"
    }

    fn name(&self) -> &str {
        "Circular references"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn default_active(&self) -> bool {
        false
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        // Global dependency graph: (Sheet, Row, Col) -> Vec<(Sheet, Row, Col)>
        let mut dependencies: HashMap<(String, u32, u32), Vec<(String, u32, u32)>> = HashMap::new();

        // 1. Build the global dependency graph
        for sheet in &workbook.sheets {
            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let refs =
                        extract_cell_references(formula, &self.cell_ref_pattern, &sheet.name);
                    dependencies.insert((sheet.name.clone(), cell.row, cell.col), refs);
                }
            }
        }

        // 2. Detect circular references using DFS on the global graph
        let cycles = find_cycles(&dependencies);

        let mut reported_cells = HashSet::new();

        for cycle in cycles {
            let is_duplicate = cycle.iter().any(|c| reported_cells.contains(c));

            if !is_duplicate {
                for cell in &cycle {
                    reported_cells.insert(cell.clone());
                }

                // Format the cycle path including sheet names
                let path_str: Vec<String> = cycle
                    .iter()
                    .map(|(s, r, c)| format!("{}!{}", s, CellReference::new(*r, *c)))
                    .collect();

                let full_path = format!("{} -> {}", path_str.join(" -> "), path_str[0]);

                // Report on the first cell
                let (sheet_name, r, c) = &cycle[0];
                let cell_ref = CellReference::new(*r, *c);

                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Cell(sheet_name.clone(), cell_ref),
                    format!("Circular reference detected: {}", full_path),
                    Severity::Error,
                ));
            }
        }

        Ok(violations)
    }
}

/// Extract cell references from a formula, resolving relative sheet names
fn extract_cell_references(
    formula: &str,
    pattern: &Regex,
    current_sheet: &str,
) -> Vec<(String, u32, u32)> {
    let mut references = Vec::new();

    for cap in pattern.captures_iter(formula) {
        // Group 2 is the cell ref (A1)
        if let Some(cell_match) = cap.get(2) {
            if let Some((row, col)) = parse_cell_reference(cell_match.as_str()) {
                // Determine sheet name
                let sheet_name = if let Some(sheet_match) = cap.get(1) {
                    let mut s = sheet_match.as_str();
                    // Remove trailing !
                    if s.ends_with('!') {
                        s = &s[..s.len() - 1];
                    }
                    // Remove surrounding single quotes if present
                    if s.starts_with('\'') && s.ends_with('\'') {
                        s = &s[1..s.len() - 1];
                    }
                    s.to_string()
                } else {
                    current_sheet.to_string()
                };

                references.push((sheet_name, row, col));
            }
        }
    }

    references
}

/// Parse a cell reference like "A1" or "$A$1" into (row, col) as 0-based indices
fn parse_cell_reference(cell_ref: &str) -> Option<(u32, u32)> {
    // Remove $ signs
    let cleaned = cell_ref.replace('$', "");

    let mut col = 0u32;
    let mut row_str = String::new();

    for ch in cleaned.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    if row_str.is_empty() {
        return None;
    }

    let row = row_str.parse::<u32>().ok()?;

    // Convert to 0-based
    Some((row.saturating_sub(1), col.saturating_sub(1)))
}

#[derive(PartialEq, Clone, Copy)]
enum VisitState {
    Unvisited,
    Visiting,
    Visited,
}

// Node type for global graph: (SheetName, Row, Col)
type Node = (String, u32, u32);

/// Find all unique elementary cycles in the graph
fn find_cycles(dependencies: &HashMap<Node, Vec<Node>>) -> Vec<Vec<Node>> {
    let mut cycles = Vec::new();
    let mut state = HashMap::new();
    let mut path = Vec::new();

    // Initialize all nodes as unvisited
    for cell in dependencies.keys() {
        state.insert(cell.clone(), VisitState::Unvisited);
    }
    for deps in dependencies.values() {
        for dep in deps {
            state.entry(dep.clone()).or_insert(VisitState::Unvisited);
        }
    }

    // Sort keys for deterministic output
    let mut keys: Vec<Node> = dependencies.keys().cloned().collect();
    keys.sort();

    for cell in keys {
        if state.get(&cell) == Some(&VisitState::Unvisited) {
            dfs_find_cycle(
                cell.clone(),
                dependencies,
                &mut state,
                &mut path,
                &mut cycles,
            );
        }
    }

    cycles
}

fn dfs_find_cycle(
    u: Node,
    dependencies: &HashMap<Node, Vec<Node>>,
    state: &mut HashMap<Node, VisitState>,
    path: &mut Vec<Node>,
    cycles: &mut Vec<Vec<Node>>,
) {
    state.insert(u.clone(), VisitState::Visiting);
    path.push(u.clone());

    if let Some(deps) = dependencies.get(&u) {
        for v in deps {
            match state.get(v) {
                Some(VisitState::Visiting) => {
                    // Cycle detected!
                    if let Some(pos) = path.iter().position(|x| x == v) {
                        let cycle = path[pos..].to_vec();
                        cycles.push(cycle);
                    }
                }
                Some(VisitState::Unvisited) => {
                    dfs_find_cycle(v.clone(), dependencies, state, path, cycles);
                }
                _ => {}
            }
        }
    }

    state.insert(u, VisitState::Visited);
    path.pop();
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_circular_reference_direct() {
        let mut cells = HashMap::new();
        // A1 = A1 (direct circular reference)
        cells.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=A1+1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = CircularReferenceRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "ERR003");
        assert!(violations[0].message.contains("Circular reference"));
    }

    #[test]
    fn test_circular_reference_indirect() {
        let mut cells = HashMap::new();
        // A1 = B1, B1 = A1 (indirect circular reference)
        cells.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=B1+1".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell { num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Formula("=A1+1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 2)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = CircularReferenceRule::new();
        let violations = rule.check(&workbook).unwrap();

        // Both cells are part of the cycle
        assert!(violations.len() >= 1);
        assert_eq!(violations[0].rule_id, "ERR003");
    }

    #[test]
    fn test_no_circular_reference() {
        let mut cells = HashMap::new();
        // A1 = B1, B1 = C1 (no circular reference)
        cells.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=B1+1".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell { num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Formula("=C1+1".to_string()),
            },
        );
        cells.insert(
            (0, 2),
            Cell { num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Formula("=10".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 3)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = CircularReferenceRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_circular_reference_cross_sheet() {
        // Sheet1!A1 = Sheet2!A1
        // Sheet2!A1 = Sheet1!A1
        let mut cells1 = HashMap::new();
        cells1.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=Sheet2!A1".to_string()),
            },
        );
        let sheet1 = Sheet {
            name: "Sheet1".to_string(),
            cells: cells1,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
        };

        let mut cells2 = HashMap::new();
        cells2.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=Sheet1!A1".to_string()),
            },
        );
        let sheet2 = Sheet {
            name: "Sheet2".to_string(),
            cells: cells2,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet1, sheet2],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = CircularReferenceRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("Sheet1!A1"));
        assert!(violations[0].message.contains("Sheet2!A1"));
    }

    #[test]
    fn test_cell_reference_parsing() {
        assert_eq!(parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(parse_cell_reference("$A$1"), Some((0, 0)));
        assert_eq!(parse_cell_reference("Z10"), Some((9, 25)));
    }
}
