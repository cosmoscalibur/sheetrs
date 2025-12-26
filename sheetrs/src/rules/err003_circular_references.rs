//! FORM005: Circular reference detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use regex::Regex;
use std::collections::{HashMap, HashSet};

pub struct CircularReferenceRule {
    cell_ref_pattern: Regex,
    config: crate::config::LinterConfig,
}

impl CircularReferenceRule {
    pub fn new(config: &crate::config::LinterConfig) -> Self {
        // Regex to match cell references (e.g., A1, $A$1, Sheet1!A1, 'Sheet Name'!A1, A1:B2)
        // Group 1: Sheet name (optional) - either quoted or unquoted
        // Group 4: Start column
        // Group 5: Start row
        // Group 6: End column (optional)
        // Group 7: End row (optional)
        let cell_ref_pattern = Regex::new(
            r"(?:('([^']+)'|([A-Za-z0-9_\.]+))!)?\$?([A-Za-z]+)\$?([0-9]+)(?::\$?([A-Za-z]+)\$?([0-9]+))?",
        )
        .unwrap();

        Self {
            cell_ref_pattern,
            config: config.clone(),
        }
    }
}

impl Default for CircularReferenceRule {
    fn default() -> Self {
        Self::new(&crate::config::LinterConfig::default())
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
        RuleCategory::UnresolvedErrors
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        // Global dependency graph: (Sheet, Row, Col) -> Vec<(Sheet, Row, Col)>
        let mut dependencies: HashMap<(String, u32, u32), Vec<(String, u32, u32)>> = HashMap::new();

        // 1. Build the global dependency graph
        for sheet in &workbook.sheets {
            let expand_ranges = self
                .config
                .get_param_bool("expand_ranges_in_dependencies", Some(&sheet.name))
                .unwrap_or(false);

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let refs = extract_cell_references(
                        formula,
                        &self.cell_ref_pattern,
                        &sheet.name,
                        expand_ranges,
                    );
                    dependencies.insert((sheet.name.clone(), cell.row, cell.col), refs);
                }
            }
        }

        // 2. Detect circular references using DFS on the global graph
        let cycles = find_cycles(&dependencies);
        // eprintln!("Total cycles found: {}", cycles.len());

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

/// Extract cell references from a formula, resolving relative sheet names and expanding ranges
fn extract_cell_references(
    formula: &str,
    pattern: &Regex,
    current_sheet: &str,
    expand_ranges: bool,
) -> Vec<(String, u32, u32)> {
    let mut references = Vec::new();

    for cap in pattern.captures_iter(formula) {
        // Determine sheet name
        // Group 1: Sheet name wrapper
        // Group 2: Quoted content
        // Group 3: Unquoted content
        let sheet_name = if cap.get(1).is_some() {
            if let Some(quoted) = cap.get(2) {
                quoted.as_str().to_string()
            } else if let Some(unquoted) = cap.get(3) {
                unquoted.as_str().to_string()
            } else {
                // Fallback (shouldn't happen with valid regex)
                current_sheet.to_string()
            }
        } else {
            current_sheet.to_string()
        };

        // Group 4: Start Col (Alpha)
        // Group 5: Start Row (Numeric)
        if let (Some(col_match), Some(row_match)) = (cap.get(4), cap.get(5)) {
            let col_str = col_match.as_str();
            let row_str = row_match.as_str();

            let (start_row, start_col) = match parse_components(row_str, col_str) {
                Some(coords) => coords,
                None => continue,
            };

            // Check for range end
            // Group 6: End Col
            // Group 7: End Row
            if let (Some(end_col_match), Some(end_row_match)) = (cap.get(6), cap.get(7)) {
                let end_col_str = end_col_match.as_str();
                let end_row_str = end_row_match.as_str();

                if let Some((end_row, end_col)) = parse_components(end_row_str, end_col_str) {
                    // Expand range
                    let min_r = start_row.min(end_row);
                    let max_r = start_row.max(end_row);
                    let min_c = start_col.min(end_col);
                    let max_c = start_col.max(end_col);

                    if expand_ranges && (max_r - min_r + 1) * (max_c - min_c + 1) <= 100_000 {
                        for r in min_r..=max_r {
                            for c in min_c..=max_c {
                                references.push((sheet_name.clone(), r, c));
                            }
                        }
                    } else {
                        references.push((sheet_name.clone(), start_row, start_col));
                        references.push((sheet_name.clone(), end_row, end_col));
                    }
                }
            } else {
                // Single cell reference
                references.push((sheet_name, start_row, start_col));
            }
        }
    }

    references
}

fn parse_components(row_str: &str, col_str: &str) -> Option<(u32, u32)> {
    let row = row_str.parse::<u32>().ok()?;
    let mut col = 0u32;
    for ch in col_str.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }
    }
    if col == 0 {
        return None;
    }

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

    // Initialize all nodes as unvisited
    for cell in dependencies.keys() {
        state.insert(cell.clone(), VisitState::Unvisited);
    }
    for deps in dependencies.values() {
        for dep in deps {
            if !state.contains_key(dep) {
                state.insert(dep.clone(), VisitState::Unvisited);
            }
        }
    }

    // Sort keys for deterministic output
    let mut keys: Vec<Node> = state.keys().cloned().collect();
    keys.sort();

    for start_node in keys {
        if state.get(&start_node) == Some(&VisitState::Unvisited) {
            // Iterative DFS
            let mut stack = vec![(start_node.clone(), 0)]; // (node, next_dep_idx)
            let mut path = Vec::new();
            let mut in_path = HashSet::new();

            while let Some((u, dep_idx)) = stack.pop() {
                if dep_idx == 0 {
                    // First time visiting this node in this path
                    if in_path.contains(&u) {
                        // Cycle detected!
                        if let Some(pos) = path.iter().position(|x| x == &u) {
                            let cycle = path[pos..].to_vec();
                            cycles.push(cycle);
                        }
                        continue;
                    }
                    if state.get(&u) == Some(&VisitState::Visited) {
                        continue;
                    }

                    state.insert(u.clone(), VisitState::Visiting);
                    in_path.insert(u.clone());
                    path.push(u.clone());
                }

                if let Some(deps) = dependencies.get(&u) {
                    if dep_idx < deps.len() {
                        // Push current node back with next index
                        stack.push((u, dep_idx + 1));
                        // Push neighbor to visit
                        stack.push((deps[dep_idx].clone(), 0));
                    } else {
                        // Finished visiting all neighbors
                        state.insert(u.clone(), VisitState::Visited);
                        in_path.remove(&u);
                        path.pop();
                    }
                } else {
                    // No dependencies
                    state.insert(u.clone(), VisitState::Visited);
                    in_path.remove(&u);
                    path.pop();
                }
            }
        }
    }

    cycles
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    fn create_test_workbook(sheet_name: &str, cells: HashMap<(u32, u32), Cell>) -> Workbook {
        let sheet = Sheet {
            name: sheet_name.to_string(),
            cells,
            used_range: Some((10, 10)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
        };

        Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
            external_workbooks: Vec::new(),
        }
    }

    #[test]
    fn test_circular_reference_direct() {
        let mut cells = HashMap::new();
        // A1 = A1+1
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=A1+1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let rule = CircularReferenceRule::new(&crate::config::LinterConfig::default());
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "ERR003");
    }

    #[test]
    fn test_circular_reference_indirect() {
        let mut cells = HashMap::new();
        // A1 = B1, B1 = A1
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=B1".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::formula("=A1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let rule = CircularReferenceRule::new(&crate::config::LinterConfig::default());
        let violations = rule.check(&workbook).unwrap();

        assert!(violations.len() >= 1);
    }

    #[test]
    fn test_circular_reference_range() {
        let mut cells = HashMap::new();
        // A1 = SUM(B1:B3)
        // B2 = A1
        // Cycle: A1 -> B2 -> A1
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(B1:B3)".to_string()),
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                num_fmt: None,
                row: 1,
                col: 1,
                value: CellValue::formula("=A1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        // By default expand_ranges is FALSE.
        // Global expand is set to true for this test.
        let mut config = crate::config::LinterConfig::default();
        config.global.params.insert(
            "expand_ranges_in_dependencies".to_string(),
            toml::Value::Boolean(true),
        );

        let rule = CircularReferenceRule::new(&config);
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("Sheet1!A1"));
        assert!(violations[0].message.contains("Sheet1!B2"));
    }

    #[test]
    fn test_circular_reference_range_no_expand() {
        let mut cells = HashMap::new();
        // A1 = SUM(B1:B3)
        // B2 = A1
        // Cycle: A1 -> B2 -> A1 (ONLY if A1 is expanded to depend on B2)
        // If no expand: A1 -> B1, B3. B2 -> A1. No cycle.
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(B1:B3)".to_string()),
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                num_fmt: None,
                row: 1,
                col: 1,
                value: CellValue::formula("=A1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        // Default config has expand=false
        let rule = CircularReferenceRule::new(&crate::config::LinterConfig::default());
        let violations = rule.check(&workbook).unwrap();

        // Should NOT detect cycle
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_circular_reference_range_sheet_override() {
        let mut cells = HashMap::new();
        // A1 = SUM(B1:B3)
        // B2 = A1
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(B1:B3)".to_string()),
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                num_fmt: None,
                row: 1,
                col: 1,
                value: CellValue::formula("=A1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);

        // Global false, but Sheet1 true
        let mut config = crate::config::LinterConfig::default();
        config.global.params.insert(
            "expand_ranges_in_dependencies".to_string(),
            toml::Value::Boolean(false),
        );
        let mut sheet_conf = crate::config::SheetConfig::default();
        sheet_conf.params.insert(
            "expand_ranges_in_dependencies".to_string(),
            toml::Value::Boolean(true),
        );
        config.sheets.insert("Sheet1".to_string(), sheet_conf);

        let rule = CircularReferenceRule::new(&config);
        let violations = rule.check(&workbook).unwrap();

        // Should detect cycle because Sheet1 expanded
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_huge_range_limit() {
        let mut cells = HashMap::new();
        // A1 = SUM(A2:A10000) - this range is okay
        // But if A1:XFD1048576 was present it would be skipped
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(A2:A3)".to_string()),
            },
        );
        cells.insert(
            (2, 0),
            Cell {
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::formula("=A1".to_string()),
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let rule = CircularReferenceRule::new(&crate::config::LinterConfig::default());
        let violations = rule.check(&workbook).unwrap();

        // Cycle: A1 -> A3 -> A1
        assert_eq!(violations.len(), 1);
    }
}
