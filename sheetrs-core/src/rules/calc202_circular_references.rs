//! CALC202: Circular Reference
//!
//! A formula that depends on its own result, directly or through other cells, cannot be calculated
//! consistently. The spreadsheet stops with a circular-reference warning, shows 0 or an error, or
//! iterates to a value that depends on the order of calculation, so results downstream are wrong.
//!
//! # Reports
//!
//! One violation per distinct loop in the cell dependency graph (unit: the loop, placed on its
//! first cell). The message lists the cells of the loop across sheets. A cell that is already part
//! of a reported loop is not reported again. References are taken from the parsed formula and
//! resolved across sheets by name; a formula the parser cannot read falls back to a textual scan.
//!
//! The values saved in the file decide how serious a loop is. A loop in which at least one cell
//! holds an error in the cache is reported as active: that is how a spreadsheet stores a circular
//! result (LibreOffice writes an error, Excel writes 0, so Excel files show a warning). Active
//! loops are also withheld from [`err102_error_cells`](super::err102_error_cells) and
//! [`err103_ref_to_error`](super::err103_ref_to_error), so the same cells are not reported twice.
//! A loop whose cache shows no error is a potential loop: it may never be evaluated (an unchosen
//! `IF` branch, the unselected cells of a `SUMIF` sum range), but another input can make it real.
//!
//! # Does not report
//!
//! - Loops that pass through the interior of a range. To bound memory a range is tracked by its
//!   two corner cells only, so `A1 = SUM(B1:B3)` with `B2 = A1` is not found.
//! - Whole column and whole row references (`A:A`, `1:1`); these are flagged by REF307.
//! - References to sheets that do not exist in the workbook and to external workbooks; neither can
//!   close a loop inside this workbook.
//! - Cells with no formula, and defined names (they are not followed).
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! `Error` for an active loop (a cell of it holds an error in the saved values), `Warning` for a
//! potential loop.

use super::helpers::{CELL_REF_PATTERN, parse_cell_coords};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet, Workbook};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::{HashMap, HashSet};

/// Rule that detects circular references in formulas.
///
/// Circular references occur when a formula refers back to its own cell, either directly or indirectly.
/// This can cause calculation errors and infinite loops.
///
/// # Range Handling
///
/// To prevent excessive memory usage with large spreadsheets, range references (e.g., `A1:B10`)
/// are represented by their corner cells (start and end) rather than being fully expanded.
/// This means circular references that pass through the middle of a range may not be detected.
///
/// Note: Whole column/row references (e.g., `A:A`, `1:1`) are not matched by the cell reference
/// pattern and are flagged by rule REF307.
pub struct CircularReferenceRule {
    /// Formula cells whose saved value is an error: how a spreadsheet stores a circular
    /// result, so a loop through one of them is active.
    errored: HashSet<(u16, u32, u32)>,
}

impl CircularReferenceRule {
    pub fn new() -> Self {
        Self {
            errored: HashSet::new(),
        }
    }
}

impl Default for CircularReferenceRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for CALC202 (walker path).
#[derive(Debug)]
pub struct CircularReferenceData {
    /// Cycle path as (sheet_index, row, col) tuples.
    pub cycle: Vec<(u16, u32, u32)>,
    /// The saved values show the loop evaluated as circular (a cell of it holds an error).
    pub active: bool,
}

impl ViolationData for CircularReferenceData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let path_str: Vec<String> = self
            .cycle
            .iter()
            .map(|(idx, r, c)| {
                let sheet_name = ctx.workbook.sheet_name_by_index(*idx).unwrap_or("Unknown");
                format!("{}!{}", sheet_name, CellReference::new(*r, *c))
            })
            .collect();
        let first = path_str.first().map_or("", |s| s.as_str());
        if self.active {
            format!(
                "Circular reference detected: {} -> {}",
                path_str.join(" -> "),
                first
            )
        } else {
            format!(
                "Potential circular reference (the dependency graph closes a loop; the saved values show no circular error): {} -> {}",
                path_str.join(" -> "),
                first
            )
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for CircularReferenceRule {
    fn id(&self) -> RuleId {
        RuleId::Calc202
    }

    fn name(&self) -> &'static str {
        "Circular Reference"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Calculations
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let Some(formula) = cell.as_formula() {
            let here = (sheet.sheet_index, at.0, at.1);
            if cell.value.is_error() {
                self.errored.insert(here);
            }
            let refs = match dependencies_of(formula, sheet.sheet_index, &ctx.name_to_index) {
                Some(found) => found,
                // A formula the parser does not read: every reference counts.
                None => extract_cell_references(
                    formula,
                    &CELL_REF_PATTERN,
                    sheet.sheet_index,
                    &ctx.name_to_index,
                ),
            };

            // The links of the sheet with other sheets, for REF306 (whole columns and rows and
            // defined names included, which the cell dependencies do not hold).
            match super::helpers::formula_links(formula, sheet.sheet_index, &ctx.name_to_index) {
                Some(links) => {
                    for to in links.sheets {
                        if to != sheet.sheet_index {
                            ctx.sheet_links.insert((sheet.sheet_index, to));
                        }
                    }
                    for name in links.names {
                        ctx.name_uses.insert((sheet.sheet_index, name));
                    }
                }
                None => {
                    for (to, _, _) in &refs {
                        if *to != sheet.sheet_index {
                            ctx.sheet_links.insert((sheet.sheet_index, *to));
                        }
                    }
                }
            }

            ctx.cell_dependencies
                .insert((sheet.sheet_index, at.0, at.1), refs);
        }
    }

    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let cycles = find_cycles(&ctx.cell_dependencies);
        let mut reported_cells = HashSet::new();

        for cycle in cycles {
            let is_duplicate = cycle.iter().any(|c| reported_cells.contains(c));

            if !is_duplicate {
                let active = cycle.iter().any(|c| self.errored.contains(c));
                for cell in &cycle {
                    reported_cells.insert(*cell);
                    // Expose circular cells for ERR102/ERR103 dominance logic (loops the
                    // saved values show as circular errors).
                    if active {
                        ctx.circular_cells.insert(*cell);
                    }
                }

                // Report on the first cell
                let (sheet_index, r, c) = &cycle[0];
                let cell_ref = CellReference::new(*r, *c);

                out.push(Violation::with_data(
                    RuleId::Calc202,
                    ViolationScope::Cell(*sheet_index, cell_ref),
                    CircularReferenceData {
                        cycle: cycle.clone(),
                        active,
                    },
                    if active {
                        Severity::Error
                    } else {
                        Severity::Warning
                    },
                ));
            }
        }
    }
}

/// References of a formula as nodes (ranges by their corners, as the rule always did).
/// `None` when the formula cannot be parsed.
fn dependencies_of(
    formula: &str,
    sheet_index: u16,
    name_to_index: &HashMap<String, u16>,
) -> Option<Vec<(u16, u32, u32)>> {
    use sheetrs_formula::ast::{Expr, RefArea};
    fn walk(
        e: &Expr,
        sheet_index: u16,
        names: &HashMap<String, u16>,
        out: &mut Vec<(u16, u32, u32)>,
    ) {
        match e {
            Expr::Ref(r) => {
                let sheet = match &r.sheet {
                    None => sheet_index,
                    Some(n) => match names.get(n.as_str()).copied() {
                        Some(i) => i,
                        None => return,
                    },
                };
                match r.area {
                    RefArea::Cell(c) => out.push((sheet, c.row, c.col)),
                    RefArea::Range(a, b) => {
                        out.push((sheet, a.row, a.col));
                        out.push((sheet, b.row, b.col));
                    }
                    // Whole columns and rows are not tracked (REF307 flags them).
                    RefArea::Cols(..) | RefArea::Rows(..) => {}
                }
            }
            Expr::Call(c) => {
                for a in c.args.iter().flatten() {
                    walk(a, sheet_index, names, out);
                }
            }
            Expr::Array(rows) => {
                for x in rows.iter().flatten() {
                    walk(x, sheet_index, names, out);
                }
            }
            Expr::Unary(_, x) | Expr::Percent(x) | Expr::Paren(x) => {
                walk(x, sheet_index, names, out);
            }
            Expr::Binary(_, a, b) => {
                walk(a, sheet_index, names, out);
                walk(b, sheet_index, names, out);
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Name(_) => {}
        }
    }
    let ast = sheetrs_formula::parse(formula.strip_prefix('=').unwrap_or(formula)).ok()?;
    let mut out = Vec::new();
    walk(&ast, sheet_index, name_to_index, &mut out);
    Some(out)
}

/// Extract cell references from a formula, resolving sheet names to indices.
///
/// Range references are represented by their corner cells (start and end) to prevent
/// excessive memory usage with large spreadsheets. External workbook references are
/// skipped since they cannot cause circular references within the current workbook.
fn extract_cell_references(
    formula: &str,
    pattern: &Regex,
    current_sheet_index: u16,
    name_to_index: &HashMap<String, u16>,
) -> Vec<(u16, u32, u32)> {
    let mut references = Vec::new();

    for cap in pattern.captures_iter(formula) {
        let sheet_index = if cap.get(1).is_some() {
            if let Some(quoted) = cap.get(2) {
                match name_to_index.get(quoted.as_str()).copied() {
                    Some(idx) => idx,
                    None => continue,
                }
            } else if let Some(unquoted) = cap.get(3) {
                match name_to_index.get(unquoted.as_str()).copied() {
                    Some(idx) => idx,
                    None => continue,
                }
            } else {
                current_sheet_index
            }
        } else {
            current_sheet_index
        };

        if let (Some(col_match), Some(row_match)) = (cap.get(4), cap.get(5)) {
            let col_str = col_match.as_str();
            let row_str = row_match.as_str();

            let Some((start_row, start_col)) = parse_cell_coords(row_str, col_str) else {
                continue;
            };

            if let (Some(end_col_match), Some(end_row_match)) = (cap.get(6), cap.get(7)) {
                let end_col_str = end_col_match.as_str();
                let end_row_str = end_row_match.as_str();

                if let Some((end_row, end_col)) = parse_cell_coords(end_row_str, end_col_str) {
                    // Only track corner cells to prevent memory issues
                    references.push((sheet_index, start_row, start_col));
                    references.push((sheet_index, end_row, end_col));
                }
            } else {
                references.push((sheet_index, start_row, start_col));
            }
        }
    }

    references
}

#[derive(PartialEq, Clone, Copy)]
enum VisitState {
    Unvisited,
    Visiting,
    Visited,
}

// Node type for global graph: (SheetIndex, Row, Col)
type Node = (u16, u32, u32);

/// Find all unique elementary cycles in the graph
fn find_cycles(dependencies: &HashMap<Node, Vec<Node>>) -> Vec<Vec<Node>> {
    let mut cycles = Vec::new();
    let mut state = HashMap::new();

    // Initialize all nodes as unvisited
    for cell in dependencies.keys() {
        state.insert(*cell, VisitState::Unvisited);
    }
    for deps in dependencies.values() {
        for dep in deps {
            if !state.contains_key(dep) {
                state.insert(*dep, VisitState::Unvisited);
            }
        }
    }

    // Sort keys for deterministic output
    let mut keys: Vec<Node> = state.keys().copied().collect();
    keys.sort();

    for start_node in keys {
        if state.get(&start_node) == Some(&VisitState::Unvisited) {
            // Iterative DFS
            let mut stack = vec![(start_node, 0)]; // (node, next_dep_idx)
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

                    state.insert(u, VisitState::Visiting);
                    in_path.insert(u);
                    path.push(u);
                }

                if let Some(deps) = dependencies.get(&u) {
                    if dep_idx < deps.len() {
                        // Push current node back with next index
                        stack.push((u, dep_idx + 1));
                        // Push neighbor to visit
                        stack.push((deps[dep_idx], 0));
                    } else {
                        // Finished visiting all neighbors
                        state.insert(u, VisitState::Visited);
                        in_path.remove(&u);
                        path.pop();
                    }
                } else {
                    // No dependencies
                    state.insert(u, VisitState::Visited);
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
    use crate::rules::walker::WorkbookWalker;
    use std::collections::HashMap;
    use std::path::PathBuf;

    fn create_test_workbook(sheet_name: &str, cells: HashMap<(u32, u32), Cell>) -> Workbook {
        let sheet = Sheet {
            name: sheet_name.to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((10, 10)),
            ..Default::default()
        };

        Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        }
    }

    fn run_rule(workbook: &Workbook) -> Vec<Violation> {
        let rule = CircularReferenceRule::new();
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(rule)];
        let walker = WorkbookWalker::new(workbook, rules);
        walker.walk()
    }

    #[test]
    fn test_circular_reference_direct() {
        let mut cells = HashMap::new();
        // A1 = A1+1
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A1+1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let violations = run_rule(&workbook);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Calc202);
    }

    #[test]
    fn test_circular_reference_indirect() {
        let mut cells = HashMap::new();
        // A1 = B1, B1 = A1
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=B1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=A1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let violations = run_rule(&workbook);

        assert!(!violations.is_empty());
    }

    #[test]
    fn test_circular_reference_range() {
        let mut cells = HashMap::new();
        // A1 = SUM(B1:B3)
        // B2 = A1
        // With non-expanding mode: A1 -> B1, B3. B2 -> A1. No cycle detected.
        // This test verifies the current behavior (no expansion)
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(B1:B3)")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                formula: Some(<Box<str>>::from("=A1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let violations = run_rule(&workbook);

        // Should NOT detect cycle with non-expanding mode
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_huge_range_limit() {
        let mut cells = HashMap::new();
        // A1 = SUM(A2:A10000) - this range is okay
        // But if A1:XFD1048576 was present it would be skipped
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(A2:A3)")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (2, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let workbook = create_test_workbook("Sheet1", cells);
        let violations = run_rule(&workbook);

        // Cycle: A1 -> A3 -> A1
        assert_eq!(violations.len(), 1);
    }

    fn formula_cell(row: u32, col: u32, f: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                formula: Some(<Box<str>>::from(f)),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        )
    }

    /// `errored` lists cells whose saved value is an error.
    fn run_cells_with(
        rule: CircularReferenceRule,
        cells: &[(u32, u32, &str)],
        errored: &[(u32, u32)],
    ) -> Vec<Violation> {
        let map = cells
            .iter()
            .map(|(r, c, f)| {
                let (k, mut cell) = formula_cell(*r, *c, f);
                if errored.contains(&(*r, *c)) {
                    cell.value = CellValue::Error(crate::violation::ExcelError::Value);
                }
                (k, cell)
            })
            .collect();
        let workbook = create_test_workbook("Sheet1", map);
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(rule)];
        WorkbookWalker::new(&workbook, rules).walk()
    }

    #[test]
    fn the_saved_values_decide_the_severity() {
        // The graph closes a loop, the saved values show nothing: a warning.
        let cells = [(0, 0, "=IF(B1>0,C1,0)"), (0, 2, "=A1+1")];
        let v = run_cells_with(CircularReferenceRule::new(), &cells, &[]);
        assert_eq!(v.len(), 1);
        assert_eq!(v[0].severity, Severity::Warning);
        // A cell of the loop holds an error in the file: an error.
        let v = run_cells_with(CircularReferenceRule::new(), &cells, &[(0, 2)]);
        assert_eq!(v.len(), 1);
        assert_eq!(v[0].severity, Severity::Error);
        // A loop through a function's argument counts the same, whichever argument.
        let v = run_cells_with(
            CircularReferenceRule::new(),
            &[(0, 0, "=SUMIF(B1:B3,\"x\",C1:C3)"), (0, 2, "=A1")],
            &[],
        );
        assert_eq!(v[0].severity, Severity::Warning);
        let v = run_cells_with(
            CircularReferenceRule::new(),
            &[(0, 0, "=SUMIF(B1:B3,\"x\",C1:C3)"), (0, 1, "=A1")],
            &[(0, 1)],
        );
        assert_eq!(v[0].severity, Severity::Error);
    }
}
