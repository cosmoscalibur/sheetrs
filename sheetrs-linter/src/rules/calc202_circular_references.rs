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
//! result (`LibreOffice` writes an error, Excel writes 0, so Excel files show a warning). Active
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
use super::{CellDependencies, LinterContext, RuleCategory, WalkerRule};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use sheetrs_io::reader::{Cell, Sheet, Workbook};
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
#[derive(Debug)]
pub struct CircularReferenceRule {
    /// Formula cells whose saved value is an error: how a spreadsheet stores a circular
    /// result, so a loop through one of them is active.
    errored: HashSet<(u16, u32, u32)>,
    /// References of the formula being read, reused by every cell.
    scratch: Vec<Node>,
}

impl CircularReferenceRule {
    #[must_use]
    pub fn new() -> Self {
        Self {
            errored: HashSet::new(),
            scratch: Vec::new(),
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
    /// Cycle path as (`sheet_index`, row, col) tuples.
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
            self.scratch.clear();
            // One parse gives the references (the cells the formula reads) and the links of
            // the sheet with other sheets, for REF306 (whole columns and rows and defined names
            // included, which the cell dependencies do not hold).
            if !read_formula(
                formula,
                sheet.sheet_index,
                &ctx.name_to_index,
                &mut self.scratch,
                &mut ctx.sheet_links,
                &mut ctx.name_uses,
            ) {
                // A formula the parser does not read: every reference counts.
                self.scratch.extend(extract_cell_references(
                    formula,
                    &CELL_REF_PATTERN,
                    sheet.sheet_index,
                    &ctx.name_to_index,
                ));
                for (to, _, _) in &self.scratch {
                    if *to != sheet.sheet_index {
                        ctx.sheet_links.insert((sheet.sheet_index, *to));
                    }
                }
            }

            ctx.cell_dependencies.insert(here, self.scratch.drain(..));
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

/// Read a formula once: push the cells it reads into `refs` (ranges by their corners, as the rule
/// always did) and record its links with other sheets and the defined names it uses. `false`
/// when the formula cannot be parsed (nothing is recorded then).
fn read_formula(
    formula: &str,
    sheet_index: u16,
    name_to_index: &HashMap<String, u16>,
    refs: &mut Vec<Node>,
    sheet_links: &mut HashSet<(u16, u16)>,
    name_uses: &mut HashSet<(u16, String)>,
) -> bool {
    use sheetrs_formula::ast::{Expr, RefArea};
    struct Walk<'a> {
        sheet_index: u16,
        names: &'a HashMap<String, u16>,
        refs: &'a mut Vec<Node>,
        links: &'a mut HashSet<(u16, u16)>,
        name_uses: &'a mut HashSet<(u16, String)>,
        /// The sheet last linked, so a formula that names it again does not hash it again.
        last_link: Option<u16>,
    }
    fn walk(e: &Expr, w: &mut Walk<'_>) {
        match e {
            Expr::Ref(r) => {
                let sheet = match &r.sheet {
                    None => w.sheet_index,
                    Some(n) => match w.names.get(n.as_str()).copied() {
                        Some(i) => i,
                        None => return,
                    },
                };
                if sheet != w.sheet_index && w.last_link != Some(sheet) {
                    w.last_link = Some(sheet);
                    w.links.insert((w.sheet_index, sheet));
                }
                match r.area {
                    RefArea::Cell(c) => w.refs.push((sheet, c.row, c.col)),
                    RefArea::Range(a, b) => {
                        w.refs.push((sheet, a.row, a.col));
                        w.refs.push((sheet, b.row, b.col));
                    }
                    // Whole columns and rows are not tracked (REF307 flags them).
                    RefArea::Cols(..) | RefArea::Rows(..) => {}
                }
            }
            Expr::Name(n) => {
                w.name_uses.insert((w.sheet_index, n.to_uppercase()));
            }
            Expr::Call(c) => {
                for a in c.args.iter().flatten() {
                    walk(a, w);
                }
            }
            Expr::Array(rows) => {
                for x in rows.iter().flatten() {
                    walk(x, w);
                }
            }
            Expr::Unary(_, x) | Expr::Percent(x) | Expr::Paren(x) => walk(x, w),
            Expr::Binary(_, a, b) => {
                walk(a, w);
                walk(b, w);
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) => {}
        }
    }
    let Ok(ast) = sheetrs_formula::parse(formula.strip_prefix('=').unwrap_or(formula)) else {
        return false;
    };
    walk(
        &ast,
        &mut Walk {
            sheet_index,
            names: name_to_index,
            refs,
            links: sheet_links,
            name_uses,
            last_link: None,
        },
    );
    true
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

/// Find all unique elementary cycles in the graph.
///
/// The search runs over the position of each formula cell among the sorted ones, with its state in
/// a vector: only a formula can be on a loop, so a reference to a cell without one is not
/// followed. The order of the search (formula cells by location, references in the order of the
/// formula) is that of the search over every referenced cell, so the loops found are the same.
fn find_cycles(graph: &CellDependencies) -> Vec<Vec<Node>> {
    let mut cycles = Vec::new();
    let mut state = vec![VisitState::Unvisited; graph.len()];
    // (position of the cell, index of the next reference to follow)
    let mut stack: Vec<(usize, usize)> = Vec::new();
    let mut path: Vec<usize> = Vec::new();

    for start in 0..graph.len() {
        if state[start] != VisitState::Unvisited {
            continue;
        }
        stack.push((start, 0));

        while let Some((u, dep_idx)) = stack.pop() {
            if dep_idx == 0 {
                // First time visiting this node in this path
                match state[u] {
                    VisitState::Visiting => {
                        // Cycle detected!
                        if let Some(pos) = path.iter().position(|&x| x == u) {
                            cycles.push(path[pos..].iter().map(|&i| graph.cell(i)).collect());
                        }
                        continue;
                    }
                    VisitState::Visited => continue,
                    VisitState::Unvisited => {}
                }
                state[u] = VisitState::Visiting;
                path.push(u);
            }

            if let Some(dep) = graph.references_at(u).get(dep_idx) {
                // Push current node back with next index
                stack.push((u, dep_idx + 1));
                // Push neighbor to visit, if it is a formula
                if let Some(next) = graph.index_of(dep) {
                    stack.push((next, 0));
                }
            } else {
                // Finished visiting all neighbors
                state[u] = VisitState::Visited;
                path.pop();
            }
        }
    }

    cycles
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::walker::WorkbookWalker;
    use sheetrs_io::reader::workbook::{Cell, CellValue, Sheet};
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

    #[test]
    fn every_loop_is_found_once_in_the_order_of_the_search() {
        // Two separate loops (A1 <-> B1, D1 -> E1 -> F1 -> D1), a formula that reads the loop
        // without being on it (H1), and references to cells with no formula (Z9).
        let cells = [
            (0, 0, "=B1+Z9"),
            (0, 1, "=A1"),
            (0, 3, "=E1"),
            (0, 4, "=F1"),
            (0, 5, "=D1+Z9"),
            (0, 7, "=A1+D1"),
        ];
        let v = run_cells_with(CircularReferenceRule::new(), &cells, &[]);
        assert_eq!(v.len(), 2);
        let first = v[0].data::<CircularReferenceData>().expect("data");
        let second = v[1].data::<CircularReferenceData>().expect("data");
        assert_eq!(first.cycle, vec![(0, 0, 0), (0, 0, 1)]);
        assert_eq!(second.cycle, vec![(0, 0, 3), (0, 0, 4), (0, 0, 5)]);
    }

    #[test]
    fn a_formula_is_read_once_for_references_and_links() {
        let names: HashMap<String, u16> = [("Data".to_string(), 1), ("Here".to_string(), 0)].into();
        let mut refs = Vec::new();
        let mut links = HashSet::new();
        let mut uses = HashSet::new();
        assert!(read_formula(
            "=SUM(Data!A1:B2)+C3+Rate+Gone!D4+Data!E5+Here!F6+A:A",
            0,
            &names,
            &mut refs,
            &mut links,
            &mut uses,
        ));
        // Ranges by their corners; an unknown sheet and a whole column hold no cell.
        assert_eq!(
            refs,
            vec![(1, 0, 0), (1, 1, 1), (0, 2, 2), (1, 4, 4), (0, 5, 5)]
        );
        // Only the other sheet is a link (a whole column of the sheet itself is not one).
        assert_eq!(links, HashSet::from([(0, 1)]));
        assert_eq!(uses, HashSet::from([(0, "RATE".to_string())]));
        // A formula the parser does not read records nothing.
        let mut refs = Vec::new();
        assert!(!read_formula(
            "=((", 0, &names, &mut refs, &mut links, &mut uses
        ));
        assert!(refs.is_empty());
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
                    cell.value = CellValue::Error(sheetrs_io::reader::ExcelError::Value);
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
