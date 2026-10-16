//! INT401–403: Formula Interruption detection
//!
//! Description: Detects cells that break a consistent formula pattern in a row or column.
//! - INT401: Interrupted by data (value cell breaks a formula sequence)
//! - INT402: Interrupted by empty cell
//! - INT403: Interrupted by a different formula pattern
//!
//! All three rules share a single `InterruptionRule` struct parameterized by the
//! `InterruptionKind`s it reports. One instance covers every enabled kind, so cells
//! are classified once regardless of how many of the three rules are enabled.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::parser_utils::{column_letters, column_letters_to_index};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::{Captures, Regex, Replacer};
use std::collections::HashMap;
use std::fmt::Write;
use std::sync::LazyLock;

/// Which type of interruption this rule instance detects.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InterruptionKind {
    /// INT401: formula sequence interrupted by a data (value) cell.
    Data,
    /// INT402: formula sequence interrupted by an empty cell.
    Empty,
    /// INT403: formula sequence interrupted by a different formula pattern.
    Other,
}

impl InterruptionKind {
    /// Returns the kind reported under `rule_id`, or `None` for non-INT rules.
    pub fn from_rule_id(rule_id: RuleId) -> Option<Self> {
        match rule_id {
            RuleId::Int401 => Some(Self::Data),
            RuleId::Int402 => Some(Self::Empty),
            RuleId::Int403 => Some(Self::Other),
            _ => None,
        }
    }

    /// Returns the rule ID under which this kind is reported.
    pub fn rule_id(self) -> RuleId {
        match self {
            Self::Data => RuleId::Int401,
            Self::Empty => RuleId::Int402,
            Self::Other => RuleId::Int403,
        }
    }
}

/// Classification of a cell for interruption analysis.
#[derive(Debug, Clone)]
enum CellType {
    /// Cell has a formula with the given normalized pattern.
    Formula(Box<str>),
    /// Cell has a value but no formula.
    Value,
    /// Cell is empty (no value, no formula).
    Empty,
}

/// Per-sheet collected cell data: sheet_index → Vec<(row, col, CellType)>.
type SheetCellData = HashMap<u16, Vec<(u32, u32, CellType)>>;

/// An interruption detected during sequence scanning.
#[derive(Debug)]
struct Interruption {
    /// Row of the interrupting cell.
    row: u32,
    /// Column of the interrupting cell.
    col: u32,
    /// Kind of interruption.
    kind: InterruptionKind,
    /// Whether the interruption is in a column (true) or row (false).
    is_column: bool,
    /// The axis index (column index if is_column, row index if !is_column).
    axis_index: u32,
    /// Start of the formula sequence (row if column, col if row).
    seq_start: u32,
    /// End of the formula sequence (row if column, col if row).
    seq_end: u32,
}

/// Rule that detects formula interruptions.
pub struct InterruptionRule {
    /// Kinds of interruption this instance reports; the first one gives its
    /// [`WalkerRule::id`] and [`WalkerRule::name`].
    kinds: Vec<InterruptionKind>,
    /// Minimum number of formula cells for a sequence to be considered.
    min_formula_sequence: usize,
    /// Per-sheet collected cell classifications.
    sheet_data: SheetCellData,
}

/// Match cell references: optional $ + 1-3 uppercase letters + optional $ + digits.
/// The {1,3} limit matches valid Excel columns (A–XFD).
static ABS_REF_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(\$?)([A-Z]{1,3})(\$?)([0-9]+)").unwrap());

/// `ABS_REF_RE` replacer that writes R1C1-style offsets relative to the
/// formula's cell straight into the output buffer.
struct R1C1Replacer<'a> {
    formula_bytes: &'a [u8],
    cell_row: i64,
    cell_col: i64,
}

impl Replacer for R1C1Replacer<'_> {
    fn replace_append(&mut self, caps: &Captures<'_>, dst: &mut String) {
        let m = caps.get(0).unwrap();

        // Skip if preceded by a letter — part of a function or sheet name
        // (e.g. LOG10 → LOG is preceded by nothing but "10" is the row part;
        //  the full match LOG10 would be LOG as col + 10 as row)
        if m.start() > 0 && self.formula_bytes[m.start() - 1].is_ascii_alphabetic() {
            dst.push_str(m.as_str());
            return;
        }

        // Skip if followed by ! (sheet qualifier, e.g. S1!A1 → skip S1)
        // or ( (function call, e.g. IF(...) — unlikely with {1,3} but safe)
        if matches!(self.formula_bytes.get(m.end()), Some(b'!' | b'(')) {
            dst.push_str(m.as_str());
            return;
        }

        let col_abs = &caps[1] == "$";
        let row_abs = &caps[3] == "$";

        let ref_col = i64::from(
            column_letters_to_index(&caps[2]).expect("ABS_REF_RE captures column letters"),
        );
        let ref_row = caps[4].parse::<i64>().unwrap_or(1) - 1;

        let _ = if row_abs {
            write!(dst, "R{ref_row}")
        } else {
            write!(dst, "R[{}]", ref_row - self.cell_row)
        };
        let _ = if col_abs {
            write!(dst, "C{ref_col}")
        } else {
            write!(dst, "C[{}]", ref_col - self.cell_col)
        };
    }
}

impl InterruptionRule {
    /// Creates a rule reporting the given interruption kinds.
    ///
    /// # Panics
    ///
    /// Panics if `kinds` is empty.
    pub fn new(kinds: Vec<InterruptionKind>, config: &LinterConfig) -> Self {
        assert!(
            !kinds.is_empty(),
            "InterruptionRule needs at least one kind"
        );
        let min_seq = config
            .get_param_int("int_min_formula_sequence", None)
            .unwrap_or(3) as usize;
        Self {
            kinds,
            min_formula_sequence: min_seq,
            sheet_data: HashMap::new(),
        }
    }

    /// Normalize a formula by converting cell references to R1C1-style offsets.
    /// Relative references become relative offsets; absolute ($) references stay absolute.
    /// Skips false matches on function names (e.g. LOG10) and sheet qualifiers (e.g. SHEET1!).
    pub(crate) fn normalize_formula(formula: &str, cell_row: u32, cell_col: u32) -> Box<str> {
        let normalized = ABS_REF_RE.replace_all(
            formula,
            R1C1Replacer {
                formula_bytes: formula.as_bytes(),
                cell_row: i64::from(cell_row),
                cell_col: i64::from(cell_col),
            },
        );
        // Copied to an exact-size allocation: patterns are retained for every
        // formula cell of the sheet, and shrinking the oversized replacement
        // buffer in place fragments the heap and raises peak RSS.
        Box::from(&*normalized)
    }

    /// Scan a sequence of cells along one axis and find interruptions.
    ///
    /// The algorithm iterates over sorted cells (O(n)), using a HashMap for
    /// O(1) position lookups. Only a gap of exactly 1 absent position is
    /// bridged — wider gaps are treated as separate formula contexts.
    ///
    /// **Gap policy (gap=1):** A single absent cell between formula cells is
    /// likely an oversight (missing formula, formatting artifact) and is
    /// reported as an Empty interruption. Two or more absent cells indicate a
    /// deliberate section boundary (headers, spacing) and break the run.
    fn find_interruptions(
        cells: &[(u32, CellType)],
        is_column: bool,
        axis_index: u32,
        min_seq: usize,
    ) -> Vec<Interruption> {
        if cells.len() < min_seq {
            return Vec::new();
        }

        let mut interruptions = Vec::new();

        // O(1) lookup from position → cell type
        let cell_map: HashMap<u32, &CellType> = cells.iter().map(|(pos, ct)| (*pos, ct)).collect();

        /// Classify a position from the cell_map (absent = Empty).
        fn classify<'a>(cell_map: &'a HashMap<u32, &'a CellType>, pos: u32) -> &'a CellType {
            cell_map.get(&pos).copied().unwrap_or(&CellType::Empty)
        }

        // Walk through sorted cells, building formula runs.
        let mut i = 0;
        while i < cells.len() {
            // Find the start of a formula run
            let pattern = match &cells[i].1 {
                CellType::Formula(p) => p.clone(),
                _ => {
                    i += 1;
                    continue;
                }
            };

            let run_start = cells[i].0;
            let mut run_end = run_start;
            let mut run_interruptions: Vec<(u32, InterruptionKind)> = Vec::new();

            // Extend the run by scanning forward through cells.
            let mut j = i + 1;
            while j < cells.len() {
                let (pos, ct) = &cells[j];

                // Gap policy: only bridge a gap of exactly 1 absent position.
                // Wider gaps break the run (new formula context).
                if *pos > run_end + 2 {
                    break;
                }

                // If there is exactly 1 absent position between run_end and
                // this cell, record it as an Empty interruption.
                if *pos == run_end + 2 && !cell_map.contains_key(&(run_end + 1)) {
                    run_interruptions.push((run_end + 1, InterruptionKind::Empty));
                }

                match ct {
                    CellType::Formula(p) if p == &pattern => {
                        run_end = *pos;
                        j += 1;
                    }
                    CellType::Formula(_) => {
                        run_interruptions.push((*pos, InterruptionKind::Other));
                        run_end = *pos;
                        j += 1;
                    }
                    CellType::Value => {
                        run_interruptions.push((*pos, InterruptionKind::Data));
                        run_end = *pos;
                        j += 1;
                    }
                    CellType::Empty => {
                        run_interruptions.push((*pos, InterruptionKind::Empty));
                        run_end = *pos;
                        j += 1;
                    }
                }
            }

            // For each interruption, check that at least one side has ≥ min_seq
            // consecutive matching formulas.
            for &(int_pos, int_kind) in &run_interruptions {
                // Count formulas before the interruption
                let mut before_count = 0usize;
                let mut check = int_pos;
                while check > run_start {
                    check -= 1;
                    match classify(&cell_map, check) {
                        CellType::Formula(p) if p == &pattern => before_count += 1,
                        _ => break,
                    }
                }

                // Count formulas after the interruption
                let mut after_count = 0usize;
                check = int_pos + 1;
                while check <= run_end {
                    match classify(&cell_map, check) {
                        CellType::Formula(p) if p == &pattern => after_count += 1,
                        _ => break,
                    }
                    check += 1;
                }

                if before_count >= min_seq || after_count >= min_seq {
                    let (row, col) = if is_column {
                        (int_pos, axis_index)
                    } else {
                        (axis_index, int_pos)
                    };
                    interruptions.push(Interruption {
                        row,
                        col,
                        kind: int_kind,
                        is_column,
                        axis_index,
                        seq_start: run_start,
                        seq_end: run_end,
                    });
                }
            }

            // Advance past this run
            i = if j > i { j } else { i + 1 };
        }

        interruptions
    }
}

/// Incident data for INT401–403.
#[derive(Debug)]
pub struct InterruptionData {
    /// Kind of interruption.
    pub kind: InterruptionKind,
    /// Whether the interruption is in a column (true) or row (false).
    pub is_column: bool,
    /// The axis index (column for vertical, row for horizontal).
    pub axis_index: u32,
    /// Start of the formula sequence.
    pub seq_start: u32,
    /// End of the formula sequence.
    pub seq_end: u32,
}

impl ViolationData for InterruptionData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let kind_label = match self.kind {
            InterruptionKind::Data => "Value cell",
            InterruptionKind::Empty => "Empty cell",
            InterruptionKind::Other => "Different formula",
        };
        let axis_label = if self.is_column {
            format!("column {}", column_letters(self.axis_index))
        } else {
            format!("row {}", self.axis_index + 1)
        };
        format!(
            "{} interrupts a formula sequence in {} (positions {}–{}).",
            kind_label,
            axis_label,
            self.seq_start + 1,
            self.seq_end + 1,
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for InterruptionRule {
    fn id(&self) -> RuleId {
        self.kinds[0].rule_id()
    }

    fn name(&self) -> &'static str {
        match self.kinds[0] {
            InterruptionKind::Data => "Interrupted by Data",
            InterruptionKind::Empty => "Interrupted by Empty",
            InterruptionKind::Other => "Interrupted by Other",
        }
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Interruptions
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        let cell_type = if let Some(formula_upper) = ctx.formula_upper(sheet, cell) {
            CellType::Formula(Self::normalize_formula(formula_upper, cell.row, cell.col))
        } else if cell.value.is_empty() {
            CellType::Empty
        } else {
            CellType::Value
        };

        self.sheet_data
            .entry(sheet.sheet_index)
            .or_default()
            .push((cell.row, cell.col, cell_type));
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(cells) = self.sheet_data.remove(&sheet.sheet_index) else {
            return;
        };

        // Group by column → scan vertically
        let mut by_col: HashMap<u32, Vec<(u32, CellType)>> = HashMap::new();
        for (row, col, ct) in &cells {
            by_col.entry(*col).or_default().push((*row, ct.clone()));
        }
        for (col, mut col_cells) in by_col {
            col_cells.sort_by_key(|(row, _)| *row);
            let ints = Self::find_interruptions(&col_cells, true, col, self.min_formula_sequence);
            for int in ints {
                if !self.kinds.contains(&int.kind) {
                    continue;
                }
                out.push(Violation::with_data(
                    int.kind.rule_id(),
                    ViolationScope::Cell(sheet.sheet_index, CellReference::new(int.row, int.col)),
                    InterruptionData {
                        kind: int.kind,
                        is_column: int.is_column,
                        axis_index: int.axis_index,
                        seq_start: int.seq_start,
                        seq_end: int.seq_end,
                    },
                    Severity::Warning,
                ));
            }
        }

        // Group by row → scan horizontally
        let mut by_row: HashMap<u32, Vec<(u32, CellType)>> = HashMap::new();
        for (row, col, ct) in &cells {
            by_row.entry(*row).or_default().push((*col, ct.clone()));
        }
        for (row, mut row_cells) in by_row {
            row_cells.sort_by_key(|(col, _)| *col);
            let ints = Self::find_interruptions(&row_cells, false, row, self.min_formula_sequence);
            for int in ints {
                if !self.kinds.contains(&int.kind) {
                    continue;
                }
                out.push(Violation::with_data(
                    int.kind.rule_id(),
                    ViolationScope::Cell(sheet.sheet_index, CellReference::new(int.row, int.col)),
                    InterruptionData {
                        kind: int.kind,
                        is_column: int.is_column,
                        axis_index: int.axis_index,
                        seq_start: int.seq_start,
                        seq_end: int.seq_end,
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
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> Cell {
        Cell {
            formula: Some(<Box<str>>::from(formula)),
            is_array: false,
            num_fmt: None,
            row,
            col,
            value: CellValue::Empty,
        }
    }

    fn make_value_cell(row: u32, col: u32, value: f64) -> Cell {
        Cell {
            formula: None,
            is_array: false,
            num_fmt: None,
            row,
            col,
            value: CellValue::Number(value),
        }
    }

    fn make_sheet(cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for cell in cells {
            cell_map.insert((cell.row, cell.col), cell);
        }
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map,
            ..Default::default()
        }
    }

    fn run_rule(kind: InterruptionKind, sheet: &Sheet) -> Vec<Violation> {
        let config = LinterConfig::default();
        let mut rule = InterruptionRule::new(vec![kind], &config);
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(sheet, cell, &mut ctx, &mut Vec::new());
        }
        collect_violations(|out| rule.on_sheet_end(sheet, &mut ctx, out))
    }

    // --- INT401: Interrupted by Data ---

    #[test]
    fn test_value_interrupts_formula_column() {
        // Column A: F,F,F,V,F,F,F — the V at row 3 interrupts
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_value_cell(3, 0, 42.0),
            make_formula_cell(4, 0, "B5+C5"),
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
        ]);
        let violations = run_rule(InterruptionKind::Data, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Int401);
    }

    #[test]
    fn test_value_interrupts_formula_row() {
        // Row 0: F,F,F,V,F,F,F — the V at col 3 interrupts
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "A2+A3"),
            make_formula_cell(0, 1, "B2+B3"),
            make_formula_cell(0, 2, "C2+C3"),
            make_value_cell(0, 3, 42.0),
            make_formula_cell(0, 4, "E2+E3"),
            make_formula_cell(0, 5, "F2+F3"),
            make_formula_cell(0, 6, "G2+G3"),
        ]);
        let violations = run_rule(InterruptionKind::Data, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_too_short_sequence() {
        // Column A: F,F,V,F,F — neither segment is ≥3
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_value_cell(2, 0, 42.0),
            make_formula_cell(3, 0, "B4+C4"),
            make_formula_cell(4, 0, "B5+C5"),
        ]);
        let violations = run_rule(InterruptionKind::Data, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_no_interruption() {
        // Column A: F,F,F,F — all formulas, no interruption
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_formula_cell(3, 0, "B4+C4"),
        ]);
        let violations = run_rule(InterruptionKind::Data, &sheet);

        assert_eq!(violations.len(), 0);
    }

    // --- INT402: Interrupted by Empty ---

    #[test]
    fn test_empty_interrupts_formula() {
        // Column A: F,F,F,_,F,F,F — empty at row 3 interrupts
        // (row 3 simply absent from cell map = empty)
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            // row 3 is absent → empty
            make_formula_cell(4, 0, "B5+C5"),
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
        ]);
        let violations = run_rule(InterruptionKind::Empty, &sheet);

        // Empty cell at row 3 is not in the HashMap, so it won't be collected
        // by on_cell(). The algorithm needs to detect gaps between formula cells.
        // With the current implementation that only sees populated cells,
        // gaps are detected if the row/col indices are non-consecutive.
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Int402);
    }

    #[test]
    fn test_empty_at_end() {
        // Column A: F,F,F,_ — empty at end is not inside a run
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
        ]);
        let violations = run_rule(InterruptionKind::Empty, &sheet);

        assert_eq!(violations.len(), 0);
    }

    // --- INT403: Interrupted by Other ---

    #[test]
    fn test_different_formula_interrupts() {
        // Column A: same pattern, one different formula
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_formula_cell(3, 0, "SUM(B4:C4)"), // different pattern
            make_formula_cell(4, 0, "B5+C5"),
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
        ]);
        let violations = run_rule(InterruptionKind::Other, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Int403);
    }

    #[test]
    fn test_same_pattern_no_violation() {
        // Column A: all same pattern
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
        ]);
        let violations = run_rule(InterruptionKind::Other, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_normalization() {
        // Verify that formulas with different absolute refs but same relative structure normalize identically
        let n1 = InterruptionRule::normalize_formula("B1+C1", 0, 0);
        let n2 = InterruptionRule::normalize_formula("B2+C2", 1, 0);
        let n3 = InterruptionRule::normalize_formula("B3+C3", 2, 0);
        assert_eq!(n1, n2);
        assert_eq!(n2, n3);
    }

    #[test]
    fn test_normalization_absolute_and_mixed_refs() {
        assert_eq!(
            &*InterruptionRule::normalize_formula("$B$3+B$3+$B3+B3", 1, 0),
            "R2C1+R2C[1]+R[1]C1+R[1]C[1]"
        );
    }

    #[test]
    fn test_normalization_skips_function_names_and_sheet_qualifiers() {
        assert_eq!(
            &*InterruptionRule::normalize_formula("LOG10(S1!A1)", 0, 0),
            "LOG10(S1!R[0]C[0])"
        );
    }

    // --- Requested kinds ---

    /// Column A breaks one run per kind: a value at row 3, an absent cell at
    /// row 8 and a different formula at row 13.
    fn make_sheet_with_every_kind() -> Sheet {
        let mut cells = Vec::new();
        for row in 0..18 {
            match row {
                3 => cells.push(make_value_cell(row, 0, 42.0)),
                8 => {}
                13 => cells.push(make_formula_cell(row, 0, "SUM(B1:C1)")),
                _ => cells.push(make_formula_cell(row, 0, &format!("B{0}+C{0}", row + 1))),
            }
        }
        make_sheet(cells)
    }

    fn run_rule_with_kinds(kinds: Vec<InterruptionKind>, sheet: &Sheet) -> Vec<RuleId> {
        let mut rule = InterruptionRule::new(kinds, &LinterConfig::default());
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(sheet, cell, &mut ctx, &mut Vec::new());
        }
        let mut ids: Vec<RuleId> =
            collect_violations(|out| rule.on_sheet_end(sheet, &mut ctx, out))
                .iter()
                .map(|v| v.rule_id)
                .collect();
        ids.sort();
        ids
    }

    #[test]
    fn test_reports_every_requested_kind_once() {
        let ids = run_rule_with_kinds(
            vec![
                InterruptionKind::Data,
                InterruptionKind::Empty,
                InterruptionKind::Other,
            ],
            &make_sheet_with_every_kind(),
        );
        assert_eq!(ids, [RuleId::Int401, RuleId::Int402, RuleId::Int403]);
    }

    #[test]
    fn test_reports_only_requested_kinds() {
        let ids = run_rule_with_kinds(vec![InterruptionKind::Empty], &make_sheet_with_every_kind());
        assert_eq!(ids, [RuleId::Int402]);
    }

    #[test]
    fn test_id_is_first_requested_kind() {
        let rule = InterruptionRule::new(
            vec![InterruptionKind::Other, InterruptionKind::Data],
            &LinterConfig::default(),
        );
        assert_eq!(rule.id(), RuleId::Int403);
    }

    #[test]
    #[should_panic(expected = "at least one kind")]
    fn test_new_rejects_empty_kinds() {
        InterruptionRule::new(Vec::new(), &LinterConfig::default());
    }

    // --- Config ---

    #[test]
    fn test_min_sequence_config() {
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "int_min_formula_sequence".to_string(),
            toml::Value::Integer(5),
        );
        let mut rule = InterruptionRule::new(vec![InterruptionKind::Data], &config);

        // Column: F×5,V,F×5 — each side has 5 formulas, which meets min=5
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_formula_cell(3, 0, "B4+C4"),
            make_formula_cell(4, 0, "B5+C5"),
            make_value_cell(5, 0, 42.0),
            make_formula_cell(6, 0, "B7+C7"),
            make_formula_cell(7, 0, "B8+C8"),
            make_formula_cell(8, 0, "B9+C9"),
            make_formula_cell(9, 0, "B10+C10"),
            make_formula_cell(10, 0, "B11+C11"),
        ]);
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        // With min=5, each side has 5 formulas → interruption is valid
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_min_sequence_config_below_threshold() {
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "int_min_formula_sequence".to_string(),
            toml::Value::Integer(5),
        );
        let mut rule = InterruptionRule::new(vec![InterruptionKind::Data], &config);

        // Column: F×3,V,F×3 — each side has only 3, below min=5
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_value_cell(3, 0, 42.0),
            make_formula_cell(4, 0, "B5+C5"),
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
        ]);
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        // With min=5, neither side has ≥5 → no violation
        assert_eq!(violations.len(), 0);
    }

    // --- Gap policy: gap=1 boundary ---

    #[test]
    fn test_gap_1_absent_cell_is_bridged() {
        // Column A: F,F,F,_,F,F,F — gap of 1 absent cell is bridged
        // (row 3 absent → empty interruption)
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            // row 3 absent
            make_formula_cell(4, 0, "B5+C5"),
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
        ]);
        let violations = run_rule(InterruptionKind::Empty, &sheet);
        assert_eq!(
            violations.len(),
            1,
            "gap=1 should be bridged as interruption"
        );
    }

    #[test]
    fn test_gap_2_absent_cells_breaks_run() {
        // Column A: F,F,F,_,_,F,F,F — gap of 2 absent cells breaks the run
        // (rows 3,4 absent → separate contexts → no interruption)
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            // rows 3,4 absent
            make_formula_cell(5, 0, "B6+C6"),
            make_formula_cell(6, 0, "B7+C7"),
            make_formula_cell(7, 0, "B8+C8"),
        ]);
        let violations = run_rule(InterruptionKind::Empty, &sheet);
        assert_eq!(
            violations.len(),
            0,
            "gap=2 should break the run (separate contexts)"
        );
    }

    #[test]
    fn test_sparse_data_no_false_merge() {
        // Column A: formulas at rows 0–2 and rows 1000–1002 — far apart
        // Should be treated as two separate contexts, not one interrupted run
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "B1+C1"),
            make_formula_cell(1, 0, "B2+C2"),
            make_formula_cell(2, 0, "B3+C3"),
            make_formula_cell(1000, 0, "B1001+C1001"),
            make_formula_cell(1001, 0, "B1002+C1002"),
            make_formula_cell(1002, 0, "B1003+C1003"),
        ]);
        let violations = run_rule(InterruptionKind::Empty, &sheet);
        assert_eq!(violations.len(), 0, "sparse data should not be merged");
    }
}
