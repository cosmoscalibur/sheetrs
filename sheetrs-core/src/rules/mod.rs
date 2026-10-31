//! Linter rule system.
//!
//! Every rule is a module named after its identifier (`dat702_numeric_formats` is DATA702) whose
//! documentation (`cargo doc -p sheetrs-core --open`) is the definition of the rule: what it decides
//! and why. A rule module's `//!` follows this layout, which a test (`rule_modules_are_documented`)
//! checks:
//!
//! ```text
//! //! RULEID: Title
//! //!
//! //! One paragraph: the problem the rule points at and why it matters.
//! //!
//! //! # Reports
//! //! What is reported, one item per violation (the unit: a cell, a range, a sheet, the book).
//! //!
//! //! # Does not report
//! //! What looks similar and is out of scope, with the reason, and which rule covers it.
//! //!
//! //! # Parameters
//! //! The configuration keys that change the result and their defaults (or "None").
//! //!
//! //! # Severity
//! //! The severity and when it changes.
//! ```
//!
//! The reason for a boundary belongs here, in the module that implements it, and not in the
//! changelog or a benchmark note.

pub mod helpers;
pub mod registry;
pub mod walker;

// Rule implementations - Excel Errors (1xx)
pub mod err101_broken_named_ranges;
pub mod err102_error_cells;
pub mod err103_ref_to_error;

// Rule implementations - Unreliable Calculations (2xx)
pub mod calc201_hardcoded_values;
pub mod calc202_circular_references;
pub mod calc203_double_operator;
pub mod calc204_approximate_lookup;
pub mod calc205_double_count;

// Rule implementations - Reference Issues (3xx)
pub mod ref301_unused_named_ranges;
pub mod ref302_duplicate_names;
pub mod ref303_empty_sheets;
pub mod ref304_large_used_range;
pub mod ref305_blank_rows_columns;
pub mod ref306_unused_sheets;
pub mod ref307_whole_column_row_refs;
pub mod ref309_ref_to_empty_cell;
pub mod ref310_longer_ref_expected;

// Rule implementations - Formula Interruptions (4xx)
pub mod int401_interrupted_by_data;
pub mod int402_interrupted_by_empty;
pub mod int403_interrupted_by_other;

// Rule implementations - Complexity (5xx)
pub mod cpx501_sheet_counts;
pub mod cpx502_merged_cells;
pub mod cpx503_excessive_conditional_formatting;
pub mod cpx504_deep_if_nesting;
pub mod cpx505_deep_formula_nesting;
pub mod cpx506_many_operations;
pub mod cpx507_multiple_sheet_ref;
pub mod cpx508_many_references;
pub mod cpx509_long_formula;

// Rule implementations - Vulnerable Formulas (6xx)
pub mod vul601_duplicate_formulas;
pub mod vul602_volatile_functions;
pub mod vul603_empty_string_test;
pub mod vul604_error_prone_functions;
pub mod vul605_legacy_array;
pub mod vul606_deprecated_func;

// Rule implementations - Data Issues (7xx)
pub mod dat701_sheet_names;
pub mod dat702_numeric_formats;
pub mod dat703_date_formats;
pub mod dat704_long_text;
pub mod dat705_unnecessary_space;
pub mod dat706_numeric_text_calc;

// Rule implementations - External References (8xx)
pub mod ext801_external_workbook;
pub mod ext802_web_urls;

// Rule implementations - Hidden Information (9xx)
pub mod hid901_hidden_worksheet;
pub mod hid902_hidden_columns_rows;

// Rule implementations - Files & Settings (10xx)
pub mod file1001_large_file_size;
pub mod file1002_old_spreadsheet;
pub mod file1003_date_system_1904;

// Rule implementations - VBA Issues (11xx)
pub mod vba1101_has_macros;

use crate::reader::{Cell, Sheet, Workbook};
use crate::violation::{ExcelError, RuleId, Violation};
use std::collections::{HashMap, HashSet};

/// Type alias for cell dependency graph: (SheetIndex, Row, Col) -> Vec<(SheetIndex, Row, Col)>
pub type CellDependencyMap = HashMap<(u16, u32, u32), Vec<(u16, u32, u32)>>;

/// Shared context for walker rules during single-pass execution
#[derive(Default)]
pub struct LinterContext {
    /// Map sheet names to indices for reference resolution
    pub name_to_index: HashMap<String, u16>,
    /// Links between sheets: `(from, to)` for every formula of sheet `from` that names a cell,
    /// range, whole column or whole row of a different sheet `to` (populated by calc202, used by
    /// ref306). A reference of a sheet to itself is not a link.
    pub sheet_links: HashSet<(u16, u16)>,
    /// Defined names used by the formulas of a sheet: `(sheet, NAME)`, uppercase (populated by
    /// calc202; ref306 resolves them to the sheets their definitions name).
    pub name_uses: HashSet<(u16, String)>,
    /// Cell dependency graph for circular reference detection
    pub cell_dependencies: CellDependencyMap,
    /// Error cells found during walk: location → error type (populated by err102).
    pub error_cells: HashMap<(u16, u32, u32), ExcelError>,
    /// Cells on a circular reference path (populated by calc202 `on_workbook_end`).
    pub circular_cells: HashSet<(u16, u32, u32)>,
    /// Cells whose saved value is `#NAME?` and that ERR102 does not count as errors: their
    /// formula calls a function written with the `_xlfn.` prefix (the application that saved the
    /// file did not know it), or reads such a cell (populated by err102).
    pub compat_name_cells: HashSet<(u16, u32, u32)>,
    /// Cells containing numeric-looking text values (populated by dat706).
    pub numeric_text_cells: HashSet<(u16, u32, u32)>,
    /// Uppercased formula of the last cell requested through
    /// [`Self::formula_upper`], keyed by `(sheet_index, row, col)`.
    formula_upper: Option<((u16, u32, u32), Box<str>)>,
}

impl LinterContext {
    /// Returns the cell's formula in uppercase, or `None` for non-formula cells.
    ///
    /// The result is cached for the cell's location, so every rule visiting the
    /// same cell shares one conversion. The location key assumes one context per
    /// workbook, as [`walker::WorkbookWalker`] creates.
    /// Rules that inspect string literals must use the original formula.
    pub fn formula_upper(&mut self, sheet: &Sheet, at: (u32, u32), cell: &Cell) -> Option<&str> {
        let formula = cell.as_formula()?;
        let location = (sheet.sheet_index, at.0, at.1);
        if self
            .formula_upper
            .as_ref()
            .is_none_or(|(cached, _)| *cached != location)
        {
            self.formula_upper = Some((location, formula.to_uppercase().into_boxed_str()));
        }
        self.formula_upper.as_ref().map(|(_, upper)| &**upper)
    }
}

/// Trait for optimized single-pass walker rules
pub trait WalkerRule: Send {
    /// Unique rule identifier
    fn id(&self) -> RuleId;

    /// Human-readable rule name
    fn name(&self) -> &'static str;

    /// Rule category
    fn category(&self) -> RuleCategory;

    /// Called once at the start of workbook processing
    fn on_workbook_start(
        &mut self,
        _workbook: &Workbook,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
    }

    /// Called at the start of each sheet
    fn on_sheet_start(
        &mut self,
        _sheet: &Sheet,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
    }

    /// Called for each cell in the workbook, with its `(row, column)` in the sheet, 0-based.
    fn on_cell(
        &mut self,
        _sheet: &Sheet,
        _at: (u32, u32),
        _cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
    }

    /// Called at the end of each sheet
    fn on_sheet_end(
        &mut self,
        _sheet: &Sheet,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
    }

    /// Called once after all sheets have been processed
    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
    }
}

/// Runs one `WalkerRule` hook against a fresh sink and returns what it emitted.
#[cfg(test)]
pub(crate) fn collect_violations(hook: impl FnOnce(&mut Vec<Violation>)) -> Vec<Violation> {
    let mut out = Vec::new();
    hook(&mut out);
    out
}

/// Rule categories matching the comparative report sections
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleCategory {
    /// Excel Errors (1xx)
    ExcelErrors,
    /// Unreliable Calculations (2xx)
    Calculations,
    /// Reference Issues (3xx)
    Reference,
    /// Formula Interruptions (4xx)
    Interruptions,
    /// Complexity (5xx)
    Complexity,
    /// Vulnerable Formulas (6xx)
    Vulnerability,
    /// Data Issues (7xx)
    Data,
    /// External References (8xx)
    External,
    /// Hidden Information (9xx)
    Hidden,
    /// Files & Settings (10xx)
    File,
    /// VBA Issues (11xx)
    VBA,
}

impl RuleCategory {
    /// Get the human-readable string representation of the category
    pub fn as_str(&self) -> &'static str {
        match self {
            RuleCategory::ExcelErrors => "Excel Errors",
            RuleCategory::Calculations => "Unreliable Calculations",
            RuleCategory::Reference => "Reference Issues",
            RuleCategory::Interruptions => "Formula Interruptions",
            RuleCategory::Complexity => "Complexity",
            RuleCategory::Vulnerability => "Vulnerable Formulas",
            RuleCategory::Data => "Data Issues",
            RuleCategory::External => "External References",
            RuleCategory::Hidden => "Hidden Information",
            RuleCategory::File => "Files & Settings",
            RuleCategory::VBA => "VBA Issues",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::CellValue;

    fn make_cell(row: u32, col: u32, formula: Option<&str>) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                formula: formula.map(Into::into),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        )
    }

    #[test]
    fn test_formula_upper_non_formula_cell() {
        let mut ctx = LinterContext::default();
        let sheet = Sheet::default();
        assert_eq!(
            ctx.formula_upper(&sheet, (0, 0), &make_cell(0, 0, None).1),
            None
        );
    }

    #[test]
    fn test_formula_upper_unicode() {
        let mut ctx = LinterContext::default();
        let sheet = Sheet::default();
        let (at, cell) = make_cell(0, 0, Some("sum('año'!a1)"));
        assert_eq!(ctx.formula_upper(&sheet, at, &cell), Some("SUM('AÑO'!A1)"));
    }

    #[test]
    fn test_formula_upper_recomputes_per_cell() {
        let mut ctx = LinterContext::default();
        let sheet = Sheet::default();
        let other_sheet = Sheet {
            sheet_index: 1,
            ..Default::default()
        };
        assert_eq!(
            ctx.formula_upper(&sheet, (0, 0), &make_cell(0, 0, Some("now()")).1),
            Some("NOW()")
        );
        assert_eq!(
            ctx.formula_upper(&sheet, (1, 0), &make_cell(1, 0, Some("rand()")).1),
            Some("RAND()")
        );
        assert_eq!(
            ctx.formula_upper(&other_sheet, (1, 0), &make_cell(1, 0, Some("today()")).1),
            Some("TODAY()")
        );
    }
}

#[cfg(test)]
mod documentation {
    /// Every rule module documents what it reports, what it does not, its parameters and its
    /// severity (the layout is in the documentation of this module).
    #[test]
    fn rule_modules_are_documented() {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/src/rules");
        let mut missing = Vec::new();
        for entry in std::fs::read_dir(dir).unwrap() {
            let path = entry.unwrap().path();
            let name = path.file_name().unwrap().to_string_lossy().into_owned();
            let is_rule = name.strip_suffix(".rs").is_some_and(|n| {
                n.trim_start_matches(|c: char| c.is_ascii_lowercase())
                    .starts_with(|c: char| c.is_ascii_digit())
                    && n.contains('_')
            });
            if !is_rule {
                continue;
            }
            let text = std::fs::read_to_string(&path).unwrap();
            let docs: String = text
                .lines()
                .take_while(|l| l.starts_with("//!") || l.trim().is_empty())
                .collect::<Vec<_>>()
                .join("\n");
            for heading in [
                "# Reports",
                "# Does not report",
                "# Parameters",
                "# Severity",
            ] {
                if !docs.contains(&format!("//! {heading}")) {
                    missing.push(format!("{name}: `{heading}`"));
                }
            }
        }
        assert!(
            missing.is_empty(),
            "rule modules without documentation:\n{}",
            missing.join("\n")
        );
    }
}
