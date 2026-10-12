//! Linter rule system

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
    /// Sheets referenced by formulas (populated by calc202, used by ref306)
    pub referenced_sheets: HashSet<u16>,
    /// Cell dependency graph for circular reference detection
    pub cell_dependencies: CellDependencyMap,
    /// Error cells found during walk: location → error type (populated by err102).
    pub error_cells: HashMap<(u16, u32, u32), ExcelError>,
    /// Cells on a circular reference path (populated by calc202 `on_workbook_end`).
    pub circular_cells: HashSet<(u16, u32, u32)>,
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
    pub fn formula_upper(&mut self, sheet: &Sheet, cell: &Cell) -> Option<&str> {
        let formula = cell.as_formula()?;
        let location = (sheet.sheet_index, cell.row, cell.col);
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
    ) -> Vec<Violation> {
        Vec::new()
    }

    /// Called at the start of each sheet
    fn on_sheet_start(&mut self, _sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        Vec::new()
    }

    /// Called for each cell in the workbook
    fn on_cell(
        &mut self,
        _sheet: &Sheet,
        _cell: &Cell,
        _ctx: &mut LinterContext,
    ) -> Vec<Violation> {
        Vec::new()
    }

    /// Called at the end of each sheet
    fn on_sheet_end(&mut self, _sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        Vec::new()
    }

    /// Called once after all sheets have been processed
    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        _ctx: &mut LinterContext,
    ) -> Vec<Violation> {
        Vec::new()
    }
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

    fn make_cell(row: u32, col: u32, formula: Option<&str>) -> Cell {
        Cell {
            formula: formula.map(Into::into),
            is_array: false,
            num_fmt: None,
            row,
            col,
            value: CellValue::Empty,
        }
    }

    #[test]
    fn test_formula_upper_non_formula_cell() {
        let mut ctx = LinterContext::default();
        let sheet = Sheet::default();
        assert_eq!(ctx.formula_upper(&sheet, &make_cell(0, 0, None)), None);
    }

    #[test]
    fn test_formula_upper_unicode() {
        let mut ctx = LinterContext::default();
        let sheet = Sheet::default();
        let cell = make_cell(0, 0, Some("sum('año'!a1)"));
        assert_eq!(ctx.formula_upper(&sheet, &cell), Some("SUM('AÑO'!A1)"));
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
            ctx.formula_upper(&sheet, &make_cell(0, 0, Some("now()"))),
            Some("NOW()")
        );
        assert_eq!(
            ctx.formula_upper(&sheet, &make_cell(1, 0, Some("rand()"))),
            Some("RAND()")
        );
        assert_eq!(
            ctx.formula_upper(&other_sheet, &make_cell(1, 0, Some("today()"))),
            Some("TODAY()")
        );
    }
}
