//! DATA703: Inconsistent date formats
//!
//! Dates shown in different layouts within one workbook (`dd-mm-yyyy` next to `mm/dd/yyyy`) are
//! easy to misread, and `03/04` means two different days depending on the reader. The rule compares
//! every date-formatted cell with the one format the configuration requires.
//!
//! # Reports
//!
//! One violation per cell (unit: the cell) that has a number format recognised as a date format and
//! whose format differs from the required one. Only cells whose value is a number or text and that
//! carry an explicit number format are examined; a formula cell is judged by its saved value and
//! format like any other. A format counts as a date format when, lowercased and without `[red]` and
//! `[blue]`, it contains `d` or `y`, or contains `m` and neither `0` nor `#`, and it does not contain
//! `general`. This is a letter heuristic: any format with a literal `d` or `y` in it is treated as a
//! date. The comparison ignores backslashes (`mm\/dd\/yyyy` equals `mm/dd/yyyy`) and is otherwise an
//! exact, case-sensitive string comparison.
//!
//! # Does not report
//!
//! - Cells without a number format, with `General`, or with a numeric format (`0.00`, `#,##0`).
//! - Empty, boolean and error cells.
//! - Dates stored as plain integers or as text under a `General` format: the rule reads formats, not
//!   values, so it cannot tell that `44000` is a date. Numbers stored as text are DATA702.
//! - Cells whose date format is an equivalent spelling of the required one (for example `MM/DD/YYYY`):
//!   the comparison is case-sensitive, so these are reported.
//!
//! # Parameters
//!
//! - `date_format` (string, default `mm/dd/yyyy`): the required format. It can be set per sheet.
//!
//! # Severity
//!
//! Always `Warning`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, CellValue, NumFmt, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};

/// Rule that detects date-formatted cells whose format differs from the required one.
pub struct InconsistentDateFormatRule {
    /// Default date format from config.
    default_date_format: String,
    /// Rule configuration.
    config: LinterConfig,
}

impl InconsistentDateFormatRule {
    /// Create a new instance with configuration.
    pub fn new(config: &LinterConfig) -> Self {
        let default_date_format = config
            .get_param_str("date_format", None)
            .unwrap_or("mm/dd/yyyy")
            .to_string();

        Self {
            default_date_format,
            config: config.clone(),
        }
    }

    /// Check if a format string represents a date.
    fn is_date_format(fmt: &str) -> bool {
        let lower = fmt.to_lowercase();
        let lower_no_color = lower.replace("[red]", "").replace("[blue]", "");

        (lower_no_color.contains('d')
            || lower_no_color.contains('y')
            || (lower_no_color.contains('m')
                && !lower_no_color.contains('0')
                && !lower_no_color.contains('#')))
            && !lower_no_color.contains("general")
    }
}

/// Incident data for DATA703.
#[derive(Debug)]
pub struct DateFormatData {
    /// The cell's `num_fmt`, shared.
    pub num_fmt: NumFmt,
}

impl ViolationData for DateFormatData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let normalized = self.num_fmt.replace('\\', "");
        format!("Date format '{normalized}' does not match required format")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for InconsistentDateFormatRule {
    fn id(&self) -> RuleId {
        RuleId::Data703
    }

    fn name(&self) -> &'static str {
        "Inconsistent Date Format"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Data
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let is_candidate = matches!(cell.value, CellValue::Number(_) | CellValue::Text(_));

        if is_candidate && let Some(fmt) = &cell.num_fmt {
            let normalized_fmt = fmt.replace('\\', "");
            let required_format = self
                .config
                .get_param_str("date_format", Some(&sheet.name))
                .unwrap_or(&self.default_date_format);

            if Self::is_date_format(&normalized_fmt)
                && normalized_fmt != required_format.replace('\\', "")
            {
                out.push(Violation::with_data(
                    RuleId::Data703,
                    ViolationScope::Cell(
                        sheet.sheet_index,
                        CellReference {
                            row: at.0,
                            col: at.1,
                        },
                    ),
                    DateFormatData {
                        num_fmt: fmt.clone(),
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
    use crate::reader::Sheet;
    use std::collections::HashMap;

    #[test]
    fn test_date_format_check() {
        let mut cells = HashMap::new();
        // Correct format
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: Some("mm/dd/yyyy".into()),
                value: CellValue::Number(44000.0),
            },
        );
        // Incorrect format (d-m-y)
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: Some("dd-mm-yyyy".into()),
                value: CellValue::Number(44000.0),
            },
        );
        // Not a date (General)
        cells.insert(
            (0, 2),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: Some("General".into()),
                value: CellValue::Number(123.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let config = LinterConfig::default();
        let mut rule = InconsistentDateFormatRule::new(&config);
        let mut ctx = LinterContext::default();

        let mut violations = Vec::new();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut violations);
        }

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data703);
    }
}
