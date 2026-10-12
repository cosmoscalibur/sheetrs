//! DAT703: Inconsistent date formatting detection
//!
//! Description: Identifies date values formatted as integers or other drift from standard temporal formats.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, CellValue, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::sync::Arc;

/// Rule that detects inconsistent date formats within contiguous ranges.
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
    /// The cell's `num_fmt`, shared via `Arc<str>`.
    pub num_fmt: Arc<str>,
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

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
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
                return vec![Violation::with_data(
                    RuleId::Data703,
                    ViolationScope::Cell(
                        sheet.sheet_index,
                        CellReference {
                            row: cell.row,
                            col: cell.col,
                        },
                    ),
                    DateFormatData {
                        num_fmt: Arc::clone(fmt),
                    },
                    Severity::Warning,
                )];
            }
        }

        Vec::new()
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
                num_fmt: Some(Arc::from("mm/dd/yyyy")),
                row: 0,
                col: 0,
                value: CellValue::Number(44000.0),
            },
        );
        // Incorrect format (d-m-y)
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: Some(Arc::from("dd-mm-yyyy")),
                row: 0,
                col: 1,
                value: CellValue::Number(44000.0),
            },
        );
        // Not a date (General)
        cells.insert(
            (0, 2),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: Some(Arc::from("General")),
                row: 0,
                col: 2,
                value: CellValue::Number(123.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let config = LinterConfig::default();
        let rule = InconsistentDateFormatRule::new(&config);
        let mut ctx = LinterContext::default();

        let mut violations = Vec::new();
        for cell in sheet.all_cells() {
            violations.extend(rule.on_cell(&sheet, cell, &mut ctx));
        }

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data703);
    }
}
