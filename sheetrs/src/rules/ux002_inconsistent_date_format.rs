use crate::config::LinterConfig;
use crate::reader::{CellValue, Workbook};
use crate::rules::LinterRule;
use crate::violation::{Severity, Violation, ViolationScope};

/// UX002: Inconsistent date format
///
/// Detects if date cells do not use the specified format.
/// To enforce consistency, all dates should use the same format string.
///
/// Configuration:
/// - `date_format`: The required format string (default: "mm/dd/yyyy")
pub struct InconsistentDateFormatRule {
    default_date_format: String,
    // Config is stored to look up per-sheet overrides.
    // Actually, as discussed, I'll store the object or just use default for now
    // and rely on `check` to look up (Wait, I can't look up without config).
    // I'll stick to a simple struct and assume global config for now,
    // OR I'll modify `LinterRule` trait to pass config to `check`?
    // Changing trait is a big refactor.
    // I'll store `config` in the rule struct.
    config: LinterConfig,
}

impl InconsistentDateFormatRule {
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

    /// Check if a format string represents a date
    fn is_date_format(fmt: &str) -> bool {
        // Simple heuristic: contains date characters
        // Exclude simple number formats
        let lower = fmt.to_lowercase();

        // Exclude "Red", "Blue" etc colors
        let lower_no_color = lower.replace("[red]", "").replace("[blue]", "");

        // Check for d, m, y. Note 'm' can be minutes, but usually in context of time.
        // Should ANY date/time that doesn't match the specific format be flagged?
        // Or only Dates?
        // User request: "Detect if date not formatted as a specific format setup"
        // So if it IS a date, it must match.
        // Excel date formats: d, m, y.
        (lower_no_color.contains('d')
            || lower_no_color.contains('y')
            || (lower_no_color.contains('m')
                && !lower_no_color.contains('0')
                && !lower_no_color.contains('#')))
            && !lower_no_color.contains("general") // General is not a date
    }
}

impl LinterRule for InconsistentDateFormatRule {
    fn id(&self) -> &'static str {
        "UX002"
    }

    fn name(&self) -> &'static str {
        "InconsistentDateFormatRule"
    }

    fn category(&self) -> crate::rules::RuleCategory {
        crate::rules::RuleCategory::FormattingAndUsability
    }

    fn check(&self, workbook: &Workbook) -> anyhow::Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            // Get sheet-specific format if exists
            let required_format = self
                .config
                .get_param_str("date_format", Some(&sheet.name))
                .unwrap_or(&self.default_date_format);

            for ((row, col), cell) in &sheet.cells {
                // Only date cells are relevant.
                // In Excel, dates are numbers with a date format.
                // Only date cells are relevant.
                // In Excel, dates are numbers. In ODS, they might be stored as text (ISO strings) with a style.
                // Formulas can also result in dates.
                let is_candidate = match cell.value {
                    CellValue::Number(_) => true,
                    CellValue::Text(_) => true,
                    CellValue::Formula { .. } => true,
                    _ => false,
                };

                if is_candidate
                    && let Some(fmt) = &cell.num_fmt {
                        // Normalize format: remove escape backslashes common in XLSX (e.g. "mm\-dd\-yyyy" -> "mm-dd-yyyy")
                        let normalized_fmt = fmt.replace('\\', "");

                        // Check if it's a date format (using original check, but on normalized or raw? usually safe to check raw)
                        if Self::is_date_format(&normalized_fmt)
                            && normalized_fmt != required_format.replace('\\', "") {
                                violations.push(Violation::new(
                                    self.id(),
                                    ViolationScope::Cell(
                                        sheet.name.clone(),
                                        crate::violation::CellReference {
                                            row: *row,
                                            col: *col,
                                        },
                                    ),
                                    format!(
                                        "Date format '{}' does not match required format '{}'",
                                        normalized_fmt, required_format
                                    ),
                                    Severity::Warning,
                                ));
                            }
                    }
            }
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::{Cell, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_date_format_check() {
        let mut cells = HashMap::new();
        // Correct format
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: Some("mm/dd/yyyy".to_string()),
                row: 0,
                col: 0,
                value: CellValue::Number(44000.0),
            },
        );
        // Incorrect format (d-m-y)
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: Some("dd-mm-yyyy".to_string()),
                row: 0,
                col: 1,
                value: CellValue::Number(44000.0),
            },
        );
        // Not a date (General)
        cells.insert(
            (0, 2),
            Cell {
                num_fmt: Some("General".to_string()),
                row: 0,
                col: 2,
                value: CellValue::Number(123.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 3)),
            hidden_columns: vec![],
            hidden_rows: vec![],
            merged_cells: vec![],
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
            sheet_path: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: vec![],
            has_macros: false,
            external_workbooks: Vec::new(),
        };

        let config = LinterConfig::default();
        // Set default to mm/dd/yyyy
        // LinterConfig default uses empty hashmaps, so "mm/dd/yyyy" logic inside rule handles it.
        // Params can be injected to test override.

        let rule = InconsistentDateFormatRule::new(&config);
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("dd-mm-yyyy"));
    }
}
