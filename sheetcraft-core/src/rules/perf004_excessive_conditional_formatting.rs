//! PERF004: Excessive conditional formatting detection

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct ExcessiveConditionalFormattingRule {
    threshold: u32,
}

impl ExcessiveConditionalFormattingRule {
    pub fn new(config: &LinterConfig) -> Self {
        let threshold = config
            .get_rule_config("PERF004")
            .and_then(|c| c.get_int("max_conditional_formatting"))
            .unwrap_or(5) as u32;

        Self { threshold }
    }
}

impl Default for ExcessiveConditionalFormattingRule {
    fn default() -> Self {
        Self { threshold: 5 }
    }
}

impl LinterRule for ExcessiveConditionalFormattingRule {
    fn id(&self) -> &str {
        "PERF004"
    }

    fn name(&self) -> &str {
        "Excessive conditional formatting"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Performance
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // Only check XLSX files
        if workbook.path.extension().and_then(|s| s.to_str()) != Some("xlsx") {
            return Ok(violations);
        }

        // Open the XLSX file and count conditional formatting rules
        use std::fs::File;
        use std::io::BufReader;

        let file = File::open(&workbook.path)?;
        let reader = BufReader::new(file);
        let mut archive = zip::ZipArchive::new(reader)?;

        for (index, sheet) in workbook.sheets.iter().enumerate() {
            let cf_count =
                crate::reader::xml_parser::count_conditional_formatting(&mut archive, index)?;

            if cf_count > self.threshold as usize {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Sheet(sheet.name.clone()),
                    format!(
                        "Sheet has {} conditional formatting rules (threshold: {})",
                        cf_count, self.threshold
                    ),
                    Severity::Warning,
                ));
            }
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_non_xlsx_file() {
        // For non-XLSX files, should return no violations
        let workbook = Workbook {
            path: PathBuf::from("test.ods"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                formula_parsing_error: None,
            }],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = ExcessiveConditionalFormattingRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}
