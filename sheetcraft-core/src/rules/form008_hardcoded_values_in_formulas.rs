use crate::config::LinterConfig;
use crate::reader::{CellValue, Workbook};
use crate::rules::{LinterRule, RuleCategory};
use crate::violation::{Severity, Violation, ViolationScope};
use regex::Regex;

/// FORM008: Hardcoded values in formulas
///
/// Detects hardcoded numeric values in formulas.
/// Hardcoded values make maintenance difficult and hide business logic.
///
/// Configuration:
/// - `ignore_hardcoded_num_values`: List of specific numbers to ignore (e.g. [1.5])
/// - `ignore_hardcoded_int_values`: If true, ignore all integer hardcoded values.
/// - `ignore_hardcoded_power_of_ten`: If true, ignore all power of ten hardcoded values (10, 100, 0.1, etc).
pub struct HardcodedValuesInFormulasRule {
    ignored_values: Vec<f64>,
    ignore_ints: bool,
    ignore_pow10: bool,
}

impl HardcodedValuesInFormulasRule {
    pub fn new(config: &LinterConfig) -> Self {
        let ignored_values = config
            .get_param_float_array("ignore_hardcoded_num_values", None)
            .unwrap_or_default();

        let ignore_ints = config
            .get_param_bool("ignore_hardcoded_int_values", None)
            .unwrap_or(false);

        let ignore_pow10 = config
            .get_param_bool("ignore_hardcoded_power_of_ten", None)
            .unwrap_or(false);

        Self {
            ignored_values,
            ignore_ints,
            ignore_pow10,
        }
    }

    fn is_ignored(&self, val: f64) -> bool {
        // Check exact match (with epsilon)
        if self
            .ignored_values
            .iter()
            .any(|&x| (x - val).abs() < f64::EPSILON)
        {
            return true;
        }

        // Check if integer
        if self.ignore_ints {
            if val.fract().abs() < f64::EPSILON {
                return true;
            }
        }

        // Check if power of 10
        if self.ignore_pow10 {
            // Power of 10 must be positive
            if val > 0.0 {
                let log = val.log10();
                if log.fract().abs() < f64::EPSILON {
                    return true;
                }
            }
        }

        false
    }
}

impl LinterRule for HardcodedValuesInFormulasRule {
    fn id(&self) -> &'static str {
        "FORM008"
    }

    fn name(&self) -> &'static str {
        "HardcodedValuesInFormulasRule"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> anyhow::Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // Regex to match quoted strings (to ignore them)
        let string_regex = Regex::new(r#""[^"]*""#).unwrap();

        // Regex to match numeric literals
        // Matches integers and decimals
        // \b ensures we match complete words. Since digits are word characters, \b prevents matching
        // digits preceded or followed by other word characters (like letters or underscores).
        // e.g., matches "123" in "123 + 456", but not "1" in "A1" or "10" in "LOG10".
        // Note: The regex crate does not support look-around/look-behind.
        let number_regex = Regex::new(r"\b(\d+(\.\d+)?)\b").unwrap();

        for sheet in &workbook.sheets {
            // Note: Ideally we would load sheet-specific config here if overriding is needed per sheet.
            // Currently using global/constructor config for simplicity and performance.
            // To support per-sheet overrides fully, we'd need to update `new` or passing logic,
            // or store `LinterConfig` and look up here.

            for ((row, col), cell) in &sheet.cells {
                if let CellValue::Formula(formula) = &cell.value {
                    // Remove strings first
                    let formula_no_strings = string_regex.replace_all(formula, "");

                    for cap in number_regex.captures_iter(&formula_no_strings) {
                        if let Some(match_str) = cap.get(1) {
                            let val_str = match_str.as_str();
                            if let Ok(val) = val_str.parse::<f64>() {
                                if !self.is_ignored(val) {
                                    violations.push(Violation::new(
                                        self.id(),
                                        ViolationScope::Cell(
                                            sheet.name.clone(),
                                            crate::violation::CellReference {
                                                row: *row,
                                                col: *col,
                                            },
                                        ),
                                        format!("Hardcoded value found in formula: {}", val),
                                        Severity::Warning,
                                    ));
                                }
                            }
                        }
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
    use toml::Value;

    #[test]
    fn test_hardcoded_values() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=123+A1".to_string()),
            },
        ); // 123 (int)
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Formula("=0+1.5".to_string()),
            },
        ); // 0 (int), 1.5 (float)
        cells.insert(
            (0, 2),
            Cell {
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Formula(r#"=IF(A1>10, "Value: 5", 100)"#.to_string()),
            },
        ); // 10 (int, pow10), 5 (string), 100 (int, pow10)

        cells.insert(
            (0, 3),
            Cell {
                num_fmt: None,
                row: 0,
                col: 3,
                value: CellValue::Formula("=0.1+0.01".to_string()),
            },
        ); // 0.1 (pow10), 0.01 (pow10)

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 4)),
            hidden_columns: vec![],
            hidden_rows: vec![],
            merged_cells: vec![],
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: vec![],
            has_macros: false,
        };

        // Case 1: Ignore ints = true, others default (false/empty)
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "ignore_hardcoded_int_values".to_string(),
            Value::Boolean(true),
        );

        // This should ignore: 123, 0, 10, 100
        // Should flag: 1.5, 0.1, 0.01

        let rule = HardcodedValuesInFormulasRule::new(&config);
        let violations = rule.check(&workbook).unwrap();

        let msgs: Vec<String> = violations.iter().map(|v| v.message.clone()).collect();
        assert!(msgs.contains(&"Hardcoded value found in formula: 1.5".to_string()));
        assert!(msgs.contains(&"Hardcoded value found in formula: 0.1".to_string()));
        assert!(msgs.contains(&"Hardcoded value found in formula: 0.01".to_string()));
        // Ints ignored
        assert!(!msgs.contains(&"Hardcoded value found in formula: 123".to_string()));
        assert!(!msgs.contains(&"Hardcoded value found in formula: 0".to_string()));

        // Case 2: Ignore pow10 = true
        let mut config2 = LinterConfig::default();
        config2.global.params.insert(
            "ignore_hardcoded_power_of_ten".to_string(),
            Value::Boolean(true),
        );

        // Should ignore: 10, 100, 0.1, 0.01
        // Should flag: 123, 0 (not strict pow10?), 1.5
        // 0 log10 is -inf. Not integer. So 0 is NOT ignored by pow10 logic.

        let rule2 = HardcodedValuesInFormulasRule::new(&config2);
        let violations2 = rule2.check(&workbook).unwrap();
        let msgs2: Vec<String> = violations2.iter().map(|v| v.message.clone()).collect();

        assert!(msgs2.contains(&"Hardcoded value found in formula: 123".to_string()));
        assert!(msgs2.contains(&"Hardcoded value found in formula: 1.5".to_string()));
        assert!(msgs2.contains(&"Hardcoded value found in formula: 0".to_string()));

        assert!(!msgs2.contains(&"Hardcoded value found in formula: 10".to_string()));
        assert!(!msgs2.contains(&"Hardcoded value found in formula: 100".to_string()));
        assert!(!msgs2.contains(&"Hardcoded value found in formula: 0.1".to_string()));
        assert!(!msgs2.contains(&"Hardcoded value found in formula: 0.01".to_string()));

        // Case 3: Specific list
        let mut config3 = LinterConfig::default();
        config3.global.params.insert(
            "ignore_hardcoded_num_values".to_string(),
            Value::Array(vec![Value::Float(1.5)]),
        );

        // Should ignore: 1.5
        // Should flag: everything else (defaults are false)

        let rule3 = HardcodedValuesInFormulasRule::new(&config3);
        let violations3 = rule3.check(&workbook).unwrap();
        let msgs3: Vec<String> = violations3.iter().map(|v| v.message.clone()).collect();

        assert!(!msgs3.contains(&"Hardcoded value found in formula: 1.5".to_string()));
        assert!(msgs3.contains(&"Hardcoded value found in formula: 123".to_string()));
    }
}
