//! CALC201: Hardcoded values in formulas detection
//!
//! Description: Detects static numeric constants embedded within formula logic.
//! Reports one violation per cell, aggregating all hardcoded constants found in
//! the formula.

use super::{LinterContext, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::sync::LazyLock;

/// Time and calendar constants excluded by `ignore_hardcoded_time_constants`.
///
/// | Value | Meaning                          |
/// |------:|----------------------------------|
/// |     7 | days per week                    |
/// |    12 | months per year                  |
/// |    24 | hours per day                    |
/// |    30 | days per month (approx)          |
/// |    31 | days per month (max)             |
/// |    52 | weeks per year                   |
/// |    53 | weeks per year (max)             |
/// |    60 | minutes per hour / seconds per minute |
/// |   365 | days per year                    |
/// |   366 | days per leap year               |
/// |  3600 | seconds per hour                 |
const TIME_CONSTANTS: &[f64] = &[7.0, 12.0, 24.0, 30.0, 31.0, 52.0, 53.0, 60.0, 365.0, 366.0, 3600.0];

/// Matches quoted strings (to strip before number extraction).
static STRING_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#""[^"]*""#).expect("CALC201 string regex must compile"));

/// Matches external workbook references like `[1]`, `[2]`.
static EXTERNAL_REF_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\[(\d+)\]").expect("CALC201 external ref regex must compile"));

/// Matches numeric literals (integers and decimals).
/// `\b` prevents matching digits inside cell references (e.g. "A1") or
/// function names (e.g. "LOG10").
static NUMBER_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b(\d+(\.\d+)?)\b").expect("CALC201 number regex must compile"));

/// Matches function calls like `ROUND(`, `VLOOKUP(`, `RANK.EQ(`.
static FUNC_CALL_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"([A-Za-z][A-Za-z0-9._]*)\(").expect("CALC201 func call regex must compile")
});

/// Function argument positions (0-indexed) considered structural.
///
/// Numbers in these positions are mechanical parameters — they define function
/// behavior rather than encoding business logic. They are excluded from
/// violation reports when `ignore_hardcoded_func_args` is enabled.
///
/// Excluded from this list (business logic even in function args):
/// - `PERCENTILE`/`QUARTILE` k value — statistical threshold choice
/// - `CEILING`/`FLOOR`/`MROUND` significance — rounding granularity
/// - `LARGE`/`SMALL` k — "top N" selection
/// - `IF` comparison values — business rule thresholds
const STRUCTURAL_ARGS: &[(&str, &[usize])] = &[
    ("ROUND",     &[1]),     // num_digits
    ("ROUNDUP",   &[1]),     // num_digits
    ("ROUNDDOWN", &[1]),     // num_digits
    ("VLOOKUP",   &[2, 3]), // col_index, range_lookup
    ("HLOOKUP",   &[2, 3]), // row_index, range_lookup
    ("INDEX",     &[1, 2]), // row_num, col_num
    ("MATCH",     &[2]),    // match_type
    ("CHOOSE",    &[0]),    // index_num
    ("LEFT",      &[1]),    // num_chars
    ("RIGHT",     &[1]),    // num_chars
    ("MID",       &[1, 2]), // start_num, num_chars
    ("WEEKDAY",   &[1]),    // return_type
    ("YEARFRAC",  &[2]),    // basis
    ("RANK",      &[2]),    // order
    ("RANK.EQ",   &[2]),    // order
    ("RANK.AVG",  &[2]),    // order
    ("SUBTOTAL",  &[0]),    // function_num
];

/// Bundled configuration for which hardcoded values to ignore.
struct IgnoreConfig<'a> {
    /// Specific numeric values to skip.
    values: &'a [f64],
    /// Skip 0 and 1.
    zero_one: bool,
    /// Skip time/calendar constants.
    time: bool,
    /// Skip all integer literals.
    ints: bool,
    /// Skip powers of 10.
    pow10: bool,
    /// Skip structural function argument literals.
    func_args: bool,
}

/// Rule that detects hardcoded numeric values in formulas.
///
/// Walks each cell; for cells containing formulas, extracts all numeric
/// constants, filters out ignored values, and emits a single violation
/// per cell listing every unique hardcoded constant found.
pub struct HardcodedValuesInFormulasRule {
    config: LinterConfig,
}

impl HardcodedValuesInFormulasRule {
    /// Create a new rule instance.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            config: config.clone(),
        }
    }

    /// Check whether a value should be ignored based on configuration.
    fn is_ignored(val: f64, cfg: &IgnoreConfig<'_>) -> bool {
        // Check 0 and 1
        if cfg.zero_one && (val.abs() < f64::EPSILON || (val - 1.0).abs() < f64::EPSILON) {
            return true;
        }

        // Check time/calendar constants
        if cfg.time
            && TIME_CONSTANTS
                .iter()
                .any(|&x| (x - val).abs() < f64::EPSILON)
        {
            return true;
        }

        // Check exact match against user-provided list (with epsilon)
        if cfg.values
            .iter()
            .any(|&x| (x - val).abs() < f64::EPSILON)
        {
            return true;
        }

        // Check if integer
        if cfg.ints && val.fract().abs() < f64::EPSILON {
            return true;
        }

        // Check if power of 10 (including fractional powers like 0.1, 0.01)
        if cfg.pow10 {
            // Power of 10 must be positive
            if val > 0.0 {
                let log = val.log10();
                if log.fract().abs() < 1e-10 {
                    return true;
                }
            }
        }

        false
    }

    /// Extract byte ranges `(start, end)` for each argument of a function call.
    ///
    /// `paren_pos` is the index of the opening `(`. Returns `None` if
    /// parentheses are unbalanced. Each range is `[start, end)` in the
    /// formula byte string.
    fn extract_arg_byte_ranges(
        formula: &str,
        paren_pos: usize,
    ) -> Option<Vec<(usize, usize)>> {
        let bytes = formula.as_bytes();
        if paren_pos >= bytes.len() || bytes[paren_pos] != b'(' {
            return None;
        }

        let mut depth: u32 = 0;
        let mut args = Vec::new();
        let mut arg_start = paren_pos + 1;
        let mut in_string = false;

        for (i, &b) in bytes[paren_pos..].iter().enumerate() {
            let pos = paren_pos + i;
            match b {
                b'"' => in_string = !in_string,
                b'(' if !in_string => {
                    depth += 1;
                    if depth == 1 {
                        arg_start = pos + 1;
                    }
                }
                b')' if !in_string => {
                    depth -= 1;
                    if depth == 0 {
                        // Push last argument (even if empty — consistent with extract_args)
                        args.push((arg_start, pos));
                        return Some(args);
                    }
                }
                b',' if !in_string && depth == 1 => {
                    args.push((arg_start, pos));
                    arg_start = pos + 1;
                }
                _ => {}
            }
        }

        None // Unbalanced
    }

    /// Compute byte ranges of numeric literals that are sole structural
    /// function arguments.
    ///
    /// For each function call in `formula` whose name is in `STRUCTURAL_ARGS`,
    /// checks whether the structural argument positions contain a bare numeric
    /// literal (the full trimmed content is a single number). If so, the byte
    /// range of that literal is added to the returned set.
    fn find_structural_arg_ranges(formula: &str) -> Vec<(usize, usize)> {
        let mut excluded = Vec::new();

        for cap in FUNC_CALL_RE.captures_iter(formula) {
            let func_name = cap.get(1).unwrap().as_str();
            // Position of the opening '(' is the last char of the full match
            let paren_pos = cap.get(0).unwrap().end() - 1;

            let structural_positions = STRUCTURAL_ARGS
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case(func_name))
                .map(|(_, positions)| *positions);

            if let Some(positions) = structural_positions
                && let Some(arg_ranges) = Self::extract_arg_byte_ranges(formula, paren_pos) {
                    for &pos in positions {
                        if let Some(&(start, end)) = arg_ranges.get(pos) {
                            let content = formula[start..end].trim();
                            // Check if the argument is a bare numeric literal
                            if !content.is_empty()
                                && let Some(m) = NUMBER_RE.find(content)
                                && m.start() == 0
                                && m.end() == content.len()
                            {
                                // Map back to formula positions
                                let offset = start
                                    + formula[start..end]
                                        .find(content)
                                        .unwrap_or(0);
                                excluded.push((offset, offset + content.len()));
                            }
                        }
                    }
                }
        }

        excluded
    }

    /// Extract non-ignored hardcoded constants from a formula string.
    fn extract_hardcoded_values(
        formula: &str,
        cfg: &IgnoreConfig<'_>,
    ) -> Vec<f64> {
        // Remove strings first
        let formula_no_strings = STRING_RE.replace_all(formula, "");

        // Collect positions of external workbook references to exclude
        let excluded_ranges: Vec<(usize, usize)> = EXTERNAL_REF_RE
            .captures_iter(&formula_no_strings)
            .filter_map(|cap| cap.get(0).map(|m| (m.start(), m.end())))
            .collect();

        // Collect positions of structural function arguments to exclude
        let structural_ranges = if cfg.func_args {
            Self::find_structural_arg_ranges(&formula_no_strings)
        } else {
            Vec::new()
        };

        let mut values = Vec::new();

        for cap in NUMBER_RE.captures_iter(&formula_no_strings) {
            if let Some(match_str) = cap.get(1) {
                let match_start = match_str.start();
                let match_end = match_str.end();

                // Skip numbers inside external reference brackets
                let is_external_ref = excluded_ranges
                    .iter()
                    .any(|(start, end)| match_start >= *start && match_end <= *end);

                // Skip numbers that are sole structural function arguments
                let is_structural = structural_ranges
                    .iter()
                    .any(|(start, end)| match_start >= *start && match_end <= *end);

                if !is_external_ref
                    && !is_structural
                    && let Ok(val) = match_str.as_str().parse::<f64>()
                    && !Self::is_ignored(val, cfg)
                    && !values.contains(&val)
                {
                    values.push(val);
                }
            }
        }

        values
    }
}

/// Incident data for CALC201.
#[derive(Debug)]
pub struct HardcodedValuesData {
    /// The deduplicated hardcoded values found in the formula.
    pub values: Box<[f64]>,
}

impl ViolationData for HardcodedValuesData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let formatted: Vec<String> = self.values.iter().map(|v| v.to_string()).collect();
        format!(
            "Hardcoded values found in formula: {}",
            formatted.join(", ")
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for HardcodedValuesInFormulasRule {
    fn id(&self) -> RuleId {
        RuleId::Calc201
    }

    fn name(&self) -> &'static str {
        "Hardcoded Number"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::Calculations
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula) = cell.as_formula() {
            let ignored_values = self
                .config
                .get_param_float_array("ignore_hardcoded_num_values", Some(&sheet.name))
                .unwrap_or_default();

            let cfg = IgnoreConfig {
                values: &ignored_values,
                zero_one: self.config
                    .get_param_bool("ignore_hardcoded_zero_one", Some(&sheet.name))
                    .unwrap_or(true),
                time: self.config
                    .get_param_bool("ignore_hardcoded_time_constants", Some(&sheet.name))
                    .unwrap_or(true),
                ints: self.config
                    .get_param_bool("ignore_hardcoded_int_values", Some(&sheet.name))
                    .unwrap_or(false),
                pow10: self.config
                    .get_param_bool("ignore_hardcoded_power_of_ten", Some(&sheet.name))
                    .unwrap_or(true),
                func_args: self.config
                    .get_param_bool("ignore_hardcoded_func_args", Some(&sheet.name))
                    .unwrap_or(true),
            };

            let values = Self::extract_hardcoded_values(formula, &cfg);

            if !values.is_empty() {
                return vec![Violation::with_data(
                    RuleId::Calc201,
                    ViolationScope::Cell(
                        sheet.sheet_index,
                        CellReference {
                            row: cell.row,
                            col: cell.col,
                        },
                    ),
                    HardcodedValuesData {
                        values: values.into_boxed_slice(),
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
    use crate::reader::Workbook;
    use crate::reader::workbook::CellValue;
    use crate::reader::{Cell, Sheet};
    use crate::rules::walker::WorkbookWalker;
    use std::collections::HashMap;
    use std::path::PathBuf;
    use toml::Value;

    /// Helper: run the walker for CALC201 and return violations.
    fn run_calc201(config: &LinterConfig, workbook: &Workbook) -> Vec<Violation> {
        let rules: Vec<Box<dyn WalkerRule>> =
            vec![Box::new(HardcodedValuesInFormulasRule::new(config))];
        let walker = WorkbookWalker::new(workbook, rules);
        walker.walk()
    }

    #[test]
    fn test_hardcoded_values() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=123+A1")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        ); // 123 (int)
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=0+1.5")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        ); // 0 (ignored by zero_one), 1.5 (float)
        cells.insert(
            (0, 2),
            Cell {
                formula: Some(<Box<str>>::from(r#"=IF(A1>10, "Value: 5", 100)"#)),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Empty,
            },
        ); // 10 (pow10), 5 (in string, stripped), 100 (pow10)
        cells.insert(
            (0, 3),
            Cell {
                formula: Some(<Box<str>>::from("=0.1+0.01")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 3,
                value: CellValue::Empty,
            },
        ); // 0.1 (pow10), 0.01 (pow10)

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 4)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Case 1: Default config
        // ignore_zero_one = true, ignore_time = true, ignore_ints = false, ignore_pow10 = true
        // Should flag: 123 (cell A1), 1.5 (cell B1)
        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);

        // One violation per cell with hardcoded values
        assert_eq!(violations.len(), 2);
        let msgs: Vec<String> = violations.iter().map(|v| v.message()).collect();
        assert!(msgs.iter().any(|m| m.contains("123")));
        assert!(msgs.iter().any(|m| m.contains("1.5")));

        // Case 2: Override ints = true, pow10 = false
        let mut config2 = LinterConfig::default();
        config2.global.params.insert(
            "ignore_hardcoded_int_values".to_string(),
            Value::Boolean(true),
        );
        config2.global.params.insert(
            "ignore_hardcoded_power_of_ten".to_string(),
            Value::Boolean(false),
        );
        config2.global.params.insert(
            "ignore_hardcoded_num_values".to_string(),
            Value::Array(vec![]),
        );

        // Should flag: 1.5 (cell B1), 0.1+0.01 aggregated (cell D1)
        let violations2 = run_calc201(&config2, &workbook);
        let msgs2: Vec<String> = violations2.iter().map(|v| v.message()).collect();

        // Integers are ignored, so 123, 0, 10, 100, 5 filtered out
        assert!(!msgs2.iter().any(|m| m.contains("123")));
        assert!(msgs2.iter().any(|m| m.contains("1.5")));
        assert!(msgs2.iter().any(|m| m.contains("0.1")));
        assert!(msgs2.iter().any(|m| m.contains("0.01")));

        // Case 3: Sheet-specific override
        let mut config3 = LinterConfig::default();
        let mut sheet_config = crate::config::SheetConfig::default();
        sheet_config.params.insert(
            "ignore_hardcoded_int_values".to_string(),
            Value::Boolean(true),
        );
        config3.sheets.insert("Sheet1".to_string(), sheet_config);

        // 123 is integer, ignored on Sheet1
        let violations3 = run_calc201(&config3, &workbook);
        let msgs3: Vec<String> = violations3.iter().map(|v| v.message()).collect();
        assert!(!msgs3.iter().any(|m| m.contains("123")));
    }

    #[test]
    fn test_aggregation_multiple_constants_per_cell() {
        let mut cells = HashMap::new();
        // Single cell with multiple hardcoded constants
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A2*1.5+3.14+42")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);

        // Only one violation for the cell, not three
        assert_eq!(violations.len(), 1);
        let msg = &violations[0].message();
        assert!(msg.contains("1.5"));
        assert!(msg.contains("3.14"));
        assert!(msg.contains("42"));
    }

    #[test]
    fn test_external_link_indices_not_flagged() {
        let mut cells = HashMap::new();

        // External link formulas - indices should NOT be flagged
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=[1]Sheet1!A1")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=[2]Data!B5")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        );

        // External link with actual constant - constant SHOULD be flagged
        // 5 is in the default ignore list, so let's use 6
        cells.insert(
            (0, 2),
            Cell {
                formula: Some(<Box<str>>::from("=[1]Sheet1!A1+6")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 3)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Disable all ignore flags (except external)
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "ignore_hardcoded_int_values".to_string(),
            Value::Boolean(false),
        );
        config.global.params.insert(
            "ignore_hardcoded_power_of_ten".to_string(),
            Value::Boolean(false),
        );
        config.global.params.insert(
            "ignore_hardcoded_num_values".to_string(),
            Value::Array(vec![]),
        );
        config.global.params.insert(
            "ignore_hardcoded_zero_one".to_string(),
            Value::Boolean(false),
        );
        config.global.params.insert(
            "ignore_hardcoded_time_constants".to_string(),
            Value::Boolean(false),
        );
        config.global.params.insert(
            "ignore_hardcoded_func_args".to_string(),
            Value::Boolean(false),
        );

        let violations = run_calc201(&config, &workbook);

        let msgs: Vec<String> = violations.iter().map(|v| v.message()).collect();

        // Indices 1 and 2 should NOT be flagged
        assert!(
            !msgs
                .iter()
                .any(|m| m == "Hardcoded values found in formula: 1")
        );
        assert!(
            !msgs
                .iter()
                .any(|m| m == "Hardcoded values found in formula: 2")
        );

        // The constant 6 SHOULD be flagged
        assert!(msgs.iter().any(|m| m.contains("6")));

        // Should have exactly 1 violation (the 6)
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_ignore_zero_one_param() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=0+A1")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=A1*1")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 2)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Default: ignore_zero_one = true → no violations for 0 and 1
        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);
        assert_eq!(violations.len(), 0);

        // Disable: ignore_zero_one = false, ignore_pow10 = false → 0 and 1 flagged
        // (1 is also 10^0, so pow10 must be disabled to test zero_one in isolation)
        let mut config2 = LinterConfig::default();
        config2.global.params.insert(
            "ignore_hardcoded_zero_one".to_string(),
            Value::Boolean(false),
        );
        config2.global.params.insert(
            "ignore_hardcoded_power_of_ten".to_string(),
            Value::Boolean(false),
        );
        let violations2 = run_calc201(&config2, &workbook);
        assert_eq!(violations2.len(), 2);
        let msgs: Vec<String> = violations2.iter().map(|v| v.message()).collect();
        assert!(msgs.iter().any(|m| m.contains("0")));
        assert!(msgs.iter().any(|m| m.contains("1")));
    }

    #[test]
    fn test_ignore_time_constants_param() {
        let mut cells = HashMap::new();
        // Formula with various time constants
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A1*24+60")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=A1/365")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 2)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Default: ignore_time = true → no violations for 24, 60, 365
        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);
        assert_eq!(violations.len(), 0);

        // Disable: ignore_time = false → 24, 60, 365 flagged
        let mut config2 = LinterConfig::default();
        config2.global.params.insert(
            "ignore_hardcoded_time_constants".to_string(),
            Value::Boolean(false),
        );
        let violations2 = run_calc201(&config2, &workbook);
        let msgs: Vec<String> = violations2.iter().map(|v| v.message()).collect();
        assert!(msgs.iter().any(|m| m.contains("24")));
        assert!(msgs.iter().any(|m| m.contains("60")));
        assert!(msgs.iter().any(|m| m.contains("365")));
    }

    #[test]
    fn test_structural_func_args_ignored() {
        let mut cells = HashMap::new();

        // ROUND: 2nd arg (num_digits) is structural
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=ROUND(A1, 6)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        // VLOOKUP: 3rd arg (col_index) is structural
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=VLOOKUP(A1, B:G, 6, FALSE)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        );

        // LEFT: 2nd arg (num_chars) is structural
        cells.insert(
            (0, 2),
            Cell {
                formula: Some(<Box<str>>::from("=LEFT(A1, 8)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Empty,
            },
        );

        // MID: 2nd and 3rd args are structural
        cells.insert(
            (0, 3),
            Cell {
                formula: Some(<Box<str>>::from("=MID(A1, 6, 9)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 3,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 4)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Default config: func_args = true → structural args excluded
        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);
        assert_eq!(violations.len(), 0, "Structural func args should not be flagged");
    }

    #[test]
    fn test_business_logic_func_args_flagged() {
        let mut cells = HashMap::new();

        // PERCENTILE: k=0.95 is business logic, not structural
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=PERCENTILE(A1:A10, 0.95)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        // CEILING: significance=0.05 is business logic
        cells.insert(
            (0, 1),
            Cell {
                formula: Some(<Box<str>>::from("=CEILING(A1, 0.05)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 2)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);
        let msgs: Vec<String> = violations.iter().map(|v| v.message()).collect();

        assert!(msgs.iter().any(|m| m.contains("0.95")), "PERCENTILE k should be flagged");
        assert!(msgs.iter().any(|m| m.contains("0.05")), "CEILING significance should be flagged");
    }

    #[test]
    fn test_nested_func_structural_args() {
        let mut cells = HashMap::new();

        // Nested: ROUND wraps VLOOKUP. Both structural args should be excluded,
        // but the hardcoded 1.21 in the first ROUND arg should be flagged.
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=ROUND(VLOOKUP(A1, B:H, 7, FALSE)*1.21, 6)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let config = LinterConfig::default();
        let violations = run_calc201(&config, &workbook);

        assert_eq!(violations.len(), 1, "Only business logic constant should be flagged");
        let msg = &violations[0].message();
        assert!(msg.contains("1.21"), "Tax rate 1.21 should be flagged");
        assert!(!msg.contains("7"), "VLOOKUP col index should not be flagged");
        assert!(!msg.contains("6"), "ROUND num_digits should not be flagged");
    }

    #[test]
    fn test_func_args_disabled() {
        let mut cells = HashMap::new();

        // ROUND with structural arg — should be flagged when func_args is disabled
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=ROUND(A1, 6)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Disable func_args → 6 should be flagged
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "ignore_hardcoded_func_args".to_string(),
            Value::Boolean(false),
        );
        let violations = run_calc201(&config, &workbook);
        assert_eq!(violations.len(), 1);
        assert!(violations[0].message().contains("6"));
    }
}
