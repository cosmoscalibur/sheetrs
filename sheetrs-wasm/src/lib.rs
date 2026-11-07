use serde::Serialize;
use sheetrs_io::reader::{Format, Source, Workbook};
use sheetrs_linter::report::{group_by_scope, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{FormatContext, Linter, LinterConfig, Severity, Violation};
use std::collections::HashMap;
use std::fmt::Write as _;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

#[derive(Serialize)]
struct RuleInfo {
    id: String,
    name: String,
    category: String,
    is_default: bool,
}

/// # Errors
///
/// A message when the result cannot be serialized.
#[wasm_bindgen]
pub fn lint_rules() -> Result<JsValue, JsValue> {
    let config = LinterConfig::default();
    let metadata = registry::rule_metadata(&config);

    let rules_info: Vec<RuleInfo> = metadata
        .iter()
        .map(|m| RuleInfo {
            id: m.id.to_string(),
            name: m.name.clone(),
            category: m.category.as_str().to_string(),
            is_default: m.is_default,
        })
        .collect();

    serde_wasm_bindgen::to_value(&rules_info)
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

/// The workbook of the bytes a page hands over, whose format its extension names.
fn source_of(file_data: &[u8], extension: &str) -> Result<Source, JsValue> {
    let format: Format = extension
        .parse()
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;
    Ok(Source::from_bytes(file_data, format))
}

/// The configuration of the text of a `sheetlint.toml`; the default rules when it is empty.
fn config_of(config_toml: &str) -> Result<LinterConfig, JsValue> {
    if config_toml.trim().is_empty() {
        Ok(LinterConfig::default())
    } else {
        LinterConfig::from_toml(config_toml)
            .map_err(|e| JsValue::from_str(&format!("Config error: {e}")))
    }
}

/// # Errors
///
/// A message when the workbook cannot be read or the configuration is not valid.
#[wasm_bindgen]
pub fn lint_workbook(
    file_data: &[u8],
    extension: &str,
    config_toml: &str,
) -> Result<String, JsValue> {
    let workbook = source_of(file_data, extension)?
        .read_workbook()
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;
    let config = config_of(config_toml)?;

    let linter = Linter::with_config(config.clone());
    let violations = linter.lint_workbook(&workbook);

    // Create a map of rule_id -> rule_name for better output
    let metadata = registry::rule_metadata(&config);
    let rule_names: HashMap<String, String> = metadata
        .iter()
        .map(|m| (m.id.to_string(), m.name.clone()))
        .collect();

    Ok(format_violations_human(&violations, &rule_names, &workbook))
}

/// Lints a workbook and returns the violations as an array of objects
/// (`rule_id`, `rule_name`, `severity`, `message`, `scope`, `sheet`, `cell`, `row`, `col`), the
/// structured counterpart of [`lint_workbook`]. `config_toml` is the text of a `sheetlint.toml`
/// (empty for the default rules).
///
/// # Errors
///
/// A message when the workbook cannot be read, the configuration is not valid or the result cannot be serialized.
#[wasm_bindgen]
pub fn lint_workbook_records(
    file_data: &[u8],
    extension: &str,
    config_toml: &str,
) -> Result<JsValue, JsValue> {
    let workbook = source_of(file_data, extension)?
        .read_workbook()
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;
    let linter = Linter::with_config(config_of(config_toml)?);
    let violations = linter.lint_workbook(&workbook);
    serde_wasm_bindgen::to_value(&violation_records(&violations, &workbook, linter.config()))
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

/// Statistics of a workbook as an object (the fields of `sheetstats --format json`), the
/// structured counterpart of [`get_workbook_stats`]. `total_file_size` is 0: bytes carry no size.
///
/// # Errors
///
/// A message when the workbook cannot be read or the result cannot be serialized.
#[wasm_bindgen]
pub fn workbook_stats(file_data: &[u8], extension: &str) -> Result<JsValue, JsValue> {
    let stats = sheetrs_stats::collect(&source_of(file_data, extension)?)
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;
    serde_wasm_bindgen::to_value(&stats)
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

fn format_violations_human(
    violations: &[Violation],
    rule_names: &HashMap<String, String>,
    workbook: &Workbook,
) -> String {
    let mut output = String::new();

    if violations.is_empty() {
        output.push_str("✓ No violations found!\n");
        return output;
    }

    let grouped = group_by_scope(violations, workbook);
    let ctx = FormatContext { workbook };

    if !grouped.book.is_empty() {
        output.push_str("📚 Book-level violations:\n");
        for violation in &grouped.book {
            output.push_str(&format_violation(violation, rule_names, 1, &ctx));
        }
        output.push('\n');
    }

    // One header per sheet, then its own violations and its cells.
    for (sheet_name, data) in &grouped.sheets {
        let _ = writeln!(output, "📄 Sheet: {sheet_name}");
        for violation in &data.sheet {
            output.push_str(&format_violation(violation, rule_names, 1, &ctx));
        }
        for (cell_ref, violations) in &data.cells {
            let _ = writeln!(output, "  📍 Cell: {cell_ref}");
            for violation in violations {
                output.push_str(&format_violation(violation, rule_names, 2, &ctx));
            }
        }
        output.push('\n');
    }

    // Print summary
    let error_count = violations
        .iter()
        .filter(|v| v.severity == Severity::Error)
        .count();
    let warning_count = violations
        .iter()
        .filter(|v| v.severity == Severity::Warning)
        .count();
    let info_count = violations
        .iter()
        .filter(|v| v.severity == Severity::Info)
        .count();

    output.push_str("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    output.push_str("📊 Summary:\n");
    if error_count > 0 {
        let _ = writeln!(output, "  ❌ Errors: {error_count}");
    }
    if warning_count > 0 {
        let _ = writeln!(output, "  ⚠️  Warnings: {warning_count}");
    }
    if info_count > 0 {
        let _ = writeln!(output, "  ℹ️  Info: {info_count}");
    }

    output
}

fn format_violation(
    violation: &Violation,
    rule_names: &HashMap<String, String>,
    indent: usize,
    ctx: &FormatContext<'_>,
) -> String {
    let indent_str = "  ".repeat(indent);
    let severity_icon = match violation.severity {
        Severity::Error => "❌",
        Severity::Warning => "⚠️ ",
        Severity::Info => "ℹ️ ",
    };

    let rule_id_str = violation.rule_id.as_str();
    let rule_name = rule_names
        .get(rule_id_str)
        .map_or("Unknown rule", std::string::String::as_str);

    format!(
        "{}{} [{}] {} - {}\n",
        indent_str,
        severity_icon,
        rule_id_str,
        rule_name,
        violation.format_message(ctx)
    )
}

/// # Errors
///
/// A message when the workbook cannot be read.
#[wasm_bindgen]
pub fn get_workbook_stats(file_data: &[u8], extension: &str) -> Result<String, JsValue> {
    let workbook = source_of(file_data, extension)?
        .read_workbook()
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;

    let mut output = String::new();

    output.push_str("📊 Workbook Statistics\n");
    output.push_str("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n");

    let _ = writeln!(output, "📄 Total Sheets: {}", workbook.sheets.len());
    output.push_str("\nSheet Names:\n");
    for (i, sheet) in workbook.sheets.iter().enumerate() {
        let _ = writeln!(output, "  {}. {}", i + 1, sheet.name);
    }

    let _ = writeln!(
        output,
        "\n🏷️  Named Ranges: {}",
        workbook.defined_names.len()
    );
    if !workbook.defined_names.is_empty() {
        output.push_str("\nNamed Range List:\n");
        for (i, (name, _formula)) in workbook.defined_names.iter().enumerate() {
            let _ = writeln!(output, "  {}. {}", i + 1, name);
        }
    }

    let _ = writeln!(
        output,
        "\n🔧 Contains Macros: {}",
        if workbook.has_macros { "Yes" } else { "No" }
    );

    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_io::reader::Sheet;
    use sheetrs_linter::ViolationScope;
    use sheetrs_linter::violation::CellReference;

    #[test]
    fn test_human_format_orders_cells_numerically() {
        let workbook = Workbook {
            sheets: vec![Sheet::new("S0".to_string(), 0)],
            ..Default::default()
        };
        let violations: Vec<Violation> = [(9, 0), (1, 0), (0, 1), (0, 26)]
            .iter()
            .map(|&(row, col)| {
                Violation::new(
                    sheetrs_linter::RuleId::Cpx509,
                    ViolationScope::Cell(0, CellReference::new(row, col)),
                    "Long formula",
                    Severity::Warning,
                )
            })
            .collect();
        let output = format_violations_human(&violations, &HashMap::new(), &workbook);
        let positions: Vec<_> = ["Cell: B1", "Cell: A2", "Cell: A10", "Cell: AA1"]
            .iter()
            .map(|label| output.find(label).unwrap_or_else(|| panic!("{label}")))
            .collect();
        // Numeric order is by row, then column: B1, AA1 (row 0), A2, A10.
        assert!(positions[0] < positions[3]);
        assert!(positions[3] < positions[1]);
        assert!(positions[1] < positions[2]);
    }

    #[test]
    fn a_sheet_with_sheet_and_cell_violations_has_one_header() {
        let workbook = Workbook {
            sheets: vec![Sheet::new("Data".to_string(), 0)],
            ..Default::default()
        };
        let violations = vec![
            Violation::new(
                sheetrs_linter::RuleId::Cpx502,
                ViolationScope::Sheet(0),
                "Merged cells",
                Severity::Warning,
            ),
            Violation::new(
                sheetrs_linter::RuleId::Cpx509,
                ViolationScope::Cell(0, CellReference::new(0, 0)),
                "Long formula",
                Severity::Warning,
            ),
        ];
        let output = format_violations_human(&violations, &HashMap::new(), &workbook);

        assert_eq!(output.matches("Sheet: Data").count(), 1);
        let sheet_level = output.find("Merged cells").expect("sheet violation");
        let cell_level = output.find("Cell: A1").expect("cell violation");
        assert!(sheet_level < cell_level);
    }
}
