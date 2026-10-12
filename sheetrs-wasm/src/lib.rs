use serde::Serialize;
use sheetrs::reader::read_workbook_from_reader;
use sheetrs::rules::registry;
use sheetrs::{FormatContext, Linter, LinterConfig, violation::Violation};
use sheetrs::{Severity, ViolationScope};
use std::collections::{BTreeMap, HashMap};
use std::fmt::Write as _;
use std::io::Cursor;
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

#[wasm_bindgen]
pub fn get_rules_definition() -> Result<JsValue, JsValue> {
    let config = LinterConfig::default();
    let metadata = registry::get_all_rule_metadata(&config);

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

#[wasm_bindgen]
pub fn lint_workbook(
    file_data: &[u8],
    extension: &str,
    config_toml: &str,
) -> Result<String, JsValue> {
    let cursor = Cursor::new(file_data);
    let workbook = read_workbook_from_reader(cursor, Some(extension))
        .map_err(|e| JsValue::from_str(&format!("Reader error: {e}")))?;

    let config = if config_toml.trim().is_empty() {
        LinterConfig::default()
    } else {
        LinterConfig::from_toml(config_toml)
            .map_err(|e| JsValue::from_str(&format!("Config error: {e}")))?
    };

    let linter = Linter::with_config(config.clone());
    let violations = linter
        .lint_workbook(&workbook)
        .map_err(|e| JsValue::from_str(&format!("Linter error: {e}")))?;

    // Create a map of rule_id -> rule_name for better output
    let metadata = registry::get_all_rule_metadata(&config);
    let rule_names: HashMap<String, String> = metadata
        .iter()
        .map(|m| (m.id.to_string(), m.name.clone()))
        .collect();

    Ok(format_violations_human(&violations, &rule_names, &workbook))
}

fn format_violations_human(
    violations: &[Violation],
    rule_names: &HashMap<String, String>,
    workbook: &sheetrs::reader::Workbook,
) -> String {
    let mut output = String::new();

    if violations.is_empty() {
        output.push_str("✓ No violations found!\n");
        return output;
    }

    // Group violations by scope for hierarchical display
    let mut book_violations = Vec::new();
    let mut sheet_violations: BTreeMap<String, Vec<&Violation>> = BTreeMap::new();
    let mut cell_violations: BTreeMap<String, BTreeMap<String, Vec<&Violation>>> = BTreeMap::new();

    for violation in violations {
        match &violation.scope {
            ViolationScope::Book => book_violations.push(violation),
            ViolationScope::Sheet(sheet_idx) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                sheet_violations
                    .entry(sheet_name)
                    .or_default()
                    .push(violation);
            }
            ViolationScope::Cell(sheet_idx, cell_ref) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                cell_violations
                    .entry(sheet_name)
                    .or_default()
                    .entry(cell_ref.to_string())
                    .or_default()
                    .push(violation);
            }
        }
    }

    let ctx = FormatContext { workbook };

    // Print book-level violations
    if !book_violations.is_empty() {
        output.push_str("📚 Book-level violations:\n");
        for violation in book_violations {
            output.push_str(&format_violation(violation, rule_names, 1, &ctx));
        }
        output.push('\n');
    }

    // Print sheet-level violations
    for (sheet_name, violations) in &sheet_violations {
        let _ = writeln!(output, "📄 Sheet: {sheet_name}");
        for violation in violations {
            output.push_str(&format_violation(violation, rule_names, 1, &ctx));
        }
        output.push('\n');
    }

    // Print cell-level violations
    for (sheet_name, cells) in &cell_violations {
        let _ = writeln!(output, "📄 Sheet: {sheet_name}");
        for (cell_ref, violations) in cells {
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
        .map(|s| s.as_str())
        .unwrap_or("Unknown rule");

    format!(
        "{}{} [{}] {} - {}\n",
        indent_str,
        severity_icon,
        rule_id_str,
        rule_name,
        violation.format_message(ctx)
    )
}

#[wasm_bindgen]
pub fn get_workbook_stats(file_data: &[u8], extension: &str) -> Result<String, JsValue> {
    let cursor = Cursor::new(file_data);
    let workbook = read_workbook_from_reader(cursor, Some(extension))
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
