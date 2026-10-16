//! Output formatters for violations

use anyhow::Result;
use colored::*;
use sheetrs::violation::CellReference;
use sheetrs::{FormatContext, Severity, Violation, ViolationScope, reader::Workbook};
use std::collections::BTreeMap;
use std::path::Path;

/// Print violations in human-readable format with colors and hierarchy
pub fn print_human(file_path: &Path, violations: &[Violation], workbook: &Workbook) {
    println!("{}", format!("Linting: {}", file_path.display()).bold());
    println!();

    if violations.is_empty() {
        println!("{}", "✓ No violations found!".green().bold());
        return;
    }

    let ctx = FormatContext { workbook };

    // Group violations by scope for hierarchical display
    let mut book_violations = Vec::new();
    // BTreeMap<SheetName, (SheetLevelViolations, BTreeMap<CellRef, CellLevelViolations>)>
    #[allow(clippy::type_complexity)]
    let mut sheet_data: BTreeMap<
        String,
        (Vec<&Violation>, BTreeMap<CellReference, Vec<&Violation>>),
    > = BTreeMap::new();

    for violation in violations {
        match &violation.scope {
            ViolationScope::Book => book_violations.push(violation),
            ViolationScope::Sheet(sheet_idx) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                let (sheet_v, _) = sheet_data.entry(sheet_name).or_default();
                sheet_v.push(violation);
            }
            ViolationScope::Cell(sheet_idx, cell_ref) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                let (_, cell_v) = sheet_data.entry(sheet_name).or_default();
                cell_v.entry(cell_ref.clone()).or_default().push(violation);
            }
        }
    }

    // Print book-level violations
    if !book_violations.is_empty() {
        println!("{}", "Book-level violations:".bold().underline());
        for violation in book_violations {
            print_violation(violation, 1, &ctx);
        }
        println!();
    }

    // Print sheet violations (consolidated)
    for (sheet_name, (sheet_violations, cell_violations)) in &sheet_data {
        println!("{} {}", "Sheet:".bold(), sheet_name.cyan().bold());

        // Print sheet-level violations first
        for violation in sheet_violations {
            print_violation(violation, 1, &ctx);
        }

        // Print cell-level violations
        for (cell_ref, violations) in cell_violations {
            println!("  {} {}", "Cell:".bold(), cell_ref.to_string().yellow());
            for violation in violations {
                print_violation(violation, 2, &ctx);
            }
        }
        println!();
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

    println!("{}", "Summary:".bold().underline());
    if error_count > 0 {
        println!("  {} {}", "Errors:".red().bold(), error_count);
    }
    if warning_count > 0 {
        println!("  {} {}", "Warnings:".yellow().bold(), warning_count);
    }
    if info_count > 0 {
        println!("  {} {}", "Info:".blue().bold(), info_count);
    }
}

fn print_violation(violation: &Violation, indent: usize, ctx: &FormatContext<'_>) {
    let indent_str = "  ".repeat(indent);
    let severity_str = match violation.severity {
        Severity::Error => "ERROR".red().bold(),
        Severity::Warning => "WARN".yellow().bold(),
        Severity::Info => "INFO".blue().bold(),
    };

    println!(
        "{}{} [{}] {}",
        indent_str,
        severity_str,
        violation.rule_id.as_str().bright_black(),
        violation.format_message(ctx)
    );
}

/// Print violations in JSON format
pub fn print_json(file_path: &Path, violations: &[Violation], workbook: &Workbook) -> Result<()> {
    let ctx = FormatContext { workbook };

    // Pre-render messages for JSON output so data-backed violations
    // produce the same human-readable strings as the legacy path.
    let rendered: Vec<serde_json::Value> = violations
        .iter()
        .map(|v| {
            serde_json::json!({
                "rule_id": v.rule_id,
                "scope": v.scope,
                "message": v.format_message(&ctx),
                "severity": v.severity,
            })
        })
        .collect();

    let output = serde_json::json!({
        "file": file_path.display().to_string(),
        "violations": rendered,
        "summary": {
            "total": violations.len(),
            "errors": violations.iter().filter(|v| v.severity == Severity::Error).count(),
            "warnings": violations.iter().filter(|v| v.severity == Severity::Warning).count(),
            "info": violations.iter().filter(|v| v.severity == Severity::Info).count(),
        }
    });

    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}
