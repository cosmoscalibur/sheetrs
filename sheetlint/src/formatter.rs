//! Output formatters for violations

use anyhow::Result;
use colored::*;
use sheetrs::{Severity, Violation, ViolationScope};
use std::collections::BTreeMap;
use std::path::Path;

/// Print violations in human-readable format with colors and hierarchy
pub fn print_human(file_path: &Path, violations: &[Violation]) {
    println!("{}", format!("Linting: {}", file_path.display()).bold());
    println!();

    if violations.is_empty() {
        println!("{}", "✓ No violations found!".green().bold());
        return;
    }

    // Group violations by scope for hierarchical display
    let mut book_violations = Vec::new();
    let mut sheet_violations: BTreeMap<String, Vec<&Violation>> = BTreeMap::new();
    let mut cell_violations: BTreeMap<String, BTreeMap<String, Vec<&Violation>>> = BTreeMap::new();

    for violation in violations {
        match &violation.scope {
            ViolationScope::Book => book_violations.push(violation),
            ViolationScope::Sheet(sheet) => {
                sheet_violations
                    .entry(sheet.clone())
                    .or_default()
                    .push(violation);
            }
            ViolationScope::Cell(sheet, cell_ref) => {
                cell_violations
                    .entry(sheet.clone())
                    .or_default()
                    .entry(cell_ref.to_string())
                    .or_default()
                    .push(violation);
            }
        }
    }

    // Print book-level violations
    if !book_violations.is_empty() {
        println!("{}", "Book-level violations:".bold().underline());
        for violation in book_violations {
            print_violation(violation, 1);
        }
        println!();
    }

    // Print sheet-level violations
    for (sheet_name, violations) in &sheet_violations {
        println!("{} {}", "Sheet:".bold(), sheet_name.cyan().bold());
        for violation in violations {
            print_violation(violation, 1);
        }
        println!();
    }

    // Print cell-level violations
    for (sheet_name, cells) in &cell_violations {
        println!("{} {}", "Sheet:".bold(), sheet_name.cyan().bold());
        for (cell_ref, violations) in cells {
            println!("  {} {}", "Cell:".bold(), cell_ref.yellow());
            for violation in violations {
                print_violation(violation, 2);
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

fn print_violation(violation: &Violation, indent: usize) {
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
        violation.rule_id.bright_black(),
        violation.message
    );
}

/// Print violations in JSON format
pub fn print_json(file_path: &Path, violations: &[Violation]) -> Result<()> {
    let output = serde_json::json!({
        "file": file_path.display().to_string(),
        "violations": violations,
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
