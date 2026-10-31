//! Output formatters for violations

use anyhow::Result;
use colored::*;
use serde::Serialize;
use sheetrs_core::violation::CellReference;
use sheetrs_core::{FormatContext, RuleId, Severity, Violation, ViolationScope, reader::Workbook};
use std::collections::BTreeMap;
use std::io::{BufWriter, Write};
use std::path::Path;

/// The violations reported in one sheet: those of the sheet itself and those of each of its cells.
#[derive(Default)]
struct SheetViolations<'a> {
    sheet: Vec<&'a Violation>,
    cells: BTreeMap<CellReference, Vec<&'a Violation>>,
}

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
    let mut sheet_data: BTreeMap<String, SheetViolations> = BTreeMap::new();

    for violation in violations {
        match &violation.scope {
            ViolationScope::Book => book_violations.push(violation),
            ViolationScope::Sheet(sheet_idx) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                sheet_data
                    .entry(sheet_name)
                    .or_default()
                    .sheet
                    .push(violation);
            }
            ViolationScope::Cell(sheet_idx, cell_ref) => {
                // Lookup sheet name from index
                let sheet_name = workbook
                    .sheet_name_by_index(*sheet_idx)
                    .unwrap_or("Unknown Sheet")
                    .to_string();
                sheet_data
                    .entry(sheet_name)
                    .or_default()
                    .cells
                    .entry(cell_ref.clone())
                    .or_default()
                    .push(violation);
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
    for (sheet_name, data) in &sheet_data {
        println!("{} {}", "Sheet:".bold(), sheet_name.cyan().bold());

        // Print sheet-level violations first
        for violation in &data.sheet {
            print_violation(violation, 1, &ctx);
        }

        // Print cell-level violations
        for (cell_ref, violations) in &data.cells {
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

/// Print violations in JSON format.
///
/// The document is written violation by violation: building it first (a `serde_json::Value`
/// tree and then its text) costs about 4 KB per violation, 2.2 GB for the 600,000 violations of a
/// large workbook. The keys are in alphabetical order, as the document has always been written.
pub fn print_json(file_path: &Path, violations: &[Violation], workbook: &Workbook) -> Result<()> {
    let ctx = FormatContext { workbook };
    let mut out = BufWriter::new(std::io::stdout().lock());
    write_json(&mut out, &file_path.display().to_string(), violations, &ctx)?;
    out.flush()?;
    Ok(())
}

fn write_json(
    out: &mut impl Write,
    file: &str,
    violations: &[Violation],
    ctx: &FormatContext<'_>,
) -> Result<()> {
    let count = |severity: Severity| violations.iter().filter(|v| v.severity == severity).count();
    let output = JsonReport {
        file,
        // Messages are rendered here, so data-backed violations produce the same human-readable
        // strings as the legacy path.
        summary: JsonSummary {
            errors: count(Severity::Error),
            info: count(Severity::Info),
            total: violations.len(),
            warnings: count(Severity::Warning),
        },
        violations: JsonViolations { violations, ctx },
    };
    serde_json::to_writer_pretty(&mut *out, &output)?;
    writeln!(out)?;
    Ok(())
}

#[derive(Serialize)]
struct JsonReport<'a> {
    file: &'a str,
    summary: JsonSummary,
    violations: JsonViolations<'a>,
}

#[derive(Serialize)]
struct JsonSummary {
    errors: usize,
    info: usize,
    total: usize,
    warnings: usize,
}

struct JsonViolations<'a> {
    violations: &'a [Violation],
    ctx: &'a FormatContext<'a>,
}

impl Serialize for JsonViolations<'_> {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.collect_seq(self.violations.iter().map(|v| JsonViolation {
            message: v.format_message(self.ctx),
            rule_id: &v.rule_id,
            scope: JsonScope(&v.scope),
            severity: &v.severity,
        }))
    }
}

#[derive(Serialize)]
struct JsonViolation<'a> {
    message: String,
    rule_id: &'a RuleId,
    scope: JsonScope<'a>,
    severity: &'a Severity,
}

/// A scope as the document has always had it: the externally tagged form of the enum, with the
/// keys of a cell in alphabetical order.
struct JsonScope<'a>(&'a ViolationScope);

impl Serialize for JsonScope<'_> {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        #[derive(Serialize)]
        struct Cell {
            col: u32,
            row: u32,
        }
        match self.0 {
            ViolationScope::Book => serializer.serialize_str("Book"),
            ViolationScope::Sheet(sheet) => {
                serializer.serialize_newtype_variant("ViolationScope", 1, "Sheet", sheet)
            }
            ViolationScope::Cell(sheet, at) => serializer.serialize_newtype_variant(
                "ViolationScope",
                2,
                "Cell",
                &(
                    sheet,
                    Cell {
                        col: at.col,
                        row: at.row,
                    },
                ),
            ),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_core::violation::CellReference;

    fn json(violations: &[Violation]) -> String {
        let workbook = Workbook::default();
        let ctx = FormatContext {
            workbook: &workbook,
        };
        let mut out = Vec::new();
        write_json(&mut out, "book.xlsx", violations, &ctx).unwrap();
        String::from_utf8(out).unwrap()
    }

    #[test]
    fn an_empty_report() {
        assert_eq!(
            json(&[]),
            r#"{
  "file": "book.xlsx",
  "summary": {
    "errors": 0,
    "info": 0,
    "total": 0,
    "warnings": 0
  },
  "violations": []
}
"#
        );
    }

    /// The keys are in alphabetical order whatever the features of `serde_json` the build gets
    /// (the document used to be a `Value` tree, whose order follows them).
    #[test]
    fn every_scope_with_its_keys_in_alphabetical_order() {
        let violations = vec![
            Violation::new(
                RuleId::Calc201,
                ViolationScope::Book,
                "book",
                Severity::Info,
            ),
            Violation::new(
                RuleId::Calc201,
                ViolationScope::Sheet(3),
                "sheet",
                Severity::Warning,
            ),
            Violation::new(
                RuleId::Calc201,
                ViolationScope::Cell(2, CellReference::new(7, 4)),
                "cell \"quoted\" é",
                Severity::Error,
            ),
        ];
        let text = json(&violations);
        let value: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(value["summary"]["total"], 3);
        assert_eq!(value["violations"][1]["scope"]["Sheet"], 3);
        assert_eq!(value["violations"][2]["scope"]["Cell"][1]["row"], 7);
        assert_eq!(value["violations"][2]["message"], "cell \"quoted\" é");
        let at = |needle: &str| text.find(needle).unwrap();
        assert!(at("\"message\"") < at("\"rule_id\"") && at("\"rule_id\"") < at("\"scope\""));
        assert!(at("\"scope\"") < at("\"severity\""));
        assert!(at("\"col\": 4") < at("\"row\": 7"));
        assert!(text.contains("\"scope\": \"Book\""));
    }
}
