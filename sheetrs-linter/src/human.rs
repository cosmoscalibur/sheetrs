//! The findings of a run as the report of `sheetrs lint` for a person: the book, then each sheet
//! with its own violations and those of its cells, and a summary by severity.
//!
//! It is plain text unless `color` is asked for; with `color` the text carries the terminal
//! styles, which [the `colored` crate](https://docs.rs/colored) still drops when the output is not
//! a terminal or `NO_COLOR` is set.

use crate::report::group_by_scope;
use crate::violation::{FormatContext, Severity, Violation};
use colored::{ColoredString, Colorize};
use sheetrs_io::reader::Workbook;
use std::io::{self, Write};

/// Writes the report of `violations` of the workbook `file` (the path as the caller gave it) to
/// `out`, with the terminal styles when `color` is set.
///
/// # Errors
///
/// A write error of `out`.
pub fn write_human(
    out: &mut impl Write,
    file: &str,
    violations: &[Violation],
    workbook: &Workbook,
    color: bool,
) -> io::Result<()> {
    let paint = |text: &str, style: fn(ColoredString) -> ColoredString| -> String {
        if color {
            style(text.normal()).to_string()
        } else {
            text.to_string()
        }
    };
    let bold = |text: &str| paint(text, ColoredString::bold);
    writeln!(out, "{}", bold(&format!("Linting: {file}")))?;
    writeln!(out)?;

    if violations.is_empty() {
        writeln!(
            out,
            "{}",
            paint("✓ No violations found!", |s| s.green().bold())
        )?;
        return Ok(());
    }

    let ctx = FormatContext { workbook };
    let grouped = group_by_scope(violations, workbook);

    if !grouped.book.is_empty() {
        writeln!(
            out,
            "{}",
            paint("Book-level violations:", |s| s.bold().underline())
        )?;
        for violation in &grouped.book {
            write_violation(out, violation, 1, &ctx, color)?;
        }
        writeln!(out)?;
    }

    for (sheet_name, data) in &grouped.sheets {
        writeln!(
            out,
            "{} {}",
            bold("Sheet:"),
            paint(sheet_name, |s| s.cyan().bold())
        )?;
        for violation in &data.sheet {
            write_violation(out, violation, 1, &ctx, color)?;
        }
        for (cell_ref, violations) in &data.cells {
            writeln!(
                out,
                "  {} {}",
                bold("Cell:"),
                paint(&cell_ref.to_string(), ColoredString::yellow)
            )?;
            for violation in violations {
                write_violation(out, violation, 2, &ctx, color)?;
            }
        }
        writeln!(out)?;
    }

    let count = |severity: Severity| violations.iter().filter(|v| v.severity == severity).count();
    writeln!(out, "{}", paint("Summary:", |s| s.bold().underline()))?;
    for (severity, label, style) in [
        (
            Severity::Error,
            "Errors:",
            (|s: ColoredString| s.red().bold()) as fn(ColoredString) -> ColoredString,
        ),
        (Severity::Warning, "Warnings:", |s| s.yellow().bold()),
        (Severity::Info, "Info:", |s| s.blue().bold()),
    ] {
        let n = count(severity);
        if n > 0 {
            writeln!(out, "  {} {n}", paint(label, style))?;
        }
    }
    Ok(())
}

fn write_violation(
    out: &mut impl Write,
    violation: &Violation,
    indent: usize,
    ctx: &FormatContext<'_>,
    color: bool,
) -> io::Result<()> {
    let paint = |text: &str, style: fn(ColoredString) -> ColoredString| -> String {
        if color {
            style(text.normal()).to_string()
        } else {
            text.to_string()
        }
    };
    let severity = match violation.severity {
        Severity::Error => paint("ERROR", |s| s.red().bold()),
        Severity::Warning => paint("WARN", |s| s.yellow().bold()),
        Severity::Info => paint("INFO", |s| s.blue().bold()),
    };
    writeln!(
        out,
        "{}{} [{}] {}",
        "  ".repeat(indent),
        severity,
        paint(violation.rule_id.as_str(), ColoredString::bright_black),
        violation.format_message(ctx)
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::violation::{CellReference, RuleId, ViolationScope};
    use sheetrs_io::reader::Sheet;

    fn workbook() -> Workbook {
        Workbook {
            sheets: vec![Sheet::new("Data".to_string(), 0)],
            ..Default::default()
        }
    }

    fn render(violations: &[Violation]) -> String {
        let mut out = Vec::new();
        write_human(&mut out, "book.xlsx", violations, &workbook(), false).unwrap();
        String::from_utf8(out).unwrap()
    }

    #[test]
    fn no_violations() {
        assert_eq!(
            render(&[]),
            "Linting: book.xlsx\n\n✓ No violations found!\n"
        );
    }

    #[test]
    fn the_hierarchy_and_the_summary_without_styles() {
        let violations = vec![
            Violation::new(
                RuleId::Calc201,
                ViolationScope::Book,
                "book",
                Severity::Info,
            ),
            Violation::new(
                RuleId::Calc201,
                ViolationScope::Cell(0, CellReference::new(6, 1)),
                "cell",
                Severity::Error,
            ),
        ];
        let text = render(&violations);
        assert!(text.contains("Book-level violations:\n  INFO [CALC201] book\n"));
        assert!(text.contains("Sheet: Data\n  Cell: B7\n    ERROR [CALC201] cell\n"));
        assert!(text.contains("Summary:\n  Errors: 1\n  Info: 1\n"));
        assert!(!text.contains('\u{1b}'), "no escape codes without color");
    }

    #[test]
    fn cells_are_in_row_then_column_order_under_one_header_per_sheet() {
        let cell = |row, col, rule| {
            Violation::new(
                rule,
                ViolationScope::Cell(0, CellReference::new(row, col)),
                "cell",
                Severity::Warning,
            )
        };
        let mut violations: Vec<Violation> = [(9, 0), (1, 0), (0, 1), (0, 26)]
            .into_iter()
            .map(|(row, col)| cell(row, col, RuleId::Cpx509))
            .collect();
        violations.push(Violation::new(
            RuleId::Cpx502,
            ViolationScope::Sheet(0),
            "sheet",
            Severity::Warning,
        ));
        let text = render(&violations);
        assert_eq!(text.matches("Sheet: Data").count(), 1);
        let positions: Vec<usize> = ["Cell: B1", "Cell: AA1", "Cell: A2", "Cell: A10"]
            .iter()
            .map(|label| text.find(label).unwrap_or_else(|| panic!("{label}")))
            .collect();
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
        assert!(
            text.find("sheet").unwrap() < positions[0],
            "the sheet's own violations come first"
        );
    }
}
