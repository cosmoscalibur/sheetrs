//! The findings of a run as the JSON document of `sheetrs lint --format json`: the file, a summary
//! with the counts by severity and the violations with their rule, severity, scope and rendered
//! message. It is the document for an interface of your own; [`crate::sarif`] is the standard one.
//!
//! The document is written violation by violation: building it first (a `serde_json::Value`
//! tree and then its text) costs about 4 KB per violation, 2.2 GB for the 600,000 violations of a
//! large workbook. The keys are in alphabetical order, as the document has always been written.

use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationScope};
use serde::Serialize;
use sheetrs_io::reader::Workbook;
use std::io::Write;

/// Writes the JSON document of `violations` of the workbook `file` (the path as the caller gave
/// it) to `out`.
///
/// # Errors
///
/// A write error of `out`.
pub fn write_json(
    out: &mut impl Write,
    file: &str,
    violations: &[Violation],
    workbook: &Workbook,
) -> serde_json::Result<()> {
    let ctx = &FormatContext { workbook };
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
    writeln!(out).map_err(serde_json::Error::io)?;
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
    use crate::violation::CellReference;

    fn json(violations: &[Violation]) -> String {
        let workbook = Workbook::default();
        let mut out = Vec::new();
        write_json(&mut out, "book.xlsx", violations, &workbook).unwrap();
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
