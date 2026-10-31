//! A flat, serializable form of the linter's findings, for consumers that are not Rust: the
//! language bindings (`sheetrs-py`, `sheetrs-wasm`) and any caller that wants JSON-like data
//! without rendering the human report.

use crate::LinterConfig;
use crate::reader::Workbook;
use crate::reader::parser_utils::column_letters;
use crate::rules::registry;
use crate::violation::{FormatContext, Severity, Violation, ViolationScope};
use serde::Serialize;
use std::collections::HashMap;

/// One violation with everything resolved: the rule's name, the sheet's name, the cell in A1
/// notation and the rendered message.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ViolationRecord {
    pub rule_id: String,
    pub rule_name: String,
    /// `"error"`, `"warning"` or `"info"`.
    pub severity: &'static str,
    pub message: String,
    /// `"book"`, `"sheet"` or `"cell"`.
    pub scope: &'static str,
    /// Name of the sheet; `None` for a book-level violation.
    pub sheet: Option<String>,
    /// Cell in A1 notation (`B7`); `None` unless the scope is `"cell"`.
    pub cell: Option<String>,
    /// 0-based row of the cell.
    pub row: Option<u32>,
    /// 0-based column of the cell.
    pub col: Option<u32>,
}

/// Name of a severity as it appears in [`ViolationRecord::severity`].
pub const fn severity_name(severity: Severity) -> &'static str {
    match severity {
        Severity::Error => "error",
        Severity::Warning => "warning",
        Severity::Info => "info",
    }
}

/// Resolves `violations` of `workbook` into records, keeping their order. `config` is the
/// configuration the linter ran with (rule names do not depend on it, but rules are built from it).
pub fn violation_records(
    violations: &[Violation],
    workbook: &Workbook,
    config: &LinterConfig,
) -> Vec<ViolationRecord> {
    let rule_names: HashMap<&'static str, String> = registry::get_all_rule_metadata(config)
        .into_iter()
        .map(|meta| (meta.id.as_str(), meta.name))
        .collect();
    let ctx = FormatContext { workbook };
    let sheet_name = |index: u16| workbook.sheet_name_by_index(index).map(str::to_owned);

    violations
        .iter()
        .map(|violation| {
            let rule_id = violation.rule_id.as_str();
            let (scope, sheet, cell, row, col) = match &violation.scope {
                ViolationScope::Book => ("book", None, None, None, None),
                ViolationScope::Sheet(index) => ("sheet", sheet_name(*index), None, None, None),
                ViolationScope::Cell(index, at) => (
                    "cell",
                    sheet_name(*index),
                    Some(format!("{}{}", column_letters(at.col), at.row + 1)),
                    Some(at.row),
                    Some(at.col),
                ),
            };
            ViolationRecord {
                rule_id: rule_id.to_owned(),
                rule_name: rule_names.get(rule_id).cloned().unwrap_or_default(),
                severity: severity_name(violation.severity),
                message: violation.format_message(&ctx),
                scope,
                sheet,
                cell,
                row,
                col,
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::Sheet;
    use crate::violation::{CellReference, RuleId};

    fn workbook() -> Workbook {
        Workbook {
            sheets: vec![Sheet::new("Data".to_string(), 0)],
            ..Default::default()
        }
    }

    #[test]
    fn every_scope_is_resolved() {
        let violations = vec![
            Violation::new(RuleId::Cpx501, ViolationScope::Book, "many", Severity::Info),
            Violation::new(
                RuleId::Cpx502,
                ViolationScope::Sheet(0),
                "merged",
                Severity::Warning,
            ),
            Violation::new(
                RuleId::Cpx509,
                ViolationScope::Cell(0, CellReference::new(6, 27)),
                "long",
                Severity::Error,
            ),
        ];
        let records = violation_records(&violations, &workbook(), &LinterConfig::default());

        assert_eq!(records.len(), 3);
        assert_eq!(
            (
                records[0].scope,
                records[0].sheet.as_deref(),
                records[0].severity
            ),
            ("book", None, "info")
        );
        assert_eq!(
            (records[1].scope, records[1].sheet.as_deref()),
            ("sheet", Some("Data"))
        );
        let cell = &records[2];
        assert_eq!(cell.cell.as_deref(), Some("AB7"));
        assert_eq!((cell.row, cell.col), (Some(6), Some(27)));
        assert_eq!(cell.severity, "error");
        assert_eq!(cell.rule_id, "CPX509");
        assert!(!cell.rule_name.is_empty());
        assert_eq!(cell.message, "long");
    }

    #[test]
    fn a_sheet_that_does_not_exist_has_no_name() {
        let violations = vec![Violation::new(
            RuleId::Cpx502,
            ViolationScope::Sheet(9),
            "merged",
            Severity::Warning,
        )];
        let records = violation_records(&violations, &workbook(), &LinterConfig::default());
        assert_eq!(records[0].sheet, None);
    }

    #[test]
    fn records_serialize_to_json_with_null_for_the_missing_parts() {
        let violations = vec![Violation::new(
            RuleId::Cpx501,
            ViolationScope::Book,
            "many",
            Severity::Info,
        )];
        let records = violation_records(&violations, &workbook(), &LinterConfig::default());
        let json = serde_json::to_value(&records).unwrap();
        assert_eq!(json[0]["sheet"], serde_json::Value::Null);
        assert_eq!(json[0]["scope"], "book");
    }
}
