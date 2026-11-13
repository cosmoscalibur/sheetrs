//! The findings of a run as a SARIF 2.1.0 log (Static Analysis Results Interchange Format,
//! OASIS standard), for the tools that read it: GitHub code scanning, the editors, the CI.
//!
//! The log is written result by result, as [`crate::report`] resolves them: a tree of the whole
//! document costs about 4 KB a violation, and a large workbook has hundreds of thousands.
//!
//! # Mapping
//!
//! | SARIF | From |
//! |---|---|
//! | `tool.driver` | `sheetrs lint`, the version of this crate, the repository as `informationUri` |
//! | `tool.driver.rules[]` | every registered rule, in the order of its ID; `ruleIndex` points into it |
//! | rule `name`, `shortDescription` | the name of the rule |
//! | rule `helpUri` | the rule reference of the linter README |
//! | rule `defaultConfiguration.enabled` | whether the rule is active by default |
//! | rule `properties.category`, `tags` | the category of the rule |
//! | result `ruleId` | the ID of the rule (`CALC201`) |
//! | result `level` | `Error` as `error`, `Warning` as `warning`, `Info` as `note` |
//! | result `message.text` | the rendered message |
//! | `locations[0].physicalLocation.artifactLocation` | the workbook file, as a URI reference |
//! | `locations[0].physicalLocation.region` | `startLine` the row and `startColumn` the column of the cell, 1-based (see below) |
//! | `locations[0].logicalLocations[0]` | the cell (`kind` `cell`, `Data!B7`), the sheet (`sheet`) or the workbook (`workbook`) |
//! | result `properties` | `scope` (`book`, `sheet` or `cell`), `sheet` and `cell` when they apply |
//!
//! A workbook has no lines: the region of a cell is its place in the table, the row as the line
//! and the column as the column of the character (`B7` is line 7, column 2). A violation of a
//! sheet or of the whole workbook has no cell, and its region is the nominal first position
//! (line 1, column 1) so that the tools that require one (GitHub code scanning does) accept it;
//! `properties.scope` and the logical location say what it really is.
//!
//! The severity is not a property of a rule (CALC202 is an error or a warning according to what
//! the file saved), so the rule descriptors carry no `defaultConfiguration.level`. No
//! `partialFingerprints` are written: they are optional, and the tools compute their own.

use crate::config::LinterConfig;
use crate::rules::registry;
use crate::violation::{Severity, Violation, ViolationScope};
use serde::Serialize;
use serde::ser::SerializeSeq;
use sheetrs_io::reader::Workbook;
use sheetrs_io::reader::parser_utils::column_letters;
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::Write;

const SCHEMA: &str =
    "https://docs.oasis-open.org/sarif/sarif/v2.1.0/errata01/os/schemas/sarif-schema-2.1.0.json";
const HELP_URI: &str =
    "https://github.com/cosmoscalibur/sheetrs/blob/main/sheetrs-linter/README.md#rule-reference";

/// Writes the SARIF log of `violations` of the workbook `file` (the path as the caller gave it,
/// relative to where the log is read) to `out`. `config` is the configuration the linter ran with.
///
/// # Errors
///
/// A write error of `out`.
pub fn write_sarif(
    out: &mut impl Write,
    file: &str,
    violations: &[Violation],
    workbook: &Workbook,
    config: &LinterConfig,
) -> serde_json::Result<()> {
    let metadata = registry::rule_metadata(config);
    let index_of: HashMap<&'static str, usize> = metadata
        .iter()
        .enumerate()
        .map(|(index, rule)| (rule.id.as_str(), index))
        .collect();
    let uri = uri_reference(file);
    let log = Log {
        schema: SCHEMA,
        version: "2.1.0",
        runs: [Run {
            tool: Tool {
                driver: Driver {
                    name: "sheetrs lint",
                    version: env!("CARGO_PKG_VERSION"),
                    semantic_version: env!("CARGO_PKG_VERSION"),
                    information_uri: env!("CARGO_PKG_REPOSITORY"),
                    rules: metadata
                        .iter()
                        .map(|rule| Rule {
                            id: rule.id.as_str(),
                            name: &rule.name,
                            short_description: Text { text: &rule.name },
                            help_uri: HELP_URI,
                            default_configuration: RuleConfiguration {
                                enabled: rule.is_default,
                            },
                            properties: RuleProperties {
                                category: rule.category.as_str(),
                                tags: [rule.category.as_str()],
                            },
                        })
                        .collect(),
                },
            },
            artifacts: [Artifact {
                location: ArtifactLocation { uri: &uri },
            }],
            invocations: [Invocation {
                execution_successful: true,
            }],
            results: Results {
                violations,
                workbook,
                file,
                uri: &uri,
                index_of: &index_of,
            },
        }],
    };
    serde_json::to_writer_pretty(&mut *out, &log)?;
    writeln!(out).map_err(serde_json::Error::io)?;
    Ok(())
}

/// The path of a file as a URI reference: `/` between the segments, and the characters a URI
/// does not allow written as `%XX` of their UTF-8 bytes. An absolute path is a `file:` URI.
#[must_use]
pub fn uri_reference(path: &str) -> String {
    let path = path.replace('\\', "/");
    let (prefix, rest) = if let Some(rest) = path.strip_prefix('/') {
        ("file:///", rest)
    } else if path.as_bytes().get(1) == Some(&b':') && path.as_bytes()[0].is_ascii_alphabetic() {
        ("file:///", path.as_str())
    } else {
        ("", path.strip_prefix("./").unwrap_or(&path))
    };
    let mut uri = String::from(prefix);
    for byte in rest.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~' | b'/' | b':') {
            uri.push(char::from(byte));
        } else {
            let _ = write!(uri, "%{byte:02X}");
        }
    }
    uri
}

/// The SARIF level of a severity.
#[must_use]
pub const fn level(severity: Severity) -> &'static str {
    match severity {
        Severity::Error => "error",
        Severity::Warning => "warning",
        Severity::Info => "note",
    }
}

#[derive(Serialize)]
struct Log<'a> {
    #[serde(rename = "$schema")]
    schema: &'static str,
    version: &'static str,
    runs: [Run<'a>; 1],
}

#[derive(Serialize)]
struct Run<'a> {
    tool: Tool<'a>,
    artifacts: [Artifact<'a>; 1],
    invocations: [Invocation; 1],
    results: Results<'a>,
}

#[derive(Serialize)]
struct Tool<'a> {
    driver: Driver<'a>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Driver<'a> {
    name: &'static str,
    version: &'static str,
    semantic_version: &'static str,
    information_uri: &'static str,
    rules: Vec<Rule<'a>>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Rule<'a> {
    id: &'static str,
    name: &'a str,
    short_description: Text<'a>,
    help_uri: &'static str,
    default_configuration: RuleConfiguration,
    properties: RuleProperties,
}

#[derive(Serialize)]
struct Text<'a> {
    text: &'a str,
}

#[derive(Serialize)]
struct RuleConfiguration {
    enabled: bool,
}

#[derive(Serialize)]
struct RuleProperties {
    category: &'static str,
    tags: [&'static str; 1],
}

#[derive(Serialize)]
struct Artifact<'a> {
    location: ArtifactLocation<'a>,
}

#[derive(Serialize)]
struct ArtifactLocation<'a> {
    uri: &'a str,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Invocation {
    execution_successful: bool,
}

/// The results, resolved one at a time while they are written.
struct Results<'a> {
    violations: &'a [Violation],
    workbook: &'a Workbook,
    file: &'a str,
    uri: &'a str,
    index_of: &'a HashMap<&'static str, usize>,
}

impl Serialize for Results<'_> {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let ctx = crate::violation::FormatContext {
            workbook: self.workbook,
        };
        let mut seq = serializer.serialize_seq(Some(self.violations.len()))?;
        for violation in self.violations {
            let message = violation.format_message(&ctx);
            seq.serialize_element(&self.result(violation, &message))?;
        }
        seq.end()
    }
}

impl Results<'_> {
    fn result<'m>(&'m self, violation: &Violation, message: &'m str) -> SarifResult<'m> {
        let sheet = violation
            .scope
            .sheet_index()
            .and_then(|index| self.workbook.sheet_name_by_index(index));
        let (scope, cell, line, column, logical) = match &violation.scope {
            ViolationScope::Book => ("book", None, 1, 1, ("workbook", self.file.to_owned())),
            ViolationScope::Sheet(_) => (
                "sheet",
                None,
                1,
                1,
                ("sheet", sheet.unwrap_or_default().to_owned()),
            ),
            ViolationScope::Cell(_, at) => {
                let cell = format!("{}{}", column_letters(at.col), at.row + 1);
                let qualified = format!("{}!{cell}", sheet.unwrap_or_default());
                (
                    "cell",
                    Some(cell),
                    at.row + 1,
                    at.col + 1,
                    ("cell", qualified),
                )
            }
        };
        let name = match (scope, &cell) {
            ("cell", Some(cell)) => cell.clone(),
            _ => logical.1.clone(),
        };
        SarifResult {
            rule_id: violation.rule_id.as_str(),
            rule_index: self.index_of.get(violation.rule_id.as_str()).copied(),
            level: level(violation.severity),
            message: Text { text: message },
            locations: [Location {
                physical_location: PhysicalLocation {
                    artifact_location: ArtifactLocation { uri: self.uri },
                    region: Region {
                        start_line: line,
                        start_column: column,
                    },
                },
                logical_locations: [LogicalLocation {
                    name,
                    fully_qualified_name: logical.1,
                    kind: logical.0,
                }],
            }],
            properties: ResultProperties {
                scope,
                sheet: sheet.map(str::to_owned),
                cell,
            },
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SarifResult<'a> {
    rule_id: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    rule_index: Option<usize>,
    level: &'static str,
    message: Text<'a>,
    locations: [Location<'a>; 1],
    properties: ResultProperties,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Location<'a> {
    physical_location: PhysicalLocation<'a>,
    logical_locations: [LogicalLocation; 1],
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct PhysicalLocation<'a> {
    artifact_location: ArtifactLocation<'a>,
    region: Region,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Region {
    start_line: u32,
    start_column: u32,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct LogicalLocation {
    name: String,
    fully_qualified_name: String,
    kind: &'static str,
}

#[derive(Serialize)]
struct ResultProperties {
    scope: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    sheet: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cell: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::violation::{CellReference, RuleId};
    use sheetrs_io::reader::Sheet;

    fn workbook() -> Workbook {
        Workbook {
            sheets: vec![Sheet::new("My Data".to_string(), 0)],
            ..Default::default()
        }
    }

    fn log(violations: &[Violation], file: &str) -> serde_json::Value {
        let mut out = Vec::new();
        write_sarif(
            &mut out,
            file,
            violations,
            &workbook(),
            &LinterConfig::default(),
        )
        .unwrap();
        serde_json::from_slice(&out).unwrap()
    }

    #[test]
    fn an_empty_run_has_the_tool_the_rules_and_no_results() {
        let log = log(&[], "book.xlsx");
        assert_eq!(log["version"], "2.1.0");
        assert_eq!(log["$schema"], SCHEMA);
        let run = &log["runs"][0];
        assert_eq!(run["tool"]["driver"]["name"], "sheetrs lint");
        assert_eq!(run["results"].as_array().unwrap().len(), 0);
        assert_eq!(run["invocations"][0]["executionSuccessful"], true);
        assert_eq!(run["artifacts"][0]["location"]["uri"], "book.xlsx");
        let rules = run["tool"]["driver"]["rules"].as_array().unwrap();
        assert_eq!(
            rules.len(),
            registry::rule_metadata(&LinterConfig::default()).len()
        );
        let ids: Vec<&str> = rules.iter().map(|r| r["id"].as_str().unwrap()).collect();
        let mut sorted = ids.clone();
        sorted.sort_unstable();
        assert_eq!(ids, sorted, "ruleIndex follows the order of the IDs");
    }

    #[test]
    fn every_scope_is_a_result_with_its_level_and_location() {
        let violations = vec![
            Violation::new(RuleId::Cpx501, ViolationScope::Book, "many", Severity::Info),
            Violation::new(
                RuleId::Cpx502,
                ViolationScope::Sheet(0),
                "merged \"x\" é",
                Severity::Warning,
            ),
            Violation::new(
                RuleId::Cpx509,
                ViolationScope::Cell(0, CellReference::new(6, 27)),
                "long",
                Severity::Error,
            ),
        ];
        let log = log(&violations, "dir/My Book.xlsx");
        let results = log["runs"][0]["results"].as_array().unwrap();
        assert_eq!(results.len(), 3);
        let rules = log["runs"][0]["tool"]["driver"]["rules"]
            .as_array()
            .unwrap();

        for (result, rule, level) in [
            (&results[0], "CPX501", "note"),
            (&results[1], "CPX502", "warning"),
            (&results[2], "CPX509", "error"),
        ] {
            assert_eq!(result["ruleId"], rule);
            assert_eq!(result["level"], level);
            let index = usize::try_from(result["ruleIndex"].as_u64().unwrap()).unwrap();
            assert_eq!(rules[index]["id"], rule);
            assert_eq!(
                result["locations"][0]["physicalLocation"]["artifactLocation"]["uri"],
                "dir/My%20Book.xlsx"
            );
        }
        assert_eq!(results[1]["message"]["text"], "merged \"x\" é");

        let book = &results[0]["locations"][0];
        assert_eq!(book["logicalLocations"][0]["kind"], "workbook");
        assert_eq!(book["physicalLocation"]["region"]["startLine"], 1);
        assert_eq!(results[0]["properties"]["scope"], "book");

        let sheet = &results[1]["locations"][0];
        assert_eq!(sheet["logicalLocations"][0]["kind"], "sheet");
        assert_eq!(sheet["logicalLocations"][0]["name"], "My Data");
        assert_eq!(results[1]["properties"]["sheet"], "My Data");

        // Row 7 (0-based 6) and column AB (0-based 27): line 7, column 28.
        let cell = &results[2]["locations"][0];
        assert_eq!(cell["physicalLocation"]["region"]["startLine"], 7);
        assert_eq!(cell["physicalLocation"]["region"]["startColumn"], 28);
        assert_eq!(cell["logicalLocations"][0]["kind"], "cell");
        assert_eq!(cell["logicalLocations"][0]["name"], "AB7");
        assert_eq!(
            cell["logicalLocations"][0]["fullyQualifiedName"],
            "My Data!AB7"
        );
        assert_eq!(results[2]["properties"]["cell"], "AB7");
    }

    #[test]
    fn a_path_is_a_uri_reference() {
        assert_eq!(uri_reference("a/b.xlsx"), "a/b.xlsx");
        assert_eq!(uri_reference("./a b/é.ods"), "a%20b/%C3%A9.ods");
        assert_eq!(uri_reference("dir\\book.xlsx"), "dir/book.xlsx");
        assert_eq!(uri_reference("/tmp/book.xlsx"), "file:///tmp/book.xlsx");
        assert_eq!(
            uri_reference("C:\\books\\a.xlsx"),
            "file:///C:/books/a.xlsx"
        );
    }

    #[test]
    fn the_levels_of_the_severities() {
        assert_eq!(level(Severity::Error), "error");
        assert_eq!(level(Severity::Warning), "warning");
        assert_eq!(level(Severity::Info), "note");
    }
}
