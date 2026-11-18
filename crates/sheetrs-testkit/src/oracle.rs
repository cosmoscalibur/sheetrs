//! Runs cases in `LibreOffice` through the Python UNO bridge.

use std::io::{ErrorKind, Write};
use std::process::{Command, ExitStatus, Stdio};

use serde::Deserialize;
use thiserror::Error;

use crate::case::{Case, CellInput, Observed};

const SCRIPT: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/oracle/oracle.py");
const EXPECTED_VERSION_PREFIX: &str = "26.8.";

/// A failure while running the oracle.
#[derive(Debug, Error)]
pub enum Error {
    #[error("oracle process I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("oracle JSON error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("the oracle script failed ({status}): {stderr}")]
    Script { status: ExitStatus, stderr: String },
    #[error("case {name:?} failed in the oracle: {message}")]
    Case { name: String, message: String },
    #[error("the oracle is LibreOffice {found}, expected {EXPECTED_VERSION_PREFIX}x")]
    Version { found: String },
    #[error("case {case:?}, sheet {sheet:?}, cell {cell}: input numbers must be finite")]
    NonFinite {
        case: String,
        sheet: String,
        cell: String,
    },
}

/// What one run of the oracle returned.
#[derive(Debug, Deserialize)]
pub struct Report {
    /// The `LibreOffice` version that produced the values.
    pub office: String,
    pub results: Vec<Outcome>,
}

/// The result of one case, in the order the cases were given.
#[derive(Debug, Deserialize)]
#[serde(untagged)]
pub enum Outcome {
    Observed(Observed),
    Failed { name: String, error: String },
}

impl Report {
    /// Returns the observed values of every case.
    ///
    /// # Errors
    ///
    /// [`Error::Case`] for the first case that failed inside the oracle.
    pub fn into_observed(self) -> Result<Vec<Observed>, Error> {
        self.results
            .into_iter()
            .map(|outcome| match outcome {
                Outcome::Observed(observed) => Ok(observed),
                Outcome::Failed { name, error } => Err(Error::Case {
                    name,
                    message: error,
                }),
            })
            .collect()
    }
}

/// Applies the cases to `LibreOffice` in a single process with a fresh user profile.
///
/// # Errors
///
/// [`Error::Io`] when the script cannot be started, [`Error::Script`] when it
/// exits with a failure and [`Error::Json`] when its output cannot be read.
pub fn run(cases: &[Case]) -> Result<Report, Error> {
    ensure_finite(cases)?;
    let input = serde_json::to_vec(&serde_json::json!({ "cases": cases }))?;
    let mut child = Command::new("python3")
        .arg(SCRIPT)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    let Some(mut stdin) = child.stdin.take() else {
        unreachable!("stdin is piped");
    };
    // A script that dies before reading its input breaks the pipe; its exit
    // status and stderr below explain why, so that write error is not reported.
    if let Err(error) = stdin.write_all(&input)
        && error.kind() != ErrorKind::BrokenPipe
    {
        return Err(error.into());
    }
    drop(stdin);

    let output = child.wait_with_output()?;
    if !output.status.success() {
        return Err(Error::Script {
            status: output.status,
            stderr: String::from_utf8_lossy(&output.stderr).into_owned(),
        });
    }
    let report: Report = serde_json::from_slice(&output.stdout)?;
    ensure_version(&report.office)?;
    Ok(report)
}

fn ensure_version(found: &str) -> Result<(), Error> {
    if found.starts_with(EXPECTED_VERSION_PREFIX) {
        Ok(())
    } else {
        Err(Error::Version {
            found: found.to_owned(),
        })
    }
}

fn ensure_finite(cases: &[Case]) -> Result<(), Error> {
    for case in cases {
        for sheet in &case.sheets {
            for (cell, content) in &sheet.cells {
                if matches!(content, CellInput::Number(number) if !number.is_finite()) {
                    return Err(Error::NonFinite {
                        case: case.name.clone(),
                        sheet: sheet.name.clone(),
                        cell: cell.clone(),
                    });
                }
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use serde_json::json;

    use crate::case::{CellInput, SheetSpec};

    use super::*;

    fn case_with(content: CellInput) -> Case {
        Case {
            name: "c".to_owned(),
            sheets: vec![SheetSpec {
                name: "S1".to_owned(),
                cells: BTreeMap::from([("A1".to_owned(), content)]),
            }],
            ops: vec![],
            reads: vec![],
        }
    }

    #[test]
    fn only_the_expected_office_version_is_accepted() {
        assert!(ensure_version("26.8.0.3").is_ok());
        assert!(matches!(
            ensure_version("25.2.1.2"),
            Err(Error::Version { .. })
        ));
    }

    #[test]
    fn non_finite_input_numbers_are_rejected_before_running() {
        assert!(ensure_finite(&[case_with(CellInput::Number(1.5))]).is_ok());
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(
                matches!(
                    ensure_finite(&[case_with(CellInput::Number(value))]),
                    Err(Error::NonFinite { .. })
                ),
                "{value}"
            );
        }
    }

    #[test]
    fn a_failed_outcome_deserializes_and_fails_into_observed() -> Result<(), serde_json::Error> {
        let report: Report = serde_json::from_value(json!({
            "office": "26.8.0.3",
            "results": [{"name": "x", "error": "boom"}]
        }))?;
        assert!(matches!(
            report.into_observed(),
            Err(Error::Case { name, message }) if name == "x" && message == "boom"
        ));
        Ok(())
    }
}
