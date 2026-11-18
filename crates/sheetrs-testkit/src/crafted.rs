//! Hand-written packages opened by the oracle and by a reader, cell by cell.

use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::fs;
use std::sync::{Mutex, PoisonError};

use sheetrs_core::{CellPos, CellRef, Visibility, Workbook};

use crate::case::{Case, Observed, Op, Read, SheetSpec, TextFilter};
use crate::compare::{Leniency, matches};
use crate::oracle;

/// A package with the range of one sheet to compare.
#[derive(Debug, Clone)]
pub struct Package {
    /// The case name that a difference is reported under.
    pub name: String,
    /// The bytes of the file.
    pub bytes: Vec<u8>,
    /// The sheet to read.
    pub sheet: String,
    /// The range of the sheet to read, such as `A1:B2`.
    pub range: String,
}

/// Serializes the runs against `LibreOffice`: each starts a process that the others would
/// disturb.
static ORACLE: Mutex<()> = Mutex::new(());

/// Opens each package with the oracle and with `read`, and fails with every cell of its sheet
/// where they differ, as far as `leniency` allows them to.
///
/// Files are written to `dir` as `{group}_{index}.{extension}`. A text `filter` is the import
/// filter the oracle opens them with.
///
/// # Errors
///
/// When the packages cannot be written or the oracle cannot run; differences are a panic of
/// the assertion.
///
/// # Panics
///
/// When a package is read differently by the oracle and by `read`.
pub fn assert_reads_as_the_oracle(
    dir: &str,
    group: &str,
    extension: &str,
    filter: Option<&TextFilter>,
    packages: &[Package],
    read: impl Fn(&str) -> Result<Workbook, Box<dyn StdError>>,
    leniency: impl Fn(&crate::case::CellValue, Option<&sheetrs_core::Cell>) -> Leniency,
) -> Result<(), Box<dyn StdError>> {
    let _guard = ORACLE.lock().unwrap_or_else(PoisonError::into_inner);
    let filter = filter.cloned();
    let path = |index: usize| format!("{dir}/{group}_{index}.{extension}");
    let mut cases = Vec::new();
    for (index, package) in packages.iter().enumerate() {
        fs::write(path(index), &package.bytes)?;
        cases.push(Case {
            name: package.name.clone(),
            sheets: vec![SheetSpec {
                name: "Dummy".to_owned(),
                cells: BTreeMap::new(),
            }],
            ops: vec![Op::Open {
                path: path(index),
                filter: filter.clone(),
            }],
            reads: vec![Read {
                sheet: package.sheet.clone(),
                range: package.range.clone(),
            }],
        });
    }
    let report = oracle::run(&cases)?;

    let mut differences = Vec::new();
    for (index, (outcome, package)) in report.results.iter().zip(packages).enumerate() {
        let name = &package.name;
        let ours = read(&path(index));
        let line = match (outcome, &ours) {
            (oracle::Outcome::Observed(observed), Ok(workbook)) => {
                let sheet = workbook.sheet(&package.sheet);
                let mut found = Vec::new();
                // The cells of the range are numbered from its first cell.
                let first = package
                    .range
                    .split(':')
                    .next()
                    .and_then(|corner| CellRef::from_a1(corner).ok())
                    .map_or((0, 0), |cell| (cell.row.index, cell.col.index));
                for (row, cells) in (first.0..).zip(&observed.reads[0].cells) {
                    for (col, expected) in (first.1..).zip(cells) {
                        let actual = sheet.and_then(|sheet| {
                            sheet.get(CellPos { row, col }).map(|view| view.to_cell())
                        });
                        let agrees = match leniency(expected, actual.as_ref()) {
                            Leniency::Mismatch => false,
                            Leniency::FormulaOptional => {
                                matches(expected, actual.as_ref(), true, true)
                            }
                            Leniency::Strict => matches(expected, actual.as_ref(), true, false),
                            Leniency::Values => matches(expected, actual.as_ref(), false, false),
                        };
                        if !agrees {
                            found.push(format!(
                                "{name} at {row},{col}: oracle {expected:?}, read {actual:?}"
                            ));
                        }
                    }
                }
                found.extend(workbook_differences(name, workbook, observed));
                (!found.is_empty()).then(|| found.join("\n"))
            }
            (oracle::Outcome::Observed(observed), Err(error)) => Some(format!(
                "{name}: oracle {:?}, read failed: {error}",
                observed.reads[0].cells.first().and_then(|row| row.first())
            )),
            (oracle::Outcome::Failed { error, .. }, Ok(_)) => {
                Some(format!("{name}: oracle failed ({error}), read succeeded"))
            }
            (oracle::Outcome::Failed { .. }, Err(_)) => None,
        };
        differences.extend(line);
    }
    assert!(differences.is_empty(), "\n{}", differences.join("\n"));
    Ok(())
}

/// How the oracle names the document at `target`: a Windows path is a `file` address, and
/// spaces are escaped.
fn oracle_name(target: &str) -> String {
    let mut characters = target.chars();
    let windows = matches!(
        (characters.next(), characters.next(), characters.next()),
        (Some(drive), Some(':'), Some('\\' | '/')) if drive.is_ascii_alphabetic()
    );
    let named = if windows {
        format!("file:///{}", target.replace('\\', "/"))
    } else {
        target.to_owned()
    };
    named.replace(' ', "%20")
}

/// Whether the documents the workbook refers to are those the oracle names, each document
/// matching a different name. The oracle names a document by its address, which for a path
/// relative to the workbook is the path as it finds it.
fn same_links(workbook: &Workbook, observed: &[String]) -> bool {
    let mut unused: Vec<&String> = observed.iter().collect();
    for link in workbook.external_links() {
        let target = oracle_name(&link.target);
        let suffix = format!("/{}", target.trim_start_matches("./"));
        let Some(position) = unused
            .iter()
            .position(|name| **name == target || name.ends_with(&suffix))
        else {
            return false;
        };
        unused.remove(position);
    }
    unused.is_empty()
}

/// What differs between the workbook a reader returned and what the oracle observed of the whole
/// workbook: its sheets, the macro modules it holds and the documents it refers to.
fn workbook_differences(name: &str, workbook: &Workbook, observed: &Observed) -> Vec<String> {
    let mut found = Vec::new();
    // The sheets keep their order and names, and a hidden sheet stays hidden.
    let ours: Vec<_> = workbook
        .sheets()
        .map(|sheet| {
            (
                sheet.name().to_owned(),
                sheet.visibility() == Visibility::Visible,
            )
        })
        .collect();
    let theirs: Vec<_> = observed
        .sheets
        .iter()
        .map(|sheet| (sheet.name.clone(), sheet.visible))
        .collect();
    if ours != theirs {
        found.push(format!("{name}: oracle sheets {theirs:?}, read {ours:?}"));
    }
    let mut modules: Vec<(String, String)> = workbook
        .macro_modules()
        .iter()
        .map(|module| (module.library.clone(), module.name.clone()))
        .collect();
    modules.sort();
    if modules != observed.basic_modules {
        found.push(format!(
            "{name}: oracle modules {:?}, read {modules:?}",
            observed.basic_modules
        ));
    }
    if !same_links(workbook, &observed.external_links) {
        found.push(format!(
            "{name}: oracle links {:?}, read {:?}",
            observed.external_links,
            workbook.external_links()
        ));
    }
    found
}
