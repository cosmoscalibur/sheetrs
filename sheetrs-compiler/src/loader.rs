//! Frontend for XLSX/ODS files, built on the `sheetrs-core` readers.
//!
//! `sheetrs-core` already resolves the container formats (shared strings, shared
//! formulas, ODS table repetition) and normalizes ODS references and argument
//! separators to the Excel syntax, so the only format knowledge left here is
//! the syntax of function names.

use crate::addr::CellRef;
use crate::cellstore::CellStore;
use crate::model::{Model, ModelCell, ModelSettings};
use crate::options::{DateSystem, Iteration, Search, SearchMode};
use crate::value::{ErrorKind, Value};
use sheetrs_core::reader::{CellValue, DocumentSettings, Workbook};
use std::path::Path;

pub fn load(path: &Path) -> Result<Model, String> {
    let ods = path
        .extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("ods"));
    // The cells go straight into the model's map as the reader finds them: no workbook of
    // cells is built first. What the file carries (settings, names) comes with the workbook
    // the reader returns, whose sheets are empty.
    let mut sink = ModelSink::default();
    let wb = sheetrs_core::reader::read_workbook_into(path, &mut sink)
        .map_err(|e| format!("{}: {e:#}", path.display()))?;
    // The file's format decides how it is read (settings the document carries, names,
    // reference syntax); what computes it is LibreOffice's behavior.
    let mut m = assemble(wb, ods, sink.cells);
    m.from_xlsx = !ods;
    m.file_name = path.file_name().map(|n| n.to_string_lossy().into_owned());
    m.source = sheetrs_engine::ir::SourceInfo {
        file: m.file_name.clone().unwrap_or_default(),
        format: if ods { "ods" } else { "xlsx" }.to_string(),
        sha256: sha256_of(path).map_err(|e| format!("{}: {e}", path.display()))?,
    };
    Ok(m)
}

/// SHA-256 of the file, in hexadecimal.
fn sha256_of(path: &Path) -> std::io::Result<String> {
    use sha2::{Digest, Sha256};
    let mut hasher = Sha256::new();
    std::io::copy(&mut std::fs::File::open(path)?, &mut hasher)?;
    Ok(hasher
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect())
}

/// The model's cells, filled by the reader.
#[derive(Default)]
struct ModelSink {
    cells: CellStore,
}

impl sheetrs_core::reader::CellSink for ModelSink {
    fn hint(&mut self, sheet: u16, cells: usize) {
        self.cells.reserve(sheet, cells);
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: sheetrs_core::reader::Cell) {
        if let Some(c) = model_cell(cell.value, cell.formula, cell.is_array) {
            self.cells.insert(CellRef::new(sheet, at.0, at.1), c);
        }
    }
}

/// What the model keeps of a cell: nothing for one that holds no value and no formula.
fn model_cell(value: CellValue, formula: Option<Box<str>>, array: bool) -> Option<ModelCell> {
    let value = convert(&value);
    let formula = formula.map(String::from);
    (formula.is_some() || value != Value::Empty).then_some(ModelCell {
        value,
        formula,
        array,
    })
}

/// Convert a workbook that holds its cells, consuming it: values and formula text move
/// instead of being copied, and each sheet's cells are released as soon as they are converted.
pub fn from_workbook_owned(mut wb: Workbook, ods: bool) -> Model {
    let mut cells = CellStore::default();
    for (i, sheet) in wb.sheets.iter_mut().enumerate() {
        // In row order, so that each cell is an append.
        let sheet_cells = std::mem::take(&mut sheet.cells);
        cells.reserve(i as u16, sheet_cells.len());
        for ((row, col), cell) in sheet_cells {
            if let Some(c) = model_cell(cell.value, cell.formula, cell.is_array) {
                cells.insert(CellRef::new(i as u16, row, col), c);
            }
        }
    }
    assemble(wb, ods, cells)
}

/// The model of a workbook whose cells were read apart.
fn assemble(wb: Workbook, ods: bool, cells: CellStore) -> Model {
    let mut m = header(&wb, ods);
    m.cells = cells;
    for (i, sheet) in wb.sheets.into_iter().enumerate() {
        m.sheets.push(sheet.name);
        m.hidden_rows
            .extend(sheet.hidden_rows.iter().map(|&(a, b)| (i as u16, a, b)));
    }
    m
}

/// Workbook-level settings: date system, document settings, defined names.
fn header(wb: &Workbook, ods: bool) -> Model {
    let mut m = Model::new();
    m.from_file = true;
    m.macros = wb.has_macros;
    m.properties = wb.properties.clone();
    // The null date is the document's: an ODS stores any date, an XLSX says `date1904` or
    // gets LibreOffice's default, 1899-12-30, on the real calendar.
    m.date_system = match wb.settings.null_date {
        Some((y, mo, d)) if ods => DateSystem::from_ymd(y, i64::from(mo), i64::from(d))
            .unwrap_or_else(|| {
                m.diagnostics.push(format!(
                    "the null date {y}-{mo}-{d} is not a date; using 1899-12-30"
                ));
                DateSystem::NULL_1899
            }),
        _ if wb.date1904 => DateSystem::NULL_1904,
        _ => DateSystem::NULL_1899,
    };
    m.settings = settings(&wb.settings, ods, &mut m.unsupported_options);
    m.names = wb
        .defined_names
        .iter()
        .map(|(k, v)| (k.clone(), if ods { ods_name(v) } else { v.clone() }))
        .collect();
    m
}

pub fn from_workbook(wb: &Workbook, ods: bool) -> Model {
    let mut m = header(wb, ods);
    for (i, sheet) in wb.sheets.iter().enumerate() {
        m.sheets.push(sheet.name.clone());
        m.hidden_rows
            .extend(sheet.hidden_rows.iter().map(|&(a, b)| (i as u16, a, b)));
        for ((row, col), cell) in sheet.cells.iter() {
            let value = convert(&cell.value);
            let formula = cell.formula.as_deref().map(str::to_string);
            if formula.is_none() && value == Value::Empty {
                continue;
            }
            m.cells.insert(
                CellRef::new(i as u16, row, col),
                ModelCell {
                    value,
                    formula,
                    array: cell.is_array,
                },
            );
        }
    }
    m
}

/// Resolve the settings the document states into engine terms. ODS applies the
/// ODF defaults to what it leaves out (case sensitive, whole-cell, regular
/// expressions on, wildcards off); LibreOffice itself omits an attribute that
/// equals the default. Excel files carry no search settings and use Excel's
/// behavior.
fn settings(s: &DocumentSettings, ods: bool, unsupported: &mut Vec<String>) -> ModelSettings {
    let search = ods.then(|| Search {
        mode: if s.use_wildcards == Some(true) {
            SearchMode::Wildcards
        } else if s.use_regular_expressions.unwrap_or(true) {
            SearchMode::Regex
        } else {
            SearchMode::Literal
        },
        whole_cell: s.search_whole_cell.unwrap_or(true),
        case_sensitive: s.case_sensitive.unwrap_or(true),
    });
    if s.precision_as_shown == Some(true) {
        unsupported.push(
            "the document uses 'precision as shown' (values rounded to their displayed precision)"
                .into(),
        );
    }
    ModelSettings {
        iteration: s
            .iteration
            .as_ref()
            .filter(|i| i.enabled)
            .map(|i| Iteration {
                max_iterations: i.max_iterations,
                max_change: i.max_change,
            }),
        search,
        two_digit_year_start: s.null_year,
        // An XLSX import stores `SyntaxStringRef` 2 (Excel A1) in the ODS it is saved as;
        // a document created in LibreOffice stores nothing: Calc A1.
        string_ref_excel: Some(!ods || matches!(s.string_ref_syntax, Some(2 | 4))),
        locale: s.locale.clone(),
        manual_calc: s.auto_calculate == Some(false),
    }
}

fn convert(v: &CellValue) -> Value {
    match v {
        CellValue::Empty => Value::Empty,
        CellValue::Number(n) => Value::Number(*n),
        CellValue::Text(t) => Value::Text(t.clone()),
        CellValue::Boolean(b) => Value::Bool(*b),
        CellValue::Error(e) => Value::Error(error_kind(&e.to_string())),
    }
}

/// Excel error literal or LibreOffice `Err:NNN` code.
fn error_kind(s: &str) -> ErrorKind {
    if let Some(k) = ErrorKind::parse(s) {
        return k;
    }
    match s.strip_prefix("Err:").and_then(|c| c.parse::<u16>().ok()) {
        Some(522 | 523) => ErrorKind::Circ,
        Some(524) => ErrorKind::Ref,
        Some(525 | 508 | 509 | 510 | 511) => ErrorKind::Name,
        Some(502 | 503) => ErrorKind::Num,
        Some(532) => ErrorKind::DivZero,
        _ => ErrorKind::Value,
    }
}

/// ODS stores a defined range as `Sheet7.C3:Sheet7.C10` or
/// `$'My Sheet'.$A$1:.$B$2`; rewrite it as `Sheet7!C3:C10`. Anything that is
/// not a plain reference is returned unchanged.
fn ods_name(def: &str) -> String {
    let mut sheet: Option<String> = None;
    let mut cells = Vec::new();
    for part in def.split(':') {
        let part = part.strip_prefix('$').unwrap_or(part);
        let (name, cell) = if let Some(rest) = part.strip_prefix('\'') {
            match rest.split_once("'.") {
                Some((n, c)) => (Some(n.replace("''", "'")), c),
                None => return def.to_string(),
            }
        } else {
            match part.rsplit_once('.') {
                Some((n, c)) if !n.is_empty() => (Some(n.to_string()), c),
                Some((_, c)) => (None, c),
                None => return def.to_string(),
            }
        };
        if crate::addr::parse_area(cell).is_none() {
            return def.to_string();
        }
        sheet = sheet.or(name);
        cells.push(cell);
    }
    let area = cells.join(":");
    match sheet {
        Some(n) if n.chars().all(|c| c.is_alphanumeric() || c == '_') => format!("{n}!{area}"),
        Some(n) => format!("'{}'!{area}", n.replace('\'', "''")),
        None => area,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ods_defined_ranges() {
        assert_eq!(ods_name("Sheet7.C3:Sheet7.C10"), "Sheet7!C3:C10");
        assert_eq!(ods_name("$'My Sheet'.$A$1:.$B$2"), "'My Sheet'!$A$1:$B$2");
        assert_eq!(ods_name("$Sheet1.$A$1"), "Sheet1!$A$1");
        assert_eq!(ods_name("SUM(1;2)"), "SUM(1;2)");
    }

    #[test]
    fn libreoffice_error_codes() {
        assert_eq!(error_kind("Err:522"), ErrorKind::Circ);
        assert_eq!(error_kind("#N/A"), ErrorKind::Na);
        assert_eq!(error_kind("Err:999"), ErrorKind::Value);
    }
}
