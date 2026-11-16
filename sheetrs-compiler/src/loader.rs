//! Frontend for XLSX/ODS files, built on the `sheetrs-io` readers.
//!
//! `sheetrs-io` already resolves the container formats (shared strings, shared
//! formulas, ODS table repetition). The readers give each file's own formula syntax; the loader
//! reads through the analysis form (`Book::analysis_into`, the Excel notation: ODS references
//! and argument separators converted), so the only format knowledge left here is the syntax of
//! function names.

use crate::cellstore::CellStore;
use crate::error::{Error, Result};
use crate::model::{Model, ModelCell, ModelSettings};
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::options::{Iteration, Search, SearchMode};
use sheetrs_engine::value::{ErrorKind, Value};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, DocumentSettings, Format, Workbook};
use std::fmt::Write as _;
use std::path::Path;

/// The model of the workbook `source` refers to: its cells go straight into the model as the reader
/// finds them, with no workbook of cells built first.
///
/// # Errors
///
/// When the workbook cannot be read.
pub fn load(source: &Book) -> Result<Model> {
    let ods = source.format() == Format::Ods;
    let mut sink = ModelSink::default();
    // What the file carries (settings, names) comes with the workbook the reader returns, whose
    // sheets are empty.
    let wb = source
        .analysis_into(&mut sink)
        .map_err(|error| Error::Workbook {
            what: source.to_string(),
            source: error,
        })?;
    // The file's format decides how it is read (settings the document carries, names,
    // reference syntax); what computes it is LibreOffice's behavior.
    let mut m = assemble(wb, ods, sink.cells);
    m.array_members = sink.array_members;
    m.from_xlsx = !ods;
    // The locale of the run: the one given to the workbook, else the system's (an ODS does
    // not carry the regional setting that matters: see `locale::resolve`).
    m.settings.book_locale = source.locale().map(ToString::to_string);
    m.settings.system_locale = sheetrs_io::locale::system().map(|tag| tag.to_string());
    m.file_name = source
        .path()
        .and_then(Path::file_name)
        .map(|n| n.to_string_lossy().into_owned());
    m.source = sheetrs_engine::ir::SourceInfo {
        file: m.file_name.clone().unwrap_or_default(),
        format: if ods { "ods" } else { "xlsx" }.to_string(),
        sha256: sha256_of(source)?,
    };
    Ok(m)
}

/// SHA-256 of the workbook, in hexadecimal.
fn sha256_of(source: &Book) -> Result<String> {
    use sha2::{Digest, Sha256};
    use std::io::Read;
    let mut hasher = Sha256::new();
    if let Some(path) = source.path() {
        let mut file = std::fs::File::open(path).map_err(|e| Error::io(path, e))?;
        let mut buf = vec![0u8; 64 * 1024];
        loop {
            let n = file.read(&mut buf).map_err(|e| Error::io(path, e))?;
            if n == 0 {
                break;
            }
            hasher.update(&buf[..n]);
        }
    } else if let Some(data) = source.bytes() {
        hasher.update(data);
    }
    Ok(hasher
        .finalize()
        .iter()
        .fold(String::new(), |mut hex, byte| {
            let _ = write!(hex, "{byte:02x}");
            hex
        }))
}

/// The model's cells, filled by the reader.
#[derive(Default)]
struct ModelSink {
    cells: CellStore,
    array_members: std::collections::HashMap<CellRef, (u32, u32)>,
}

impl sheetrs_io::reader::CellSink for ModelSink {
    fn hint(&mut self, sheet: u16, cells: usize) {
        self.cells.reserve(sheet, cells);
    }

    fn array_member(&mut self, sheet: u16, at: (u32, u32), offset: (u32, u32)) {
        self.array_members
            .insert(CellRef::new(sheet, at.0, at.1), offset);
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: sheetrs_io::reader::Cell) {
        if let Some(c) = model_cell(&cell.value, cell.formula, cell.is_array) {
            self.cells.insert(CellRef::new(sheet, at.0, at.1), c);
        }
    }
}

/// What the model keeps of a cell: nothing for one that holds no value and no formula.
fn model_cell(value: &CellValue, formula: Option<Box<str>>, array: bool) -> Option<ModelCell> {
    let value = convert(value);
    let formula = formula.map(String::from);
    (formula.is_some() || value != Value::Empty).then_some(ModelCell {
        value,
        formula,
        array,
    })
}

/// Convert a workbook that holds its cells, consuming it: values and formula text move
/// instead of being copied, and each sheet's cells are released as soon as they are converted.
#[must_use]
pub fn from_workbook_owned(mut wb: Workbook, format: Format) -> Model {
    let mut cells = CellStore::default();
    for (i, sheet) in wb.sheets.iter_mut().enumerate() {
        // In row order, so that each cell is an append.
        let sheet_cells = std::mem::take(&mut sheet.cells);
        cells.reserve(i as u16, sheet_cells.len());
        for ((row, col), cell) in sheet_cells {
            if let Some(c) = model_cell(&cell.value, cell.formula, cell.is_array) {
                cells.insert(CellRef::new(i as u16, row, col), c);
            }
        }
    }
    assemble(wb, format == Format::Ods, cells)
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
    m.properties.clone_from(&wb.properties);
    // The date system is the document's, one of the three of LibreOffice: the reader said which (an
    // ODS with another null date is an error of the reader, before the compiler sees the workbook).
    m.date_system = wb.date_system;
    m.settings = settings(&wb.settings, ods, &mut m.unsupported_options);
    m.names = wb
        .defined_names
        .iter()
        .map(|(k, v)| (k.clone(), if ods { ods_name(v) } else { v.clone() }))
        .collect();
    m
}

/// The model of a workbook that holds its cells, of the given format, leaving the workbook as it is.
#[must_use]
pub fn from_workbook(wb: &Workbook, format: Format) -> Model {
    let mut m = header(wb, format == Format::Ods);
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
/// expressions on, wildcards off); `LibreOffice` itself omits an attribute that
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
        book_locale: None,
        system_locale: None,
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

/// An error of the workbook as the engine has it: the Excel errors, the circular reference of
/// `LibreOffice` (`Err:522`) and the other codes only `LibreOffice` has (`Err:502`, `Err:523`, …),
/// which keep their code.
fn error_kind(s: &str) -> ErrorKind {
    ErrorKind::parse(s).unwrap_or(ErrorKind::Value)
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
        if sheetrs_engine::addr::parse_area(cell).is_none() {
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
        assert_eq!(error_kind("Err:999"), ErrorKind::Libre(999));
        assert_eq!(error_kind("Err:502"), ErrorKind::Libre(502));
        assert_eq!(error_kind("garbage"), ErrorKind::Value);
    }
}
