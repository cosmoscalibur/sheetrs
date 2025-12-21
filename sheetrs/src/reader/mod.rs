//! Excel/ODS file reader using custom XML parsers

use anyhow::{Context, Result};

use std::collections::HashMap;
use std::fs::File;
use std::path::Path;
use zip::ZipArchive;

pub mod ods_parser;
pub mod parser_utils;
pub mod workbook;
pub mod xlsx_parser;

use self::ods_parser::OdsReader;
use self::xlsx_parser::XlsxReader;
pub use workbook::{Cell, CellValue, Sheet, Workbook};

/// Trait for spreadsheet format readers
pub trait WorkbookReader {
    fn read_sheets(&mut self) -> Result<Vec<Sheet>>;
    fn read_defined_names(&mut self) -> Result<HashMap<String, String>>;
    fn read_hidden_sheets(&mut self) -> Result<Vec<String>>;
    fn has_macros(&mut self) -> Result<bool>;
    fn read_external_links(&mut self) -> Result<Vec<String>>;
}

/// Read a workbook from a file path
pub fn read_workbook<P: AsRef<Path>>(path: P) -> Result<Workbook> {
    let path_ref = path.as_ref();
    let file = File::open(path_ref)
        .with_context(|| format!("Failed to open file: {}", path_ref.display()))?;
    let mut archive = ZipArchive::new(file).context("Failed to open zip archive")?;

    let is_xlsx = path_ref
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("xlsx") || s.eq_ignore_ascii_case("xlsm"))
        .unwrap_or(false);

    let is_ods = path_ref
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("ods"))
        .unwrap_or(false);

    let (sheets, defined_names, hidden_sheets, has_macros, external_links) = if is_xlsx {
        let mut reader = XlsxReader::new(&mut archive)?;
        (
            reader.read_sheets()?,
            reader.read_defined_names()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_links()?,
        )
    } else if is_ods {
        let mut reader = OdsReader::new(&mut archive)?;
        (
            reader.read_sheets()?,
            reader.read_defined_names()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_links()?,
        )
    } else {
        return Err(anyhow::anyhow!("Unsupported file format"));
    };

    Ok(Workbook {
        path: path_ref.to_path_buf(),
        sheets,
        defined_names,
        hidden_sheets,
        has_macros,
        external_links,
    })
}
