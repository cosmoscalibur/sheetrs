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
pub use workbook::{Cell, CellValue, ExternalWorkbook, Sheet, Workbook};

/// Trait for spreadsheet format readers
pub trait WorkbookReader {
    fn read_sheets(&mut self) -> Result<Vec<Sheet>>;
    fn read_defined_names(&mut self) -> Result<HashMap<String, String>>;
    fn read_hidden_sheets(&mut self) -> Result<Vec<String>>;
    fn has_macros(&mut self) -> Result<bool>;
    fn read_external_links(&mut self) -> Result<Vec<String>>;
    /// Read external workbook references with indices
    fn read_external_workbooks(&mut self) -> Result<Vec<ExternalWorkbook>>;
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

    let (sheets, defined_names, hidden_sheets, has_macros, external_workbooks) = if is_xlsx {
        let mut reader = XlsxReader::new(&mut archive)?;
        (
            reader.read_sheets()?,
            reader.read_defined_names()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
        )
    } else if is_ods {
        let mut reader = OdsReader::new(&mut archive)?;
        let sheets = reader.read_sheets()?;
        let defined_names = reader.read_defined_names()?;
        (
            sheets,
            defined_names,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
        )
    } else {
        return Err(anyhow::anyhow!("Unsupported file format"));
    };

    // Maintain backward compatibility: external_links derived from external_workbooks
    let external_links = external_workbooks.iter().map(|e| e.path.clone()).collect();

    Ok(Workbook {
        path: path_ref.to_path_buf(),
        sheets,
        defined_names,
        hidden_sheets,
        has_macros,
        external_links,
        external_workbooks,
    })
}

#[cfg(test)]
mod date_format_parity_tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_date_format_parity_ods_xlsx() {
        // Truth values for verification
        const EXPECTED_D7_FORMAT: &str = "m/d/yyyy";
        const EXPECTED_D7_VALUE: f64 = 45139.0; // 2023/08/01
        const EXPECTED_D8_FORMAT: &str = "dd/mm/yy hh:mm";

        // Load ODS
        const TEST_ODS: &[u8] = include_bytes!("../../../tests/minimal_test.ods");
        let cursor_ods = Cursor::new(TEST_ODS);
        let mut archive_ods = ZipArchive::new(cursor_ods).unwrap();
        let mut reader_ods = OdsReader::new(&mut archive_ods).unwrap();
        let sheets_ods = reader_ods.read_sheets().unwrap();

        // Load XLSX
        const TEST_XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");
        let cursor_xlsx = Cursor::new(TEST_XLSX);
        let mut archive_xlsx = ZipArchive::new(cursor_xlsx).unwrap();
        let mut reader_xlsx = XlsxReader::new(&mut archive_xlsx).unwrap();
        let sheets_xlsx = reader_xlsx.read_sheets().unwrap();

        // Find "Indexing tests" sheet in both formats
        let sheet_ods = sheets_ods
            .iter()
            .find(|s| s.name == "Indexing tests")
            .expect("Should have 'Indexing tests' sheet in ODS");
        let sheet_xlsx = sheets_xlsx
            .iter()
            .find(|s| s.name == "Indexing tests")
            .expect("Should have 'Indexing tests' sheet in XLSX");

        // STEP 1: Verify ODS against truth values

        // Verify expected date cells exist in ODS
        assert!(
            sheet_ods.cells.contains_key(&(6, 3)),
            "ODS should have cell D7"
        );
        assert!(
            sheet_ods.cells.contains_key(&(7, 3)),
            "ODS should have cell D8"
        );

        // Verify D7 truth values (Indexing tests)
        let d7_ods = sheet_ods.cells.get(&(6, 3)).unwrap();
        assert_eq!(
            d7_ods.num_fmt.as_ref().unwrap(),
            EXPECTED_D7_FORMAT,
            "ODS D7 format should match truth value"
        );
        if let CellValue::Number(val) = d7_ods.value {
            assert!(
                (val - EXPECTED_D7_VALUE).abs() < 0.1,
                "ODS D7 value should be ~{}, got {}",
                EXPECTED_D7_VALUE,
                val
            );
        }

        // Verify D8 truth values (Indexing tests) - THE FIX TARGET
        let d8_ods = sheet_ods.cells.get(&(7, 3)).unwrap();
        assert_eq!(
            d8_ods.num_fmt.as_ref().unwrap(),
            EXPECTED_D8_FORMAT,
            "ODS D8 format should be 'dd/mm/yy hh:mm' (two-digit), not 'd/m/yy hh:mm'"
        );
        // Verify it's a formula or numeric value
        assert!(
            matches!(
                d8_ods.value,
                CellValue::Formula { .. } | CellValue::Number(_)
            ),
            "ODS D8 should be formula or number"
        );

        // STEP 2: Verify XLSX against truth values

        // Verify expected date cells exist in XLSX
        assert!(
            sheet_xlsx.cells.contains_key(&(6, 3)),
            "XLSX should have cell D7"
        );
        assert!(
            sheet_xlsx.cells.contains_key(&(7, 3)),
            "XLSX should have cell D8"
        );

        // Verify D7 truth values
        let d7_xlsx = sheet_xlsx.cells.get(&(6, 3)).unwrap();
        assert_eq!(
            d7_xlsx.num_fmt.as_ref().unwrap(),
            EXPECTED_D7_FORMAT,
            "XLSX D7 format should match truth value"
        );
        if let CellValue::Number(val) = d7_xlsx.value {
            assert!(
                (val - EXPECTED_D7_VALUE).abs() < 0.1,
                "XLSX D7 value should be ~{}, got {}",
                EXPECTED_D7_VALUE,
                val
            );
        }

        // Verify D8 truth values
        let d8_xlsx = sheet_xlsx.cells.get(&(7, 3)).unwrap();
        assert_eq!(
            d8_xlsx.num_fmt.as_ref().unwrap(),
            EXPECTED_D8_FORMAT,
            "XLSX D8 format should match truth value"
        );

        // STEP 3: Verify ODS == XLSX parity

        // Collect all date cells from both formats
        let mut date_cells_ods: Vec<_> = sheet_ods
            .cells
            .iter()
            .filter(|(_, cell)| is_date_cell(cell))
            .collect();
        date_cells_ods.sort_by_key(|(pos, _)| *pos);

        let mut date_cells_xlsx: Vec<_> = sheet_xlsx
            .cells
            .iter()
            .filter(|(_, cell)| is_date_cell(cell))
            .collect();
        date_cells_xlsx.sort_by_key(|(pos, _)| *pos);

        // Verify same number of date cells
        assert_eq!(
            date_cells_ods.len(),
            date_cells_xlsx.len(),
            "Should have same number of date cells: ODS={}, XLSX={}",
            date_cells_ods.len(),
            date_cells_xlsx.len()
        );

        // Verify each date cell matches (position, value, style)
        for ((pos_ods, cell_ods), (pos_xlsx, cell_xlsx)) in
            date_cells_ods.iter().zip(date_cells_xlsx.iter())
        {
            // Same cell positions
            assert_eq!(pos_ods, pos_xlsx, "Date cells should be at same positions");

            // Same cell values
            match (&cell_ods.value, &cell_xlsx.value) {
                (CellValue::Number(v1), CellValue::Number(v2)) => {
                    assert!(
                        (v1 - v2).abs() < 0.0001,
                        "Date values should match at {:?}: ODS={}, XLSX={}",
                        pos_ods,
                        v1,
                        v2
                    );
                }
                _ => {
                    // For formulas or other types, just ensure both are the same type
                    assert_eq!(
                        std::mem::discriminant(&cell_ods.value),
                        std::mem::discriminant(&cell_xlsx.value),
                        "Date cell value types should match at {:?}",
                        pos_ods
                    );
                }
            }

            // Same date format styles
            assert_eq!(
                cell_ods.num_fmt, cell_xlsx.num_fmt,
                "Date format should match at {:?}: ODS={:?}, XLSX={:?}",
                pos_ods, cell_ods.num_fmt, cell_xlsx.num_fmt
            );
        }
    }

    // Helper: Detect if a cell has a date format
    fn is_date_cell(cell: &Cell) -> bool {
        cell.num_fmt
            .as_ref()
            .map(|fmt| {
                let lower = fmt.to_lowercase();
                (lower.contains('d')
                    || lower.contains('y')
                    || (lower.contains('m') && !lower.contains('0') && !lower.contains('#')))
                    && !lower.contains("general")
            })
            .unwrap_or(false)
    }
}
