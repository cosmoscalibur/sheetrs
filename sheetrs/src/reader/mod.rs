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
pub use workbook::{
    Cell, CellValue, DocumentSettings, ExternalWorkbook, IterationSettings, Sheet, Workbook,
};

/// Trait for spreadsheet format readers
///
/// This trait defines the common interface that different file format parsers (XLSX, ODS)
/// must implement. It allows the main logic to abstract over specific file formats.
pub trait WorkbookReader {
    /// Read all sheets from the workbook
    fn read_sheets(&mut self) -> Result<Vec<Sheet>>;

    /// Read defined names (named ranges) from the workbook
    fn read_defined_names(&mut self) -> Result<HashMap<String, String>>;

    /// Read the list of hidden sheet names
    fn read_hidden_sheets(&mut self) -> Result<Vec<String>>;

    /// Check if the workbook contains macros (VBA, Basic, Scripts)
    fn has_macros(&mut self) -> Result<bool>;

    /// Read external file links (generic links)
    fn read_external_links(&mut self) -> Result<Vec<String>>;

    /// Read external workbook references with their indices
    ///
    /// The indices correspond to the reference ID used in formulas (e.g., `[1]Sheet1!A1`).
    /// For XLSX, this maps to the externalLink relations.
    /// For ODS, this maps to the order of appearance or explicit file references.
    fn read_external_workbooks(&mut self) -> Result<Vec<ExternalWorkbook>>;

    /// Read the last modification date (parsed from file metadata)
    fn read_modified_date(&mut self) -> Result<Option<chrono::DateTime<chrono::Utc>>>;

    /// Check if the workbook uses the 1904 date system
    fn read_date1904(&mut self) -> Result<bool>;

    /// Read the calculation settings stored in the document, as stored.
    fn read_settings(&mut self) -> Result<DocumentSettings>;
}

/// Read a workbook from a file path
///
/// This function detects the file format based on extension and delegates to the appropriate reader.
///
/// # Arguments
///
/// * `path` - Path to the spreadsheet file
///
/// # Returns
///
/// * `Result<Workbook>` - The parsed workbook structure
///
/// # Examples
///
/// ```no_run
/// use sheetrs::reader::read_workbook;
/// let params = read_workbook("tests/minimal_test.xlsx").unwrap();
/// ```
/// Read a workbook from a file path
///
/// This function detects the file format based on extension and delegates to `read_workbook_from_reader`.
pub fn read_workbook<P: AsRef<Path>>(path: P) -> Result<Workbook> {
    let path_ref = path.as_ref();
    let file = File::open(path_ref)
        .with_context(|| format!("Failed to open file: {}", path_ref.display()))?;

    // Extract file size from metadata (Parse-Dont-Validate)
    let file_size_bytes = std::fs::metadata(path_ref).map(|m| m.len()).unwrap_or(0);

    let extension = path_ref.extension().and_then(|s| s.to_str());
    let mut workbook = read_workbook_from_reader(file, extension)?;

    // Preserve the original path and file size
    workbook.path = path_ref.to_path_buf();
    workbook.file_size_bytes = file_size_bytes;
    Ok(workbook)
}

/// Read a workbook from a generic reader (must implement Read + Seek)
///
/// # Arguments
///
/// * `reader` - Any source that implements `std::io::Read` and `std::io::Seek` (e.g., `File`, `Cursor`)
/// * `extension` - Optional hint for the file format ("xlsx", "xlsm", "ods")
pub fn read_workbook_from_reader<R: std::io::Read + std::io::Seek>(
    reader: R,
    extension: Option<&str>,
) -> Result<Workbook> {
    let mut archive = ZipArchive::new(reader).context("Failed to open zip archive")?;

    let is_xlsx = extension
        .map(|s| s.eq_ignore_ascii_case("xlsx") || s.eq_ignore_ascii_case("xlsm"))
        .unwrap_or(false);

    let is_ods = extension
        .map(|s| s.eq_ignore_ascii_case("ods"))
        .unwrap_or(false);

    // If extension is unknown, we could try to detect by contents,
    // but for now we'll stick to extension-based or default to one if Zip format allows.
    // In practice, we usually know the extension.

    let (
        sheets,
        defined_names,
        hidden_sheets,
        has_macros,
        external_workbooks,
        modified_date,
        date1904,
        settings,
    ) = if is_xlsx {
        let mut reader = XlsxReader::new(&mut archive)?;
        (
            reader.read_sheets()?,
            reader.read_defined_names()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
            reader.read_modified_date()?,
            reader.read_date1904()?,
            reader.read_settings()?,
        )
    } else if is_ods || (!is_xlsx && archive.by_name("content.xml").is_ok()) {
        // Simple heuristic for ODS if no extension: check for content.xml
        let mut reader = OdsReader::new(&mut archive)?;
        let sheets = reader.read_sheets()?;
        let defined_names = reader.read_defined_names()?;
        (
            sheets,
            defined_names,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
            reader.read_modified_date()?,
            reader.read_date1904()?,
            reader.read_settings()?,
        )
    } else if is_xlsx || archive.by_name("[Content_Types].xml").is_ok() {
        // Simple heuristic for XLSX if no extension: check for [Content_Types].xml
        let mut reader = XlsxReader::new(&mut archive)?;
        (
            reader.read_sheets()?,
            reader.read_defined_names()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
            reader.read_modified_date()?,
            reader.read_date1904()?,
            reader.read_settings()?,
        )
    } else {
        return Err(anyhow::anyhow!("Unsupported or unrecognizable file format"));
    };

    Ok(Workbook {
        path: Path::new("").to_path_buf(), // Default empty path for reader
        sheets,
        defined_names,
        hidden_sheets,
        has_macros,
        external_workbooks,
        modified_date,
        date1904,
        settings,
        file_size_bytes: 0, // Will be set by read_workbook when path is known
    })
}

#[cfg(test)]
mod date_format_parity_tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_date_format_parity_ods_xlsx() {
        // Truth values for verification (2023/08/01)
        const EXPECTED_D7_FORMAT: &str = "m/d/yyyy";
        const EXPECTED_D7_ODS_VALUE: f64 = 45139.0; // days since the null date 1899-12-30
        const EXPECTED_D7_XLSX_VALUE: f64 = 45139.0; // Excel serial (same from 1900-03-01)
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
            d7_ods.num_fmt.as_deref().unwrap(),
            EXPECTED_D7_FORMAT,
            "ODS D7 format should match truth value"
        );
        if let CellValue::Number(val) = d7_ods.value {
            assert!(
                (val - EXPECTED_D7_ODS_VALUE).abs() < 0.1,
                "ODS D7 value should be ~{EXPECTED_D7_ODS_VALUE}, got {val}"
            );
        }

        // Verify D8 truth values (Indexing tests) - THE FIX TARGET
        let d8_ods = sheet_ods.cells.get(&(7, 3)).unwrap();
        assert_eq!(
            d8_ods.num_fmt.as_deref().unwrap(),
            EXPECTED_D8_FORMAT,
            "ODS D8 format should be 'dd/mm/yy hh:mm' (two-digit), not 'd/m/yy hh:mm'"
        );
        // Verify it's a formula or numeric value
        assert!(
            d8_ods.is_formula() || matches!(d8_ods.value, CellValue::Number(_)),
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
            d7_xlsx.num_fmt.as_deref().unwrap(),
            EXPECTED_D7_FORMAT,
            "XLSX D7 format should match truth value"
        );
        if let CellValue::Number(val) = d7_xlsx.value {
            assert!(
                (val - EXPECTED_D7_XLSX_VALUE).abs() < 0.1,
                "XLSX D7 value should be ~{EXPECTED_D7_XLSX_VALUE}, got {val}"
            );
        }

        // Verify D8 truth values
        let d8_xlsx = sheet_xlsx.cells.get(&(7, 3)).unwrap();
        assert_eq!(
            d8_xlsx.num_fmt.as_deref().unwrap(),
            EXPECTED_D8_FORMAT,
            "XLSX D8 format should match truth value"
        );

        // STEP 3: Verify ODS == XLSX parity (same calendar date, not same serial)

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

        // Verify each date cell represents the same calendar date.
        // ODS counts days from the null date 1899-12-30 and XLSX from Excel's
        // epoch; they agree from 1900-03-01 on, which holds for the fixture.
        for ((pos_ods, cell_ods), (pos_xlsx, cell_xlsx)) in
            date_cells_ods.iter().zip(date_cells_xlsx.iter())
        {
            // Same cell positions
            assert_eq!(pos_ods, pos_xlsx, "Date cells should be at same positions");

            // Same cell values (skip tight comparison for formula cells — cached values
            // may differ between ODS and XLSX due to epoch and recalculation differences)
            if cell_ods.is_formula() || cell_xlsx.is_formula() {
                assert_eq!(
                    std::mem::discriminant(&cell_ods.value),
                    std::mem::discriminant(&cell_xlsx.value),
                    "Date cell value types should match at {pos_ods:?}"
                );
            } else {
                match (&cell_ods.value, &cell_xlsx.value) {
                    (CellValue::Number(v_ods), CellValue::Number(v_xlsx)) => {
                        assert!(
                            (v_ods - v_xlsx).abs() < 0.0001,
                            "Date cells should represent the same date at {pos_ods:?}: ODS serial={v_ods}, XLSX serial={v_xlsx}"
                        );
                    }
                    _ => {
                        assert_eq!(
                            std::mem::discriminant(&cell_ods.value),
                            std::mem::discriminant(&cell_xlsx.value),
                            "Date cell value types should match at {pos_ods:?}"
                        );
                    }
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

#[cfg(test)]
mod external_workbook_parity_tests {
    use super::*;

    #[test]
    fn test_parity_external_workbook() {
        let root = std::env::var("CARGO_MANIFEST_DIR").unwrap();
        let ods_path = format!("{root}/../tests/minimal_test.ods");
        let xlsx_path = format!("{root}/../tests/minimal_test.xlsx");

        let ods_wb = read_workbook(&ods_path).unwrap();
        let xlsx_wb = read_workbook(&xlsx_path).unwrap();

        // 1. Both files has external workbooks count 1
        assert_eq!(
            ods_wb.external_workbooks.len(),
            1,
            "ODS should have exactly 1 external workbook"
        );
        assert_eq!(
            xlsx_wb.external_workbooks.len(),
            1,
            "XLSX should have exactly 1 external workbook"
        );

        // 2. External workbook unique element is equal between both (basename check)
        let ods_ext = std::path::Path::new(&ods_wb.external_workbooks[0].path)
            .file_name()
            .unwrap();
        let xlsx_ext = std::path::Path::new(&xlsx_wb.external_workbooks[0].path)
            .file_name()
            .unwrap();
        assert_eq!(ods_ext, xlsx_ext, "External workbook basename should match");

        // 3. Sheet7!J6 in ODS contain `[1]`
        // Sheet7 is at index 5 (0-indexed) based on `examples/list_sheets.rs` output
        let sheet_idx = 5;
        let row = 5; // J6 -> row 6 -> 0-indexed 5
        let col = 9; // J -> 10th col -> 0-indexed 9

        // Helper to get formula
        let get_formula = |wb: &Workbook, sheet_idx: usize, row: u32, col: u32| {
            wb.sheets
                .get(sheet_idx)
                .and_then(|s| s.cells.get(&(row, col)))
                .and_then(|c| c.formula.clone())
        };

        let ods_formula =
            get_formula(&ods_wb, sheet_idx, row, col).expect("ODS Sheet7!J6 should have a formula");

        assert!(
            ods_formula.contains("[1]"),
            "ODS formula at Sheet7!J6 should contain '[1]', found: {ods_formula}"
        );

        // 4. ODS formula = XLSX formula in Sheet7!J6
        let xlsx_formula = get_formula(&xlsx_wb, sheet_idx, row, col)
            .expect("XLSX Sheet7!J6 should have a formula");

        assert_eq!(
            ods_formula, xlsx_formula,
            "Formulas at Sheet7!J6 should match exactly"
        );
    }
}

#[cfg(test)]
mod escaped_text_parity_tests {
    use super::*;
    use std::io::Cursor;

    fn read(bytes: &[u8], extension: &str) -> Workbook {
        read_workbook_from_reader(Cursor::new(bytes), Some(extension)).unwrap()
    }

    /// Cells whose XLSX markup holds `&quot;`, `&apos;`, `&lt;` or `&gt;`
    /// read the same from both formats.
    #[test]
    fn test_escaped_text_parity() {
        let xlsx = read(include_bytes!("../../../tests/minimal_test.xlsx"), "xlsx");
        let ods = read(include_bytes!("../../../tests/minimal_test.ods"), "ods");

        let formulas = [
            ("Sheet7", (6, 7), r#"IF(I7="", "Y","N")"#),
            (
                "Sheet7",
                (6, 11),
                r#"IF(H7="Y", 1,IF(H8="Y",2,IF(C4<0,3,4)))"#,
            ),
            ("Sheet7", (7, 8), r#""""#),
            ("hidden_formula_error", (3, 3), "'Indexing tests'!H12"),
        ];
        let texts = [(
            "Summary",
            (26, 3),
            "Indexing tests!H12 -> hidden_formula_error!B7 -> hidden_formula_error!D4",
        )];

        for (format, workbook) in [("xlsx", &xlsx), ("ods", &ods)] {
            let cell = |sheet: &str, (row, col): (u32, u32)| {
                workbook
                    .get_sheet(sheet)
                    .and_then(|s| s.get_cell(row, col))
                    .unwrap_or_else(|| panic!("{format}: missing {sheet} {row},{col}"))
            };
            for (sheet, pos, expected) in formulas {
                let formula = cell(sheet, pos).as_formula();
                assert_eq!(formula, Some(expected), "{format}: {sheet} {pos:?}");
            }
            for (sheet, pos, expected) in texts {
                let value = &cell(sheet, pos).value;
                assert_eq!(
                    value,
                    &CellValue::Text(expected.into()),
                    "{format}: {sheet} {pos:?}"
                );
            }
        }
    }
}
