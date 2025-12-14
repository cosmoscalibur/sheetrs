//! Excel/ODS file reader using calamine

use anyhow::{Context, Result};
use calamine::{Data, Range, Reader, Sheets, open_workbook_auto};
use std::collections::HashMap;
use std::path::Path;

pub mod ods_parser;
pub mod workbook;
pub mod xml_parser;

pub use workbook::{Cell, CellValue, Sheet, Workbook};

/// Read a workbook from a file path
pub fn read_workbook<P: AsRef<Path>>(path: P) -> Result<Workbook> {
    let path = path.as_ref();
    // Open workbook with calamine
    let mut excel: Sheets<_> = open_workbook_auto(path)
        .with_context(|| format!("Failed to open workbook: {}", path.display()))?;

    // For manual parsing, open the file as a zip archive if it is an XLSX or ODS
    let is_xlsx = path.extension().and_then(|s| s.to_str()) == Some("xlsx");
    let is_ods = path.extension().and_then(|s| s.to_str()) == Some("ods");

    let mut archive = if is_xlsx || is_ods {
        use std::fs::File;
        use std::io::BufReader;
        let file =
            File::open(path).with_context(|| format!("Failed to open file: {}", path.display()))?;
        let reader = BufReader::new(file);
        Some(zip::ZipArchive::new(reader)?)
    } else {
        None
    };

    let sheet_names = excel.sheet_names();
    let mut sheets = Vec::new();

    for (index, sheet_name) in sheet_names.iter().enumerate() {
        // Get both values and formulas
        let range = excel.worksheet_range(sheet_name).ok();

        let (formula_range, formula_error) = match excel.worksheet_formula(sheet_name) {
            Ok(range) => (Some(range), None),
            Err(e) => (None, Some(format!("{:?}", e))),
        };

        let mut sheet = parse_sheet(
            sheet_name,
            range.as_ref(),
            formula_range.as_ref(),
            formula_error,
        );

        // Extract hidden columns/rows and merged cells for XLSX/ODS files
        if let Some(ref mut archive_ref) = archive {
            if is_xlsx {
                if let Ok((hidden_cols, hidden_rows)) =
                    xml_parser::extract_hidden_columns_rows_from_xlsx(archive_ref, index)
                {
                    sheet.hidden_columns = hidden_cols;
                    sheet.hidden_rows = hidden_rows;
                }
                if let Ok(merged) = xml_parser::extract_merged_cells_from_xlsx(archive_ref, index) {
                    sheet.merged_cells = merged;
                }
            }
            // Extract hidden columns/rows and merged cells for ODS files
            else if is_ods {
                if let Ok((hidden_cols, hidden_rows)) =
                    ods_parser::extract_hidden_columns_rows_from_ods(archive_ref, index)
                {
                    sheet.hidden_columns = hidden_cols;
                    sheet.hidden_rows = hidden_rows;
                }
                if let Ok(merged) = ods_parser::extract_merged_cells_from_ods(archive_ref, index) {
                    sheet.merged_cells = merged;
                }
            }
        }

        sheets.push(sheet);
    }

    // Extract defined names, hidden sheets, macros, AND STYLES
    let (defined_names, hidden_sheets, has_macros, styles) =
        if let Some(ref mut archive_ref) = archive {
            if is_xlsx {
                let names = xml_parser::extract_defined_names_from_xlsx(archive_ref)?;
                let hidden = xml_parser::extract_hidden_sheets_from_xlsx(archive_ref)?;
                let macros = xml_parser::has_vba_project_xlsx(archive_ref)?;
                let styles = xml_parser::parse_styles(archive_ref).unwrap_or_default();
                (names, hidden, macros, styles)
            } else if is_ods {
                let hidden = ods_parser::extract_hidden_sheets_from_ods(archive_ref)?;
                let macros = ods_parser::has_macros_ods(archive_ref)?;
                (HashMap::new(), hidden, macros, Vec::new())
            } else {
                (HashMap::new(), Vec::new(), false, Vec::new())
            }
        } else {
            (HashMap::new(), Vec::new(), false, Vec::new())
        };

    // If we have styles, apply them to cells
    if !styles.is_empty() && is_xlsx {
        if let Some(ref mut archive_ref) = archive {
            for (index, sheet) in sheets.iter_mut().enumerate() {
                if let Ok(cell_styles) =
                    xml_parser::extract_cell_style_indices_from_xlsx(archive_ref, index)
                {
                    for ((row, col), style_idx) in cell_styles {
                        if let Some(fmt) = styles.get(style_idx) {
                            if let Some(cell) = sheet.cells.get_mut(&(row, col)) {
                                cell.num_fmt = Some(fmt.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(Workbook {
        path: path.to_path_buf(),
        sheets,
        defined_names,
        hidden_sheets,
        has_macros,
    })
}

fn parse_sheet(
    name: &str,
    range: Option<&Range<Data>>,
    formula_range: Option<&Range<String>>,
    formula_parsing_error: Option<String>,
) -> Sheet {
    let mut cells = HashMap::new();

    // Determine valid bounds for values
    let (r_start, r_end) = if let Some(r) = range {
        (r.start().unwrap_or((0, 0)), r.end().unwrap_or((0, 0)))
    } else {
        ((u32::MAX, u32::MAX), (0, 0))
    };

    // Determine valid bounds for formulas
    let (f_start, f_end) = if let Some(f) = formula_range {
        (f.start().unwrap_or((0, 0)), f.end().unwrap_or((0, 0)))
    } else {
        ((u32::MAX, u32::MAX), (0, 0))
    };

    // Calculate global bounding box (union of both ranges)
    let min_row = r_start.0.min(f_start.0);
    let min_col = r_start.1.min(f_start.1);
    let max_row = r_end.0.max(f_end.0);
    let max_col = r_end.1.max(f_end.1);

    // If no valid range, return early
    if min_row > max_row || min_col > max_col {
        return Sheet {
            name: name.to_string(),
            cells,
            used_range: None,
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error,
        };
    }

    for row in min_row..=max_row {
        for col in min_col..=max_col {
            let mut cell_value = None;
            let mut formula = None;

            // Get the calculated value if within range
            if let Some(r) = range {
                let (r_rows, r_cols) = r.get_size();
                if row >= r_start.0
                    && row < r_start.0 + r_rows as u32
                    && col >= r_start.1
                    && col < r_start.1 + r_cols as u32
                {
                    let rel_row = (row - r_start.0) as usize;
                    let rel_col = (col - r_start.1) as usize;

                    if let Some(cell_data) = r.get((rel_row, rel_col)) {
                        if !matches!(cell_data, Data::Empty) {
                            cell_value = Some(parse_cell_value(cell_data));
                        }
                    }
                }
            }

            // Get the formula if within range
            if let Some(f) = formula_range {
                let (f_rows, f_cols) = f.get_size();
                if row >= f_start.0
                    && row < f_start.0 + f_rows as u32
                    && col >= f_start.1
                    && col < f_start.1 + f_cols as u32
                {
                    let rel_row = (row - f_start.0) as usize;
                    let rel_col = (col - f_start.1) as usize;

                    if let Some(formula_str) = f.get((rel_row, rel_col)) {
                        if !formula_str.is_empty() {
                            formula = Some(formula_str.clone());
                        }
                    }
                }
            }

            // Create cell if we have either a value or a formula
            if cell_value.is_some() || formula.is_some() {
                let value = if let Some(f) = formula {
                    CellValue::Formula(f)
                } else {
                    cell_value.unwrap_or(CellValue::Empty)
                };

                let cell = Cell {
                    row,
                    col,
                    value,
                    num_fmt: None,
                };
                cells.insert((row, col), cell);
            }
        }
    }

    Sheet {
        name: name.to_string(),
        cells,
        used_range: if min_row <= max_row && min_col <= max_col {
            Some((max_row + 1, max_col + 1))
        } else {
            None
        },
        hidden_columns: Vec::new(),
        hidden_rows: Vec::new(),
        merged_cells: Vec::new(),
        formula_parsing_error,
    }
}

fn parse_cell_value(data: &Data) -> CellValue {
    match data {
        Data::Int(i) => CellValue::Number(*i as f64),
        Data::Float(f) => CellValue::Number(*f),
        Data::String(s) => CellValue::Text(s.clone()),
        Data::Bool(b) => CellValue::Boolean(*b),
        Data::Error(e) => CellValue::Error(format!("{:?}", e)),
        Data::Empty => CellValue::Empty,
        Data::DateTime(dt) => CellValue::Number(dt.as_f64()),
        Data::DateTimeIso(s) => CellValue::Text(s.clone()),
        Data::DurationIso(s) => CellValue::Text(s.clone()),
    }
}
