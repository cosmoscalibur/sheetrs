// ! Writer module for manipulating Excel/ODS files

mod xlsx_writer;

pub use xlsx_writer::{WorkbookModifications, modify_workbook_xlsx};

use anyhow::Result;

use std::path::Path;

/// Modify a workbook file (supports multiple operations)
pub fn modify_workbook<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let input = input_path.as_ref();

    // Determine file type by extension
    match input.extension().and_then(|s| s.to_str()) {
        Some("xlsx") => modify_workbook_xlsx(input, output_path.as_ref(), modifications),
        Some("ods") => {
            anyhow::bail!("ODS format not yet supported for modification")
        }
        _ => anyhow::bail!("Unsupported file format"),
    }
}

/// Legacy wrapper for removing sheets
pub fn remove_sheets<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    sheet_names: &[String],
) -> Result<()> {
    let modifications = WorkbookModifications {
        remove_sheets: Some(sheet_names.iter().cloned().collect()),
        remove_named_ranges: None,
    };
    modify_workbook(input_path, output_path, &modifications)
}

/// Legacy wrapper for removing named ranges
pub fn remove_named_ranges<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    range_names: &[String],
) -> Result<()> {
    let modifications = WorkbookModifications {
        remove_sheets: None,
        remove_named_ranges: Some(range_names.iter().cloned().collect()),
    };
    modify_workbook(input_path, output_path, &modifications)
}
