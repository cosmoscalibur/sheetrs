use sheetrs::reader::read_workbook;
use std::path::Path;

fn main() {
    let path = Path::new("minimal_test.ods");
    match read_workbook(path) {
        Ok(workbook) => {
            for sheet in &workbook.sheets {
                println!("Sheet: {}", sheet.name);
                println!(
                    "  Conditional Formatting Count: {}",
                    sheet.conditional_formatting_count
                );
                println!(
                    "  Conditional Formatting Ranges: {:?}",
                    sheet.conditional_formatting_ranges
                );
            }
        }
        Err(e) => {
            eprintln!("Error reading workbook: {}", e);
        }
    }
}
