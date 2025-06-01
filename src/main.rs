use std::collections::HashMap;
use std::env;
use xlsutils::{XlsxWorkbook, invalid_formulas_by_sheet_path};

fn main() {
    println!("XLSUtils Project");
    let xls_path: String;
    if let Some(arg1) = env::args().nth(1) {
        xls_path = arg1;
    } else {
        println!("File path required");
        return;
    }
    let mut workbook = XlsxWorkbook::open(&xls_path).unwrap();
    dbg!(&workbook.sheets);
    dbg!(&workbook.defined_names);
    let mut cells_with_errors_insheet: Vec<String>;
    let mut cells_with_errors: HashMap<String, Vec<String>> = HashMap::new();
    for (sheet_name, sheet_path) in workbook.sheets.clone() {
        println!("Analyzed {sheet_name}");
        cells_with_errors_insheet = invalid_formulas_by_sheet_path(&mut workbook, &sheet_path);

        if cells_with_errors_insheet.len() > 0 {
            cells_with_errors.insert(sheet_name.clone(), cells_with_errors_insheet.clone());
            cells_with_errors_insheet.clear();
        }
    }
    println!("Cells with detected errors: {:#?}", cells_with_errors);
    return;
}
