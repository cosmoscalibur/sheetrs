use std::env;
use xlsutils::{XlsxWorkbook, invalid_formulas_all};

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
    let cells_with_errors = invalid_formulas_all(&mut workbook);
    println!("Cells with detected errors: {:#?}", cells_with_errors);
    return;
}
