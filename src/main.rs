use std::env;
use xlschecks::{XlsxWorkbook, invalid_formulas_all};

fn main() {
    println!("XLSChecks Project");
    let xls_path: String;
    if let Some(arg1) = env::args().nth(1) {
        xls_path = arg1;
    } else {
        println!("File path required");
        return;
    }
    let mut workbook = XlsxWorkbook::open(&xls_path).unwrap();
    println!(
        "Cells with detected errors: {:#?}",
        invalid_formulas_all(&mut workbook)
    );
    println!("Defined name errors: {:#?}", workbook.defined_name_errors());
    return;
}
