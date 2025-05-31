use std::env;
use xlsutils::XlsxWorkbook;

fn main() {
    println!("XLSUtils Project");
    let xls_path: String;
    if let Some(arg1) = env::args().nth(1) {
        xls_path = dbg!(arg1);
    } else {
        println!("File path required");
        return;
    }
    let workbook = XlsxWorkbook::open(&xls_path).unwrap();
    for (sheet_name, sheet_path) in workbook.sheets.clone() {
        dbg!(&sheet_path);
        dbg!(&sheet_name);
    }
    return;
}
