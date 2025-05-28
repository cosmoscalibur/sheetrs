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
    let mut workbook = XlsxWorkbook::open(&xls_path).unwrap();
    for (sheet_id, sheet_name) in workbook.sheets.clone() {
        dbg!(&sheet_id);
        dbg!(&sheet_name);
    }
    return;
}
