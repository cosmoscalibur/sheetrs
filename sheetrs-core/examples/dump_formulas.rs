use sheetrs_core::reader::CellValue;
use sheetrs_core::reader::read_workbook;
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        println!("Usage: dump_formulas <file>");
        return;
    }
    let path = &args[1];
    let workbook = read_workbook(path).unwrap();

    for sheet in workbook.sheets {
        println!("Sheet: {}", sheet.name);
        if let Some(err) = &sheet.formula_parsing_error {
            println!("  [FORMULA PARSING ERROR]: {err}");
        }
        for ((row, col), cell) in sheet.cells {
            if let Some(formula) = &cell.formula {
                if cell.value.is_error() {
                    println!(
                        "  ({}, {}) [ERROR]: {} (Formula: {})",
                        row,
                        col,
                        cell.value
                            .as_error()
                            .map_or("Unknown".to_string(), |e| e.to_string()),
                        formula
                    );
                } else {
                    println!("  ({row}, {col}) [FORMULA]: {formula}");
                }
            } else {
                match &cell.value {
                    CellValue::Text(t) => println!("  ({row}, {col}) [TEXT]: {t}"),
                    CellValue::Number(n) => println!("  ({row}, {col}) [NUMBER]: {n}"),
                    CellValue::Boolean(b) => println!("  ({row}, {col}) [BOOL]: {b}"),
                    CellValue::Error(e) => println!("  ({row}, {col}) [ERROR]: {e}"),
                    CellValue::Empty => println!("  ({row}, {col}) [EMPTY]"),
                }
            }
        }
    }
}
