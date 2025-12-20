use sheetcraft_core::reader::CellValue;
use sheetcraft_core::reader::read_workbook;
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
            println!("  [FORMULA PARSING ERROR]: {}", err);
        }
        for ((row, col), cell) in sheet.cells {
            match &cell.value {
                CellValue::Formula {
                    formula,
                    cached_error,
                } => {
                    if let Some(err) = cached_error {
                        println!(
                            "  ({}, {}) [ERROR]: {} (Formula: {})",
                            row, col, err, formula
                        );
                    } else {
                        println!("  ({}, {}) [FORMULA]: {}", row, col, formula);
                    }
                }
                CellValue::Text(t) => println!("  ({}, {}) [TEXT]: {}", row, col, t),
                CellValue::Number(n) => println!("  ({}, {}) [NUMBER]: {}", row, col, n),
                CellValue::Boolean(b) => println!("  ({}, {}) [BOOL]: {}", row, col, b),
                CellValue::Empty => println!("  ({}, {}) [EMPTY]", row, col),
            }
        }
    }
}
