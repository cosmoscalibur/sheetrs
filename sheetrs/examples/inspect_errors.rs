use anyhow::Result;
use sheetrs::reader::read_workbook;
use std::env;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <file.xlsx|file.ods> [sheet_name]", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    let target_sheet = args.get(2).map(|s| s.as_str());

    println!("Reading workbook: {file_path}");
    let workbook = read_workbook(file_path)?;

    for sheet in &workbook.sheets {
        if let Some(target) = target_sheet
            && sheet.name != target
        {
            continue;
        }

        println!("\n=== Sheet: {} ===", sheet.name);

        let mut error_cells = Vec::new();

        for cell in sheet.all_cells() {
            // Check for error values
            if let Some(error) = cell.value.as_error() {
                let formula = cell.as_formula().unwrap_or("");
                error_cells.push((cell.row, cell.col, formula.to_string(), error.to_string()));
            } else if let Some(formula) = cell.as_formula() {
                // Check if formula text contains error literals
                let error_literals = [
                    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!",
                    "#CALC!",
                ];

                for err_lit in error_literals {
                    if formula.contains(err_lit) {
                        error_cells.push((
                            cell.row,
                            cell.col,
                            formula.to_string(),
                            err_lit.to_string(),
                        ));
                        break;
                    }
                }
            }
        }

        if error_cells.is_empty() {
            println!("  No error cells found");
        } else {
            println!("  Found {} error cells:", error_cells.len());
            for (row, col, formula, error) in error_cells {
                let col_letter = num_to_col(col as usize);
                println!("    {}{}:", col_letter, row + 1);
                println!("      Error: {error}");
                println!(
                    "      Formula: {}",
                    if formula.is_empty() {
                        "(empty)"
                    } else {
                        &formula
                    }
                );
            }
        }
    }

    Ok(())
}

fn num_to_col(mut num: usize) -> String {
    let mut result = String::new();
    while num >= 26 {
        result.push((b'A' + (num % 26) as u8) as char);
        num = num / 26 - 1;
    }
    result.push((b'A' + num as u8) as char);
    result.chars().rev().collect()
}
