use sheetrs::reader::{CellValue, read_workbook};
use std::env;

fn main() -> anyhow::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <file.ods|file.xlsx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let workbook = read_workbook(path)?;

    println!("File: {}", path);
    println!("Sheets: {}", workbook.sheets.len());

    for sheet in &workbook.sheets {
        println!("\n=== Sheet: {} ===", sheet.name);

        // Look for Sheet7 specifically
        if sheet.name == "Sheet7" {
            println!("Found Sheet7! Inspecting cells...");

            // Check cell C3 (row 2, col 2 in 0-based indexing)
            if let Some(cell) = sheet.cells.get(&(2, 2)) {
                println!("Cell C3 found:");
                println!("  Row: {}, Col: {}", cell.row, cell.col);
                println!("  Value: {:?}", cell.value);
                println!("  Num Format: {:?}", cell.num_fmt);

                match &cell.value {
                    CellValue::Text(text) => {
                        println!("  -> Text content: '{}'", text);
                        println!("  -> Trimmed: '{}'", text.trim());
                        println!(
                            "  -> Can parse as f64: {}",
                            text.trim().parse::<f64>().is_ok()
                        );
                    }
                    CellValue::Number(num) => {
                        println!("  -> Number value: {}", num);
                    }
                    CellValue::Boolean(b) => {
                        println!("  -> Boolean value: {}", b);
                    }
                    CellValue::Formula {
                        formula,
                        cached_error,
                    } => {
                        println!("  -> Formula: {}", formula);
                        println!("  -> Cached error: {:?}", cached_error);
                    }
                    CellValue::Empty => {
                        println!("  -> Empty cell");
                    }
                }
            } else {
                println!("Cell C3 NOT FOUND in cells HashMap!");
            }

            // Show all cells in Sheet7
            println!("\nAll cells in Sheet7:");
            let mut cells_vec: Vec<_> = sheet.cells.iter().collect();
            cells_vec.sort_by_key(|(pos, _)| *pos);

            for ((row, col), cell) in cells_vec {
                println!("  ({}, {}) = {:?}", row, col, cell.value);
            }
        }
    }

    Ok(())
}
