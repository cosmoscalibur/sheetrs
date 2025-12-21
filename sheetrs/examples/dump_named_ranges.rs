use sheetrs::reader::read_workbook;
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        println!("Usage: dump_named_ranges <file>");
        return;
    }
    let path = &args[1];

    match read_workbook(path) {
        Ok(workbook) => {
            println!("Named ranges in {}:", path);
            println!("Total: {}", workbook.defined_names.len());
            for (name, reference) in &workbook.defined_names {
                println!("  {} -> {}", name, reference);
            }
        }
        Err(e) => {
            eprintln!("Error reading workbook: {}", e);
        }
    }
}
