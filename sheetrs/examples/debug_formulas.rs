use sheetrs::reader::read_workbook;

use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    let path = &args[1];
    let sheet_name = &args[2];
    let cells_to_check = &args[3..];

    let wb = read_workbook(path).unwrap();
    let sheet = wb.sheets.iter().find(|s| s.name == *sheet_name).unwrap();

    for cell_ref_str in cells_to_check {
        let mut col = 0u32;
        let mut row_str = String::new();
        for ch in cell_ref_str.chars() {
            if ch.is_ascii_alphabetic() {
                col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
            } else {
                row_str.push(ch);
            }
        }
        let row = row_str.parse::<u32>().unwrap() - 1;
        col -= 1;

        if let Some(cell) = sheet.cells.get(&(row, col)) {
            println!("{}: {:?}", cell_ref_str, cell.value);
        } else {
            println!("{}: NOT FOUND", cell_ref_str);
        }
    }
}
