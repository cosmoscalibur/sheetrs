use regex::Regex;
use sheetcraft_core::reader::read_workbook;
use std::env;

fn extract_cell_references(
    formula: &str,
    _sheet_names: &[String],
    current_sheet: &str,
    expand_ranges: bool,
) -> Vec<(String, u32, u32)> {
    let mut references = Vec::new();
    // Regex to match cell references (e.g., A1, $A$1, Sheet1!A1, 'Sheet Name'!A1, A1:B2)
    let re = Regex::new(r"(?:('([^']+)'|([A-Za-z0-9_\.]+))!)?\$?([A-Za-z]+)\$?([0-9]+)(?::\$?([A-Za-z]+)\$?([0-9]+))?").unwrap();

    for cap in re.captures_iter(formula) {
        let mut sheet_name = current_sheet.to_string();

        // Group 1: Outer sheet name wrapper
        if let Some(_) = cap.get(1) {
            if let Some(quoted) = cap.get(2) {
                sheet_name = quoted.as_str().to_string();
            } else if let Some(unquoted) = cap.get(3) {
                sheet_name = unquoted.as_str().to_string();
            }
            // Resolve sheet name alias/canonicalization if needed, but here we just look it up
        }

        let start_col_str = cap.get(4).unwrap().as_str();
        let start_row_str = cap.get(5).unwrap().as_str();

        let start_col = parse_col(start_col_str);
        let start_row = start_row_str.parse::<u32>().unwrap_or(1) - 1;

        if let Some(end_col_match) = cap.get(6) {
            let end_row_match = cap.get(7).unwrap();

            let end_col = parse_col(end_col_match.as_str());
            let end_row = end_row_match.as_str().parse::<u32>().unwrap_or(1) - 1;

            let min_r = start_row.min(end_row);
            let max_r = start_row.max(end_row);
            let min_c = start_col.min(end_col);
            let max_c = start_col.max(end_col);

            if expand_ranges && (max_r - min_r + 1) * (max_c - min_c + 1) <= 100_000 {
                for r in min_r..=max_r {
                    for c in min_c..=max_c {
                        references.push((sheet_name.clone(), r, c));
                    }
                }
            } else {
                references.push((sheet_name.clone(), start_row, start_col));
                references.push((sheet_name.clone(), end_row, end_col));
            }
        } else {
            references.push((sheet_name, start_row, start_col));
        }
    }
    references
}

fn parse_col(col_str: &str) -> u32 {
    let mut col = 0u32;
    for ch in col_str.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }
    }
    col.saturating_sub(1)
}

fn parse_cell_ref(r: &str) -> (u32, u32) {
    // Keep this for main() arg parsing
    let mut col = 0u32;
    let mut row_str = String::new();
    for ch in r.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }
    let row = row_str.parse::<u32>().unwrap_or(1).saturating_sub(1);
    (row, col.saturating_sub(1))
}

fn main() {
    let args: Vec<String> = env::args().collect();
    let path = &args[1];
    let s_name = &args[2];
    let r_name = &args[3];

    let wb = read_workbook(path).unwrap();
    let sheet_names: Vec<String> = wb.sheets.iter().map(|s| s.name.clone()).collect();
    let sheet = wb.sheets.iter().find(|s| s.name == *s_name).unwrap();

    let (r, c) = parse_cell_ref(r_name);
    let cell = sheet.cells.get(&(r, c)).expect("Cell not found");
    let formula = cell.value.as_formula().expect("Not a formula");

    println!("Formula: {}", formula);
    let refs = extract_cell_references(formula, &sheet_names, s_name, false);
    for (sn, row, col) in refs {
        // Convert back to ref
        let mut c = col + 1;
        let mut col_letter = String::new();
        while c > 0 {
            let m = (c - 1) % 26;
            col_letter.insert(0, (b'A' + m as u8) as char);
            c = (c - m) / 26;
        }
        println!("  -> {}!{}{}", sn, col_letter, row + 1);
    }
}
