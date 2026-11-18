//! Prints aggregate statistics on how a workbook is shaped, to size storage designs: how many
//! distinct formula templates there are once references are made relative, how long the runs
//! of identical templates down a column are, and how much text repeats. Never prints content.

use std::collections::{BTreeMap, HashSet};
use std::env;
use std::error::Error;
use std::fs::File;

use sheetrs_core::{CellPos, CellRef, FormulaRef, SheetView, ValueRef};

/// Rewrites the A1 references of a formula as offsets from `base`, so that formulas filled down
/// a column normalize to the same text. A rough tokenizer: it is meant for statistics only.
fn relative_template(formula: &str, base: CellPos) -> String {
    let bytes = formula.as_bytes();
    let mut template = String::with_capacity(formula.len());
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'"' {
            let start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    if bytes.get(index + 1) == Some(&b'"') {
                        index += 2;
                        continue;
                    }
                    index += 1;
                    break;
                }
                index += 1;
            }
            template.push_str(&formula[start..index]);
        } else if byte == b'$' || byte.is_ascii_alphabetic() {
            let start = index;
            while index < bytes.len()
                && (bytes[index] == b'$'
                    || bytes[index] == b'_'
                    || bytes[index] == b'.'
                    || bytes[index].is_ascii_alphanumeric())
            {
                index += 1;
            }
            let token = &formula[start..index];
            let is_call = bytes.get(index) == Some(&b'(');
            match CellRef::from_a1(token) {
                Ok(cell) if !is_call => template.push_str(&cell.to_r1c1(base)),
                _ => template.push_str(token),
            }
        } else {
            template.push(char::from(byte));
            index += 1;
        }
    }
    template
}

#[derive(Default)]
struct Shape {
    cells: usize,
    numbers: usize,
    texts: usize,
    others: usize,
    formulas: usize,
    formula_text_bytes: usize,
    runs: usize,
    runs_of_100_or_more_cells: usize,
    columns_with_data: usize,
    bounding_cells: u64,
    text_value_bytes: usize,
    exact_formula_texts: HashSet<String>,
    templates: HashSet<String>,
    text_values: HashSet<String>,
}

fn template_of(sheet: &SheetView<'_>, pos: CellPos, formula: FormulaRef<'_>) -> Option<String> {
    match formula {
        FormulaRef::Text(text) => Some(relative_template(text, pos)),
        FormulaRef::Shared { index } => {
            let anchor = sheet.shared_anchor(index)?;
            match sheet.get(anchor)?.formula {
                Some(FormulaRef::Text(text)) => Some(relative_template(text, anchor)),
                _ => None,
            }
        }
    }
}

/// Counts the runs of consecutive rows with the same template in one column.
fn count_runs(entries: &[(u32, Option<String>)], shape: &mut Shape) {
    let mut previous: Option<(u32, &str)> = None;
    let mut length = 0_usize;
    let close = |length: &mut usize, shape: &mut Shape| {
        if *length > 0 {
            shape.runs += 1;
            if *length >= 100 {
                shape.runs_of_100_or_more_cells += *length;
            }
            *length = 0;
        }
    };
    for (row, template) in entries {
        let Some(template) = template.as_deref() else {
            close(&mut length, shape);
            previous = None;
            continue;
        };
        let continues = previous.is_some_and(|(previous_row, previous_template)| {
            *row == previous_row + 1 && previous_template == template
        });
        if !continues {
            close(&mut length, shape);
        }
        length += 1;
        previous = Some((*row, template));
    }
    close(&mut length, shape);
}

fn collect_sheet(sheet: &SheetView<'_>, shape: &mut Shape) {
    let mut by_column: BTreeMap<u32, Vec<(u32, Option<String>)>> = BTreeMap::new();
    let (mut max_row, mut max_col) = (0_u32, 0_u32);
    for (pos, cell) in sheet.cells_by_column() {
        shape.cells += 1;
        max_row = max_row.max(pos.row);
        max_col = max_col.max(pos.col);
        match &cell.value {
            ValueRef::Number(_) => shape.numbers += 1,
            ValueRef::Text(text) => {
                shape.texts += 1;
                shape.text_value_bytes += text.len();
                shape.text_values.insert(text.to_string());
            }
            _ => shape.others += 1,
        }
        let template = cell.formula.and_then(|formula| {
            shape.formulas += 1;
            if let FormulaRef::Text(text) = formula {
                shape.formula_text_bytes += text.len();
                shape.exact_formula_texts.insert(text.to_owned());
            }
            let template = template_of(sheet, pos, formula);
            if let Some(template) = &template {
                shape.templates.insert(template.clone());
            }
            template
        });
        by_column
            .entry(pos.col)
            .or_default()
            .push((pos.row, template));
    }
    shape.columns_with_data += by_column.len();
    shape.bounding_cells +=
        u64::from(max_row.saturating_add(1)) * u64::from(max_col.saturating_add(1));
    for entries in by_column.values() {
        count_runs(entries, shape);
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let path = env::args().nth(1).ok_or("usage: shape_stats <file.xlsx>")?;
    let workbook = sheetrs_xlsx::read(File::open(path)?)?;
    let mut shape = Shape::default();
    let mut sheet_sizes: Vec<usize> = Vec::new();
    for sheet in workbook.sheets() {
        collect_sheet(&sheet, &mut shape);
        sheet_sizes.push(sheet.len());
    }
    sheet_sizes.sort_unstable_by(|a, b| b.cmp(a));
    let largest_per_mille: Vec<usize> = sheet_sizes
        .iter()
        .take(3)
        .map(|size| size * 1000 / shape.cells.max(1))
        .collect();
    println!("largest_three_sheets_per_mille_of_cells={largest_per_mille:?}");

    let distinct_text_bytes: usize = shape.text_values.iter().map(String::len).sum();
    println!(
        "cells={} numbers={} texts={} others={} formulas={}",
        shape.cells, shape.numbers, shape.texts, shape.others, shape.formulas
    );
    println!(
        "formula_text_bytes={} distinct_exact_texts={} distinct_relative_templates={}",
        shape.formula_text_bytes,
        shape.exact_formula_texts.len(),
        shape.templates.len()
    );
    println!(
        "formula_runs_down_columns={} (tenths of cells per run: {})",
        shape.runs,
        shape.formulas * 10 / shape.runs.max(1)
    );
    println!(
        "formula_cells_in_runs_of_100_or_more={}",
        shape.runs_of_100_or_more_cells
    );
    println!(
        "text_values={} distinct_text_values={} text_bytes={} distinct_text_bytes={distinct_text_bytes}",
        shape.texts,
        shape.text_values.len(),
        shape.text_value_bytes
    );
    println!(
        "columns_with_data={} density_per_mille={}",
        shape.columns_with_data,
        u64::try_from(shape.cells)? * 1000 / shape.bounding_cells.max(1)
    );
    Ok(())
}
