//! Compares the time of the operations a workbook needs on two storage layouts of synthetic
//! cells: an ordered map of 64-byte cells, as the model stores them today, and one sorted
//! vector of rows with typed channels per column. The sizes follow a large production sheet.

use std::collections::BTreeMap;
use std::hint::black_box;
use std::time::Instant;

use sheetrs_core::CellPos;

const ROWS: u32 = 20_000;
const COLUMNS: u32 = 100;
const KEPT_FROM_COLUMN: u32 = 97;

/// A cell as large as the one of the model: a value, a formula and flags.
#[derive(Clone, Copy)]
struct FatCell {
    payload: [u64; 6],
    has_formula: bool,
}

/// One column: sorted rows and one channel per attribute.
#[derive(Default)]
struct Column {
    rows: Vec<u32>,
    payloads: Vec<u64>,
    formula_ids: Vec<u32>,
}

fn timed<T>(label: &str, work: impl FnOnce() -> T) -> T {
    let started = Instant::now();
    let result = work();
    println!("{label:<44} {:>8} us", started.elapsed().as_micros());
    result
}

fn fat_cell(row: u32, column: u32) -> FatCell {
    FatCell {
        payload: [u64::from(row) ^ u64::from(column); 6],
        has_formula: !(row + column).is_multiple_of(3),
    }
}

fn build_tree() -> BTreeMap<CellPos, FatCell> {
    let mut tree = BTreeMap::new();
    for row in 0..ROWS {
        for col in 0..COLUMNS {
            tree.insert(CellPos { row, col }, fat_cell(row, col));
        }
    }
    tree
}

fn build_columns() -> Vec<Column> {
    let mut columns: Vec<Column> = (0..COLUMNS).map(|_| Column::default()).collect();
    for row in 0..ROWS {
        for (col, column) in (0..COLUMNS).zip(columns.iter_mut()) {
            column.rows.push(row);
            column.payloads.push(u64::from(row) ^ u64::from(col));
            column
                .formula_ids
                .push(u32::from(!(row + col).is_multiple_of(3)) * (row % 500 + 1));
        }
    }
    columns
}

/// A cheap deterministic sequence of cells to look up.
fn lookups() -> Vec<(u32, u32)> {
    let mut state = 0x2545_F491_4F6C_DD1D_u64;
    (0..1_000_000)
        .map(|_| {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1);
            let row = u32::try_from((state >> 33) % u64::from(ROWS)).unwrap_or(0);
            let col = u32::try_from((state >> 13) % u64::from(COLUMNS)).unwrap_or(0);
            (row, col)
        })
        .collect()
}

fn main() {
    println!("{ROWS} rows x {COLUMNS} columns = {} cells", ROWS * COLUMNS);

    println!("-- build, cells arriving in row order");
    let mut tree = timed("ordered map: insert", build_tree);
    let mut columns = timed("columns: push", build_columns);

    println!("-- convert every formula to a value");
    timed("ordered map: iterate and modify", || {
        for cell in tree.values_mut() {
            cell.has_formula = false;
        }
    });
    timed("columns: clear the formula channel", || {
        for column in &mut columns {
            column.formula_ids.fill(0);
        }
    });

    println!("-- 1 000 000 lookups of a cell by row and column");
    let cells = lookups();
    let found = timed("ordered map: get", || {
        cells
            .iter()
            .filter(|(row, col)| {
                tree.contains_key(&CellPos {
                    row: *row,
                    col: *col,
                })
            })
            .count()
    });
    let found_in_columns = timed("columns: binary search in the column", || {
        cells
            .iter()
            .filter(|(row, col)| {
                usize::try_from(*col)
                    .ok()
                    .and_then(|index| columns.get(index))
                    .is_some_and(|column| column.rows.binary_search(row).is_ok())
            })
            .count()
    });
    black_box((found, found_in_columns));

    println!("-- read one column over all rows (a filter criterion)");
    let sum = timed("ordered map: one lookup per row", || {
        (0..ROWS)
            .filter_map(|row| tree.get(&CellPos { row, col: 36 }))
            .map(|cell| cell.payload[0])
            .fold(0_u64, u64::wrapping_add)
    });
    let sum_in_columns = timed("columns: scan the column", || {
        columns.get(36).map_or(0, |column| {
            column
                .payloads
                .iter()
                .fold(0_u64, |a, b| a.wrapping_add(*b))
        })
    });
    black_box((sum, sum_in_columns));

    println!("-- delete columns 1..{KEPT_FROM_COLUMN}, the rest shifts left");
    let shifted = timed("ordered map: rebuild without them", || {
        let mut rebuilt = BTreeMap::new();
        for (pos, cell) in tree {
            if pos.col == 0 {
                rebuilt.insert(pos, cell);
            } else if pos.col >= KEPT_FROM_COLUMN {
                rebuilt.insert(
                    CellPos {
                        row: pos.row,
                        col: pos.col - KEPT_FROM_COLUMN + 1,
                    },
                    cell,
                );
            }
        }
        rebuilt
    });
    timed("columns: remove the column vectors", || {
        columns.drain(1..usize::try_from(KEPT_FROM_COLUMN).unwrap_or(1));
    });
    black_box((shifted.len(), columns.len()));
}
