//! Reads an XLSX file, on one thread or on the given number of workers, and prints aggregate
//! counts and the elapsed time, never cell content.
//!
//! Run it under `/usr/bin/time -v` to measure the peak resident set size.

use std::env;
use std::error::Error;
use std::fs::File;
use std::time::Instant;

use sheetrs_core::Formula;

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let path = arguments
        .next()
        .ok_or("usage: read_stats <file.xlsx> [workers]")?;
    let workers: usize = arguments.next().map_or(Ok(1), |workers| workers.parse())?;
    let started = Instant::now();
    let workbook = sheetrs_xlsx::read_parallel(File::open(&path)?, workers)?;
    let elapsed = started.elapsed();

    let (mut cells, mut formulas) = (0_usize, 0_usize);
    for sheet in workbook.sheets() {
        cells += sheet.len();
        formulas += sheet
            .cells()
            .filter(|(_, cell)| {
                matches!(
                    cell.formula,
                    Some(Formula::Text(_) | Formula::Shared { .. })
                )
            })
            .count();
    }
    println!(
        "sheets={} cells={cells} formulas={formulas} read_ms={}",
        workbook.sheets().len(),
        elapsed.as_millis()
    );
    Ok(())
}
