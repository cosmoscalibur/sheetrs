//! Times the writer of a workbook alone: `Book::modify_to_bytes` with a `CellWrites` of one cached number
//! for every cell of a rectangle, no engine and no JSON in between. It is the harness to judge a change
//! of the writer (XLSX or ODS) by the writer alone.
//!
//! ```text
//! cargo build --release -p sheetrs-io --example writer_bench
//! target/release/examples/writer_bench TEMPLATE SHEET FIRST_ROW LAST_ROW FIRST_COL LAST_COL [RUNS] [--out FILE]
//! ```
//!
//! Rows and columns are 0-based and inclusive: the formulas of the benchmark workbook of
//! `sheetrs-compiler/scripts/gen_large_fixture.py` are `Calc`, rows 1 to 300000, columns 0 to 7. The
//! template must hold a formula in each of those cells (a cell without one is still written, as a value).
//! Prints the best of the runs, the peak memory and the size of the output; `--out` keeps the last output.

use sheetrs_io::Book;
use sheetrs_io::write::{CachedResult, CellWrites, WorkbookModifications};
use std::time::Instant;

fn peak_mb() -> u64 {
    std::fs::read_to_string("/proc/self/status")
        .ok()
        .and_then(|s| {
            s.lines()
                .find(|l| l.starts_with("VmHWM"))
                .and_then(|l| l.split_whitespace().nth(1)?.parse::<u64>().ok())
        })
        .map_or(0, |kb| kb / 1024)
}

fn main() {
    let mut args: Vec<String> = std::env::args().skip(1).collect();
    let out = args.iter().position(|a| a == "--out").map(|i| {
        let v = args.remove(i + 1);
        args.remove(i);
        v
    });
    if args.len() < 6 {
        eprintln!(
            "usage: writer_bench TEMPLATE SHEET FIRST_ROW LAST_ROW FIRST_COL LAST_COL [RUNS] [--out FILE]"
        );
        std::process::exit(2);
    }
    let n = |i: usize| args[i].parse::<u32>().expect("a number");
    let (sheet, r0, r1, c0, c1) = (args[1].clone(), n(2), n(3), n(4), n(5));
    let runs: usize = args
        .get(6)
        .map_or(3, |s| s.parse().expect("a number of runs"));
    let template = Book::from_file(&args[0]).expect("the template opens");

    let t = Instant::now();
    let mut writes = CellWrites::new();
    for row in r0..=r1 {
        for col in c0..=c1 {
            writes.result(
                sheet.clone(),
                row,
                col,
                CachedResult::Number(f64::from(row) * 1.25 + f64::from(col)),
            );
        }
    }
    let cells = writes.len();
    println!("writes built: {cells} cells in {:?}", t.elapsed());
    let replacements = writes.into_sheets();

    let mut best = f64::MAX;
    let mut size = 0;
    for _ in 0..runs {
        let mods = WorkbookModifications {
            replacements: replacements.clone(),
            ..WorkbookModifications::default()
        };
        let t = Instant::now();
        let bytes = template.modify_to_bytes(&mods).expect("the writer works");
        best = best.min(t.elapsed().as_secs_f64());
        size = bytes.len();
        if let Some(path) = &out {
            std::fs::write(path, &bytes).expect("the output is written");
        }
    }
    println!(
        "modify_to_bytes: best of {runs} {best:.3} s, {:.0} ns per cell, output {:.1} MB, peak {} MB",
        best * 1e9 / cells as f64,
        size as f64 / 1e6,
        peak_mb()
    );
}
