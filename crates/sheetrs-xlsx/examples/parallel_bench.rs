//! Times the sequential reader and the block-parallel prototype on one file and checks that
//! both produce the same cells. Prints timings and a verdict, never content.

use std::env;
use std::error::Error;
use std::fs::File;
use std::time::{Duration, Instant};

use sheetrs_core::Workbook;

const REPETITIONS: usize = 3;

fn best_of(
    mut read: impl FnMut() -> Result<Workbook, Box<dyn Error>>,
) -> Result<(Duration, Workbook), Box<dyn Error>> {
    let mut best = Duration::MAX;
    let mut last = None;
    for _ in 0..REPETITIONS {
        let started = Instant::now();
        let workbook = read()?;
        best = best.min(started.elapsed());
        last = Some(workbook);
    }
    Ok((best, last.ok_or("no repetitions")?))
}

fn same(left: &Workbook, right: &Workbook) -> bool {
    left.sheets().len() == right.sheets().len()
        && left
            .sheets()
            .zip(right.sheets())
            .all(|(a, b)| a.name() == b.name() && a.len() == b.len() && a.cells().eq(b.cells()))
}

/// Inflates every part of the package and drops the bytes: the floor of any reader.
fn inflate_only(path: &str) -> Result<(Duration, u64), Box<dyn Error>> {
    let mut best = Duration::MAX;
    let mut bytes = 0;
    for _ in 0..REPETITIONS {
        let mut archive = zip::ZipArchive::new(File::open(path)?)?;
        let started = Instant::now();
        bytes = 0;
        for index in 0..archive.len() {
            bytes += std::io::copy(&mut archive.by_index(index)?, &mut std::io::sink())?;
        }
        best = best.min(started.elapsed());
    }
    Ok((best, bytes))
}

fn main() -> Result<(), Box<dyn Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: parallel_bench <file.xlsx>")?;
    let (inflating, inflated_bytes) = inflate_only(&path)?;
    println!(
        "inflate only          {:>6} ms  ({} MB inflated)",
        inflating.as_millis(),
        inflated_bytes / 1_000_000
    );
    let (sequential, reference) = best_of(|| Ok(sheetrs_xlsx::read(File::open(&path)?)?))?;
    println!("sequential            {:>6} ms", sequential.as_millis());
    for workers in [2, 4, 8, 16] {
        let (elapsed, workbook) =
            best_of(|| Ok(sheetrs_xlsx::read_parallel(File::open(&path)?, workers)?))?;
        println!(
            "parallel, {workers:>2} workers {:>6} ms  speedup x{:.2}  identical cells: {}",
            elapsed.as_millis(),
            sequential.as_secs_f64() / elapsed.as_secs_f64(),
            same(&reference, &workbook)
        );
    }
    Ok(())
}
