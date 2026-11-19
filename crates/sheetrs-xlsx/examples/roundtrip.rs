//! Reads an XLSX workbook and writes it back: `roundtrip <input> <output>`; prints the times.

use std::fs::File;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let (Some(input), Some(output)) = (arguments.next(), arguments.next()) else {
        return Err("usage: roundtrip <input> <output>".into());
    };
    let started = Instant::now();
    let workbook = sheetrs_xlsx::read(File::open(input)?)?;
    let read = started.elapsed();
    let started = Instant::now();
    sheetrs_xlsx::write(&workbook, File::create(output)?)?;
    println!("read {read:?}, write {:?}", started.elapsed());
    Ok(())
}
