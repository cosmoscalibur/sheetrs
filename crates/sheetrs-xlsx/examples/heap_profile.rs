//! Reads an XLSX file under the `dhat` heap profiler and writes the profile to the given path.
//!
//! The profile holds allocation sizes and call stacks, never cell content. Build with debug
//! line tables (`CARGO_PROFILE_RELEASE_DEBUG=line-tables-only`) to see function names.

use std::env;
use std::error::Error;
use std::fs::File;

#[global_allocator]
static ALLOCATOR: dhat::Alloc = dhat::Alloc;

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let (Some(workbook_path), Some(profile_path)) = (arguments.next(), arguments.next()) else {
        return Err("usage: heap_profile <file.xlsx> <profile.json>".into());
    };
    let profiler = dhat::Profiler::builder().file_name(profile_path).build();
    let workbook = sheetrs_xlsx::read(File::open(workbook_path)?)?;
    let stats = dhat::HeapStats::get();
    println!(
        "sheets={} peak_bytes={} live_bytes={} total_bytes={} total_blocks={}",
        workbook.sheets().len(),
        stats.max_bytes,
        stats.curr_bytes,
        stats.total_bytes,
        stats.total_blocks
    );
    drop(workbook);
    drop(profiler);
    Ok(())
}
