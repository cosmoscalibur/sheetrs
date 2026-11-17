//! Whether the results a workbook caches are known to be out of date, and whether `LibreOffice` calculates something of
//! an XLSX when it opens it.
//!
//! An XLSX that was saved after a change that leaves results stale (rows deleted, a name or a sheet removed) says so: it
//! carries `fullCalcOnLoad` in `calcPr`. A conversion of formulas to values reads the results of the cells, so a workbook that
//! says its results are stale has them calculated.
//!
//! An ODS has no such mark and no scan here: `LibreOffice` calculates when it loads one the formulas whose cached result is of a
//! type and in a cell format it recalculates, and the formulas that cache no result at all, and the planner starts every plan on an
//! ODS with that calculation (the startup of the plan, `plan.rs`), in which the compiler finds the roots in the one read that
//! compiles the workbook (`sheetrs_io::reader::CellSink::recalculated_on_load` tells the cells, `sheetrs_compiler::Cache` keeps the
//! rest of what the file caches).

use crate::error::Result;
use sheetrs_io::Book;
use sheetrs_io::reader::Format;
use std::io::{Cursor, Read, Seek};

/// Whether the cached results of `source` are known to be stale: an XLSX marked to be recalculated when it is opened.
///
/// # Errors
///
/// When the file cannot be read.
pub(crate) fn input_is_stale(source: &Book) -> Result<bool> {
    // The mark of an XLSX is read in one place, the book (`Book::recalculation_marked`), which the compiler asks too.
    if source.format() == Format::Xlsx {
        return Ok(source.recalculation_marked()?);
    }
    Ok(false)
}

/// Whether `LibreOffice` calculates some formula of the XLSX `source` when it opens it, whatever result the file caches: one of the
/// functions it calculates on load (`HYPERLINK`, `CELL`, `INFO`, `FORMULATEXT`, the volatile ones, `INDIRECT`, `OFFSET`). A run that
/// keeps the cache calculates them too, as the load does.
///
/// # Errors
///
/// When the file cannot be read.
pub(crate) fn calculates_on_load(source: &Book) -> Result<bool> {
    reads_a_function_calculated_on_load(source)
}

/// The text of a formula that names a function `LibreOffice` calculates on load. A search of the bytes of the parts that hold the formulas,
/// which finds more than the calls (`SUBCELL(`): the engine decides what is stale, this only says that a pass may be needed.
const LOAD_FUNCTIONS: [&str; 11] = [
    "HYPERLINK(",
    "CELL(",
    "INFO(",
    "FORMULATEXT(",
    "FORMULA(",
    "NOW(",
    "TODAY(",
    "RAND(",
    "RANDBETWEEN(",
    "INDIRECT(",
    "OFFSET(",
];

fn reads_a_function_calculated_on_load(source: &Book) -> Result<bool> {
    fn search<R: Read + Seek>(reader: R, format: Format) -> Result<bool> {
        let finder = aho_corasick::AhoCorasick::builder()
            .ascii_case_insensitive(true)
            .build(LOAD_FUNCTIONS)
            .expect("literal patterns");
        let mut archive = zip::ZipArchive::new(reader)?;
        let names: Vec<String> = (0..archive.len())
            .filter_map(|i| {
                archive
                    .by_index(i)
                    .ok()
                    .and_then(|f| f.name().ok().map(String::from))
            })
            .filter(|name| match format {
                Format::Ods => name == "content.xml",
                _ => {
                    name.starts_with("xl/worksheets/")
                        && std::path::Path::new(name)
                            .extension()
                            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                }
            })
            .collect();
        let keep = LOAD_FUNCTIONS.iter().map(|f| f.len()).max().unwrap_or(1) - 1;
        let mut chunk = vec![0u8; 1 << 20];
        for name in names {
            let mut part = archive.by_name(&name)?;
            let mut window: Vec<u8> = Vec::new();
            loop {
                let read = part.read(&mut chunk)?;
                if read == 0 {
                    break;
                }
                window.extend_from_slice(&chunk[..read]);
                if finder.is_match(&window) {
                    return Ok(true);
                }
                let drop = window.len().saturating_sub(keep);
                window.drain(..drop);
            }
        }
        Ok(false)
    }
    match (source.path(), source.bytes()) {
        (Some(path), _) => search(std::fs::File::open(path)?, source.format()),
        (None, Some(data)) => search(Cursor::new(data.as_ref()), source.format()),
        (None, None) => unreachable!("a book is a file or bytes"),
    }
}
