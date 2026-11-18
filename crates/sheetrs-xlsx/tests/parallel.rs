use std::error::Error as StdError;
use std::io::Cursor;

use sheetrs_xlsx::{read, read_parallel};

mod common;
use common::single_sheet_package_with;

fn row(n: u32) -> String {
    format!(
        r#"<row r="{n}"><c r="A{n}"><v>{n}</v></c><c r="B{n}" t="str"><f>A{n}*2</f><v>x{n}</v></c><c r="C{n}" t="inlineStr"><is><t>t{n}</t></is></c></row>"#
    )
}

/// A worksheet of several megabytes, so that it is cut into many blocks of rows.
fn large_rows(count: u32) -> String {
    (1..=count).map(row).collect()
}

#[test]
fn the_parallel_reader_returns_the_cells_of_the_sequential_one() -> Result<(), Box<dyn StdError>> {
    let package = single_sheet_package_with(&large_rows(80_000), None, "")?.into_inner();
    let sequential = read(Cursor::new(package.clone()))?;
    for workers in [2, 3, 8] {
        let parallel = read_parallel(Cursor::new(package.clone()), workers)?;
        assert_eq!(parallel.sheets().len(), sequential.sheets().len());
        for (left, right) in parallel.sheets().iter().zip(sequential.sheets()) {
            assert_eq!(left.name(), right.name());
            assert_eq!(left.len(), 240_000, "{workers} workers");
            assert!(left.cells().eq(right.cells()), "{workers} workers");
        }
    }
    Ok(())
}
