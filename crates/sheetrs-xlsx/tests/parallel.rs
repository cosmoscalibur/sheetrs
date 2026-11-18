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
        for (left, right) in parallel.sheets().zip(sequential.sheets()) {
            assert_eq!(left.name(), right.name());
            assert_eq!(left.len(), 240_000, "{workers} workers");
            assert!(left.cells().eq(right.cells()), "{workers} workers");
        }
    }
    Ok(())
}

fn row_in_group(n: u32, anchor_ref: &str) -> String {
    let formula = if n == 1 {
        format!(r#"<f t="shared" ref="{anchor_ref}" si="0">A1*2</f>"#)
    } else {
        r#"<f t="shared" si="0"/>"#.to_owned()
    };
    format!(r#"<row r="{n}"><c r="A{n}"><v>{n}</v></c><c r="B{n}">{formula}<v>0</v></c></row>"#)
}

fn same_cells(left: &sheetrs_core::Workbook, right: &sheetrs_core::Workbook) -> bool {
    left.sheets().len() == right.sheets().len()
        && left
            .sheets()
            .zip(right.sheets())
            .all(|(a, b)| a.len() == b.len() && a.cells().eq(b.cells()))
}

#[test]
fn a_shared_group_that_spans_blocks_is_decided_over_the_whole_sheet()
-> Result<(), Box<dyn StdError>> {
    for anchor_ref in ["B1:B80000", "zz"] {
        let rows: String = (1..=80_000).map(|n| row_in_group(n, anchor_ref)).collect();
        let package = single_sheet_package_with(&rows, None, "")?.into_inner();
        let sequential = read(Cursor::new(package.clone()))?;
        for workers in [2, 8] {
            let parallel = read_parallel(Cursor::new(package.clone()), workers)?;
            assert!(same_cells(&sequential, &parallel), "{anchor_ref} {workers}");
        }
        let cells = sequential.sheets().next().map_or(0, |sheet| sheet.len());
        assert_eq!(cells, if anchor_ref == "zz" { 80_000 } else { 160_000 });
    }
    Ok(())
}

#[test]
fn a_cell_that_repeats_an_address_in_another_block_reads_as_in_sequence()
-> Result<(), Box<dyn StdError>> {
    let rows = format!(
        "{}{}",
        large_rows(80_000),
        r#"<row r="3"><c r="A3"><f>1+1</f><v>2</v></c></row>"#
    );
    let package = single_sheet_package_with(&rows, None, "")?.into_inner();
    let sequential = read(Cursor::new(package.clone()))?;
    for workers in [2, 8] {
        let parallel = read_parallel(Cursor::new(package.clone()), workers)?;
        assert!(same_cells(&sequential, &parallel), "{workers} workers");
    }
    Ok(())
}

#[test]
fn an_array_formula_that_spans_blocks_reads_as_in_sequence() -> Result<(), Box<dyn StdError>> {
    let rows = format!(
        "{}{}",
        r#"<row r="1"><c r="D1"><f t="array" ref="D1:D80000">1+1</f></c></row>"#,
        (2..=80_000).map(row).collect::<String>()
    );
    let package = single_sheet_package_with(&rows, None, "")?.into_inner();
    let sequential = read(Cursor::new(package.clone()))?;
    for workers in [2, 8] {
        let parallel = read_parallel(Cursor::new(package.clone()), workers)?;
        assert!(same_cells(&sequential, &parallel), "{workers} workers");
    }
    assert_eq!(
        sequential.sheets().next().map_or(0, |sheet| sheet.len()),
        80_000 * 4 - 3
    );
    Ok(())
}
