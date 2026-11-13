//! A search by value that also looks at the results of formulas, over a workbook whose formulas
//! refer to many different cells: the format inference reads the cells it needs, and it must not
//! read the file once per reference (it did not finish on 2.4 million formulas).

use sheetrs_io::Book;
use sheetrs_protocol::{CellRecord, Find, Query};
use sheetrs_query::BookQuery;
use std::fmt::Write as _;
use std::io::Write;

/// An XLSX with the numbers `1..=n` in `A1:A{n}` and, in column B, `=A{k}*1+0*NOW()`: volatile, so
/// each one is worked out, and each one refers to a different cell.
fn workbook(n: u32) -> Vec<u8> {
    let mut rows = String::new();
    for k in 1..=n {
        let _ = write!(
            rows,
            "<row r=\"{k}\"><c r=\"A{k}\"><v>{k}</v></c><c r=\"B{k}\"><f>A{k}*1+0*NOW()</f><v>{k}</v></c></row>"
        );
    }
    let parts = [
        (
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>".to_string(),
        ),
        (
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>".to_string(),
        ),
        (
            "xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>".to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>".to_string(),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>{rows}</sheetData></worksheet>"),
        ),
    ];
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn formulas_that_refer_to_many_cells_are_searched_by_their_results() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("many.xlsx");
    std::fs::write(&path, workbook(40)).unwrap();
    let cells = Book::from_file(&path)
        .unwrap()
        .query(&Query {
            find: Some(Find {
                include_formulas: true,
                ..Find::literal("7")
            }),
            ..Query::default()
        })
        .unwrap()
        .cells;
    let found: Vec<String> = cells.iter().map(CellRecord::cell).collect();
    // `7`, `17`, `27` and `37` in both columns.
    assert_eq!(
        found,
        ["A7", "B7", "A17", "B17", "A27", "B27", "A37", "B37"],
        "{found:?}"
    );
}
