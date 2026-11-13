//! A run with `set_recalculate(false)` keeps the results the file caches (issue #126): the cells no
//! operation touches keep what the application that wrote the file computed, and the workbook is
//! not marked to be recalculated when it is opened. By default every formula is calculated again.
//!
//! The workbook has `Sheet1!B1` = `=C1` (C1 empty) cached as 0 and `Sheet1!B2` = `=IF(B1<>"","x",0)`
//! cached as `x`, as Excel caches them (`0<>""` is TRUE there); the engine, like `LibreOffice`, gives
//! 0 for B2. `Data` has no formula and nothing that `Sheet1` refers to.

use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

fn zip(path: &Path, files: &[(&str, &str)]) {
    let mut w = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    for (name, body) in files {
        w.start_file(*name, zip::write::SimpleFileOptions::default())
            .unwrap();
        w.write_all(body.as_bytes()).unwrap();
    }
    w.finish().unwrap();
}

fn xlsx(path: &Path) {
    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="B1"><f>C1</f><v>0</v></c></row>
<row r="2"><c r="B2" t="str"><f>IF(B1&lt;&gt;"","x",0)</f><v>x</v></c></row>
</sheetData></worksheet>"#;
    let data = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>
</sheetData></worksheet>"#;
    zip(
        path,
        &[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet1),
            ("xl/worksheets/sheet2.xml", data),
        ],
    );
}

fn ods(path: &Path) {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"><office:body><office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row><table:table-cell/><table:table-cell table:formula="of:=[.C1]" office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell/><table:table-cell table:formula="of:=IF([.B1]&lt;&gt;&quot;&quot;;&quot;x&quot;;0)" office:value-type="string" office:string-value="x"><text:p>x</text:p></table:table-cell></table:table-row>
</table:table>
<table:table table:name="Data">
<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>
</table:table></office:spreadsheet></office:body></office:document-content>"#;
    zip(
        path,
        &[
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            (
                "META-INF/manifest.xml",
                r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
            ),
            ("content.xml", content),
        ],
    );
}

/// `Sheet1!B1:B2`, as the file caches them.
fn results(path: &Path) -> Vec<CellValue> {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "Sheet1").unwrap();
    (0..2)
        .map(|row| {
            sheet
                .cells
                .iter()
                .find(|(at, _)| (at.0, at.1) == (row, 1))
                .map_or(CellValue::Empty, |(_, c)| c.value.clone())
        })
        .collect()
}

fn text(value: &str) -> CellValue {
    CellValue::Text(value.into())
}

fn marked(path: &Path) -> bool {
    if path.extension().is_some_and(|e| e == "xlsx") {
        let mut archive = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut xml = String::new();
        archive
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml.contains("fullCalcOnLoad")
    } else {
        // An ODS has no mark: it drops the cached results of its formulas.
        results(path).iter().any(|v| matches!(v, CellValue::Empty))
    }
}

fn each_format(check: impl Fn(&str, &Path, &Path)) {
    for (extension, make) in [("xlsx", xlsx as fn(&Path)), ("ods", ods as fn(&Path))] {
        let dir = tempfile::tempdir().unwrap();
        let book: PathBuf = dir.path().join(format!("book.{extension}"));
        let out = dir.path().join(format!("out.{extension}"));
        make(&book);
        assert_eq!(
            results(&book),
            [CellValue::Number(0.0), text("x")],
            "{extension}: the fixture"
        );
        check(extension, &book, &out);
    }
}

#[test]
fn by_default_an_edit_calculates_every_formula() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        assert!(editor.recalculate());
        editor.delete_columns(["Data!D".parse().unwrap()]);
        editor.write(out).unwrap();
        // The engine, as LibreOffice, gives 0 where Excel cached `x`.
        assert_eq!(
            results(out),
            [CellValue::Number(0.0), CellValue::Number(0.0)],
            "{extension}"
        );
    });
}

#[test]
fn an_edit_without_recalculation_keeps_the_cached_results() {
    each_format(|extension, book, out| {
        for edit in 0..3 {
            let mut editor = Book::from_file(book).unwrap();
            editor.set_recalculate(false);
            match edit {
                0 => editor.delete_columns(["Data!D".parse().unwrap()]),
                1 => editor.delete_rows(["Data!3".parse().unwrap()]),
                _ => editor.formula_to_value(["Data".parse().unwrap()]),
            };
            let report = editor.write(out).unwrap();
            assert!(report.diagnostics.is_empty(), "{extension}: {report:?}");
            assert_eq!(
                results(out),
                [CellValue::Number(0.0), text("x")],
                "{extension}, edit {edit}"
            );
            assert!(!marked(out), "{extension}, edit {edit}: marked");
        }
    });
}

#[test]
fn a_conversion_without_recalculation_converts_the_cached_results() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        editor.set_recalculate(false);
        // The rows that the results of the formulas read are deleted, and then converted: the
        // values are the ones the file caches.
        editor
            .delete_rows(["Sheet1!1".parse().unwrap()])
            .formula_to_value(["Sheet1".parse().unwrap()]);
        editor.write(out).unwrap();
        let workbook = Book::from_file(out).unwrap().workbook().unwrap();
        let sheet = workbook.sheets.iter().find(|s| s.name == "Sheet1").unwrap();
        // B2 is in the row 1 now, a value and no formula.
        let cell = sheet
            .cells
            .iter()
            .find(|(at, _)| (at.0, at.1) == (0, 1))
            .map(|(_, c)| c.clone())
            .unwrap();
        assert_eq!(cell.value, text("x"), "{extension}");
        assert!(cell.formula.is_none(), "{extension}");
    });
}

#[test]
fn the_calculate_operation_runs_whatever_the_option() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        editor
            .set_recalculate(false)
            .delete_columns(["Data!D".parse().unwrap()]);
        editor.calculate(Calculation::default());
        editor.write(out).unwrap();
        assert_eq!(
            results(out),
            [CellValue::Number(0.0), CellValue::Number(0.0)],
            "{extension}"
        );
    });
}

#[test]
fn the_option_stays_off_in_the_explanation_of_the_plan() {
    each_format(|extension, book, _| {
        let mut editor = Book::from_file(book).unwrap();
        editor.delete_columns(["Data!D".parse().unwrap()]);
        assert!(editor.explain().contains("recalculate"), "{extension}");
        editor.set_recalculate(false);
        assert!(!editor.explain().contains("recalculate"), "{extension}");
    });
}
