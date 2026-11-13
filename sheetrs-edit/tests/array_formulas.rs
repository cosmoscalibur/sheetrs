//! An array formula whose result is a range keeps one element per member cell after a calculation
//! or an edit (issue #127): the anchor holds the top-left element and each member its own, in an
//! XLSX and in an ODS.

use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::io::Write;
use std::path::Path;

fn zip(path: &Path, files: &[(&str, &str)]) {
    let mut w = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    for (name, body) in files {
        w.start_file(*name, zip::write::SimpleFileOptions::default())
            .unwrap();
        w.write_all(body.as_bytes()).unwrap();
    }
    w.finish().unwrap();
}

/// `Sheet1!E1:F3` is the array formula `=Data!A1:B3` (members as Excel writes them, with a cached
/// value), `Data` holds `a 1 / b 2 / c 3`.
fn xlsx(path: &Path) {
    let s = |r: &str, v: &str| format!(r#"<c r="{r}" t="inlineStr"><is><t>{v}</t></is></c>"#);
    let n = |r: &str, v: &str| format!(r#"<c r="{r}"><v>{v}</v></c>"#);
    let member = |r: &str, v: &str, text: bool| {
        let t = if text { r#" t="str""# } else { "" };
        format!(r#"<c r="{r}"{t}><f ca="1"/><v>{v}</v></c>"#)
    };
    let sheet1 = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="E1" t="str" cm="1"><f t="array" ref="E1:F3">Data!A1:B3</f><v>a</v></c>{f1}</row>
<row r="2">{e2}{f2}</row>
<row r="3">{e3}{f3}</row>
</sheetData></worksheet>"#,
        f1 = member("F1", "1", false),
        e2 = member("E2", "b", true),
        f2 = member("F2", "2", false),
        e3 = member("E3", "c", true),
        f3 = member("F3", "3", false),
    );
    let data = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1">{a1}{b1}</row><row r="2">{a2}{b2}</row><row r="3">{a3}{b3}</row>
</sheetData></worksheet>"#,
        a1 = s("A1", "a"),
        b1 = n("B1", "1"),
        a2 = s("A2", "b"),
        b2 = n("B2", "2"),
        a3 = s("A3", "c"),
        b3 = n("B3", "3"),
    );
    let rel = |id: u8| {
        format!(
            r#"<Relationship Id="rId{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{id}.xml"/>"#
        )
    };
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
                &format!(
                    r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}{}</Relationships>"#,
                    rel(1),
                    rel(2)
                ),
            ),
            ("xl/worksheets/sheet1.xml", &sheet1),
            ("xl/worksheets/sheet2.xml", &data),
        ],
    );
}

/// The same workbook in an ODS: the array formula spans `E1:F3`, its members are covered cells.
fn ods(path: &Path) {
    let text = |v: &str| {
        format!(
            r#"<table:table-cell office:value-type="string" office:string-value="{v}"><text:p>{v}</text:p></table:table-cell>"#
        )
    };
    let num = |v: &str| {
        format!(
            r#"<table:table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>"#
        )
    };
    let covered_text = |v: &str| {
        format!(
            r#"<table:covered-table-cell office:value-type="string" office:string-value="{v}"><text:p>{v}</text:p></table:covered-table-cell>"#
        )
    };
    let covered_num = |v: &str| {
        format!(
            r#"<table:covered-table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:covered-table-cell>"#
        )
    };
    let blank = "<table:table-cell/>".repeat(4);
    let anchor = r#"<table:table-cell table:formula="of:=[$Data.A1:.B3]" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="3" office:value-type="string" office:string-value="a"><text:p>a</text:p></table:table-cell>"#;
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"><office:body><office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row>{blank}{anchor}{f1}</table:table-row>
<table:table-row>{blank}{e2}{f2}</table:table-row>
<table:table-row>{blank}{e3}{f3}</table:table-row>
</table:table>
<table:table table:name="Data">
<table:table-row>{a1}{b1}</table:table-row>
<table:table-row>{a2}{b2}</table:table-row>
<table:table-row>{a3}{b3}</table:table-row>
</table:table></office:spreadsheet></office:body></office:document-content>"#,
        f1 = covered_num("1"),
        e2 = covered_text("b"),
        f2 = covered_num("2"),
        e3 = covered_text("c"),
        f3 = covered_num("3"),
        a1 = text("a"),
        b1 = num("1"),
        a2 = text("b"),
        b2 = num("2"),
        a3 = text("c"),
        b3 = num("3"),
    );
    zip(
        path,
        &[
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            (
                "META-INF/manifest.xml",
                r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
            ),
            ("content.xml", &content),
        ],
    );
}

/// The values of `Sheet1!E1:F3`, row by row: text as `s:..`, numbers as `n:..`.
fn block(path: &Path) -> Vec<String> {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "Sheet1").unwrap();
    let mut out = Vec::new();
    for row in 0..3 {
        for column in 4..6 {
            let cell = sheet
                .cells
                .iter()
                .find(|(at, _)| (at.0, at.1) == (row, column))
                .map(|(_, c)| c);
            out.push(match cell.map(|c| &c.value) {
                Some(CellValue::Text(t)) => format!("s:{t}"),
                Some(CellValue::Number(n)) => format!("n:{n}"),
                other => format!("{other:?}"),
            });
        }
    }
    out
}

const WANT: [&str; 6] = ["s:a", "n:1", "s:b", "n:2", "s:c", "n:3"];

fn each_format(check: impl Fn(&str, &Path, &Path)) {
    for (extension, make) in [("xlsx", xlsx as fn(&Path)), ("ods", ods as fn(&Path))] {
        let dir = tempfile::tempdir().unwrap();
        let book = dir.path().join(format!("book.{extension}"));
        let out = dir.path().join(format!("out.{extension}"));
        make(&book);
        assert_eq!(block(&book), WANT, "{extension}: the fixture");
        check(extension, &book, &out);
    }
}

#[test]
fn a_calculation_gives_each_member_of_an_array_its_own_element() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        editor.calculate(Calculation::default());
        editor.write(out).unwrap();
        assert_eq!(block(out), WANT, "{extension}");
    });
}

#[test]
fn an_edit_elsewhere_gives_each_member_of_an_array_its_own_element() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        editor.delete_columns(["Data!D".parse().unwrap()]);
        editor.write(out).unwrap();
        assert_eq!(block(out), WANT, "{extension}");
    });
}

#[test]
fn a_conversion_elsewhere_gives_each_member_of_an_array_its_own_element() {
    each_format(|extension, book, out| {
        let mut editor = Book::from_file(book).unwrap();
        editor.formula_to_value(["Data".parse().unwrap()]);
        editor.write(out).unwrap();
        assert_eq!(block(out), WANT, "{extension}");
    });
}
