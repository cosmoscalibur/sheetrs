//! A workbook that is recalculated is its own template: every formula cell of the workbook reaches the
//! program and is written back, whatever the compiler did with the formula, in an XLSX and in an ODS.
//!
//! The workbooks are written here with results that are stale on purpose (`999`, `99`, `77`): a
//! formula that is not written keeps its stale result, and a formula that is written does not. They
//! hold a plain formula, a shared formula (XLSX), a formula that does not depend on any cell (which a
//! build that folds turns into a constant) and an array formula with its interior cell.

use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::io::Write;
use std::path::Path;
use std::process::Command;

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

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
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>999</v></c><c r="C1"><f t="shared" ref="C1:C3" si="0">A1*2</f><v>999</v></c><c r="D1"><f>1+1</f><v>99</v></c><c r="E1"><f t="array" ref="E1:E2">A1:A2*2</f><v>999</v></c></row>
<row r="2"><c r="A2"><v>2</v></c><c r="C2"><f t="shared" si="0"/><v>999</v></c><c r="E2"><v>77</v></c></row>
<row r="3"><c r="A3"><v>3</v></c><c r="C3"><f t="shared" si="0"/><v>999</v></c></row>
</sheetData></worksheet>"#;
    zip(
        path,
        &[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet),
        ],
    );
}

fn ods(path: &Path) {
    let f = |formula: &str, stale: &str, extra: &str| {
        format!(
            r#"<table:table-cell table:formula="of:={formula}" office:value-type="float" office:value="{stale}"{extra}><text:p>{stale}</text:p></table:table-cell>"#
        )
    };
    let n = |v: &str| {
        format!(
            r#"<table:table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>"#
        )
    };
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"><office:body><office:spreadsheet><table:table table:name="Data">
<table:table-row>{a1}{b1}{c1}{d1}{e1}</table:table-row>
<table:table-row>{a2}<table:table-cell/>{c2}<table:table-cell/>{e2}</table:table-row>
<table:table-row>{a3}<table:table-cell/>{c3}</table:table-row>
</table:table></office:spreadsheet></office:body></office:document-content>"#,
        a1 = n("1"),
        b1 = f("[.A1]+1", "999", ""),
        c1 = f("[.A1]*2", "999", ""),
        d1 = f("1+1", "99", ""),
        e1 = f(
            "MMULT([.A1:.A2];1)",
            "999",
            r#" table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="2""#
        ),
        a2 = n("2"),
        c2 = f("[.A2]*2", "999", ""),
        e2 = r#"<table:covered-table-cell office:value-type="float" office:value="77"><text:p>77</text:p></table:covered-table-cell>"#,
        a3 = n("3"),
        c3 = f("[.A3]*2", "999", ""),
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

fn formula_cells(path: &Path) -> Vec<((u32, u32), CellValue)> {
    Book::from_file(path)
        .unwrap()
        .workbook()
        .unwrap()
        .sheets
        .into_iter()
        .flat_map(|s| s.cells.into_iter())
        .filter(|(_, c)| c.formula.is_some())
        .map(|(at, c)| (at, c.value))
        .collect()
}

#[test]
fn every_formula_cell_stays_a_formula_with_a_fresh_result_in_each_format() {
    for (ext, make) in [("xlsx", xlsx as fn(&Path)), ("ods", ods as fn(&Path))] {
        let dir = tempfile::tempdir().unwrap();
        let book = dir.path().join(format!("book.{ext}"));
        let out = dir.path().join(format!("out.{ext}"));
        make(&book);
        let before = formula_cells(&book);
        // The plain, the shared, the constant and the array cells (anchor and interior).
        assert!(before.len() >= 7, "{ext}: {before:?}");
        let run = Command::new(SHEETC)
            .args(["calculate", "-i", book.to_str().unwrap(), "--output"])
            .arg(&out)
            .output()
            .unwrap();
        assert!(
            run.status.success(),
            "{ext}: {}",
            String::from_utf8_lossy(&run.stderr)
        );
        let after = formula_cells(&out);
        // Every formula is still a formula, in the same place.
        let places =
            |v: &[((u32, u32), CellValue)]| v.iter().map(|(at, _)| *at).collect::<Vec<_>>();
        assert_eq!(places(&before), places(&after), "{ext}");
        // None keeps the stale result it had.
        let value = |at: (u32, u32)| {
            after
                .iter()
                .find(|(a, _)| *a == at)
                .map(|(_, v)| v.clone())
                .unwrap()
        };
        for (at, v) in &after {
            assert!(
                !matches!(v, CellValue::Number(n) if [999.0, 99.0, 77.0].contains(n)),
                "{ext}: {at:?} was not written: {v:?}"
            );
        }
        assert_eq!(value((0, 1)), CellValue::Number(2.0), "{ext} A1+1");
        assert_eq!(value((0, 2)), CellValue::Number(2.0), "{ext} shared C1");
        assert_eq!(value((1, 2)), CellValue::Number(4.0), "{ext} shared C2");
        assert_eq!(value((2, 2)), CellValue::Number(6.0), "{ext} shared C3");
        assert_eq!(value((0, 3)), CellValue::Number(2.0), "{ext} 1+1");
    }
}

/// The program of a workbook mode lists every formula cell of the workbook; the program of a mode
/// that gives values lists none (nothing in it is about a workbook to write).
#[test]
fn the_program_knows_every_formula_of_its_workbook() {
    use sheetrs_engine::ir::Mode;
    for (ext, make) in [("xlsx", xlsx as fn(&Path)), ("ods", ods as fn(&Path))] {
        let dir = tempfile::tempdir().unwrap();
        let book = dir.path().join(format!("book.{ext}"));
        make(&book);
        let model =
            sheetrs_compiler::loader::load(&sheetrs_io::Book::from_file(&book).unwrap()).unwrap();
        let in_the_file = formula_cells(&book).len();
        for mode in [Mode::Recalculate, Mode::Embedded] {
            let spec = sheetrs_compiler::Spec {
                mode,
                ..sheetrs_compiler::Spec::default()
            };
            let (program, _) = sheetrs_compiler::compile(&model, &spec).unwrap();
            let listed: usize = program
                .formula_cells
                .iter()
                .map(|r| r.cells().count())
                .sum();
            assert_eq!(listed, in_the_file, "{ext} {mode:?}");
        }
        let (values, _) = sheetrs_compiler::compile(
            &model,
            &sheetrs_compiler::Spec {
                outputs: vec![sheetrs_compiler::BindingSpec::parse("Data!B1")],
                ..sheetrs_compiler::Spec::default()
            },
        )
        .unwrap();
        assert!(values.formula_cells.is_empty(), "{ext}");
    }
}

/// Nothing declared, the outputs are the formula cells and nothing else, in every mode: the
/// constants are already in the workbook, so an output that held them would only read and carry them
/// again.
#[test]
fn without_declarations_the_outputs_are_the_formula_cells_in_every_mode() {
    use sheetrs_engine::ir::Mode;
    use std::collections::BTreeSet;
    for (ext, make) in [("xlsx", xlsx as fn(&Path)), ("ods", ods as fn(&Path))] {
        let dir = tempfile::tempdir().unwrap();
        let book = dir.path().join(format!("book.{ext}"));
        make(&book);
        let model =
            sheetrs_compiler::loader::load(&sheetrs_io::Book::from_file(&book).unwrap()).unwrap();
        let formulas: BTreeSet<(u32, u32)> =
            formula_cells(&book).into_iter().map(|(at, _)| at).collect();
        let cells_of = |mode: Mode| -> BTreeSet<(u32, u32)> {
            let spec = sheetrs_compiler::Spec {
                mode,
                ..sheetrs_compiler::Spec::default()
            };
            let (program, _) = sheetrs_compiler::compile(&model, &spec).unwrap();
            program
                .outputs
                .iter()
                .flat_map(|b| {
                    (b.rect.r0..=b.rect.r1)
                        .flat_map(move |r| (b.rect.c0..=b.rect.c1).map(move |c| (r, c)))
                })
                .collect()
        };
        for mode in [Mode::Values, Mode::Embedded, Mode::Recalculate] {
            assert_eq!(
                cells_of(mode),
                formulas,
                "{ext}: exactly the formula cells ({mode:?})"
            );
        }
    }
}
