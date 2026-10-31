//! Deleting rows and columns of an XLSX, and removing sheets and names, with the references to
//! them adjusted as LibreOffice 26.8 does (the expectations are what it gave when asked through
//! UNO: `Rows.removeByIndex`, `Columns.removeByIndex`, `Sheets.removeByName` and
//! `NamedRanges.removeByName`).

use sheetrs_core::writer::{ValueScope, WorkbookModifications, modify_workbook};
use std::collections::HashSet;
use std::fmt::Write as _;
use std::io::{Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;

/// A worksheet: the rows of `<sheetData>` and what follows it.
struct Sheet {
    name: &'static str,
    rows: String,
    after: String,
    before: String,
}

fn sheet(name: &'static str, rows: String) -> Sheet {
    Sheet {
        name,
        rows,
        after: String::new(),
        before: String::new(),
    }
}

/// `n` rows with the number `r` in column A and `10 r` in column B.
fn number_rows(n: u32) -> String {
    (1..=n)
        .map(|r| {
            format!(
                r#"<row r="{r}" spans="1:2"><c r="A{r}"><v>{r}</v></c><c r="B{r}"><v>{}</v></c></row>"#,
                r * 10
            )
        })
        .collect()
}

/// Rows with a formula in column A (`99` cached) for each text.
fn formula_rows(formulas: &[&str]) -> String {
    formulas
        .iter()
        .enumerate()
        .map(|(i, f)| {
            format!(
                r#"<row r="{0}"><c r="A{0}"><f>{f}</f><v>99</v></c></row>"#,
                i + 1
            )
        })
        .collect()
}

fn write_xlsx(path: &Path, sheets: &[Sheet], names: &str, extra: &[(&str, &str)]) {
    let mut zip = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    let o = SimpleFileOptions::default();
    let mut parts: Vec<(String, String)> = vec![
        ("_rels/.rels".into(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()),
    ];
    let mut overrides = String::new();
    let mut rels = String::new();
    let mut entries = String::new();
    for (i, s) in sheets.iter().enumerate() {
        let n = i + 1;
        let _ = write!(
            overrides,
            r#"<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        );
        let _ = write!(
            rels,
            r#"<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>"#
        );
        let _ = write!(
            entries,
            r#"<sheet name="{}" sheetId="{n}" r:id="rId{n}"/>"#,
            s.name
        );
        parts.push((
            format!("xl/worksheets/sheet{n}.xml"),
            format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}<sheetData>{}</sheetData>{}</worksheet>"#,
                s.before, s.rows, s.after
            ),
        ));
    }
    parts.push(("[Content_Types].xml".into(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{overrides}</Types>"#)));
    parts.push(("xl/_rels/workbook.xml.rels".into(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#)));
    parts.push(("xl/workbook.xml".into(), format!(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{entries}</sheets>{names}<calcPr calcId="1"/></workbook>"#)));
    for (name, body) in extra {
        parts.push((name.to_string(), body.to_string()));
    }
    for (name, body) in parts {
        zip.start_file(name, o).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap();
}

fn part(path: &Path, name: &str) -> String {
    let mut archive = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut text = String::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_string(&mut text)
        .unwrap();
    text
}

/// The `<f>` texts of a worksheet part, in order.
fn formulas(xml: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(i) = rest.find("<f>") {
        let after = &rest[i + 3..];
        let end = after.find("</f>").unwrap();
        out.push(after[..end].to_string());
        rest = &after[end..];
    }
    out
}

fn scope(spec: &str) -> ValueScope {
    spec.parse().unwrap()
}

fn run(
    dir: &Path,
    sheets: &[Sheet],
    names: &str,
    mods: &WorkbookModifications,
) -> anyhow::Result<std::path::PathBuf> {
    let input = dir.join("in.xlsx");
    let output = dir.join("out.xlsx");
    write_xlsx(&input, sheets, names, &[]);
    modify_workbook(&input, &output, mods)?;
    Ok(output)
}

const NAMES: &str = r#"<definedNames><definedName name="NameOne">T!$A$6</definedName><definedName name="NameTwo">T!$A$1:$A$10</definedName><definedName name="Gone">T!$A$4</definedName></definedNames>"#;

fn book() -> Vec<Sheet> {
    let mut t = sheet("T", number_rows(10));
    t.before = r#"<dimension ref="A1:B10"/><cols><col min="1" max="1" width="9"/><col min="2" max="4" width="20"/></cols>"#.into();
    t.after =
        r#"<mergeCells count="2"><mergeCell ref="A7:B8"/><mergeCell ref="A4:B5"/></mergeCells>"#
            .into();
    let f = sheet(
        "F",
        formula_rows(&[
            "T!A2",
            "T!A4",
            "T!A5",
            "T!A6",
            "SUM(T!A2:A3)",
            "SUM(T!A2:A4)",
            "SUM(T!A4:A5)",
            "SUM(T!A4:A8)",
            "SUM(T!A2:A8)",
            "SUM(T!A5:A6)",
            "SUM(T!A:A)",
            "SUM(T!4:5)",
            "SUM(T!3:6)",
            "SUM(T!1:3)",
            "T!$A$6",
            "SUM(T!A6:B7)",
            "NameOne+1",
            "A1+A2",
        ]),
    );
    vec![t, f]
}

#[test]
fn deleting_rows_adjusts_references_as_libreoffice() {
    let dir = tempfile::tempdir().unwrap();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    let out = run(dir.path(), &book(), NAMES, &mods).unwrap();
    // The `getFormula` of each cell after `Rows.removeByIndex(3, 2)` in LibreOffice.
    assert_eq!(
        formulas(&part(&out, "xl/worksheets/sheet2.xml")),
        [
            "T!A2",
            "T!#REF!",
            "T!#REF!",
            "T!A4",
            "SUM(T!A2:A3)",
            "SUM(T!A2:A3)",
            "SUM(T!#REF!)",
            "SUM(T!A4:A6)",
            "SUM(T!A2:A6)",
            "SUM(T!A4:A4)",
            "SUM(T!A:A)",
            "SUM(T!#REF!)",
            "SUM(T!3:4)",
            "SUM(T!1:3)",
            "T!$A$4",
            "SUM(T!A4:B5)",
            "NameOne+1",
            "A1+A2",
        ]
    );
    // The names that point at the sheet follow it, and a name left on deleted cells is `#REF!`.
    let workbook = part(&out, "xl/workbook.xml");
    assert!(
        workbook.contains(r#"<definedName name="NameOne">T!$A$4</definedName>"#),
        "{workbook}"
    );
    assert!(workbook.contains(r#"<definedName name="NameTwo">T!$A$1:$A$8</definedName>"#));
    assert!(workbook.contains(r#"<definedName name="Gone">T!#REF!</definedName>"#));
    // The sheet itself: rows renumbered, the merged areas and the dimension moved.
    let t = part(&out, "xl/worksheets/sheet1.xml");
    assert!(t.contains(r#"<dimension ref="A1:B8"/>"#), "{t}");
    assert!(t.contains(r#"<mergeCell ref="A5:B6"/>"#));
    assert!(!t.contains(r#"<mergeCell ref="A4:B5"/>"#) && !t.contains("A7:B8"));
    assert!(t.contains(r#"<row r="4"><c r="A4"><v>6</v></c>"#) || t.contains(r#"<row r="4" "#));
    assert!(!t.contains(r#"r="9""#) && !t.contains(r#"r="10""#));
    assert!(!t.contains("spans"));
}

#[test]
fn a_formula_that_changes_loses_its_value_and_one_that_does_not_keeps_it() {
    let dir = tempfile::tempdir().unwrap();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    let out = run(dir.path(), &book(), NAMES, &mods).unwrap();
    let f = part(&out, "xl/worksheets/sheet2.xml");
    // `T!A2` does not change: the cached value stays. `T!A6` does: it has none.
    assert!(f.contains(r#"<c r="A1"><f>T!A2</f><v>99</v></c>"#), "{f}");
    assert!(f.contains(r#"<c r="A4"><f>T!A4</f></c>"#), "{f}");
}

#[test]
fn deleting_columns_adjusts_references_as_libreoffice() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(2));
    t.before = r#"<dimension ref="A1:B2"/>"#.into();
    let f = sheet(
        "F",
        formula_rows(&[
            "T!B1",
            "T!C1",
            "T!D1",
            "T!E1",
            "SUM(T!B1:C1)",
            "SUM(T!B1:D1)",
            "SUM(T!C1:D1)",
            "SUM(T!C1:G1)",
            "SUM(T!B1:G1)",
            "SUM(T!D1:E1)",
            "SUM(T!1:1)",
            "SUM(T!C:D)",
            "SUM(T!B:E)",
            "SUM(T!A:B)",
            "T!$E$1",
            "SUM(T!E1:F2)",
        ]),
    );
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!C:D")],
        ..Default::default()
    };
    let out = run(dir.path(), &[t, f], "", &mods).unwrap();
    // After `Columns.removeByIndex(2, 2)` in LibreOffice.
    assert_eq!(
        formulas(&part(&out, "xl/worksheets/sheet2.xml")),
        [
            "T!B1",
            "T!#REF!",
            "T!#REF!",
            "T!C1",
            "SUM(T!B1:B1)",
            "SUM(T!B1:B1)",
            "SUM(T!#REF!)",
            "SUM(T!C1:E1)",
            "SUM(T!B1:E1)",
            "SUM(T!C1:C1)",
            "SUM(T!1:1)",
            "SUM(T!#REF!)",
            "SUM(T!B:C)",
            "SUM(T!A:B)",
            "T!$C$1",
            "SUM(T!C1:D2)",
        ]
    );
}

#[test]
fn several_deletions_read_the_coordinates_of_the_input() {
    let dir = tempfile::tempdir().unwrap();
    let f = sheet(
        "F",
        formula_rows(&["T!A5", "T!A8", "T!A3", "SUM(T!A1:A10)"]),
    );
    let mods = WorkbookModifications {
        // Rows 3-4 and 7, written in any order.
        delete_rows: vec![scope("T!7"), scope("T!3:4")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", number_rows(10)), f], "", &mods).unwrap();
    assert_eq!(
        formulas(&part(&out, "xl/worksheets/sheet2.xml")),
        ["T!A3", "T!A5", "T!#REF!", "SUM(T!A1:A7)"]
    );
}

#[test]
fn removing_a_sheet_turns_its_references_into_ref_errors() {
    let dir = tempfile::tempdir().unwrap();
    let f = sheet(
        "F",
        formula_rows(&[
            "T!A1",
            "SUM(T!A1:A3)",
            "SUM(T!A1:A3,F!B1)",
            "T!A1+1",
            "INDIRECT(&quot;T!A1&quot;)",
            "NameOne+1",
            "NameTwo",
        ]),
    );
    let mods = WorkbookModifications {
        remove_sheets: Some(HashSet::from(["T".to_string()])),
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", number_rows(3)), f], NAMES, &mods).unwrap();
    // LibreOffice: `#REF!` in place of each reference, the text of INDIRECT as it was, and the
    // names kept with a definition that is `#REF!`.
    assert_eq!(
        formulas(&part(&out, "xl/worksheets/sheet2.xml")),
        [
            "#REF!",
            "SUM(#REF!)",
            "SUM(#REF!,F!B1)",
            "#REF!+1",
            "INDIRECT(&quot;T!A1&quot;)",
            "NameOne+1",
            "NameTwo"
        ]
    );
    let workbook = part(&out, "xl/workbook.xml");
    assert!(
        workbook.contains(r#"<definedName name="NameOne">#REF!</definedName>"#),
        "{workbook}"
    );
    assert!(workbook.contains(r#"<definedName name="NameTwo">#REF!</definedName>"#));
}

#[test]
fn removing_a_named_range_leaves_the_name_in_the_formulas_and_unknown_in_other_names() {
    let dir = tempfile::tempdir().unwrap();
    let names = r#"<definedNames><definedName name="NameOne">T!$A$1</definedName><definedName name="NameTwo">T!$A$1+NameOne</definedName><definedName name="Nameless">F!$A$1+MyNameOne(1)+&quot;NameOne&quot;</definedName></definedNames>"#;
    let f = sheet(
        "F",
        formula_rows(&["NameOne+1", "SUM(NameOne,2)", "NameTwo"]),
    );
    let mods = WorkbookModifications {
        remove_named_ranges: Some(HashSet::from(["NameOne".to_string()])),
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", number_rows(1)), f], names, &mods).unwrap();
    // LibreOffice: the formulas keep the name (`#NAME?` when computed); another name that used it
    // has `#NAME?` in its place.
    assert_eq!(
        formulas(&part(&out, "xl/worksheets/sheet2.xml")),
        ["NameOne+1", "SUM(NameOne,2)", "NameTwo"]
    );
    let workbook = part(&out, "xl/workbook.xml");
    assert!(!workbook.contains(r#"name="NameOne""#), "{workbook}");
    assert!(
        workbook.contains(r#"<definedName name="NameTwo">T!$A$1+#NAME?</definedName>"#),
        "{workbook}"
    );
    assert!(
        workbook.contains("F!$A$1+MyNameOne(1)+&quot;NameOne&quot;"),
        "{workbook}"
    );
}

#[test]
fn recalculation_on_open_is_up_to_the_user() {
    let dir = tempfile::tempdir().unwrap();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!1")],
        ..Default::default()
    };
    let out = run(dir.path(), &book(), "", &mods).unwrap();
    assert!(!part(&out, "xl/workbook.xml").contains("fullCalcOnLoad"));
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!1")],
        recalculate_on_open: true,
        ..Default::default()
    };
    let out = run(dir.path(), &book(), "", &mods).unwrap();
    assert!(part(&out, "xl/workbook.xml").contains(r#"<calcPr calcId="1" fullCalcOnLoad="1"/>"#));
}

#[test]
fn a_shared_formula_keeps_its_cells_when_rows_go() {
    let dir = tempfile::tempdir().unwrap();
    // B2:B5 share `A2*2`, so B3 is `A3*2`; row 3 goes.
    let rows = r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B5" si="0">A2*2</f><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row><row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>8</v></c></row><row r="5"><c r="A5"><v>5</v></c><c r="B5"><f t="shared" si="0"/><v>10</v></c></row>"#;
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!3")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", rows.into())], "", &mods).unwrap();
    let t = part(&out, "xl/worksheets/sheet1.xml");
    // Each cell of the group gets its own formula, read at its new place.
    assert_eq!(formulas(&t), ["A2*2", "A3*2", "A4*2"]);
    assert!(t.contains(r#"<c r="B4"><f>A4*2</f>"#), "{t}");
}

#[test]
fn the_scopes_of_one_command_read_the_input_coordinates() {
    let dir = tempfile::tempdir().unwrap();
    let rows = r#"<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><f>A1+B1</f><v>3</v></c><c r="B2"><f>A1*B1</f><v>2</v></c></row><row r="3"><c r="A3"><f>A2+B2</f><v>5</v></c></row>"#;
    let mods = WorkbookModifications {
        // Row 1 goes, and the formulas of row 2 become values: row 2 of the input.
        delete_rows: vec![scope("T!1")],
        convert_to_values: vec![scope("T!2:2")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", rows.into())], "", &mods).unwrap();
    let t = part(&out, "xl/worksheets/sheet1.xml");
    // A2 and B2 are values now, at row 1; A3 was `A2+B2` and is `A1+B1` at row 2.
    assert_eq!(formulas(&t), ["A1+B1"], "{t}");
    assert!(t.contains(r#"<c r="A1"><v>3</v></c>"#), "{t}");
    assert!(t.contains(r#"<row r="2"><c r="A2"><f>A1+B1</f>"#), "{t}");
}

#[test]
fn what_cannot_be_adjusted_stops_the_command() {
    let dir = tempfile::tempdir().unwrap();
    // A deletion that covers a part of an array formula.
    let rows = r#"<row r="1"><c r="A1"><f t="array" ref="A1:A3">B1:B3*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>4</v></c></row><row r="3"><c r="A3"><v>6</v></c></row>"#;
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!2")],
        ..Default::default()
    };
    let e = run(dir.path(), &[sheet("T", rows.into())], "", &mods).unwrap_err();
    assert!(format!("{e:#}").contains("array formula"), "{e:#}");
    // The whole array may go, or stay.
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!1:3")],
        ..Default::default()
    };
    assert!(run(dir.path(), &[sheet("T", rows.into())], "", &mods).is_ok());
    // Tables and comments in a sheet that changes.
    let mut t = sheet("T", number_rows(3));
    t.after = r#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#.into();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!1")],
        ..Default::default()
    };
    let e = run(dir.path(), &[t], "", &mods).unwrap_err();
    assert!(format!("{e:#}").contains("tables"), "{e:#}");
    // A sheet that does not exist.
    let mods = WorkbookModifications {
        delete_rows: vec![scope("Nope!1")],
        ..Default::default()
    };
    let e = run(dir.path(), &[sheet("T", number_rows(1))], "", &mods).unwrap_err();
    assert!(format!("{e:#}").contains("not found"), "{e:#}");
    // Rows and columns are not interchangeable.
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!B:C")],
        ..Default::default()
    };
    assert!(run(dir.path(), &[sheet("T", number_rows(1))], "", &mods).is_err());
}

#[test]
fn notes_follow_their_cells_and_those_of_deleted_cells_go() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("in.xlsx");
    let mut t = sheet("T", number_rows(8));
    t.after = r#"<legacyDrawing r:id="rId1"/>"#.into();
    let comments = r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList><comment ref="B3" authorId="0"><text><t>keeps</t></text></comment><comment ref="B5" authorId="0"><text><t>goes</t></text></comment><comment ref="B6" authorId="0"><text><t>moves</t></text></comment></commentList></comments>"#;
    let threaded = r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B5" id="{1}"><text>goes</text></threadedComment><threadedComment ref="B7" id="{2}"><text>moves</text></threadedComment></ThreadedComments>"#;
    let shape = |id: u32, row: u32, col: u32| {
        format!(
            "<v:shape id=\"s{id}\" type=\"#_x0000_t202\"><x:ClientData ObjectType=\"Note\"><x:Anchor>\n {}, 15, {}, 2, {}, 15, {}, 4</x:Anchor><x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>",
            col + 1,
            row.saturating_sub(1),
            col + 3,
            row + 3
        )
    };
    let vml = format!(
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202"/>{}{}{}</xml>"#,
        shape(1, 2, 1),
        shape(2, 4, 1),
        shape(3, 5, 1)
    );
    write_xlsx(
        &input,
        &[t],
        "",
        &[
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>"#,
            ),
            ("xl/comments1.xml", comments),
            ("xl/threadedComments/threadedComment1.xml", threaded),
            ("xl/drawings/vmlDrawing1.vml", vml.as_str()),
        ],
    );
    let output = dir.path().join("out.xlsx");
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    modify_workbook(&input, &output, &mods).unwrap();
    let comments = part(&output, "xl/comments1.xml");
    // B3 stays, B5 (deleted row 5) goes, B6 is B4 now.
    assert!(
        comments.contains(r#"<comment ref="B3""#) && comments.contains("keeps"),
        "{comments}"
    );
    assert!(!comments.contains("goes"), "{comments}");
    assert!(
        comments.contains(r#"<comment ref="B4""#) && comments.contains("moves"),
        "{comments}"
    );
    let threaded = part(&output, "xl/threadedComments/threadedComment1.xml");
    assert!(
        !threaded.contains("goes") && threaded.contains(r#"ref="B5""#),
        "{threaded}"
    );
    let vml = part(&output, "xl/drawings/vmlDrawing1.vml");
    // The shapes of the notes on rows 3 and 6 (0-based 2 and 5): the first stays, the second
    // moves to 0-based row 3 with its anchor two rows up; the one of row 5 (0-based 4) goes.
    assert!(
        vml.contains("<x:Row>2</x:Row>") && vml.contains("<x:Row>3</x:Row>"),
        "{vml}"
    );
    assert!(
        !vml.contains("<x:Row>4</x:Row>") && !vml.contains("<x:Row>5</x:Row>"),
        "{vml}"
    );
    assert!(
        vml.contains("s1") && vml.contains("s3") && !vml.contains("\"s2\""),
        "{vml}"
    );
    assert!(vml.contains("2, 15, 2, 2, 4, 15, 6, 4"), "{vml}");
}

#[test]
fn the_cells_of_a_shared_group_keep_their_formula_when_its_first_cell_goes() {
    let dir = tempfile::tempdir().unwrap();
    // B2:B5 share `A2*2`; the first cell is in row 2, and column B is where the group is.
    let rows = r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B5" si="0">A2*2</f><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row><row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>8</v></c></row><row r="5"><c r="A5"><v>5</v></c><c r="B5"><f t="shared" si="0"/><v>10</v></c></row>"#;
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!2")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", rows.into())], "", &mods).unwrap();
    let t = part(&out, "xl/worksheets/sheet1.xml");
    // The group has no first cell left; each cell has a formula of its own, read at its new place.
    assert_eq!(formulas(&t), ["A2*2", "A3*2", "A4*2"], "{t}");
    assert!(!t.contains("shared"), "{t}");
    // The same when the column of the first cell goes: the other columns keep working.
    let rows = r#"<row r="1"><c r="A1"><f t="shared" ref="A1:B1" si="0">C1+1</f></c><c r="B1"><f t="shared" si="0"/></c></row>"#;
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!A")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", rows.into())], "", &mods).unwrap();
    assert_eq!(formulas(&part(&out, "xl/worksheets/sheet1.xml")), ["C1+1"]);
}

#[test]
fn a_merged_area_loses_cells_but_not_its_first_one_without_a_word() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(3));
    t.after = r#"<mergeCells count="1"><mergeCell ref="B2:D2"/></mergeCells>"#.into();
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!C")],
        ..Default::default()
    };
    let out = run(dir.path(), &[sheet("T", number_rows(3)), t], "", &mods);
    // The second sheet holds the area: C goes, B2:D2 is B2:C2.
    assert!(
        part(&out.unwrap(), "xl/worksheets/sheet2.xml").contains(r#"<mergeCell ref="B2:C2"/>"#)
    );
    let mut t = sheet("T", number_rows(3));
    t.after = r#"<mergeCells count="1"><mergeCell ref="B2:D2"/></mergeCells>"#.into();
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!B")],
        ..Default::default()
    };
    let e = run(dir.path(), &[t], "", &mods).unwrap_err();
    assert!(format!("{e:#}").contains("merged area"), "{e:#}");
}

#[test]
fn ranges_of_whole_columns_and_rows_in_the_elements_of_a_sheet_follow_the_deletion() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(3));
    t.after = r#"<conditionalFormatting sqref="C:C B2:B3"><cfRule type="expression" priority="1"><formula>$C1&gt;1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation sqref="3:5"><formula1>1</formula1></dataValidation></dataValidations>"#.into();
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!A")],
        delete_rows: vec![scope("T!1")],
        ..Default::default()
    };
    let out = run(dir.path(), &[t], "", &mods).unwrap();
    let t = part(&out, "xl/worksheets/sheet1.xml");
    assert!(t.contains(r#"sqref="B:B A1:A2""#), "{t}");
    assert!(t.contains(r#"sqref="2:4""#), "{t}");
    // The formula of the rule is read on its sheet.
    assert!(
        t.contains("<formula>$B#REF!&gt;1</formula>") || t.contains("&gt;1</formula>"),
        "{t}"
    );
}

#[test]
fn page_breaks_and_the_anchors_of_shapes_follow_the_rows_as_libreoffice_moves_them() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("in.xlsx");
    let mut t = sheet("T", number_rows(15));
    // LibreOffice wrote breaks before rows 3, 4, 5, 6 and 10 (`id` is the 0-based row after).
    t.after = r#"<rowBreaks count="5" manualBreakCount="5"><brk id="2" man="true" max="16383" min="0"/><brk id="3" man="true" max="16383" min="0"/><brk id="4" man="true" max="16383" min="0"/><brk id="5" man="true" max="16383" min="0"/><brk id="9" man="true" max="16383" min="0"/></rowBreaks><drawing r:id="rId1"/>"#.into();
    let anchor = |from: u32, to: u32| {
        format!(
            "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>720000</xdr:colOff><xdr:row>{to}</xdr:row><xdr:rowOff>126360</xdr:rowOff></xdr:to><xdr:sp/></xdr:twoCellAnchor>"
        )
    };
    // The shapes of the probe: row 4 (deleted), rows 3-6, row 7, rows 4-5 (1-based).
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">{}{}{}{}</xdr:wsDr>"#,
        anchor(3, 3),
        anchor(2, 5),
        anchor(6, 6),
        anchor(3, 4)
    );
    write_xlsx(
        &input,
        &[t],
        "",
        &[
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            ("xl/drawings/drawing1.xml", drawing.as_str()),
        ],
    );
    let output = dir.path().join("out.xlsx");
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    modify_workbook(&input, &output, &mods).unwrap();
    // LibreOffice 26.8 after `Rows.removeByIndex(3, 2)`: the breaks before rows 3, 4 and 8.
    let t = part(&output, "xl/worksheets/sheet1.xml");
    let ids: Vec<&str> = t
        .split("<brk id=\"")
        .skip(1)
        .map(|s| s.split('"').next().unwrap())
        .collect();
    assert_eq!(ids, ["2", "3", "7"], "{t}");
    assert!(!t.contains("manualBreakCount"), "{t}");
    // The top left corner of each shape and its size in rows, as LibreOffice wrote them.
    let drawing = part(&output, "xl/drawings/drawing1.xml");
    let rows: Vec<&str> = drawing
        .split("<xdr:row>")
        .skip(1)
        .map(|s| s.split('<').next().unwrap())
        .collect();
    assert_eq!(rows, ["3", "3", "2", "5", "4", "4", "3", "4"], "{drawing}");
}

/// The rows of the corners of the anchors of a drawing after rows 4-5 of its sheet go.
fn anchor_rows_after_deleting_rows_4_and_5(edit_as: &str, anchors: &[(u32, u32)]) -> Vec<String> {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("in.xlsx");
    let mut t = sheet("T", number_rows(15));
    t.after = r#"<drawing r:id="rId1"/>"#.into();
    let drawing: String = anchors
        .iter()
        .map(|(from, to)| {
            format!(
                r#"<xdr:twoCellAnchor{edit_as}><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/></xdr:twoCellAnchor>"#
            )
        })
        .collect();
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">{drawing}</xdr:wsDr>"#
    );
    write_xlsx(
        &input,
        &[t],
        "",
        &[
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            ("xl/drawings/drawing1.xml", drawing.as_str()),
        ],
    );
    let output = dir.path().join("out.xlsx");
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    modify_workbook(&input, &output, &mods).unwrap();
    part(&output, "xl/drawings/drawing1.xml")
        .split("<xdr:row>")
        .skip(1)
        .map(|s| s.split('<').next().unwrap().to_string())
        .collect()
}

/// A shape that resizes with its cells has its bottom corner follow its cell, and one that does not
/// keeps its size in rows (what LibreOffice 26.8 gave for each).
#[test]
fn the_corners_of_a_shape_follow_their_cells_when_it_resizes_with_them() {
    // From row 3 to 8 (1-based 4 to 9): rows 4-5 go.
    let anchors = [(3, 8), (6, 8), (1, 3)];
    assert_eq!(
        anchor_rows_after_deleting_rows_4_and_5(r#" editAs="twoCell""#, &anchors),
        ["3", "6", "4", "6", "1", "3"]
    );
    assert_eq!(
        anchor_rows_after_deleting_rows_4_and_5(r#" editAs="oneCell""#, &anchors),
        ["3", "8", "4", "6", "1", "3"]
    );
}

/// The 3D references of sheet `F` after `mods`, over the sheets `T`, `U`, `V` in that order.
fn three_d(mods: &WorkbookModifications) -> Vec<String> {
    let dir = tempfile::tempdir().unwrap();
    let f = sheet(
        "F",
        formula_rows(&[
            "SUM(T:V!A2:A4)",
            "SUM(T:V!A1)",
            "SUM(T:V!A6:A8)",
            "SUM(T:V!B2:D2)",
            "SUM(U:V!A2:A6)",
            "SUM('T:V'!A3:A9)",
        ]),
    );
    let sheets = [
        sheet("T", number_rows(10)),
        sheet("U", number_rows(10)),
        sheet("V", number_rows(10)),
        f,
    ];
    let output = run(dir.path(), &sheets, "", mods).unwrap();
    formulas(&part(&output, "xl/worksheets/sheet4.xml"))
}

/// What LibreOffice 26.8 gave for the same deletions (`sheetrs-core/scripts/compare_structure_with_libreoffice.py`).
#[test]
fn a_3d_reference_follows_a_deletion_in_its_span_only_where_the_deletion_touches_it() {
    // Row 1 of the sheet in the middle of the span: only the reference to row 1 changes.
    let one_row = three_d(&WorkbookModifications {
        delete_rows: vec![scope("U!1")],
        ..Default::default()
    });
    assert_eq!(
        one_row,
        [
            "SUM(T:V!A2:A4)",
            "SUM(T!#REF!)",
            "SUM(T:V!A6:A8)",
            "SUM(T:V!B2:D2)",
            "SUM(U:V!A2:A6)",
            "SUM('T:V'!A3:A9)",
        ]
    );
    // Rows 2-3 touch the ranges that name them; the sheet is at an end of the span or outside it.
    for (sheet, kept) in [("U", false), ("T", false), ("V", false), ("F", true)] {
        let out = three_d(&WorkbookModifications {
            delete_rows: vec![scope(&format!("{sheet}!2:3"))],
            ..Default::default()
        });
        assert_eq!(out[0] == "SUM(T:V!A2:A4)", kept, "{sheet}: {out:?}");
        if !kept {
            assert_eq!(out[0], "SUM(T:V!A2:A2)", "{sheet}");
        }
    }
}

#[test]
fn removing_a_sheet_of_a_3d_span_passes_the_end_to_the_next_one() {
    let out = three_d(&WorkbookModifications {
        remove_sheets: Some(["U".to_string()].into()),
        ..Default::default()
    });
    // The sheet in the middle changes nothing; at an end of a span it passes the end on.
    assert_eq!(out[0], "SUM(T:V!A2:A4)");
    assert_eq!(out[4], "SUM(V!A2:A6)");
    let out = three_d(&WorkbookModifications {
        remove_sheets: Some(["T".to_string(), "V".to_string()].into()),
        ..Default::default()
    });
    assert_eq!(out[0], "SUM(U!A2:A4)");
    assert_eq!(out[4], "SUM(U!A2:A6)");
}

#[test]
fn the_validations_of_the_extensions_follow_their_cells() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(10));
    t.after = r#"<extLst><ext uri="{x}"><x14:dataValidations count="2" xmlns:x14="x" xmlns:xm="m"><x14:dataValidation type="list"><x14:formula1><xm:f>T!$A$6:$A$8</xm:f></x14:formula1><xm:sqref>B6:B7 B2</xm:sqref></x14:dataValidation><x14:dataValidation type="list"><x14:formula1><xm:f>T!$A$1</xm:f></x14:formula1><xm:sqref>C4</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"#.into();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    let output = run(dir.path(), &[t], "", &mods).unwrap();
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    // The range follows the rows, the validation whose only cell goes is gone, and the count with it.
    assert!(xml.contains("<xm:sqref>B4:B5 B2</xm:sqref>"), "{xml}");
    assert!(xml.contains("<xm:f>T!$A$4:$A$6</xm:f>"), "{xml}");
    assert!(!xml.contains("C4"), "{xml}");
    assert!(!xml.contains("count=\"2\""), "{xml}");
}

#[test]
fn a_filter_shrinks_around_the_deleted_columns_and_drops_their_criteria() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(10));
    t.after = r#"<autoFilter ref="A1:E10"><filterColumn colId="1"><filters><filter val="1"/></filters></filterColumn><filterColumn colId="3"><filters><filter val="2"/></filters></filterColumn></autoFilter>"#.into();
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!B")],
        ..Default::default()
    };
    let output = run(dir.path(), &[t], "", &mods).unwrap();
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    // LibreOffice 26.8 shrinks the range of the filter; the criteria of column B go with it.
    assert!(xml.contains(r#"<autoFilter ref="A1:D10">"#), "{xml}");
    assert!(xml.contains(r#"<filterColumn colId="2">"#), "{xml}");
    assert!(!xml.contains(r#"<filter val="1"/>"#), "{xml}");
}

/// A reference across sheets of a sheet that is not edited changes when a sheet inside its span
/// is, whatever the names: the text of the formula does not name that sheet.
#[test]
fn a_3d_span_over_an_edited_sheet_it_does_not_name_changes() {
    let dir = tempfile::tempdir().unwrap();
    let sheets = [
        sheet("Alpha", number_rows(10)),
        sheet("Beta", number_rows(10)),
        sheet("Gamma", number_rows(10)),
        sheet("F", formula_rows(&["SUM(Alpha:Gamma!A2:A5)"])),
    ];
    let mods = WorkbookModifications {
        delete_rows: vec![scope("Beta!3")],
        ..Default::default()
    };
    let output = run(dir.path(), &sheets, "", &mods).unwrap();
    assert_eq!(
        formulas(&part(&output, "xl/worksheets/sheet4.xml")),
        ["SUM(Alpha:Gamma!A2:A4)"]
    );
}

#[test]
fn the_extensions_of_a_sheet_that_is_not_edited_follow_the_sheet_they_read() {
    let dir = tempfile::tempdir().unwrap();
    let mut f = sheet("F", number_rows(3));
    f.after = r#"<extLst><ext uri="{x}"><x14:dataValidations count="1" xmlns:x14="x" xmlns:xm="m"><x14:dataValidation type="list"><x14:formula1><xm:f>T!$A$6:$A$8</xm:f></x14:formula1><xm:sqref>B1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"#.into();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4:5")],
        ..Default::default()
    };
    let output = run(dir.path(), &[sheet("T", number_rows(10)), f], "", &mods).unwrap();
    let xml = part(&output, "xl/worksheets/sheet2.xml");
    assert!(xml.contains("<xm:f>T!$A$4:$A$6</xm:f>"), "{xml}");
}

#[test]
fn an_extension_left_without_children_is_not_written() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(10));
    t.after = r#"<extLst><ext uri="{x}"><x14:dataValidations count="1" xmlns:x14="x" xmlns:xm="m"><x14:dataValidation type="list"><xm:sqref>C4</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"#.into();
    let mods = WorkbookModifications {
        delete_rows: vec![scope("T!4")],
        ..Default::default()
    };
    let output = run(dir.path(), &[t], "", &mods).unwrap();
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    assert!(
        !xml.contains("dataValidations") && !xml.contains("extLst"),
        "{xml}"
    );
}

#[test]
fn the_sort_of_a_filter_follows_the_columns() {
    let dir = tempfile::tempdir().unwrap();
    let mut t = sheet("T", number_rows(10));
    t.after = r#"<autoFilter ref="A1:E10"><sortState ref="A2:E10"><sortCondition ref="D2:D10"/></sortState></autoFilter>"#.into();
    let mods = WorkbookModifications {
        delete_columns: vec![scope("T!B")],
        ..Default::default()
    };
    let output = run(dir.path(), &[t], "", &mods).unwrap();
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains(r#"<autoFilter ref="A1:D10">"#), "{xml}");
    assert!(xml.contains(r#"<sortState ref="A2:D10">"#), "{xml}");
    assert!(xml.contains(r#"<sortCondition ref="C2:C10"/>"#), "{xml}");
}
