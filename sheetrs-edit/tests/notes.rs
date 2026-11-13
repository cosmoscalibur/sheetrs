//! Replacing in rich notes. The expectations are what `LibreOffice` 26.8 gave when it loaded the same
//! files and ran `replaceAll` on the notes (probed with `Sheet.replaceAll`, `SearchType` 2):
//!
//! - a note of an XLSX with two runs of different fonts and a phonetic run reads as the runs one
//!   after the other (`Ana: hello old`) and is saved as one run with `LibreOffice`'s font (Arial
//!   10), the author staying;
//! - a note of an ODS with a span, a run of spaces and a line break reads as `hello  old\nx` and a
//!   second paragraph as the next line; it is saved as plain paragraphs, the creator staying.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Replace, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

fn input(ext: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/replace")
        .join(format!("input.{ext}"))
}

/// A copy of `input.<ext>` with `part` rewritten by `change`.
fn with_part(ext: &str, part: &str, dir: &Path, change: impl Fn(String) -> String) -> PathBuf {
    let mut archive = zip::ZipArchive::new(std::fs::File::open(input(ext)).unwrap()).unwrap();
    let out = dir.join(format!("rich.{ext}"));
    let mut writer = zip::ZipWriter::new(std::fs::File::create(&out).unwrap());
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().unwrap().into_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        if name == part {
            bytes = change(String::from_utf8(bytes).unwrap()).into_bytes();
        }
        writer
            .start_file(&name, zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(&bytes).unwrap();
    }
    writer.finish().unwrap();
    out
}

fn notes(path: &Path) -> Vec<(String, String)> {
    let query: sheetrs_protocol::Query =
        serde_json::from_value(serde_json::json!({"properties": ["comment"]})).unwrap();
    Book::from_file(path)
        .unwrap()
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .filter_map(|c| {
            c.comment
                .clone()
                .map(|n| (format!("{}!{}", c.sheet, c.cell()), n))
        })
        .collect()
}

fn replace_hello(path: &Path, out: &Path) {
    Book::from_file(path)
        .unwrap()
        .replace(
            vec![Scope::Workbook],
            [Replace {
                find: Find {
                    search_in: SearchIn::Comment,
                    ..Find::literal("hello")
                },
                with: "bye".into(),
            }],
        )
        .write(out)
        .unwrap();
}

#[test]
fn the_runs_of_a_rich_note_of_an_xlsx_are_one_text_and_the_phonetic_run_is_not() {
    let dir = tempfile::tempdir().unwrap();
    let rich = with_part("xlsx", "xl/comments1.xml", dir.path(), |xml| {
        let start = xml.find("<comment ref=\"A1\"").unwrap();
        let end = xml[start..].find("</comment>").unwrap() + start + "</comment>".len();
        format!(
            "{}<comment ref=\"A1\" authorId=\"0\"><text><r><rPr><b/><sz val=\"12\"/><rFont val=\"Calibri\"/></rPr><t>Ana:</t></r><r><rPr><sz val=\"9\"/><rFont val=\"Arial\"/></rPr><t xml:space=\"preserve\"> hello old</t></r><rPh sb=\"0\" eb=\"1\"><t>PHON hello</t></rPh></text></comment>{}",
            &xml[..start],
            &xml[end..]
        )
    });
    assert_eq!(
        notes(&rich)
            .into_iter()
            .find(|(at, _)| at == "T!A1")
            .unwrap()
            .1,
        "Ana: hello old"
    );
    let out = dir.path().join("out.xlsx");
    replace_hello(&rich, &out);
    let a1 = notes(&out)
        .into_iter()
        .find(|(at, _)| at == "T!A1")
        .unwrap();
    assert_eq!(a1.1, "Ana: bye old");
    // One run with the font LibreOffice writes; the other notes keep what they had.
    let xml = {
        let mut archive = zip::ZipArchive::new(std::fs::File::open(&out).unwrap()).unwrap();
        let mut s = String::new();
        archive
            .by_name("xl/comments1.xml")
            .unwrap()
            .read_to_string(&mut s)
            .unwrap();
        s
    };
    assert!(xml.contains(
        "<text><r><rPr><sz val=\"10\"/><rFont val=\"Arial\"/><family val=\"2\"/></rPr><t xml:space=\"preserve\">Ana: bye old</t></r></text>"
    ));
    assert!(!xml.contains("PHON"));
}

#[test]
fn the_spans_spaces_and_line_breaks_of_a_note_of_an_ods_are_text() {
    let dir = tempfile::tempdir().unwrap();
    let rich = with_part("ods", "content.xml", dir.path(), |xml| {
        let start = xml.find("<office:annotation ").unwrap();
        let from = xml[start..].find("<text:p>").unwrap() + start;
        let to = xml[from..].find("</office:annotation>").unwrap() + from;
        format!(
            "{}<text:p><text:span text:style-name=\"T1\">hello</text:span> <text:s/>old<text:line-break/>x</text:p><text:p>second hello</text:p>{}",
            &xml[..from],
            &xml[to..]
        )
    });
    assert_eq!(
        notes(&rich)
            .into_iter()
            .find(|(at, _)| at == "T!A1")
            .unwrap()
            .1,
        "hello  old\nx\nsecond hello"
    );
    let out = dir.path().join("out.ods");
    replace_hello(&rich, &out);
    let a1 = notes(&out)
        .into_iter()
        .find(|(at, _)| at == "T!A1")
        .unwrap();
    assert_eq!(a1.1, "bye  old\nx\nsecond bye");
}
