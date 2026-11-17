//! Workbooks made in the tests, in the two formats, from cells at positions of the A1 notation: a
//! constant, a text, or a formula with the result it caches, written in the notation of each file.
//! Nothing else is in the file (no styles, no `<dimension>`, no marks), so it is the workbook of
//! a program that caches its results and says nothing else.

// Each test file uses part of the helpers; and the XML of a few cells is built with `format!` for
// readability, which costs nothing that matters in a test.
#![allow(dead_code, clippy::format_collect, clippy::format_push_string)]

use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::io::Write;
use std::path::{Path, PathBuf};

pub const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

/// A cell of a test workbook.
pub enum Cell {
    Number(f64),
    Text(&'static str),
    /// The formula in Excel's notation and in the one of an ODS (the text after `of:=`), and the
    /// number it caches.
    Formula(&'static str, &'static str, f64),
    /// What the XLSX holds inside the `<c>` element (a shared or array formula) and the element of
    /// the ODS.
    Raw(&'static str, &'static str),
}

/// A sheet: its name, its cells (`"B2"` and what it holds) and the rows it hides (1-based).
pub struct Sheet {
    pub name: &'static str,
    pub cells: Vec<(&'static str, Cell)>,
    pub hidden_rows: Vec<u32>,
}

pub fn sheet(name: &'static str, cells: Vec<(&'static str, Cell)>) -> Sheet {
    Sheet {
        name,
        cells,
        hidden_rows: Vec::new(),
    }
}

/// A formula that reads the same in both notations.
pub fn same(text: &'static str, cached: f64) -> Cell {
    Cell::Formula(text, text, cached)
}

/// `(row, column)`, 0-based, of a cell in the A1 notation (letters then digits).
pub fn position(a1: &str) -> (u32, u32) {
    let letters = a1.chars().take_while(char::is_ascii_alphabetic).count();
    let column = a1[..letters].bytes().fold(0, |n, b| {
        n * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1)
    });
    (a1[letters..].parse::<u32>().unwrap() - 1, column - 1)
}

fn column_letters(column: u32) -> String {
    let mut n = column + 1;
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'A' + u8::try_from(n % 26).unwrap());
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap()
}

fn escaped(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn sorted(sheet: &Sheet) -> Vec<((u32, u32), &Cell)> {
    let mut cells: Vec<((u32, u32), &Cell)> = sheet
        .cells
        .iter()
        .map(|(a1, c)| (position(a1), c))
        .collect();
    cells.sort_by_key(|(at, _)| *at);
    cells
}

fn write_xlsx(path: &Path, sheets: &[Sheet], names: &[(&str, &str)]) {
    let mut zip = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    let options = zip::write::SimpleFileOptions::default();
    let mut part = |name: &str, text: &str| {
        zip.start_file(name, options).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    };
    let overrides: String = (1..=sheets.len())
        .map(|i| format!(r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#))
        .collect();
    part(
        "[Content_Types].xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{overrides}</Types>"#
        ),
    );
    part(
        "_rels/.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
    );
    let tabs: String = sheets
        .iter()
        .enumerate()
        .map(|(i, s)| {
            format!(
                r#"<sheet name="{}" sheetId="{id}" r:id="rId{id}"/>"#,
                escaped(s.name),
                id = i + 1
            )
        })
        .collect();
    let defined: String = if names.is_empty() {
        String::new()
    } else {
        let all: String = names
            .iter()
            .map(|(name, text)| {
                format!(
                    r#"<definedName name="{name}">{}</definedName>"#,
                    escaped(text)
                )
            })
            .collect();
        format!("<definedNames>{all}</definedNames>")
    };
    part(
        "xl/workbook.xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{tabs}</sheets>{defined}</workbook>"#
        ),
    );
    let rels: String = (1..=sheets.len())
        .map(|i| format!(r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>"#))
        .collect();
    part(
        "xl/_rels/workbook.xml.rels",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        ),
    );
    for (i, s) in sheets.iter().enumerate() {
        let cells = sorted(s);
        let last = cells
            .iter()
            .map(|(at, _)| at.0 + 1)
            .chain(s.hidden_rows.iter().copied())
            .max()
            .unwrap_or(0);
        let mut rows = String::new();
        for row in 0..last {
            let hidden = if s.hidden_rows.contains(&(row + 1)) {
                r#" hidden="1""#
            } else {
                ""
            };
            let inner: String = cells
                .iter()
                .filter(|(at, _)| at.0 == row)
                .map(|((r, c), cell)| {
                    let a1 = format!("{}{}", column_letters(*c), r + 1);
                    match cell {
                        Cell::Number(n) => format!(r#"<c r="{a1}"><v>{n}</v></c>"#),
                        Cell::Text(t) => format!(
                            r#"<c r="{a1}" t="inlineStr"><is><t>{}</t></is></c>"#,
                            escaped(t)
                        ),
                        Cell::Formula(text, _, cached) => {
                            format!(r#"<c r="{a1}"><f>{}</f><v>{cached}</v></c>"#, escaped(text))
                        }
                        Cell::Raw(xlsx, _) => format!(r#"<c r="{a1}">{xlsx}</c>"#),
                    }
                })
                .collect();
            if !inner.is_empty() || !hidden.is_empty() {
                rows.push_str(&format!(r#"<row r="{}"{hidden}>{inner}</row>"#, row + 1));
            }
        }
        part(
            &format!("xl/worksheets/sheet{}.xml", i + 1),
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>"#
            ),
        );
    }
    zip.finish().unwrap();
}

fn write_ods(path: &Path, sheets: &[Sheet], names: &[(&str, &str)]) {
    let mut zip = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    zip.start_file(
        "mimetype",
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
    )
    .unwrap();
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .unwrap();
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("META-INF/manifest.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#).unwrap();
    let mut tables = String::new();
    for s in sheets {
        let cells = sorted(s);
        let last = cells
            .iter()
            .map(|(at, _)| at.0 + 1)
            .chain(s.hidden_rows.iter().copied())
            .max()
            .unwrap_or(0);
        let width = cells.iter().map(|(at, _)| at.1 + 1).max().unwrap_or(0);
        let mut rows = String::new();
        for row in 0..last {
            let visibility = if s.hidden_rows.contains(&(row + 1)) {
                r#" table:visibility="collapse""#
            } else {
                ""
            };
            let mut inner = String::new();
            for column in 0..width {
                let cell = cells
                    .iter()
                    .find(|(at, _)| *at == (row, column))
                    .map(|(_, c)| c);
                inner.push_str(&match cell {
                    None => "<table:table-cell/>".to_string(),
                    Some(Cell::Number(n)) => format!(
                        r#"<table:table-cell office:value-type="float" office:value="{n}"><text:p>{n}</text:p></table:table-cell>"#
                    ),
                    Some(Cell::Text(t)) => format!(
                        r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
                        escaped(t)
                    ),
                    Some(Cell::Raw(_, ods)) => (*ods).to_string(),
                    Some(Cell::Formula(_, text, cached)) => format!(
                        r#"<table:table-cell table:formula="of:={}" office:value-type="float" office:value="{cached}"><text:p>{cached}</text:p></table:table-cell>"#,
                        escaped(text)
                    ),
                });
            }
            rows.push_str(&format!(
                "<table:table-row{visibility}>{inner}</table:table-row>"
            ));
        }
        tables.push_str(&format!(
            r#"<table:table table:name="{}">{rows}</table:table>"#,
            escaped(s.name)
        ));
    }
    let named: String = if names.is_empty() {
        String::new()
    } else {
        let all: String = names
            .iter()
            .map(|(name, text)| {
                format!(
                    r#"<table:named-expression table:name="{name}" table:base-cell-address="$'{}'.$A$1" table:expression="{}"/>"#,
                    escaped(sheets[0].name),
                    escaped(text)
                )
            })
            .collect();
        format!("<table:named-expressions>{all}</table:named-expressions>")
    };
    zip.start_file("content.xml", options).unwrap();
    zip.write_all(format!(r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.3"><office:body><office:spreadsheet>{tables}{named}</office:spreadsheet></office:body></office:document-content>"#).as_bytes()).unwrap();
    zip.finish().unwrap();
}

/// The workbook of `sheets` in the format of `extension`, in `dir`. `names` are defined names: the
/// text of the definition in Excel's notation for an XLSX and in the one of an ODS for an ODS.
pub fn workbook(dir: &Path, extension: &str, sheets: &[Sheet], names: &[(&str, &str)]) -> PathBuf {
    let path = dir.join(format!("book.{extension}"));
    if extension == "xlsx" {
        write_xlsx(&path, sheets, names);
    } else {
        write_ods(&path, sheets, names);
    }
    path
}

/// What the workbook at `path` holds in `sheet` at `a1`.
pub fn value(path: &Path, sheet: &str, a1: &str) -> CellValue {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    let at = position(a1);
    sheet
        .cells
        .iter()
        .find(|(p, _)| (p.0, p.1) == at)
        .map_or(CellValue::Empty, |(_, cell)| cell.value.clone())
}

pub fn number(path: &Path, sheet: &str, a1: &str) -> f64 {
    match value(path, sheet, a1) {
        CellValue::Number(n) => n,
        other => panic!("{sheet}!{a1} is {other:?}, not a number"),
    }
}

/// The formula of the cell at `a1`, as the reader gives it (Excel's notation in both formats).
pub fn formula(path: &Path, sheet: &str, a1: &str) -> Option<String> {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    let at = position(a1);
    sheet
        .cells
        .iter()
        .find(|(p, _)| (p.0, p.1) == at)
        .and_then(|(_, cell)| cell.formula.as_deref().map(str::to_string))
}
