//! A selective read gives what a full read gives, restricted to what was asked, and does less
//! work to get it. Both formats, the same expectations.

use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, CellValue, Properties, Selection, SheetState, Window};
use std::collections::BTreeMap;
use std::io::Write;
use std::path::{Path, PathBuf};

fn asset(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

/// What a cell holds, as something to compare.
type Held = (String, Option<String>, bool, Option<String>);

fn held(cell: &Cell) -> Held {
    (
        format!("{:?}", cell.value),
        cell.formula.as_deref().map(str::to_string),
        cell.is_array,
        cell.num_fmt.as_deref().map(str::to_string),
    )
}

/// A cell that says something: a value or a formula. The cells with only a format are left out
/// of the comparison, because a read that does not ask for formats does not report them.
fn meaningful(cell: &Cell) -> bool {
    !matches!(cell.value, CellValue::Empty) || cell.formula.is_some()
}

#[derive(Default)]
struct Collect {
    cells: Vec<(u16, (u32, u32), Cell)>,
}

impl CellSink for Collect {
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        self.cells.push((sheet, at, cell));
    }
}

type Cells = BTreeMap<(u16, u32, u32), Held>;

fn collected(sink: &Collect) -> Cells {
    sink.cells
        .iter()
        .filter(|(_, _, cell)| meaningful(cell))
        .map(|(sheet, at, cell)| ((*sheet, at.0, at.1), held(cell)))
        .collect()
}

/// What the full read holds for the selection: its cells in a window of a selected sheet, with
/// the properties that were not asked for taken away.
fn expected(source: &Book, selection: &Selection, windows: &[(String, Window)]) -> Cells {
    let props = selection.properties();
    // The full read in the syntax of the file, as the selective read gives it.
    let mut full = Collect::default();
    let workbook = source.workbook_into(&mut full).unwrap();
    let mut out = Cells::new();
    for (index, (row, col), cell) in &full.cells {
        let sheet = &workbook.sheets[usize::from(*index)];
        let wanted = windows.iter().any(|(name, window)| {
            name.eq_ignore_ascii_case(&sheet.name) && window.contains(*row, *col)
        });
        if !wanted {
            continue;
        }
        let projected = Cell {
            value: if props.value {
                cell.value.clone()
            } else {
                CellValue::Empty
            },
            formula: if props.formula {
                cell.formula.clone()
            } else {
                None
            },
            // A selective read says whether a cell is in an array formula when it reads the formulas.
            is_array: cell.is_array && props.formula,
            num_fmt: if props.number_format {
                cell.num_fmt.clone()
            } else {
                None
            },
        };
        if meaningful(&projected) {
            out.insert((sheet.sheet_index, *row, *col), held(&projected));
        }
    }
    out
}

fn all_properties() -> [Properties; 4] {
    [
        Properties::ALL,
        Properties::VALUE,
        Properties {
            value: false,
            formula: true,
            number_format: false,
            comment: false,
        },
        Properties {
            value: true,
            formula: true,
            number_format: true,
            comment: false,
        },
    ]
}

/// The windows worth trying on a sheet whose cells reach `(rows, cols)`.
fn windows_for(rows: u32, cols: u32) -> Vec<Vec<Window>> {
    vec![
        vec![Window::ALL],
        vec![Window::new(0, 0, 2, 1)],
        vec![Window::new(rows / 2, 0, rows, cols / 2)],
        vec![Window::new(1, 1, 1, 1)],
        vec![
            Window::new(0, 0, 1, 0),
            Window::new(rows.saturating_sub(1), 0, rows, cols),
        ],
        vec![Window::new(rows + 10, 0, rows + 20, cols)],
        vec![Window {
            first_row: 3,
            last_row: u32::MAX,
            first_col: 1,
            last_col: 2,
        }],
    ]
}

fn check_fixture(path: &Path) {
    let source = Book::from_file(path).unwrap();
    let workbook = source.workbook().unwrap();
    for (position, sheet) in workbook.sheets.iter().enumerate() {
        let (rows, cols) = sheet
            .cells
            .keys()
            .fold((0, 0), |(r, c), (row, col)| (r.max(row), c.max(col)));
        for windows in windows_for(rows, cols) {
            for props in all_properties() {
                let mut selection = Selection::new(props);
                let mut pairs = Vec::new();
                for window in &windows {
                    selection = selection.window(sheet.name.clone(), *window);
                    pairs.push((sheet.name.clone(), *window));
                }
                let mut sink = Collect::default();
                let infos = source.selection_into(&selection, &mut sink).unwrap();
                let context = format!(
                    "{} sheet {} {windows:?} {props:?}",
                    path.display(),
                    sheet.name
                );
                assert_eq!(
                    collected(&sink),
                    expected(&source, &selection, &pairs),
                    "{context}"
                );
                // The sheets listed are those of the workbook, in its order, up to the selected one.
                assert!(infos.len() > position, "{context}");
                for (info, full) in infos.iter().zip(&workbook.sheets) {
                    assert_eq!(
                        (&info.name, info.state),
                        (&full.name, full.state),
                        "{context}"
                    );
                    assert_eq!(info.index, full.sheet_index, "{context}");
                }
            }
        }
    }
    // Several sheets at once, and a sheet that does not exist.
    let mut selection = Selection::new(Properties::ALL).window("NoSuchSheet", Window::ALL);
    let mut pairs = Vec::new();
    for sheet in &workbook.sheets {
        selection = selection.window(sheet.name.clone(), Window::new(0, 0, 4, 4));
        pairs.push((sheet.name.clone(), Window::new(0, 0, 4, 4)));
    }
    let mut sink = Collect::default();
    source.selection_into(&selection, &mut sink).unwrap();
    assert_eq!(collected(&sink), expected(&source, &selection, &pairs));
}

#[test]
fn a_selection_gives_what_the_full_read_gives_of_it_in_xlsx() {
    check_fixture(&asset("../tests/lint_book.xlsx"));
    check_fixture(&asset("../sheetrs-edit/tests/assets/structure/book.xlsx"));
}

#[test]
fn a_selection_gives_what_the_full_read_gives_of_it_in_ods() {
    check_fixture(&asset("../tests/lint_book.ods"));
    check_fixture(&asset("../sheetrs-edit/tests/assets/structure/book.ods"));
}

#[test]
fn bytes_are_read_selectively_as_a_file_is() {
    let path = asset("../sheetrs-edit/tests/assets/structure/book.xlsx");
    let by_file = Book::from_file(&path).unwrap();
    let by_bytes = Book::from_bytes(std::fs::read(&path).unwrap()).unwrap();
    let selection = Selection::new(Properties::ALL).window("T", Window::new(0, 0, 5, 3));
    let (mut a, mut b) = (Collect::default(), Collect::default());
    by_file.selection_into(&selection, &mut a).unwrap();
    by_bytes.selection_into(&selection, &mut b).unwrap();
    assert_eq!(collected(&a), collected(&b));
    assert!(!collected(&a).is_empty());
}

// ---- the work not done ----------------------------------------------------------------------------

fn xlsx(sheets: &[(&str, &str)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let mut put = |name: &str, text: &str| {
        zip.start_file(name, options).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    };
    put(
        "[Content_Types].xml",
        r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
    );
    put(
        "_rels/.rels",
        r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
    );
    let names = sheets
        .iter()
        .enumerate()
        .map(|(i, (name, _))| {
            format!(
                r#"<sheet name="{name}" sheetId="{}" r:id="rId{}"/>"#,
                i + 1,
                i + 1
            )
        })
        .collect::<Vec<_>>()
        .concat();
    put(
        "xl/workbook.xml",
        &format!(
            r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{names}</sheets></workbook>"#
        ),
    );
    let rels = (0..sheets.len())
        .map(|i| format!(r#"<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>"#, i + 1))
        .collect::<Vec<_>>()
        .concat();
    put(
        "xl/_rels/workbook.xml.rels",
        &format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        ),
    );
    for (i, (_, body)) in sheets.iter().enumerate() {
        put(&format!("xl/worksheets/sheet{}.xml", i + 1), body);
    }
    zip.finish().unwrap().into_inner()
}

const HEAD: &str = r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#;

fn rows(upto: u32) -> String {
    (1..=upto)
        .map(|r| format!(r#"<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}"><f>A{r}*2</f><v>{}</v></c></row>"#, r * 2))
        .collect::<Vec<_>>()
        .concat()
}

#[test]
fn an_xlsx_sheet_is_left_after_the_last_row_asked_for() {
    // After row 5 the sheet is not well formed: a full read fails on it, and a read that stops at
    // row 3 never gets there.
    let broken = format!("{HEAD}{}</wrong>", rows(5));
    let source = Book::from_bytes(xlsx(&[("S", &broken)])).unwrap();
    assert!(source.workbook().is_err(), "the full read goes through it");
    let selection = Selection::new(Properties::ALL).window("S", Window::new(0, 0, 2, 1));
    let mut sink = Collect::default();
    source.selection_into(&selection, &mut sink).unwrap();
    let cells = collected(&sink);
    assert_eq!(cells.len(), 6, "rows 1 to 3, two cells each: {cells:?}");
    assert!(cells.keys().all(|(_, row, _)| *row <= 2));
}

#[test]
fn an_xlsx_sheet_that_is_not_selected_is_not_opened() {
    // The second sheet is not even a part of the archive: the workbook lists it and that is all.
    let mut zip_bytes = xlsx(&[
        ("A", &format!("{HEAD}{}</sheetData></worksheet>", rows(3))),
        ("B", ""),
    ]);
    {
        // Rebuild the archive without the part of the second sheet.
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(zip_bytes.clone())).unwrap();
        let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            if file.name().unwrap() == "xl/worksheets/sheet2.xml" {
                continue;
            }
            out.start_file(
                file.name().unwrap().into_owned(),
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
            std::io::copy(&mut file, &mut out).unwrap();
        }
        zip_bytes = out.finish().unwrap().into_inner();
    }
    let source = Book::from_bytes(zip_bytes).unwrap();
    let selection = Selection::new(Properties::VALUE).window("A", Window::ALL);
    let mut sink = Collect::default();
    let infos = source.selection_into(&selection, &mut sink).unwrap();
    assert_eq!(
        infos.iter().map(|i| i.name.as_str()).collect::<Vec<_>>(),
        ["A", "B"]
    );
    assert_eq!(collected(&sink).len(), 6);
    // Asking for it is another matter.
    let selection = Selection::new(Properties::VALUE).window("B", Window::ALL);
    assert!(
        source
            .selection_into(&selection, &mut Collect::default())
            .is_err()
    );
}

#[test]
fn an_xlsx_with_rows_out_of_order_loses_what_comes_after_the_stop() {
    // The documented limit: rows are expected in ascending order, as Excel requires. A row that
    // comes after a later one is not seen by a read that stops at the later one.
    let body = format!(
        r#"{HEAD}<row r="3"><c r="A3"><v>3</v></c></row><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let source = Book::from_bytes(xlsx(&[("S", &body)])).unwrap();
    assert_eq!(source.workbook().unwrap().sheets[0].cells.len(), 2);
    let selection = Selection::new(Properties::VALUE).window("S", Window::new(0, 0, 0, 0));
    let mut sink = Collect::default();
    source.selection_into(&selection, &mut sink).unwrap();
    assert!(
        collected(&sink).is_empty(),
        "A1 comes after row 3 and is not read"
    );
}

#[test]
fn what_is_not_asked_for_is_not_built() {
    let source = Book::from_bytes(xlsx(&[(
        "S",
        &format!("{HEAD}{}</sheetData></worksheet>", rows(3)),
    )]))
    .unwrap();
    let mut sink = Collect::default();
    source
        .selection_into(
            &Selection::new(Properties::VALUE).window("S", Window::ALL),
            &mut sink,
        )
        .unwrap();
    assert_eq!(sink.cells.len(), 6);
    assert!(
        sink.cells
            .iter()
            .all(|(_, _, c)| c.formula.is_none() && c.num_fmt.is_none())
    );
    let mut sink = Collect::default();
    source
        .selection_into(
            &Selection::new(Properties::ALL).window("S", Window::ALL),
            &mut sink,
        )
        .unwrap();
    assert_eq!(
        sink.cells
            .iter()
            .filter(|(_, _, c)| c.formula.is_some())
            .count(),
        3
    );
}

const ODS_HEAD: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet>"#;

fn ods(tables: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .unwrap();
    zip.start_file("content.xml", zip::write::SimpleFileOptions::default())
        .unwrap();
    zip.write_all(
        format!("{ODS_HEAD}{tables}</office:spreadsheet></office:body></office:document-content>")
            .as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap().into_inner()
}

fn ods_table(name: &str, upto: u32) -> String {
    let rows = (1..=upto)
        .map(|r| format!(r#"<table:table-row><table:table-cell office:value-type="float" office:value="{r}"><text:p>{r}</text:p></table:table-cell><table:table-cell table:formula="of:=[.A{r}]*2" office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell></table:table-row>"#, r * 2, r * 2))
        .collect::<Vec<_>>()
        .concat();
    format!(r#"<table:table table:name="{name}">{rows}</table:table>"#)
}

#[test]
fn an_ods_ends_with_the_last_table_selected() {
    // The second table is not well formed. A full read fails on it; a read of the first never
    // gets there, and one that asks for the second does.
    let tables = format!(
        "{}<table:table table:name=\"B\"><table:table-row></wrong>",
        ods_table("A", 4)
    );
    let source = Book::from_bytes(ods(&tables)).unwrap();
    assert!(source.workbook().is_err());
    let mut sink = Collect::default();
    let infos = source
        .selection_into(
            &Selection::new(Properties::ALL).window("A", Window::ALL),
            &mut sink,
        )
        .unwrap();
    assert_eq!(collected(&sink).len(), 8);
    assert_eq!(
        infos.iter().map(|i| i.name.as_str()).collect::<Vec<_>>(),
        ["A"]
    );
    assert!(
        source
            .selection_into(
                &Selection::new(Properties::ALL).window("B", Window::ALL),
                &mut Collect::default()
            )
            .is_err()
    );
}

#[test]
fn an_ods_table_is_left_after_the_last_row_asked_for() {
    let source =
        Book::from_bytes(ods(&format!("{}{}", ods_table("A", 50), ods_table("B", 2)))).unwrap();
    let selection = Selection::new(Properties::ALL)
        .window("A", Window::new(1, 0, 2, 1))
        .window("B", Window::new(0, 0, 0, 0));
    let mut sink = Collect::default();
    let infos = source.selection_into(&selection, &mut sink).unwrap();
    let cells = collected(&sink);
    let of_a: Vec<_> = cells.keys().filter(|(s, _, _)| *s == 0).collect();
    assert_eq!(
        of_a.len(),
        4,
        "rows 2 and 3 of A, two cells each: {cells:?}"
    );
    assert_eq!(
        cells.keys().filter(|(s, _, _)| *s == 1).count(),
        1,
        "A1 of B"
    );
    assert_eq!(infos.len(), 2);
    assert!(infos.iter().all(|i| i.state == SheetState::Visible));
}

/// The rows an autofilter hid, as the reader reports them: the workbooks are the ones `LibreOffice`
/// saved with the filter `key = x` on `A1:B6` and a row hidden by hand (row 10) under it.
#[derive(Default)]
struct Filtered {
    rows: Vec<(u16, u32, u32)>,
}

impl CellSink for Filtered {
    fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
    fn filtered_rows(&mut self, sheet: u16, first: u32, last: u32) {
        self.rows.push((sheet, first, last));
    }
}

fn filtered_rows_of(file: &str, selection: &Selection) -> Vec<u32> {
    let source = Book::from_file(asset(&format!("../sheetrs-query/tests/assets/{file}"))).unwrap();
    let mut sink = Filtered::default();
    source.selection_into(selection, &mut sink).unwrap();
    let mut rows: Vec<u32> = sink
        .rows
        .iter()
        .flat_map(|&(_, first, last)| first..=last)
        .collect();
    rows.sort_unstable();
    rows
}

#[test]
fn the_rows_a_filter_hides_are_reported_when_asked_and_not_the_ones_hidden_by_hand() {
    // Rows 1, 3 and 5 are the ones the filter hides (the header, `key`, is not `x` and
    // LibreOffice filters it out here too); row 10 was hidden by hand.
    for file in ["filtered.xlsx", "filtered.ods"] {
        // The windows stop at row 3: the rows reported up to there are the ones asked
        // (an XLSX tells the range of its filter after its rows, so it reads to the end and may
        // tell more).
        let asked = Selection::new(Properties::VALUE)
            .with_filtered_rows()
            .window("T", Window::new(0, 0, 2, 1));
        let rows: Vec<u32> = filtered_rows_of(file, &asked)
            .into_iter()
            .filter(|&r| r <= 2)
            .collect();
        assert_eq!(rows, vec![0, 2], "{file}");
        let all = Selection::new(Properties::VALUE)
            .with_filtered_rows()
            .window("T", Window::ALL);
        assert_eq!(filtered_rows_of(file, &all), vec![0, 2, 4], "{file}");
        let not_asked = Selection::new(Properties::VALUE).window("T", Window::ALL);
        assert!(filtered_rows_of(file, &not_asked).is_empty(), "{file}");
    }
}

#[test]
fn a_filter_kept_as_a_table_part_is_filtered_for_an_ods_only() {
    // LibreOffice writes the filter of a named database range into a table part of an XLSX, which it
    // does not read back as a filter; an ODS keeps the rows filtered.
    let asked = Selection::new(Properties::VALUE)
        .with_filtered_rows()
        .window("T", Window::ALL);
    assert!(filtered_rows_of("filtered_range.xlsx", &asked).is_empty());
    assert_eq!(filtered_rows_of("filtered_range.ods", &asked), vec![2, 4]);
}
