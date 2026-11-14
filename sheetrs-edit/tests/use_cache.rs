//! A run that uses the cache (`Book::with_use_cache(true)`) calculates the formulas the plan reaches,
//! and what depends on them, and leaves every other formula with the result the file caches. The
//! workbooks of `tests/assets/use_cache/` cache results the engine would not give (`from-excel` where
//! the engine gives `t`, a `ROW()` that caches 99), so a formula that kept its cache is told from one
//! that was calculated again, in each format.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use sheetrs_protocol::{Find, Replace};
use std::collections::BTreeMap;
use std::io::Read;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn asset(name: &str, extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/use_cache")
        .join(format!("{name}.{extension}"))
}

/// The values of `Calc`, by cell (`A1`).
type Values = BTreeMap<String, CellValue>;

fn calc(path: &Path) -> Values {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    sheet
        .cells
        .iter()
        .map(|((row, col), cell)| {
            let name = format!(
                "{}{}",
                char::from(b'A' + u8::try_from(col).unwrap()),
                row + 1
            );
            (name, cell.value.clone())
        })
        .collect()
}

/// The workbook after `edit`, run with the cache or without it.
fn run(
    name: &str,
    extension: &str,
    use_cache: bool,
    edit: impl Fn(&mut Book),
) -> (Values, PathBuf, tempfile::TempDir) {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let mut book = Book::from_file(asset(name, extension))
        .unwrap()
        .with_use_cache(use_cache);
    edit(&mut book);
    book.write(&out).unwrap();
    (calc(&out), out, dir)
}

fn number(n: f64) -> CellValue {
    CellValue::Number(n)
}

fn text(t: &str) -> CellValue {
    CellValue::Text(t.into())
}

fn at<'a>(values: &'a Values, cell: &str) -> &'a CellValue {
    values.get(cell).unwrap_or_else(|| panic!("no {cell}"))
}

/// The formulas no test reaches keep what the file caches.
fn untouched(values: &Values, cached: &Values, extension: &str) {
    for k in 1..=40 {
        let cell = format!("C{k}");
        assert_eq!(at(values, &cell), at(cached, &cell), "{extension} {cell}");
    }
}

/// What the file caches, as the input has it.
fn input(extension: &str, name: &str) -> Values {
    calc(&asset(name, extension))
}

#[test]
fn a_deletion_calculates_what_it_reaches_and_keeps_the_rest_as_the_file_has_it() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_rows(["Data!3".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, delete);
        let (full, ..) = run("book", extension, false, delete);
        let cached = input(extension, "book");
        untouched(&kept, &cached, extension);
        // The formulas the deletion reaches: the range that shrank (A1), what reads it (B1), the
        // whole column the row leaves (A5), `ROW()` and the name (A6), the rows a SUBTOTAL reads (B2).
        for cell in ["A1", "A5", "A6", "A8", "B1", "B2"] {
            assert_eq!(at(&kept, cell), at(&full, cell), "{extension} {cell}");
            assert_ne!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
        assert_eq!(*at(&kept, "A1"), number(180.0), "{extension}");
        assert_eq!(*at(&kept, "A5"), number(5.0), "{extension}");
        assert_eq!(*at(&kept, "B1"), number(185.0), "{extension}");
        assert_eq!(*at(&kept, "A8"), number(8.0), "{extension}");
        // The others read the same cells in other places, or nothing the deletion touched: they
        // keep what the file caches, and a full calculation changes one of them.
        for cell in ["A2", "A3", "A4", "A7", "A9"] {
            assert_eq!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
        assert_eq!(*at(&kept, "A7"), text("from-excel"), "{extension}");
        assert_eq!(*at(&full, "A7"), text("t"), "{extension}");
    }
}

#[test]
fn what_is_not_reached_is_the_cell_of_the_file_byte_for_byte() {
    let cell = |xml: &str, name: &str| {
        let start = xml.find(&format!("<c r=\"{name}\"")).unwrap();
        let end = xml[start..].find("</c>").unwrap() + "</c>".len();
        xml[start..start + end].to_string()
    };
    let sheet = |path: &Path| {
        let mut archive = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut xml = String::new();
        archive
            .by_name("xl/worksheets/sheet2.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml
    };
    let (.., out, _dir) = run("book", "xlsx", true, |book| {
        book.delete_rows(["Data!3".parse().unwrap()]);
    });
    let (got, want) = (sheet(&out), sheet(&asset("book", "xlsx")));
    for name in ["A2", "A3", "A7"] {
        assert_eq!(cell(&got, name), cell(&want, name), "{name}");
    }
}

#[test]
fn a_replacement_reaches_the_formulas_that_read_the_cell() {
    for extension in EXTENSIONS {
        let replace = |book: &mut Book| {
            book.replace(
                ["Data!A2".into()],
                [Replace {
                    find: Find::literal("20"),
                    with: "25".into(),
                }],
            );
        };
        let (kept, ..) = run("book", extension, true, replace);
        let (full, ..) = run("book", extension, false, replace);
        let cached = input(extension, "book");
        untouched(&kept, &cached, extension);
        // The cell, the range and the whole column that hold it, and what depends on them.
        for cell in ["A1", "A2", "A3", "A5", "B1", "B2"] {
            assert_eq!(at(&kept, cell), at(&full, cell), "{extension} {cell}");
        }
        assert_eq!(*at(&kept, "A2"), number(50.0), "{extension}");
        assert_eq!(*at(&kept, "A3"), number(51.0), "{extension}");
        assert_eq!(*at(&kept, "A1"), number(215.0), "{extension}");
        // A name, a position or a string that nothing changed: the file's.
        for cell in ["A4", "A6", "A7", "A8", "A9"] {
            assert_eq!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
    }
}

#[test]
fn a_deleted_name_reaches_the_formulas_that_read_it() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_names(["rate"]);
        };
        let (kept, ..) = run("book", extension, true, delete);
        let (full, ..) = run("book", extension, false, delete);
        let cached = input(extension, "book");
        untouched(&kept, &cached, extension);
        assert_eq!(at(&kept, "A6"), at(&full, "A6"), "{extension}");
        assert_ne!(at(&kept, "A6"), at(&cached, "A6"), "{extension}");
        for cell in ["A1", "A2", "A7", "A8", "B1"] {
            assert_eq!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
    }
}

#[test]
fn hidden_rows_reach_subtotal() {
    for extension in EXTENSIONS {
        let hide = |book: &mut Book| {
            book.hide_rows(["Data!2".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, hide);
        let (full, ..) = run("book", extension, false, hide);
        let cached = input(extension, "book");
        untouched(&kept, &cached, extension);
        assert_eq!(*at(&kept, "B2"), number(190.0), "{extension}");
        assert_eq!(at(&kept, "B2"), at(&full, "B2"), "{extension}");
        for cell in ["A1", "A2", "A7", "A8"] {
            assert_eq!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
    }
}

#[test]
fn a_workbook_that_reads_cells_at_run_time_is_calculated_whole() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_rows(["Data!3".parse().unwrap()]);
        };
        let (kept, ..) = run("indirect", extension, true, delete);
        let (full, ..) = run("indirect", extension, false, delete);
        assert_eq!(kept, full, "{extension}");
        assert_eq!(*at(&kept, "A7"), text("t"), "{extension}");
    }
}

#[test]
fn a_replacement_and_a_deletion_together_are_calculated_whole() {
    for extension in EXTENSIONS {
        let both = |book: &mut Book| {
            book.replace(
                ["Data!A2".into()],
                [Replace {
                    find: Find::literal("20"),
                    with: "25".into(),
                }],
            );
            book.delete_rows(["Data!5".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, both);
        let (full, ..) = run("book", extension, false, both);
        assert_eq!(kept, full, "{extension}");
    }
}

/// The report of issue 126: an edit on another sheet replaced the results `Excel` cached with the
/// engine's (`=IF(B1<>"","x",0)` is `x` for Excel and `0` for the engine).
#[test]
fn an_edit_that_reaches_no_formula_leaves_every_result_as_the_file_has_it() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let (kept, ..) = run("repro", extension, true, delete);
        let (full, ..) = run("repro", extension, false, delete);
        assert_eq!(kept, input(extension, "repro"), "{extension}");
        assert_eq!(*at(&kept, "B2"), text("x"), "{extension}");
        assert_eq!(*at(&full, "B2"), number(0.0), "{extension}");
    }
}

#[test]
fn a_workbook_the_file_says_is_stale_is_calculated_whole() {
    // An XLSX marked to be recalculated when it is opened caches nothing that is to be trusted.
    let dir = tempfile::tempdir().unwrap();
    let stale = dir.path().join("stale.xlsx");
    let mut archive =
        zip::ZipArchive::new(std::fs::File::open(asset("book", "xlsx")).unwrap()).unwrap();
    let mut out = zip::ZipWriter::new(std::fs::File::create(&stale).unwrap());
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).unwrap();
        let name = String::from_utf8_lossy(entry.name_raw()).into_owned();
        let mut data = Vec::new();
        entry.read_to_end(&mut data).unwrap();
        if name == "xl/workbook.xml" {
            let xml = String::from_utf8(data).unwrap();
            data = xml
                .replace("<calcPr ", "<calcPr fullCalcOnLoad=\"1\" ")
                .into_bytes();
        }
        out.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut out, &data).unwrap();
    }
    out.finish().unwrap();
    let run = |use_cache: bool| {
        let target = dir.path().join(format!("out{use_cache}.xlsx"));
        let mut book = Book::from_file(&stale).unwrap().with_use_cache(use_cache);
        book.delete_rows(["Data!3".parse().unwrap()]);
        book.write(&target).unwrap();
        calc(&target)
    };
    assert_eq!(run(true), run(false));
}

#[test]
fn the_plan_says_that_it_uses_the_cache() {
    let book = Book::from_file(asset("book", "xlsx")).unwrap();
    let mut with = book.clone().with_use_cache(true);
    let mut without = book;
    for editor in [&mut with, &mut without] {
        editor.delete_rows(["Data!3".parse().unwrap()]);
    }
    assert!(with.explain().contains("use the cache"));
    assert!(!without.explain().contains("use the cache"));
}

#[test]
fn a_deleted_column_reaches_the_formulas_that_read_it() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!B".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, delete);
        let (full, ..) = run("book", extension, false, delete);
        let cached = input(extension, "book");
        untouched(&kept, &cached, extension);
        // `Data!B1` is gone: the formula that read it is an error, whatever it cached.
        assert_eq!(at(&kept, "A7"), at(&full, "A7"), "{extension}");
        assert_ne!(at(&kept, "A7"), at(&cached, "A7"), "{extension}");
        // `Data!C1` is `Data!B1` now: the formulas that read it keep their results.
        assert_eq!(at(&kept, "A2"), at(&cached, "A2"), "{extension}");
    }
}

/// The cells of an array formula are calculated together: the writer drops the result of the first
/// one when the range it reads shrinks, and the members that keep theirs are not left behind.
#[test]
fn the_cells_of_an_array_formula_are_calculated_together() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_rows(["Data!2".parse().unwrap()]);
        };
        let (kept, ..) = run("array", extension, true, delete);
        let (full, ..) = run("array", extension, false, delete);
        assert_eq!(kept, full, "{extension}");
        assert_eq!(*at(&kept, "E1"), number(10.0), "{extension}");
        // The block is three rows, the range two: the third row of the block is not the old 30.
        assert_ne!(*at(&kept, "E3"), number(30.0), "{extension}");
        let cached = input(extension, "array");
        for k in 1..=8 {
            let cell = format!("H{k}");
            assert_eq!(at(&kept, &cell), at(&cached, &cell), "{extension} {cell}");
        }
    }
}
