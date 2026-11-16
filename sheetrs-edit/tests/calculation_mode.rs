//! A run that uses the cache (`Book::with_use_cache(true)`) calculates the formulas the plan reaches,
//! and what depends on them, and leaves every other formula with the result the file caches. The
//! workbooks of `tests/assets/cached/` cache results the engine would not give (`from-excel` where
//! the engine gives `t`, a `ROW()` that caches 99), so a formula that kept its cache is told from one
//! that was calculated again, in each format.

use sheetrs_edit::{BookEdit, CachedResult, Calculation, CellWrites};
use sheetrs_io::reader::CellValue;
use sheetrs_io::{Book, Volatile};
use sheetrs_protocol::{Find, Replace, SearchIn};
use std::collections::BTreeMap;
use std::io::Read;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn asset(name: &str, extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/cached")
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
    cached: bool,
    edit: impl Fn(&mut Book),
) -> (Values, PathBuf, tempfile::TempDir) {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let mut book = Book::from_file(asset(name, extension))
        .unwrap()
        .with_calculate_all(!cached);
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

/// A replacement of the text `from` by `to` in the values of `scopes`.
fn replace_in(book: &mut Book, scopes: &[&str], from: &str, to: &str) {
    book.replace(
        scopes.iter().map(|scope| (*scope).into()),
        [Replace {
            find: Find::literal(from),
            with: to.into(),
        }],
    );
}

/// What the file caches, as `LibreOffice` shows it when it opens it. An ODS that caches plain numbers in cells with no format has them
/// calculated on load (`ods_recalculation.rs`), which the run that keeps the cache does with an edit that reaches nothing.
fn input(extension: &str, name: &str) -> Values {
    if extension == "ods" {
        return run(name, extension, true, |book| {
            book.hide_columns(["Data!Z:Z".parse().unwrap()]);
        })
        .0;
    }
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
        // whole column the row leaves (A5) and the rows a SUBTOTAL reads (B2).
        for cell in ["A1", "A5", "B1", "B2"] {
            assert_eq!(at(&kept, cell), at(&full, cell), "{extension} {cell}");
            assert_ne!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
        assert_eq!(*at(&kept, "A1"), number(180.0), "{extension}");
        assert_eq!(*at(&kept, "A5"), number(5.0), "{extension}");
        assert_eq!(*at(&kept, "B1"), number(185.0), "{extension}");
        // The others read the same cells in other places, or nothing the deletion touched: they keep what the
        // file caches, as `LibreOffice` does when it saves and opens the file (the name `rate` points to
        // `Data!$C$1`, above the row; `ROW()` is in a sheet that lost no row), and a full calculation changes
        // them.
        for cell in ["A2", "A3", "A4", "A6", "A7", "A8", "A9"] {
            assert_eq!(at(&kept, cell), at(&cached, cell), "{extension} {cell}");
        }
        assert_eq!(
            *at(&kept, "A6"),
            number(if extension == "ods" { 14.0 } else { 140.0 }),
            "{extension}"
        );
        assert_eq!(
            *at(&kept, "A8"),
            number(if extension == "ods" { 8.0 } else { 99.0 }),
            "{extension}"
        );
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

/// A workbook with a formula that reads cells at run time (`INDIRECT`): `LibreOffice` calculates that formula when it opens the file, and
/// what depends on it, and every other formula keeps its result (`indirect` of the oracle), in both formats. The cells it may read hold
/// the results the file caches.
#[test]
fn a_formula_that_reads_at_run_time_is_calculated_and_the_rest_is_kept() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let (kept, ..) = run("indirect", extension, true, delete);
        let (full, ..) = run("indirect", extension, false, delete);
        let cached = input(extension, "indirect");
        // The formula that reads at run time is calculated; a text no edit reached is the one the file caches, which a full run changes.
        assert_eq!(at(&kept, "A10"), at(&full, "A10"), "{extension}");
        if extension == "xlsx" {
            assert_ne!(at(&kept, "A10"), at(&cached, "A10"), "{extension}");
        }
        assert_eq!(*at(&kept, "A7"), text("from-excel"), "{extension}");
        assert_eq!(*at(&full, "A7"), text("t"), "{extension}");
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

/// `book.xlsx` with the mark that makes the application that opens it recalculate every formula.
fn stale_xlsx(dir: &Path) -> PathBuf {
    let stale = dir.join("stale.xlsx");
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
    stale
}

#[test]
fn a_workbook_the_file_says_is_stale_is_calculated_whole() {
    // An XLSX marked to be recalculated when it is opened caches nothing that is to be trusted.
    let dir = tempfile::tempdir().unwrap();
    let stale = stale_xlsx(dir.path());
    let run = |cached: bool| {
        let target = dir.path().join(format!("out{cached}.xlsx"));
        let mut book = Book::from_file(&stale).unwrap().with_calculate_all(!cached);
        book.delete_rows(["Data!3".parse().unwrap()]);
        book.write(&target).unwrap();
        calc(&target)
    };
    assert_eq!(run(true), run(false));
}

#[test]
fn the_plan_says_that_the_cache_is_the_default() {
    let book = Book::from_file(asset("book", "xlsx")).unwrap();
    let mut with = book.clone().with_calculate_all(false);
    let mut without = book.with_calculate_all(true);
    for editor in [&mut with, &mut without] {
        editor.delete_rows(["Data!3".parse().unwrap()]);
    }
    assert!(with.explain().contains("the cache (the default)"));
    assert!(!without.explain().contains("the cache (the default)"));
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

// What follows pins each rule that decides which formulas an edit reaches: with one of them off, one
// of these fails (see `sheetrs-compiler/tests/cache.rs` for the same rules on the compiler alone).

/// What the tests below compare for a formula after an edit: with the cache, without it, and as the
/// file has it.
struct Outcome {
    kept: Values,
    full: Values,
    cached: Values,
}

impl Outcome {
    fn of(name: &str, extension: &str, edit: impl Fn(&mut Book)) -> Self {
        Self {
            kept: run(name, extension, true, &edit).0,
            full: run(name, extension, false, &edit).0,
            cached: input(extension, name),
        }
    }

    /// The formula was calculated again: it is what a run without the cache gives, and not what the
    /// file caches.
    fn calculated(&self, cell: &str, extension: &str) {
        let (kept, full, cached) = (
            at(&self.kept, cell),
            at(&self.full, cell),
            at(&self.cached, cell),
        );
        assert_eq!(kept, full, "{extension} {cell}");
        // An ODS number that `LibreOffice` calculates on load may already be what the full run gives.
        if !(extension == "ods" && matches!(cached, CellValue::Number(_) | CellValue::Boolean(_))) {
            assert_ne!(kept, cached, "{extension} {cell}");
        }
    }

    /// The formula kept the result of the file as `LibreOffice` loads it, which a run without the cache would not change (in an ODS
    /// a plain number in a cell with no format is calculated on load, see `input`).
    fn kept(&self, cell: &str, extension: &str) {
        let (kept, full, cached) = (
            at(&self.kept, cell),
            at(&self.full, cell),
            at(&self.cached, cell),
        );
        if extension == "ods" && matches!(cached, CellValue::Number(_) | CellValue::Boolean(_)) {
            // What the load calculated and the edit did not reach is the same in `full` when nothing moved it.
            assert_eq!(kept, cached, "{extension} {cell}: as it is loaded");
            return;
        }
        assert_eq!(kept, cached, "{extension} {cell}");
        assert_ne!(kept, full, "{extension} {cell}");
    }

    /// The formulas `from..=to` of `A` keep what the file has.
    fn keep_rows(&self, from: u32, to: u32, extension: &str) {
        for row in from..=to {
            self.kept(&format!("A{row}"), extension);
        }
    }
}

/// `NOW()`, `TODAY()` and `RAND()` that no edit reaches are volatile all the same: the file caches the
/// number of another moment, and they are calculated again with the clock and the seed of the book.
#[test]
fn volatile_formulas_are_calculated_though_the_edit_reaches_nothing() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            *book = book.clone().with_volatile(Volatile {
                now: Some(50000.25),
                seed: Some(7),
            });
            book.delete_columns(["Data!B".parse().unwrap()]);
        };
        let (kept, ..) = run("volatile", extension, true, delete);
        let cached = input(extension, "volatile");
        assert_eq!(*at(&kept, "A1"), number(50000.0), "{extension}");
        assert_eq!(*at(&kept, "A2"), number(50000.25), "{extension}");
        let CellValue::Number(random) = *at(&kept, "A3") else {
            panic!("{extension}: RAND is a number")
        };
        assert!((0.0..1.0).contains(&random), "{extension}: {random}");
        assert_ne!(at(&kept, "A3"), at(&cached, "A3"), "{extension}");
        // The formulas that are not volatile keep what the file caches.
        untouched(&kept, &cached, extension);
    }
}

/// A replacement of the text of a formula rewrites the formula and not the result it cached: the cell
/// is calculated again though no other formula reads it.
#[test]
fn a_replaced_formula_is_calculated_again() {
    for extension in EXTENSIONS {
        let replace = |book: &mut Book| {
            book.replace(
                ["Calc!A5".into()],
                [Replace {
                    find: Find {
                        search_in: SearchIn::Formula,
                        ..Find::literal("*2")
                    },
                    with: "*3".into(),
                }],
            );
        };
        let outcome = Outcome::of("rules", extension, replace);
        outcome.calculated("A5", extension);
        assert_eq!(*at(&outcome.kept, "A5"), number(60.0), "{extension}");
        // No other formula is reached.
        for cell in ["A1", "A2", "A3", "A4", "A6"] {
            outcome.kept(cell, extension);
        }
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// A result written into a formula (`write_cells`) is taken as current by the plan that only writes; where
/// a replacement leaves results stale, the recalculation that ends the run calculates that formula as well,
/// the same as a run without the cache, though the writer left it with the result it was given.
#[test]
fn a_formula_with_a_result_is_calculated_when_the_plan_leaves_it_stale() {
    for extension in EXTENSIONS {
        let write_and_replace = |book: &mut Book| {
            let mut writes = CellWrites::new();
            writes.result("Calc", 4, 0, CachedResult::Number(1234.0));
            book.write_cells(writes);
            replace_in(book, &["Data!B1"], "x", "y");
        };
        let outcome = Outcome::of("rules", extension, write_and_replace);
        outcome.calculated("A5", extension);
        assert_eq!(*at(&outcome.kept, "A5"), number(40.0), "{extension}");
        // The formula that nothing reaches keeps what the file caches.
        outcome.kept("A6", extension);
    }
}

/// A formula that reads whole rows of a sheet reads every column of it: it keeps its text when a column
/// goes, and its result is not the one the file caches. A whole column is not read by a column.
#[test]
fn a_deleted_column_reaches_the_formulas_that_read_whole_rows() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        outcome.calculated("A1", extension);
        assert_eq!(*at(&outcome.kept, "A1"), number(17.0), "{extension}");
        // A whole column, and what reads a cell: the file's results.
        outcome.kept("A5", extension);
        outcome.kept("A6", extension);
    }
}

/// Observed in `LibreOffice` (the file saved and opened again, both formats): a deleted column reaches what reads whole rows
/// of its sheet (A1: `SUM(Data!1:1)`), and nothing that reads other cells, whatever the function: `ROW()`, a name whose
/// target is not after the column, `SHEET()`, `ISFORMULA` and a whole column keep the result the file caches.
#[test]
fn a_deleted_column_reaches_the_whole_rows_and_keeps_positions_and_names() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        outcome.calculated("A1", extension);
        for cell in ["A2", "A3", "A4", "A5", "A6", "A7"] {
            outcome.kept(cell, extension);
        }
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// A column that goes before the target of a name moves it, and a formula that reads through the name is calculated again
/// (`rate` is `Data!$C$1`: column B moves it, column D does not).
#[test]
fn moving_the_target_of_a_name_reaches_the_formulas_that_use_it() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!B".parse().unwrap()]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        outcome.calculated("A3", extension);
        assert_eq!(*at(&outcome.kept, "A3"), number(14.0), "{extension}");
        // `ROW()` does not move with a column of another sheet.
        outcome.kept("A2", extension);
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// A sheet that goes reaches every formula of the workbook: `LibreOffice` calculates them all, the ones that read nothing of the
/// sheet, `ROW()` and the forty that no edit reaches included.
#[test]
fn a_deleted_sheet_reaches_every_formula() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_sheets(["Other"]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        for cell in ["A1", "A2", "A3", "A4", "A5", "A6"] {
            outcome.calculated(cell, extension);
        }
        assert_eq!(*at(&outcome.kept, "A2"), number(2.0), "{extension}");
        assert_eq!(*at(&outcome.kept, "A3"), number(14.0), "{extension}");
        for k in 1..=40 {
            outcome.calculated(&format!("C{k}"), extension);
        }
    }
}

/// `ISFORMULA` and `FORMULATEXT` of a cell are answered when the formula is lowered, so the formula
/// keeps no reference to the cell: what it says is stale when the cell is replaced. When the cell only moves (a deletion) the
/// result the file caches is kept, as in `LibreOffice`.
#[test]
fn formulas_that_ask_what_a_cell_holds_are_calculated_when_a_cell_changes() {
    for extension in EXTENSIONS {
        let outcome = Outcome::of("rules", extension, |book| {
            replace_in(book, &["Data!A2"], "20", "25");
        });
        outcome.calculated("A7", extension);
        // LibreOffice caches `#N/A` for `FORMULATEXT` of a constant, as the engine gives.
        assert_eq!(
            at(&outcome.kept, "A8"),
            at(&outcome.full, "A8"),
            "{extension}"
        );
        assert_eq!(
            *at(&outcome.kept, "A7"),
            CellValue::Boolean(false),
            "{extension}"
        );
        let delete = |book: &mut Book| {
            book.delete_rows(["Data!6".parse().unwrap()]);
        };
        Outcome::of("rules", extension, delete).kept("A7", extension);
    }
}

/// A formula that caches an empty result (`<v></v>` in an XLSX, `office:value=""` in an ODS) has no result,
/// as one with no `<v>` has: it is calculated, and the other formulas keep what the file caches.
#[test]
fn a_formula_that_caches_an_empty_result_is_calculated() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let outcome = Outcome::of("empty", extension, delete);
        for (cell, result) in [("A1", 40.0), ("A2", 60.0), ("A3", 80.0)] {
            assert_eq!(
                *at(&calc(&asset("empty", extension)), cell),
                CellValue::Empty,
                "{extension} {cell}"
            );
            outcome.calculated(cell, extension);
            assert_eq!(
                *at(&outcome.kept, cell),
                number(result),
                "{extension} {cell}"
            );
        }
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// However many formulas a replacement reaches (ten of the forty, or more than a quarter of them), the others keep the result
/// the file caches, as in `LibreOffice`: the share of the formulas that are stale does not change what is calculated.
#[test]
fn a_replacement_calculates_what_it_reaches_however_many_formulas_that_is() {
    for extension in EXTENSIONS {
        for cells in [9, 10, 11, 12, 39] {
            let replace = |book: &mut Book| {
                replace_in(book, &[&format!("Data!A1:A{cells}")], "11", "12");
            };
            let outcome = Outcome::of("quarter", extension, replace);
            // The formulas that read a replaced cell are the first `cells`.
            for row in 1..=cells {
                outcome.calculated(&format!("A{row}"), extension);
            }
            outcome.keep_rows(cells + 1, 40, extension);
        }
    }
}

/// The same for the formulas the writer leaves with no result (a deletion that shortens the range they read).
#[test]
fn a_deletion_calculates_every_formula_it_leaves_without_a_result() {
    for extension in EXTENSIONS {
        // The formula `k` reads `Data!A{k}:A{k+40}`: deleting the row `r` shortens the `81 - r` last ones.
        for (without_result, row) in [(9, 72), (10, 71), (11, 70), (12, 69), (39, 42)] {
            let delete = |book: &mut Book| {
                book.delete_rows([format!("Data!{row}").parse().unwrap()]);
            };
            let outcome = Outcome::of("quarter", extension, delete);
            for k in 41 - without_result..=40 {
                outcome.calculated(&format!("A{k}"), extension);
            }
            outcome.keep_rows(1, 40 - without_result, extension);
        }
    }
}

/// The cells a replacement rewrites go to the engine as rectangles, and more than 64 in a sheet are one
/// rectangle that holds them all: the formulas that read a cell between them are calculated again too (and
/// the result is what a run without the cache gives), the ones that read outside it are not.
#[test]
fn many_replaced_cells_in_a_sheet_are_one_rectangle() {
    for extension in EXTENSIONS {
        // A1, A3, ..., A129: 65 cells, no two adjacent.
        let cells: Vec<String> = (1..=129).step_by(2).map(|r| format!("Data!A{r}")).collect();
        assert_eq!(cells.len(), 65);
        let replace = |book: &mut Book| {
            let scopes: Vec<&str> = cells.iter().map(String::as_str).collect();
            replace_in(book, &scopes, "11", "12");
        };
        let outcome = Outcome::of("runs", extension, replace);
        // The formulas of the replaced cells (A1..A65 of Calc) and the two that read a cell between them.
        for row in 1..=67 {
            outcome.calculated(&format!("A{row}"), extension);
        }
        // `Data!A131` is below the last replaced cell: the file's result, and so are the others.
        outcome.keep_rows(68, 301, extension);
    }
}

/// A run of adjacent replaced cells is one rectangle: the 70 cells of a row and another far below make
/// two, and what reads between them is not reached. Counted one by one they would be 71, more than 64,
/// and the one rectangle that holds them would reach it.
#[test]
fn adjacent_replaced_cells_are_one_rectangle() {
    for extension in EXTENSIONS {
        let replace = |book: &mut Book| {
            replace_in(book, &["Data!B1:BS1", "Data!B41"], "11", "12");
        };
        let outcome = Outcome::of("runs", extension, replace);
        // `Data!B1:BS1` and `Data!B41`.
        for row in 69..=139 {
            outcome.calculated(&format!("A{row}"), extension);
        }
        // `Data!C20` and `Data!BR30` are inside the rectangle of both, and are replaced by neither.
        outcome.keep_rows(140, 141, extension);
        outcome.keep_rows(142, 301, extension);
    }
}

/// A `calculate` step has the mode of the run, as `sheetc calculate` does: by default the formulas nothing reaches keep
/// the result the file caches, and with `calculate_all` every formula is calculated.
#[test]
fn a_calculate_step_inherits_the_mode_of_the_run() {
    for extension in EXTENSIONS {
        let calculate = |book: &mut Book| {
            book.calculate(Calculation::default());
        };
        let (kept, ..) = run("book", extension, true, calculate);
        let (full, ..) = run("book", extension, false, calculate);
        // The result the file caches, where the engine would give another.
        assert_eq!(*at(&kept, "A7"), text("from-excel"), "{extension}");
        // The engine's result for every formula with `calculate_all`.
        assert_eq!(*at(&full, "A7"), text("t"), "{extension}");
        assert_eq!(*at(&full, "A8"), number(8.0), "{extension}");
    }
}

/// What a `calculate_all` step wrote is right, whatever the file cached: the edit that follows it calculates what it
/// reaches and leaves the rest as the step left it, not as the file had it. By default the step leaves the file's.
#[test]
fn a_calculation_makes_the_results_right_for_the_edit_that_follows() {
    for extension in EXTENSIONS {
        let both = |book: &mut Book| {
            book.calculate(Calculation::default());
            book.delete_rows(["Data!3".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, both);
        let (full, ..) = run("book", extension, false, both);
        assert_eq!(*at(&full, "A7"), text("t"), "{extension}");
        assert_eq!(*at(&kept, "A7"), text("from-excel"), "{extension}");
        // What the deletion reaches is the same in both.
        for cell in ["A1", "A5"] {
            assert_eq!(at(&kept, cell), at(&full, cell), "{extension} {cell}");
        }
    }
}

/// The workbook in memory (what Python and WebAssembly write) is the one a file gives, with the cache,
/// from a file and from bytes.
#[test]
fn in_memory_runs_use_the_cache_as_files_do() {
    type Edit = fn(&mut Book);
    let edits: [Edit; 4] = [
        |book| {
            book.delete_rows(["Data!3".parse().unwrap()]);
        },
        |book| replace_in(book, &["Data!A2"], "20", "25"),
        |book| {
            book.hide_rows(["Data!2".parse().unwrap()]);
        },
        |book| {
            book.calculate(Calculation::default());
            book.delete_columns(["Data!B".parse().unwrap()]);
        },
    ];
    for extension in EXTENSIONS {
        let path = asset("book", extension);
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        for edit in edits {
            let sources = [
                Book::from_file(&path).unwrap(),
                Book::from_bytes(std::fs::read(&path).unwrap()).unwrap(),
            ];
            for source in sources {
                let mut book = source.with_calculate_all(false);
                edit(&mut book);
                book.write(&out).unwrap();
                let mut bytes = Vec::new();
                book.write_to(&mut bytes).unwrap();
                assert_eq!(bytes, std::fs::read(&out).unwrap(), "{extension}");
            }
        }
        // And the cache was used: the formula nothing reaches is as the file has it.
        let mut book = Book::from_bytes(std::fs::read(&path).unwrap())
            .unwrap()
            .with_calculate_all(false);
        book.delete_rows(["Data!3".parse().unwrap()]);
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        std::fs::write(&out, bytes).unwrap();
        assert_eq!(*at(&calc(&out), "A7"), text("from-excel"), "{extension}");
    }
}

/// A workbook the file says is stale is calculated whole in memory too.
#[test]
fn a_stale_workbook_in_memory_is_calculated_whole() {
    let dir = tempfile::tempdir().unwrap();
    let stale = stale_xlsx(dir.path());
    let run = |cached: bool| {
        let mut book = Book::from_bytes(std::fs::read(&stale).unwrap())
            .unwrap()
            .with_calculate_all(!cached);
        book.delete_rows(["Data!3".parse().unwrap()]);
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        let out = dir.path().join(format!("memory{cached}.xlsx"));
        std::fs::write(&out, bytes).unwrap();
        calc(&out)
    };
    assert_eq!(run(true), run(false));
}

/// The inputs a calculation is given are the declared inputs of its build: by default only what they reach is
/// calculated (`SUM(Data!A1:A6)` reads the cell written), and what nothing reaches keeps what the file caches;
/// with `calculate_all` every formula is calculated.
#[test]
fn a_calculation_calculates_only_what_its_inputs_reach_by_default() {
    for extension in EXTENSIONS {
        let calculate = |book: &mut Book| {
            let mut inputs = serde_json::Map::new();
            inputs.insert("Data!A2".into(), serde_json::json!(25));
            book.calculate(Calculation {
                inputs,
                ..Calculation::default()
            });
        };
        let (kept, ..) = run("book", extension, true, calculate);
        let (full, ..) = run("book", extension, false, calculate);
        assert_eq!(*at(&kept, "A1"), number(215.0), "{extension}: reached");
        assert_eq!(*at(&full, "A1"), number(215.0), "{extension}: reached");
        assert_eq!(
            *at(&kept, "A7"),
            text("from-excel"),
            "{extension}: not reached"
        );
        assert_eq!(*at(&full, "A7"), text("t"), "{extension}: calculate_all");
    }
}

/// The startup of an ODS: whatever the edit is, `LibreOffice` calculates when it opens the file the formulas that cache a number in a
/// cell with no format (`rate*2` caches 140 and `ROW()` 99 in `book.ods`), and keeps a cached text (`from-excel`). The plan starts with that
/// calculation in the mode of the run, so an edit that leaves nothing stale (hiding a column) shows it too; an XLSX leaves them as they are.
/// OBSERVED in `LibreOffice` 26.8 (`oracle.json`, `book.ods`, `hide columns Data!B`, default load).
#[test]
fn an_ods_starts_with_the_calculation_of_the_load_and_an_xlsx_does_not() {
    let hide = |book: &mut Book| {
        book.hide_columns(["Data!B".parse().unwrap()]);
    };
    let (ods, ..) = run("book", "ods", true, hide);
    assert_eq!(
        *at(&ods, "A6"),
        number(14.0),
        "ods: rate*2 is calculated on load"
    );
    assert_eq!(
        *at(&ods, "A8"),
        number(8.0),
        "ods: ROW() is calculated on load"
    );
    assert_eq!(
        *at(&ods, "A7"),
        text("from-excel"),
        "ods: a cached text is kept"
    );
    let (xlsx, ..) = run("book", "xlsx", true, hide);
    assert_eq!(*at(&xlsx, "A6"), number(140.0), "xlsx");
    assert_eq!(*at(&xlsx, "A8"), number(99.0), "xlsx");
    assert_eq!(*at(&xlsx, "A7"), text("from-excel"), "xlsx");
}

/// With `calculate_all` the startup of an ODS calculates every formula: the roots are the only thing that changes with the mode.
#[test]
fn the_startup_of_an_ods_calculates_every_formula_with_calculate_all() {
    let hide = |book: &mut Book| {
        book.hide_columns(["Data!B".parse().unwrap()]);
    };
    let (ods, ..) = run("book", "ods", false, hide);
    assert_eq!(*at(&ods, "A6"), number(14.0));
    assert_eq!(*at(&ods, "A8"), number(8.0));
    assert_eq!(
        *at(&ods, "A7"),
        text("t"),
        "calculate_all calculates what the file caches as text"
    );
    // An operation after the startup applies its own rule: a deletion reaches a cone in the default mode and every formula with calculate_all.
    let delete = |book: &mut Book| {
        book.delete_rows(["Data!3".parse().unwrap()]);
    };
    let (kept, ..) = run("book", "ods", true, delete);
    let (full, ..) = run("book", "ods", false, delete);
    assert_eq!(
        *at(&kept, "A7"),
        text("from-excel"),
        "default: nothing reached it"
    );
    assert_eq!(*at(&full, "A7"), text("t"));
    assert_eq!(*at(&kept, "A6"), number(14.0));
    assert_eq!(*at(&full, "A6"), number(14.0));
}

/// The explanation of an ODS says the startup first, and where it is calculated: a pass of its own when nothing else in the plan runs the
/// engine, the pass that does otherwise (a deletion leaves results stale, a conversion needs the values).
#[test]
fn the_explanation_of_an_ods_starts_with_the_startup() {
    let mut hide = Book::from_file(asset("book", "ods")).unwrap();
    hide.hide_columns(["Data!B".parse().unwrap()]);
    let text = hide.explain();
    assert!(text.starts_with("startup of an ODS:"), "{text}");
    assert!(text.contains("in a pass of its own, the first"), "{text}");
    assert!(text.contains("pass 1 of 2:\n  recalculate"), "{text}");
    let mut delete = Book::from_file(asset("book", "ods")).unwrap();
    delete.delete_rows(["Data!3".parse().unwrap()]);
    let text = delete.explain();
    assert!(text.starts_with("startup of an ODS:"), "{text}");
    assert!(
        text.contains("in pass 2, with what the plan asks of the engine"),
        "{text}"
    );
    let mut full = Book::from_file(asset("book", "ods"))
        .unwrap()
        .with_calculate_all(true);
    full.hide_columns(["Data!B".parse().unwrap()]);
    let all = full.explain();
    assert!(
        all.contains("every formula is calculated (`calculate_all`)"),
        "{all}"
    );
    // An XLSX has no startup.
    let mut xlsx = Book::from_file(asset("book", "xlsx")).unwrap();
    xlsx.hide_columns(["Data!B".parse().unwrap()]);
    assert!(!xlsx.explain().contains("startup"));
}
