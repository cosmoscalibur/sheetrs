//! A run that uses the cache (`Book::with_use_cache(true)`) calculates the formulas the plan reaches,
//! and what depends on them, and leaves every other formula with the result the file caches. The
//! workbooks of `tests/assets/use_cache/` cache results the engine would not give (`from-excel` where
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
        assert_ne!(kept, cached, "{extension} {cell}");
    }

    /// The formula kept the result of the file, which a run without the cache would not.
    fn kept(&self, cell: &str, extension: &str) {
        let (kept, full, cached) = (
            at(&self.kept, cell),
            at(&self.full, cell),
            at(&self.cached, cell),
        );
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
fn a_formula_given_a_result_is_calculated_again_with_what_the_plan_leaves_stale() {
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

/// A formula that depends on where cells are (`ROW`, `SHEET`) or that reads through a name may give
/// another result when columns go, whatever its text says.
#[test]
fn a_deleted_column_reaches_positions_and_names() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        for cell in ["A2", "A3", "A4"] {
            outcome.calculated(cell, extension);
        }
        assert_eq!(*at(&outcome.kept, "A2"), number(2.0), "{extension}");
        assert_eq!(*at(&outcome.kept, "A3"), number(14.0), "{extension}");
        assert_eq!(*at(&outcome.kept, "A4"), number(2.0), "{extension}");
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// The same when a sheet goes: the sheets are numbered again and a name may point into it.
#[test]
fn a_deleted_sheet_reaches_positions_and_names() {
    for extension in EXTENSIONS {
        let delete = |book: &mut Book| {
            book.delete_sheets(["Other"]);
        };
        let outcome = Outcome::of("rules", extension, delete);
        for cell in ["A2", "A3", "A4"] {
            outcome.calculated(cell, extension);
        }
        assert_eq!(*at(&outcome.kept, "A2"), number(2.0), "{extension}");
        assert_eq!(*at(&outcome.kept, "A3"), number(14.0), "{extension}");
        // What reads cells of a sheet that stays: the file's results.
        for cell in ["A1", "A5", "A6"] {
            outcome.kept(cell, extension);
        }
        untouched(&outcome.kept, &outcome.cached, extension);
    }
}

/// `ISFORMULA` and `FORMULATEXT` of a cell are answered when the formula is lowered, so the formula
/// keeps no reference to the cell: what it says is stale when the cell is replaced or moves.
#[test]
fn formulas_that_ask_what_a_cell_holds_are_calculated_when_a_cell_changes_or_moves() {
    type Edit = fn(&mut Book);
    let edits: [Edit; 3] = [
        |book| replace_in(book, &["Data!A2"], "20", "25"),
        |book| {
            book.delete_columns(["Data!D".parse().unwrap()]);
        },
        |book| {
            book.delete_rows(["Data!6".parse().unwrap()]);
        },
    ];
    for extension in EXTENSIONS {
        for edit in edits {
            let outcome = Outcome::of("rules", extension, edit);
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
        }
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
                *at(&outcome.cached, cell),
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

/// Where a quarter of the formulas or less is stale (a replacement reaches ten of the forty) the cache
/// is used; where more is, the walk would only add its cost and every formula is calculated again, as
/// without the option.
#[test]
fn a_replacement_that_reaches_more_than_a_quarter_of_the_formulas_calculates_them_all() {
    for extension in EXTENSIONS {
        for cells in [9, 10, 11, 12] {
            let replace = |book: &mut Book| {
                replace_in(book, &[&format!("Data!A1:A{cells}")], "11", "12");
            };
            let outcome = Outcome::of("quarter", extension, replace);
            // The formulas that read a replaced cell are the first `cells`.
            for row in 1..=cells {
                outcome.calculated(&format!("A{row}"), extension);
            }
            if cells <= 10 {
                outcome.keep_rows(cells + 1, 40, extension);
            } else {
                for row in cells + 1..=40 {
                    outcome.calculated(&format!("A{row}"), extension);
                }
            }
        }
    }
}

/// The same limit for the formulas the writer leaves with no result (a deletion that shortens the
/// range they read): ten of the forty calculate what they reach and eleven calculate all.
#[test]
fn a_deletion_that_leaves_more_than_a_quarter_of_the_formulas_without_a_result_calculates_them_all()
{
    for extension in EXTENSIONS {
        // The formula `k` reads `Data!A{k}:A{k+40}`: deleting the row `r` shortens the `81 - r` last ones.
        for (without_result, row) in [(9, 72), (10, 71), (11, 70), (12, 69)] {
            let delete = |book: &mut Book| {
                book.delete_rows([format!("Data!{row}").parse().unwrap()]);
            };
            let outcome = Outcome::of("quarter", extension, delete);
            for k in 41 - without_result..=40 {
                outcome.calculated(&format!("A{k}"), extension);
            }
            if without_result <= 10 {
                outcome.keep_rows(1, 40 - without_result, extension);
            } else {
                for k in 1..=40 - without_result {
                    outcome.calculated(&format!("A{k}"), extension);
                }
            }
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

/// An explicit `calculate` calculates every formula, with the cache or not: it is what was asked for,
/// not an edit that leaves some results stale.
#[test]
fn a_calculate_step_calculates_every_formula_even_with_the_cache() {
    for extension in EXTENSIONS {
        let calculate = |book: &mut Book| {
            book.calculate(Calculation::default());
        };
        let (kept, ..) = run("book", extension, true, calculate);
        let (full, ..) = run("book", extension, false, calculate);
        assert_eq!(kept, full, "{extension}");
        // The engine's result, where the file caches one that the engine would not give.
        assert_eq!(*at(&kept, "A7"), text("t"), "{extension}");
    }
}

/// What a `calculate` step wrote is right, whatever the file cached: the edit that follows it calculates
/// what it reaches and leaves the rest as the step left it, not as the file had it.
#[test]
fn a_calculation_makes_the_results_right_for_the_edit_that_follows() {
    for extension in EXTENSIONS {
        let both = |book: &mut Book| {
            book.calculate(Calculation::default());
            book.delete_rows(["Data!3".parse().unwrap()]);
        };
        let (kept, ..) = run("book", extension, true, both);
        let (full, ..) = run("book", extension, false, both);
        assert_eq!(kept, full, "{extension}");
        assert_eq!(*at(&kept, "A7"), text("t"), "{extension}");
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
                let mut book = source.with_use_cache(true);
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
            .with_use_cache(true);
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
    let run = |use_cache: bool| {
        let mut book = Book::from_bytes(std::fs::read(&stale).unwrap())
            .unwrap()
            .with_use_cache(use_cache);
        book.delete_rows(["Data!3".parse().unwrap()]);
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        let out = dir.path().join(format!("memory{use_cache}.xlsx"));
        std::fs::write(&out, bytes).unwrap();
        calc(&out)
    };
    assert_eq!(run(true), run(false));
}
