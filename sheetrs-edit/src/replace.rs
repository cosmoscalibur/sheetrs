//! Find & Replace over the cells of a workbook, as `LibreOffice` does it.
//!
//! A replacement is decided against the input of its pass, in a read of the workbook that keeps
//! only the cells it changes, and then written by the same streaming writers that run the other
//! operations. What it does is what `LibreOffice` 26.8 (the reference version) does with `replaceAll`
//! on a sheet:
//!
//! - The text searched in a cell is the text of its input line, and every hit of the cell is
//!   replaced (`sheetrs_query::text`, `sheetrs_query::Matcher`).
//! - The new text is entered in the cell as if it were typed: `12` makes a number, `50%` a
//!   percentage, `03/15/2024` a date, `TRUE` a boolean, `=1+1` a formula, `'12` a text, and nothing
//!   an empty cell; a cell with the text format keeps text. A cell in `General` takes the number
//!   format `LibreOffice` gives what was typed (`0.00%` for `50%`, `MM/DD/YY` for a date, the
//!   boolean format for `TRUE`, [`EntryFormat`]), and keeps the style it has in every other
//!   respect; a cell with a format of its own keeps it, except that a boolean always takes the
//!   boolean format.
//! - A formula is searched and replaced as `LibreOffice` shows it (`=SUM($'My Sheet'.A1,B1)`, string
//!   literals included), and the result is read back and written in the notation of the file:
//!   Excel's in an XLSX, `OpenFormula`'s in an ODS. A search by value reaches a formula cell, which
//!   becomes a constant, as in `LibreOffice` (`Find::include_formulas` is for queries: a replacement
//!   always reaches the formula cells, as `LibreOffice` does).
//!
//! - A text typed as a formula is compiled as `LibreOffice` compiles it (`sheetrs_formula::calc_read::compile_calc`) and the
//!   cell keeps what it understood. A formula that is not one stays a formula in an error cell:
//!   the parentheses left open at the end are closed, the names it does not know (a function, a
//!   name, a sheet that is not there) are written in lowercase and make it `#NAME?`, and the error
//!   is the one `LibreOffice` gives for what is wrong (`Err:508` a `)` that closes nothing,
//!   `Err:509` two operands with no operator, `Err:510` an operator where an operand goes,
//!   `Err:511` a parenthesis with nothing inside, `Err:520` an operand that never comes). An XLSX
//!   writes those errors as Excel's (`#VALUE!`, `#N/A`, `#NAME?`), an ODS as `Err:5xx`.
//!
//! Where it does not reproduce `LibreOffice` yet (nobody approved these differences: each is a pending
//! decision), each such cell is left as it was or written in the simplest way, and said in the
//! [`Report`]:
//!
//! - A cell of an array formula is not replaced.

use crate::error::{Error, Result};
use sheetrs_formula::DateSystem;
use sheetrs_formula::addr::{a1, same_name};
use sheetrs_formula::calc_read::{Notation, compile_calc};
use sheetrs_io::Book;
use sheetrs_io::EntryFormat;
use sheetrs_io::reader::{Cell, CellSink, CellValue, Format, Properties, Selection, Window};
use sheetrs_io::write::{CellEdit, ReplaceSpec, SheetEdits};
use sheetrs_protocol::{Reference, Scope, SearchIn};
use sheetrs_query::Matcher;
use sheetrs_query::text::{TextContext, Typed, is_general, is_text};
use std::collections::HashMap;

/// A cell a replacement found and did not change, and why.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Skipped {
    /// The sheet of the cell.
    pub sheet: String,
    /// `B2`.
    pub cell: String,
    /// Why it was left as it was.
    pub reason: String,
}

impl Skipped {
    /// The row of the cell, 0-based.
    fn row(&self) -> Option<u32> {
        let digits: String = self.cell.chars().filter(char::is_ascii_digit).collect();
        digits.parse::<u32>().ok()?.checked_sub(1)
    }
}

/// What a run did that is more than writing the file.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Report {
    /// The cells that were replaced.
    pub replaced: usize,
    /// The cells that were found and left as they were.
    pub skipped: Vec<Skipped>,
    /// The cells written with [`CellWrites`](crate::CellWrites).
    pub written: usize,
    /// What the run did that is not what was asked for, and is not an error
    /// ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)), as structured records: a stable
    /// code, the message and the fields it was made from.
    pub diagnostics: Vec<sheetrs_protocol::Diagnostic>,
    /// The value ranges of the last `calculate` step, when the calculation asked to keep them (the
    /// `answer` of the [`Calculation`](crate::Calculation)).
    pub value_ranges: Option<sheetrs_protocol::ValueRanges>,
}

impl Report {
    pub(crate) fn absorb(&mut self, other: Self) {
        self.replaced += other.replaced;
        self.written += other.written;
        self.skipped.extend(other.skipped);
    }

    /// Writes the value ranges of the calculation as JSON to `out`, the way [`BookEdit::write`](crate::BookEdit::write) writes
    /// the workbook: `write_json(&mut file)`, or into a `Vec<u8>`.
    ///
    /// # Errors
    ///
    /// When the run kept no value ranges, or when `out` fails.
    pub fn write_json(&self, mut out: impl std::io::Write) -> crate::Result<()> {
        use crate::error::Context as _;
        let value_ranges = self
            .value_ranges
            .as_ref()
            .ok_or(crate::Error::NoCalculation)?;
        value_ranges
            .write_json(&mut out)
            .context("Cannot write the value ranges")
    }
}

struct Nothing;

impl CellSink for Nothing {
    fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
}

/// The sheets of the workbook, in order.
fn sheet_names(source: &Book) -> Result<Vec<String>> {
    let infos = source.selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?;
    let mut names = vec![
        String::new();
        infos
            .iter()
            .map(|i| usize::from(i.index) + 1)
            .max()
            .unwrap_or(0)
    ];
    for info in infos {
        names[usize::from(info.index)] = info.name;
    }
    Ok(names)
}

/// The windows of each sheet that the scopes select.
fn windows(scopes: &[Scope], names: &[String], origin: &str) -> Result<Vec<Vec<Window>>> {
    let mut by_sheet: Vec<Vec<Window>> = vec![Vec::new(); names.len()];
    let find = |name: &str| -> Result<usize> {
        names
            .iter()
            .position(|n| same_name(n, name))
            .ok_or_else(|| {
                Error::from(sheetrs_io::write::Error::SheetNotFound {
                    sheet: name.to_string(),
                    part: origin.to_string(),
                })
            })
    };
    if scopes.is_empty() {
        for w in &mut by_sheet {
            w.push(Window::ALL);
        }
    }
    for scope in scopes {
        match scope {
            Scope::Workbook => {
                for w in &mut by_sheet {
                    w.push(Window::ALL);
                }
            }
            Scope::Sheet(name) => by_sheet[find(name)?].push(Window::ALL),
            Scope::Range(Reference { sheet, area, .. }) => {
                let at = match sheet {
                    Some(name) => find(name)?,
                    None => 0,
                };
                by_sheet[at].push(Window::new(area.r0, area.c0, area.r1, area.c1));
            }
        }
    }
    Ok(by_sheet)
}

struct Decide<'a> {
    text: &'a TextContext,
    matcher: Matcher,
    spec: &'a ReplaceSpec,
    names: &'a [String],
    windows: Vec<Vec<Window>>,
    notation: Notation,
    source: &'a Book,
    defined: std::cell::OnceCell<Vec<String>>,
    /// The date system of the workbook: its dates are typed and shown in it.
    date_system: DateSystem,
    edits: HashMap<String, SheetEdits>,
    report: Report,
    /// The rows an autofilter hid, by sheet: left alone unless the replacement asks for them.
    filtered: Vec<(u16, u32, u32)>,
}

impl Decide<'_> {
    /// Takes back what was decided for the cells of the rows an autofilter hid (an XLSX tells which
    /// rows they are after the cells; an ODS before, and the same pass serves both).
    fn leave_out_filtered_rows(&mut self) {
        if self.filtered.is_empty() {
            return;
        }
        let names = self.names;
        let hidden = |sheet: &str, row: u32| {
            self.filtered.iter().any(|&(s, first, last)| {
                names.get(usize::from(s)).is_some_and(|n| n == sheet)
                    && (first..=last).contains(&row)
            })
        };
        let mut removed = 0;
        for (sheet, edits) in &mut self.edits {
            let before = edits.len();
            edits.retain(|&(row, _), _| !hidden(sheet, row));
            removed += before - edits.len();
        }
        self.edits.retain(|_, edits| !edits.is_empty());
        self.report.replaced = self.report.replaced.saturating_sub(removed);
        self.report.skipped.retain(|skipped| match skipped.row() {
            Some(row) => !hidden(&skipped.sheet, row),
            None => true,
        });
    }

    fn skip(&mut self, sheet: u16, at: (u32, u32), reason: impl Into<String>) {
        self.report.skipped.push(Skipped {
            sheet: self.names[usize::from(sheet)].clone(),
            cell: a1(at.0, at.1),
            reason: reason.into(),
        });
    }

    /// The names the workbook defines, read when a text typed as a formula names something that is
    /// neither a function nor a reference: only then is the workbook read once more.
    fn defined_names(&self) -> &Vec<String> {
        self.defined.get_or_init(|| {
            self.source
                .workbook_into(&mut Nothing)
                .map(|w| w.defined_names.keys().cloned().collect())
                .unwrap_or_default()
        })
    }

    /// What the cell becomes, with the number format `LibreOffice` gives what was typed into it.
    fn decide_in(&self, typed: Typed, input: &str, num_fmt: Option<&str>) -> CellEdit {
        let format = match typed {
            Typed::Boolean(_) => Some(EntryFormat::Boolean),
            Typed::Number(_) if is_general(num_fmt) => self.text.entry_format(input),
            _ => None,
        };
        let number = match typed {
            Typed::Number(n) => n,
            _ => 0.0,
        };
        let edit = self.decide(typed);
        match format {
            Some(format) => CellEdit::Formatted {
                edit: Box::new(edit),
                format,
                ods_value: sheetrs_io::write::ods_value(format, number, self.date_system),
                // The text the cell shows, as the application writes it.
                shown: sheetrs_engine::format::format_number(
                    number,
                    sheetrs_io::write::display_code(format),
                    &sheetrs_engine::options::Locale::en_us(),
                    self.date_system,
                ),
                style: None,
            },
            None => edit,
        }
    }

    /// What the cell becomes.
    fn decide(&self, typed: Typed) -> CellEdit {
        match typed {
            Typed::Empty => CellEdit::Clear,
            Typed::Text(t) => CellEdit::Text(t),
            Typed::Number(n) => CellEdit::Number(n),
            Typed::Boolean(b) => CellEdit::Boolean(b),
            Typed::Formula(calc) => {
                // The references to other workbooks are set apart: the compiler does not know them.
                let external = self.text.protect_externals(&calc);
                // `LibreOffice` compiles what was typed and keeps what it understood.
                let compiled = compile_calc(
                    &external.text,
                    self.notation,
                    &sheetrs_engine::functions::in_catalogue,
                    &|name| {
                        external.is_name(name)
                            || self
                                .defined_names()
                                .iter()
                                .any(|n| n.eq_ignore_ascii_case(name))
                    },
                    &|sheet| self.names.iter().any(|n| same_name(n, sheet)),
                );
                let formula = external.restore(&compiled.formula);
                match compiled.error {
                    None => CellEdit::Formula(formula),
                    Some(code) => CellEdit::Invalid { formula, code },
                }
            }
        }
    }
}

impl CellSink for Decide<'_> {
    fn filtered_rows(&mut self, sheet: u16, first: u32, last: u32) {
        self.filtered.push((sheet, first, last));
    }

    fn comment(&mut self, sheet: u16, at: (u32, u32), text: &str) {
        let Some(windows) = self.windows.get(usize::from(sheet)) else {
            return;
        };
        if !windows.iter().any(|w| w.contains(at.0, at.1)) {
            return;
        }
        // The text of a note is plain: the replacement is not entered as typed.
        let Some(new) = self.matcher.replace_all(text, &self.spec.replace.with) else {
            return;
        };
        let name = self.names[usize::from(sheet)].clone();
        self.edits
            .entry(name)
            .or_default()
            .insert(at, CellEdit::Note(new));
        self.report.replaced += 1;
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        let Some(windows) = self.windows.get(usize::from(sheet)) else {
            return;
        };
        if !windows.iter().any(|w| w.contains(at.0, at.1)) {
            return;
        }
        if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
            return;
        }
        let find = &self.spec.replace.find;
        let name = self.names[usize::from(sheet)].clone();
        let Some(text) = self
            .text
            .searched_as(&cell, &name, find.search_in, true, find.formatted)
        else {
            return;
        };
        let Some(new) = self.matcher.replace_all(&text, &self.spec.replace.with) else {
            return;
        };
        if cell.is_array {
            self.skip(sheet, at, "the cell holds an array formula");
            return;
        }
        // A formula cell searched in its formula is entered as typed text; anything else too.
        let text_format = is_text(cell.num_fmt.as_deref());
        let typed = self.text.typed_in(
            &new,
            text_format && cell.formula.is_none(),
            cell.num_fmt.as_deref(),
        );
        let edit = self.decide_in(typed, &new, cell.num_fmt.as_deref());
        self.edits.entry(name).or_default().insert(at, edit);
        self.report.replaced += 1;
    }
}

/// Decides what the replacement does to the workbook `source` refers to: the cells to rewrite by
/// sheet, and the report.
///
/// # Errors
///
/// When the file cannot be read, a scope names a sheet it does not have, or the search is one the
/// matcher does not reproduce.
pub(crate) fn decide(
    source: &Book,
    spec: &ReplaceSpec,
) -> Result<(HashMap<String, SheetEdits>, Report)> {
    let origin = source
        .path()
        .map_or_else(|| "the workbook".to_string(), |p| p.display().to_string());
    let names = sheet_names(source)?;
    let windows = windows(&spec.scopes, &names, &origin)?;
    let date_system = source.date_system()?;
    let text = TextContext::new(None, date_system)
        .map_err(|e| Error::Search(e.to_string()))?
        .for_format(source.format())
        .for_book(source);
    let matcher = Matcher::new(&spec.replace.find).map_err(|e| Error::Search(e.to_string()))?;
    let notation = match source.format() {
        Format::Ods => Notation::OpenFormula,
        _ => Notation::Excel,
    };
    let mut decide = Decide {
        text: &text,
        matcher,
        spec,
        names: &names,
        windows,
        notation,
        source,
        defined: std::cell::OnceCell::new(),
        date_system,
        edits: HashMap::new(),
        report: Report::default(),
        filtered: Vec::new(),
    };
    // A replacement leaves alone the rows an autofilter hid, unless it asks for them.
    let leaves_out_filtered = !spec.replace.find.filtered;
    if spec.replace.find.search_in == SearchIn::Comment || leaves_out_filtered {
        // Only the windows asked for are read: the notes alone, or the cells.
        let mut selection = Selection::new(if spec.replace.find.search_in == SearchIn::Comment {
            Properties {
                value: false,
                formula: false,
                number_format: false,
                comment: true,
            }
        } else {
            Properties::ALL
        });
        if leaves_out_filtered {
            selection = selection.with_filtered_rows();
        }
        for (index, windows) in decide.windows.clone().into_iter().enumerate() {
            for window in windows {
                selection = selection.window(names[index].clone(), window);
            }
        }
        source.selection_into(&selection, &mut decide)?;
    } else {
        source.workbook_into(&mut decide)?;
    }
    if leaves_out_filtered {
        decide.leave_out_filtered_rows();
    }
    Ok((decide.edits, decide.report))
}
