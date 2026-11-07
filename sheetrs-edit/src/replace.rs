//! Find & Replace over the cells of a workbook, as `LibreOffice` does it.
//!
//! A replacement is decided against the input of its pass, in a read of the workbook that keeps
//! only the cells it changes, and then written by the same streaming writers that run the other
//! operations. What it does is what `LibreOffice` 24.2.7 does with `replaceAll` on a sheet (the
//! reference version 26.8 could not be installed; see `docs/` of the crate):
//!
//! - The text searched in a cell is the text of its input line, and every hit of the cell is
//!   replaced (`sheetrs_query::text`, `sheetrs_query::Matcher`).
//! - The new text is entered in the cell as if it were typed: `12` makes a number, `50%` a
//!   percentage, `03/15/2024` a date, `TRUE` a boolean, `=1+1` a formula, `'12` a text, and nothing
//!   an empty cell; a cell with the text format keeps text.
//! - A formula is searched and replaced as `LibreOffice` shows it (`=SUM($'My Sheet'.A1,B1)`, string
//!   literals included), and the result is read back and written in the notation of the file:
//!   Excel's in an XLSX, `OpenFormula`'s in an ODS. A search by value reaches a formula cell, which
//!   becomes a constant, as in `LibreOffice` (`Find::include_formulas` is for queries: a replacement
//!   always reaches the formula cells, as `LibreOffice` does).
//!
//! Where it does not reproduce `LibreOffice` yet (nobody approved these differences: each is a pending
//! decision), each such cell is left as it was or written in the simplest way, and said in the
//! [`Report`]:
//!
//! - A replacement that leaves a formula that is not one (`=((*2`) or that names a sheet the
//!   workbook does not have is not applied: `LibreOffice` keeps an error cell with the text
//!   lowercased (`=$other.a1*2`, `#NAME?`).
//! - A cell of an array formula is not replaced.
//! - `LibreOffice` also gives the cell the number format of what was typed (a percentage, a date,
//!   a boolean); the cell here keeps its style, so a `General` cell that gets `50%` shows `0.5`.

use crate::error::{Error, Result};
use crate::writer::{CellEdit, SheetEdits};
use sheetrs_formula::addr::{a1, same_name};
use sheetrs_formula::calc_read::{Notation, from_calc};
use sheetrs_io::reader::{
    Cell, CellSink, CellValue, Format, Properties, Selection, Source, Window,
};
use sheetrs_protocol::{Reference, Replace, Scope};
use sheetrs_query::Matcher;
use sheetrs_query::text::{TextContext, Typed};
use std::collections::HashMap;
use std::path::Path;

/// A cell a replacement found and did not change, and why.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Skipped {
    pub sheet: String,
    /// `B2`.
    pub cell: String,
    pub reason: String,
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
}

impl Report {
    pub(crate) fn absorb(&mut self, other: Self) {
        self.replaced += other.replaced;
        self.written += other.written;
        self.skipped.extend(other.skipped);
    }
}

/// A replacement and where it applies.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ReplaceSpec {
    /// Empty: the whole workbook.
    pub scopes: Vec<Scope>,
    pub replace: Replace,
}

struct Nothing;

impl CellSink for Nothing {
    fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
}

/// The sheets of the workbook, in order.
fn sheet_names(source: &Source) -> Result<Vec<String>> {
    let infos = source.read_selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?;
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
fn windows(scopes: &[Scope], names: &[String], file: &Path) -> Result<Vec<Vec<Window>>> {
    let mut by_sheet: Vec<Vec<Window>> = vec![Vec::new(); names.len()];
    let find = |name: &str| -> Result<usize> {
        names
            .iter()
            .position(|n| same_name(n, name))
            .ok_or_else(|| Error::SheetNotFound {
                sheet: name.to_string(),
                part: file.display().to_string(),
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
    edits: HashMap<String, SheetEdits>,
    report: Report,
}

impl Decide<'_> {
    fn skip(&mut self, sheet: u16, at: (u32, u32), reason: impl Into<String>) {
        self.report.skipped.push(Skipped {
            sheet: self.names[usize::from(sheet)].clone(),
            cell: a1(at.0, at.1),
            reason: reason.into(),
        });
    }

    /// What the cell becomes, or why it does not.
    fn decide(&self, typed: Typed) -> std::result::Result<CellEdit, String> {
        Ok(match typed {
            Typed::Empty => CellEdit::Clear,
            Typed::Text(t) => CellEdit::Text(t),
            Typed::Number(n) => CellEdit::Number(n),
            Typed::Boolean(b) => CellEdit::Boolean(b),
            Typed::Formula(calc) => {
                let excel = from_calc(&calc, Notation::Excel).map_err(|e| e.to_string())?;
                // The sheets the formula names must be there.
                let expr = sheetrs_formula::parse(&excel).map_err(|e| e.to_string())?;
                let mut missing = None;
                expr.walk(&mut |e| {
                    if let sheetrs_formula::Expr::Ref(r) = e
                        && let Some(sheet) = &r.sheet
                        && !self.names.iter().any(|n| same_name(n, sheet))
                    {
                        missing = Some(sheet.clone());
                    }
                });
                if let Some(sheet) = missing {
                    return Err(format!(
                        "the formula names the sheet `{sheet}`, which does not exist"
                    ));
                }
                CellEdit::Formula(from_calc(&calc, self.notation).map_err(|e| e.to_string())?)
            }
        })
    }
}

impl CellSink for Decide<'_> {
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
        let Some(text) = self.text.searched(&cell, &name, find.search_in, true) else {
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
        let text_format = cell.num_fmt.as_deref() == Some("@");
        let typed = self.text.typed(&new, text_format && cell.formula.is_none());
        match self.decide(typed) {
            Ok(edit) => {
                self.edits.entry(name).or_default().insert(at, edit);
                self.report.replaced += 1;
            }
            Err(reason) => self.skip(sheet, at, reason),
        }
    }
}

/// Decides what the replacement does to the workbook at `file`: the cells to rewrite by sheet, and
/// the report.
///
/// # Errors
///
/// When the file cannot be read, a scope names a sheet it does not have, or the search is one the
/// matcher does not reproduce.
pub(crate) fn decide(
    file: &Path,
    spec: &ReplaceSpec,
) -> Result<(HashMap<String, SheetEdits>, Report)> {
    let source = Source::from_file(file)?;
    let names = sheet_names(&source)?;
    let windows = windows(&spec.scopes, &names, file)?;
    let text = TextContext::new(None, source.uses_1904_dates()?)
        .map_err(|e| Error::Search(e.to_string()))?;
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
        edits: HashMap::new(),
        report: Report::default(),
    };
    source.read_workbook_into(&mut decide)?;
    Ok((decide.edits, decide.report))
}
