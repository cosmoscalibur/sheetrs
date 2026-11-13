//! The editing of a [`Book`]: the changes to make to a workbook, and the file it is saved as.

use crate::error::{Context, Error, Result};
use crate::plan::{self, Engine};
use crate::replace::Report;
use sheetrs_io::Book;
use sheetrs_io::write::{
    Calculation, CellWrites, ColumnRange, Operation, ReplaceSpec, RowRange, ValueScope,
    VisibilityChange,
};
use sheetrs_protocol::{Replace, Scope, ValueRanges};
use std::path::Path;

/// What a [`Book`] is edited with: the changes are collected in the book, and saved as a new file
/// by [`write`](Self::write), [`write_in_place`](Self::write_in_place) or
/// [`write_to`](Self::write_to).
///
/// **Every change is written against the workbook as it is**, in its rows, its columns and its
/// sheets, and the order they are given in changes nothing: `delete_rows("S!3")` and
/// `delete_rows("S!4")` are the rows 3 and 4 of the workbook whichever comes first, and a
/// `formula_to_value` gives the values the workbook has whether it is given before or after a
/// deletion. The planner decides what runs in which order (cells written and replacements, then
/// the calculation, then conversions and deletions, in one pass over the file), so that nobody has
/// to think about what the workbook looks like after each step; the operations of one kind keep
/// their order (a sweep of replacements is a sequence of texts: each one sees what the previous one
/// left). To work step by step, save between steps and open what was saved: the saved workbook is
/// the input of the next step.
///
/// Nothing is read until a `write`, which runs the changes in the fewest streaming passes over the
/// file: the memory it needs does not depend on the size of the workbook. What it will do is
/// [`explain`](Self::explain)ed first, without touching any file.
///
/// What the changes leave stale (rows, columns, sheets or names deleted, text replaced) is calculated
/// again by the engine as the last thing the write does, so the workbook is saved with fresh results.
/// The workbook is marked to be recalculated when it is opened only if the engine cannot calculate it,
/// and then the [`Report`] says why (a diagnostic with the code [`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)): the mark is never
/// decided beforehand. To keep the results the file caches instead, with no recalculation and no mark,
/// call [`Book::set_recalculate`](sheetrs_io::Book::set_recalculate) with `false` before writing.
///
/// ```no_run
/// use sheetrs_io::Book;
/// use sheetrs_edit::BookEdit;
///
/// # fn main() -> sheetrs_edit::Result<()> {
/// let mut book = Book::from_file("in.xlsx")?;
/// book.delete_rows(["Data!3:10".parse()?])
///     .delete_sheets(["Scratch"])
///     .write("out.xlsx")?;
/// # Ok(()) }
/// ```
pub trait BookEdit {
    /// Calculate the workbook (*Data ▸ Calculate* in `LibreOffice`): the inputs of the
    /// [`Calculation`] (a map from a cell, range or defined name to the value written there) are
    /// written first and every formula is calculated again, so that the workbook is written with
    /// fresh results and without the mark to recalculate it when it is opened. The `ranges` are
    /// the ranges whose values the answer holds; `None` is the outputs the workbook declares or,
    /// with none declared, its formula cells.
    ///
    /// With `answer` the value ranges are kept in the [`Report`] the write returns
    /// ([`Report::value_ranges`], [`Report::write_json`]), and a workbook the engine cannot calculate is
    /// an error ([`Error::Calculation`]) and not a warning: there is no answer to give, and the mark
    /// to recalculate cannot stand for it. To read the value ranges without writing the workbook anywhere:
    /// `write_to(std::io::sink())`.
    ///
    /// It is an operation like the others, not a query. Where the engine cannot calculate the
    /// workbook (a function it does not have), the run goes on, marks the workbook to be
    /// recalculated when it is opened and says so with a diagnostic ([`RECALC_NOT_SUPPORTED`](crate::RECALC_NOT_SUPPORTED)).
    fn calculate(&mut self, calculation: Calculation) -> &mut Self;

    /// Delete the sheets of these names (*Delete Sheet* in `LibreOffice`).
    fn delete_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self;

    /// Delete the defined names, or database ranges (*Manage Names ▸ Delete* in `LibreOffice`).
    fn delete_names(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self;

    /// Replace the formulas of the scopes by the values they have (*Data ▸ Calculate ▸ Formula to
    /// Value* in `LibreOffice`).
    fn formula_to_value(&mut self, scopes: impl IntoIterator<Item = ValueScope>) -> &mut Self;

    /// Delete rows, in the rows of the workbook; the references to them become `#REF!` and the
    /// ranges around them shrink.
    fn delete_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self;

    /// Delete columns, in the columns of the workbook.
    fn delete_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self;

    /// Hide rows (*Format ▸ Rows ▸ Hide* in `LibreOffice`), in the rows of the workbook. The rows
    /// stay in the file; a row hidden this way counts as hidden by hand for a search with
    /// `filtered` off. Changes of visibility are applied before any deletion, so the rows are those
    /// of the input.
    fn hide_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self;

    /// Show rows (*Format ▸ Rows ▸ Show*), whatever hid them (by hand or by a filter).
    fn show_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self;

    /// Hide columns (*Format ▸ Columns ▸ Hide*), in the columns of the workbook.
    fn hide_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self;

    /// Show columns (*Format ▸ Columns ▸ Show*).
    fn show_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self;

    /// Hide the sheets of these names (*Sheet ▸ Hide Sheet*). One sheet must stay visible. A very
    /// hidden sheet of an XLSX stays very hidden.
    fn hide_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self;

    /// Show the sheets of these names (*Sheet ▸ Show Sheet*), hidden or very hidden.
    fn show_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self;

    /// Find and replace text in the cells of `scopes` (the whole workbook when there are none), as
    /// `LibreOffice`'s Find & Replace does: see [`sheetrs_protocol::Replace`] for what is searched
    /// and how, and the module documentation of the replacement for what it does to each kind of
    /// cell. What was replaced and what was left is in the [`Report`] that the write returns.
    ///
    /// A sweep gives all its replacements at once: `replaces` is a sequence and each element is an
    /// operation, in order, each one seeing what the previous one left, exactly as if `replace` were
    /// called once per element. They share the `scopes`; a replacement that needs other options (or
    /// other scopes) is another call. One replacement is `[replace]`.
    fn replace(
        &mut self,
        scopes: impl IntoIterator<Item = Scope>,
        replaces: impl IntoIterator<Item = Replace>,
    ) -> &mut Self;

    /// Write values into cells: the results a program computed into the formulas of a workbook that
    /// serves as a template, and the constants of its inputs (see [`CellWrites`]).
    fn write_cells(&mut self, writes: CellWrites) -> &mut Self;

    /// What the write will do, as text, one change per line, grouped by the pass over the file that
    /// runs it.
    fn explain(&self) -> String;

    /// Runs the changes and writes the result to `output`, which is not the file being edited. A
    /// failed run leaves no `output` behind. The [`Report`] says what the replacements did.
    ///
    /// # Errors
    ///
    /// When there are no changes (a copy of the file is not what was asked for), when a sheet
    /// addressed does not exist, when a formula has a 3D reference, when a file cannot be read or
    /// written, or when the writer cannot make a change faithfully.
    fn write(&self, output: impl AsRef<Path>) -> Result<Report>;

    /// Runs the changes and writes the bytes of the workbook to `out`, in memory: no file is
    /// written, and nothing is read but the file a [`Book::from_file`] names. It is the way to edit
    /// where there is no file system, as in a browser (`write_to(&mut bytes)` with a `Vec<u8>`). The
    /// [`Report`] says what the replacements did.
    ///
    /// # Errors
    ///
    /// For the reasons [`write`](Self::write) lists, except those of writing a file, and when `out`
    /// fails.
    fn write_to(&self, out: impl std::io::Write) -> Result<Report>;

    /// The value ranges of the last [`calculate`](Self::calculate) step, as the engine calculates
    /// them on the workbook the changes leave: the changes run with the workbook written nowhere.
    /// The step's `ranges` are the ones asked for, or the outputs the workbook declares or, with
    /// none, its formula cells.
    ///
    /// # Errors
    ///
    /// When there is no `calculate` step ([`Error::NoCalculation`]), when the engine cannot
    /// calculate the workbook (there is no answer to give, as there is no workbook to mark), and for
    /// the reasons [`write`](Self::write) lists.
    fn value_ranges(&self) -> Result<ValueRanges>;

    /// Runs the changes and replaces the file being edited with the result.
    ///
    /// The result is written to a temporary file in the same folder and renamed over the original
    /// when it is complete, so a failed run leaves the original as it was and no half-written file
    /// behind; the permissions of the original are kept. The workbook must have been given as a
    /// file ([`Book::from_file`]).
    ///
    /// # Errors
    ///
    /// When the workbook was given as bytes, and for the reasons [`write`](Self::write) lists.
    fn write_in_place(&self) -> Result<Report>;
}

impl BookEdit for Book {
    fn value_ranges(&self) -> Result<ValueRanges> {
        let mut book = self.clone();
        let last = book
            .operations()
            .iter()
            .rposition(|operation| matches!(operation, Operation::Calculate(_)))
            .ok_or(crate::Error::NoCalculation)?;
        let mut operations = book.operations().to_vec();
        if let Operation::Calculate(calculation) = &mut operations[last] {
            calculation.answer = true;
        }
        book.clear_operations();
        for operation in operations {
            book.push(operation);
        }
        book.write_to(std::io::sink())?
            .value_ranges
            .ok_or(crate::Error::NoCalculation)
    }

    fn calculate(&mut self, calculation: Calculation) -> &mut Self {
        self.push(Operation::Calculate(calculation))
    }

    fn delete_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self {
        self.push(Operation::DeleteSheets(
            names.into_iter().map(Into::into).collect(),
        ))
    }

    fn delete_names(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self {
        self.push(Operation::DeleteNames(
            names.into_iter().map(Into::into).collect(),
        ))
    }

    fn formula_to_value(&mut self, scopes: impl IntoIterator<Item = ValueScope>) -> &mut Self {
        self.push(Operation::FormulaToValue(scopes.into_iter().collect()))
    }

    fn delete_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self {
        self.push(Operation::DeleteRows(rows.into_iter().collect()))
    }

    fn delete_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self {
        self.push(Operation::DeleteColumns(columns.into_iter().collect()))
    }

    fn hide_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self {
        let changes = rows.into_iter().map(|range| VisibilityChange::Rows {
            range,
            visible: false,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn show_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self {
        let changes = rows.into_iter().map(|range| VisibilityChange::Rows {
            range,
            visible: true,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn hide_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self {
        let changes = columns.into_iter().map(|range| VisibilityChange::Columns {
            range,
            visible: false,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn show_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self {
        let changes = columns.into_iter().map(|range| VisibilityChange::Columns {
            range,
            visible: true,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn hide_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self {
        let changes = names.into_iter().map(|name| VisibilityChange::Sheet {
            name: name.into(),
            visible: false,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn show_sheets(&mut self, names: impl IntoIterator<Item = impl Into<String>>) -> &mut Self {
        let changes = names.into_iter().map(|name| VisibilityChange::Sheet {
            name: name.into(),
            visible: true,
        });
        self.push(Operation::SetVisibility(changes.collect()))
    }

    fn replace(
        &mut self,
        scopes: impl IntoIterator<Item = Scope>,
        replaces: impl IntoIterator<Item = Replace>,
    ) -> &mut Self {
        let scopes: Vec<Scope> = scopes.into_iter().collect();
        for replace in replaces {
            self.push(Operation::Replace(ReplaceSpec {
                scopes: scopes.clone(),
                replace,
            }));
        }
        self
    }

    fn write_cells(&mut self, writes: CellWrites) -> &mut Self {
        self.push(Operation::WriteCells(writes))
    }

    fn explain(&self) -> String {
        plan::explain_book(self)
    }

    fn write(&self, output: impl AsRef<Path>) -> Result<Report> {
        plan::execute(self, output.as_ref(), &Engine::default())
    }

    fn write_to(&self, mut out: impl std::io::Write) -> Result<Report> {
        let (bytes, report) = plan::execute_in_memory(self, &Engine::default())?;
        out.write_all(&bytes)
            .context("Cannot write the workbook to the destination")?;
        Ok(report)
    }

    fn write_in_place(&self) -> Result<Report> {
        let path = self.path().ok_or(Error::NoFile)?;
        let folder = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let extension = self.format().extension();
        let temporary = tempfile::Builder::new()
            .prefix(".sheetrs-")
            .suffix(&format!(".{extension}"))
            .tempfile_in(folder)
            .context("Cannot create a temporary file next to the workbook")?;
        let report = plan::execute(self, temporary.path(), &Engine::default())?;
        let permissions = std::fs::metadata(path)
            .context("Cannot read the permissions of the workbook")?
            .permissions();
        std::fs::set_permissions(temporary.path(), permissions)
            .context("Cannot keep the permissions of the workbook")?;
        temporary
            .persist(path)
            .map_err(|e| e.error)
            .context("Cannot replace the workbook")?;
        Ok(report)
    }
}
