//! The editor: the changes to make to one workbook, written as a sequence and saved as a new file.

use crate::error::{Context, Error, Result};
use crate::plan::Plan;
use crate::replace::Report;
use crate::writer::{ColumnRange, RowRange, ValueScope};
use crate::writes::CellWrites;
use sheetrs_io::reader::Source;
use sheetrs_protocol::{Replace, Scope};
use std::path::Path;

/// Changes to one workbook, saved as a new file.
///
/// An editor is bound to the [`Source`] it edits and collects what to do to it, in order: each
/// change addresses the workbook **as the changes before it leave it**, so after
/// `delete_rows("S!3")` the row 3 is what was row 4. Nothing is read until [`save`](Self::save),
/// which runs the changes in the fewest streaming passes over the file that give the same
/// result: the memory it needs does not depend on the size of the workbook. What it will do is
/// [`explain`](Self::explain)ed first, without touching any file.
///
/// Whether the saved workbook is marked to be recalculated when it is opened is the editor's
/// decision, from what the changes leave stale ([`marks_recalculation`](Self::marks_recalculation)):
/// there is nothing to ask for or to avoid.
///
/// ```no_run
/// use sheetrs_io::reader::Source;
/// use sheetrs_edit::Editor;
///
/// # fn main() -> sheetrs_edit::Result<()> {
/// let source = Source::from_file("in.xlsx")?;
/// Editor::new(&source)
///     .delete_rows(["Data!3:10".parse()?])
///     .remove_sheets(["Scratch"])
///     .save("out.xlsx")?;
/// # Ok(()) }
/// ```
#[derive(Debug, Clone)]
pub struct Editor {
    source: Source,
    plan: Plan,
}

impl Editor {
    /// An editor with no changes yet, for the workbook `source` refers to. The workbook is not
    /// read: a [`Source`] is a reference, and cloning it does not copy a file.
    #[must_use]
    pub fn new(source: &Source) -> Self {
        Self {
            source: source.clone(),
            plan: Plan::new(),
        }
    }

    /// Remove the sheets of these names.
    pub fn remove_sheets(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        self.plan.remove_sheets(names);
        self
    }

    /// Remove the defined names (or database ranges).
    pub fn remove_named_ranges(
        &mut self,
        names: impl IntoIterator<Item = impl Into<String>>,
    ) -> &mut Self {
        self.plan.remove_named_ranges(names);
        self
    }

    /// Replace the formulas of the scopes by the values they have.
    pub fn convert_to_values(&mut self, scopes: impl IntoIterator<Item = ValueScope>) -> &mut Self {
        self.plan.convert_to_values(scopes);
        self
    }

    /// Delete rows; the references to them become `#REF!` and the ranges around them shrink. The
    /// ranges of one call are simultaneous.
    pub fn delete_rows(&mut self, rows: impl IntoIterator<Item = RowRange>) -> &mut Self {
        self.plan.delete_rows(rows);
        self
    }

    /// Delete columns. The ranges of one call are simultaneous.
    pub fn delete_columns(&mut self, columns: impl IntoIterator<Item = ColumnRange>) -> &mut Self {
        self.plan.delete_columns(columns);
        self
    }

    /// Find and replace text in the cells of `scopes` (the whole workbook when there are none), as
    /// `LibreOffice`'s Find & Replace does: see [`sheetrs_protocol::Replace`] for what is searched
    /// and how, and the module documentation of the replacement for what it does to each kind of
    /// cell. What was replaced and what was left is in the [`Report`] that [`save`](Self::save)
    /// returns.
    pub fn replace(
        &mut self,
        scopes: impl IntoIterator<Item = Scope>,
        replace: Replace,
    ) -> &mut Self {
        self.plan.replace(scopes.into_iter().collect(), replace);
        self
    }

    /// Write values into cells: the results a program computed into the formulas of a workbook
    /// that serves as a template, and the constants of its inputs (see [`CellWrites`]). The cells
    /// are addressed as the changes before this one leave the workbook, and what comes after
    /// addresses the workbook this leaves.
    pub fn write_cells(&mut self, writes: CellWrites) -> &mut Self {
        self.plan.write_cells(writes);
        self
    }

    /// Whether there is nothing to do.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.plan.is_empty()
    }

    /// Whether the saved workbook will be marked to be recalculated when it is opened: when a
    /// change leaves cached results stale (rows, columns, sheets or names removed), not when the
    /// changes only convert formulas to the values they have.
    #[must_use]
    pub fn marks_recalculation(&self) -> bool {
        self.plan.marks_recalculation()
    }

    /// What [`save`](Self::save) will do, as text, one change per line, grouped by the pass over
    /// the file that runs it.
    ///
    /// # Errors
    ///
    /// When a change addresses a sheet that an earlier one removed.
    pub fn explain(&self) -> Result<String> {
        self.plan.explain()
    }

    /// Runs the changes and writes the result to `output`, which is not the file being edited. A
    /// failed run leaves no `output` behind. The [`Report`] says what the replacements did.
    ///
    /// # Errors
    ///
    /// When there are no changes (a copy of the file is not what was asked for), when a sheet
    /// addressed does not exist or was removed, when a formula has a 3D reference, when a file
    /// cannot be read or written, or when the writer cannot make a change faithfully.
    pub fn save(&self, output: impl AsRef<Path>) -> Result<Report> {
        self.plan.execute(&self.source, output.as_ref())
    }

    /// Runs the changes and replaces the file being edited with the result.
    ///
    /// The result is written to a temporary file in the same folder and renamed over the original
    /// when it is complete, so a failed run leaves the original as it was and no half-written file
    /// behind; the permissions of the original are kept. The workbook must have been given as a
    /// file ([`Source::from_file`]).
    ///
    /// # Errors
    ///
    /// When the workbook was given as bytes, and for the reasons [`save`](Self::save) lists.
    pub fn overwrite(&self) -> Result<Report> {
        let path = self.source.path().ok_or(Error::NoFile)?;
        let folder = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let extension = self.source.format().extension();
        let temporary = tempfile::Builder::new()
            .prefix(".sheetrs-")
            .suffix(&format!(".{extension}"))
            .tempfile_in(folder)
            .context("Cannot create a temporary file next to the workbook")?;
        let report = self.plan.execute(&self.source, temporary.path())?;
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
