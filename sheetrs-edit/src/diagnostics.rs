//! What a run did that the person who asked should know, and that is not an error: the
//! [`Diagnostic`]s of a [`Report`](crate::Report).

use sheetrs_compiler::{FromCacheNote, Unsupported};
use sheetrs_protocol::Diagnostic;

/// The code of the diagnostic of a calculation the engine cannot run.
pub const RECALC_NOT_SUPPORTED: &str = "RECALC_NOT_SUPPORTED";

/// The diagnostic of an `operation` (`calculate`, or the recalculation of the results the deletions
/// left stale) that the engine cannot run, and `reason` says why (a function it does not have, a
/// reference it does not read). The workbook is written as the other operations leave it and marked
/// to be recalculated when it is opened, so the recalculation that is needed is done by the
/// spreadsheet program that opens it later.
#[must_use]
pub(crate) fn not_calculated(operation: &str, reason: &str) -> Diagnostic {
    Diagnostic::warning(
        RECALC_NOT_SUPPORTED,
        format!(
            "the workbook was not calculated ({reason}); it is marked to be recalculated, and the \
             recalculation it needs will be done by the spreadsheet program that opens it"
        ),
    )
    .with("operation", operation)
    .with("reason", reason)
}

/// The diagnostic of the startup of an ODS that the engine cannot calculate (`reason`): `LibreOffice` calculates some formulas
/// when it opens the file, and the engine could not, so the workbook goes on as the file has it, with the results it caches. It
/// has the code of [`RECALC_NOT_SUPPORTED`], and nothing marks the workbook: no operation of the plan left a result stale.
#[must_use]
pub(crate) fn not_calculated_on_load(reason: &str) -> Diagnostic {
    Diagnostic::warning(
        RECALC_NOT_SUPPORTED,
        format!(
            "what LibreOffice calculates when it opens the file was not calculated ({reason}); the workbook keeps the results the file caches"
        ),
    )
    .with("operation", "startup")
    .with("reason", reason)
}

/// The code of the diagnostic of a run that left as the file caches it the cells that depend on another workbook
/// (the stored value of the link, which `LibreOffice` shows).
pub const EXTERNAL_FROM_CACHE: &str = "EXTERNAL_FROM_CACHE";
/// The code of the diagnostic of a run that left as the file caches it the cells that call a function a macro may
/// define (macros are never run).
pub const MACRO_FROM_CACHE: &str = "MACRO_FROM_CACHE";
/// The code of the diagnostic of a run that left as the file caches it the cells that call a function the engine does
/// not implement (an add-in or extension, or a function of `LibreOffice` it lacks).
pub const EXTENSION_FROM_CACHE: &str = "EXTENSION_FROM_CACHE";

/// The code of the diagnostic of a run in secure mode that made an error the cells that call what the engine cannot
/// calculate and the run calculates (a macro, a function of `LibreOffice` the engine lacks).
pub const UNSUPPORTED_AS_ERROR: &str = "UNSUPPORTED_AS_ERROR";

/// The diagnostic of an `operation` that left as the file caches it what the engine cannot calculate (`note`),
/// because no edit reached the cell: it is the result of the program that wrote the file, which `LibreOffice` shows
/// too until something dirties the cell (OBSERVED, `sheetrs-compiler/docs/FIDELITY.md`), and which a dirtied cell
/// does not keep (an add-in that is not installed is `#NAME?`, an external value is its stored one).
#[must_use]
pub(crate) fn from_cache(operation: &str, note: &FromCacheNote) -> Diagnostic {
    if let Some(error) = &note.error {
        return Diagnostic::warning(
            UNSUPPORTED_AS_ERROR,
            format!(
                "{} formula(s) call {} that the engine cannot calculate and the run calculates them: secure mode makes the call an error ({error}), and what reads it is calculated from it (first {}: {})",
                note.cells,
                note.kind.as_str(),
                note.first,
                note.what.join(", ")
            ),
        )
        .with("operation", operation)
        .with("cells", note.cells.to_string())
        .with("first", note.first.clone())
        .with("error", error.clone());
    }
    let code = match note.kind {
        Unsupported::External => EXTERNAL_FROM_CACHE,
        Unsupported::Macro => MACRO_FROM_CACHE,
        Unsupported::Extension => EXTENSION_FROM_CACHE,
    };
    Diagnostic::warning(
        code,
        format!(
            "{} formula(s) that depend on {} keep the result the file caches, as LibreOffice keeps it until something dirties the cell (first {}: {})",
            note.cells,
            note.kind.as_str(),
            note.first,
            note.what.join(", ")
        ),
    )
    .with("operation", operation)
    .with("cells", note.cells.to_string())
    .with("first", note.first.clone())
}
