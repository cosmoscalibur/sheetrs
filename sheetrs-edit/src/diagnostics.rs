//! What a run did that the person who asked should know, and that is not an error: the
//! [`Diagnostic`]s of a [`Report`](crate::Report).

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

/// The code of the diagnostic of a conversion that took the results the file caches.
pub const CONVERTED_FROM_CACHE: &str = "CONVERTED_FROM_CACHE";

/// The diagnostic of a conversion of formulas to values (`operation`) that the engine could not calculate
/// (`reason`: a function it does not have, a reference it does not read) and that took the results the
/// file caches instead, because nothing the run did before it leaves them stale. They are the results of the
/// program that wrote the file, which can differ from what `LibreOffice` calculates when it opens it.
#[must_use]
pub(crate) fn converted_from_cache(operation: &str, reason: &str) -> Diagnostic {
    Diagnostic::warning(
        CONVERTED_FROM_CACHE,
        format!(
            "the formulas were converted to the results the file caches, not to the ones calculated ({reason})"
        ),
    )
    .with("operation", operation)
    .with("reason", reason)
}
