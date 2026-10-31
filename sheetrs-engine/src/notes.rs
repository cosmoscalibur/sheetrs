//! What the engine met that LibreOffice does and it does not reproduce: a result that is only an
//! error value here. The compiler turns what happens while it folds the workbook into a build
//! error that names the cell; at run time it stays an error value.

use std::cell::RefCell;

/// A date serial outside 0001-01-01 … 32768-01-01.
pub const DATE_RANGE: &str =
    "handle a date serial outside 0001-01-01 … 32768-01-01, the range LibreOffice supports";
/// A `TEXT` code with the language tag of a language the engine has no names for.
pub const TEXT_LANGUAGE: &str =
    "use a TEXT code that names a language the engine has no month and day names for";

/// A regular expression outside the subset of ICU's syntax that is reproduced.
pub const REGEX: &str =
    "use a regular expression outside the subset of ICU's syntax the engine reproduces";

/// A regular expression run on text with a letter whose case folding is several characters.
pub const REGEX_FOLD: &str = "search with a regular expression in text that has letters whose case folding is several characters (`ß`, `ﬁ`, `İ`…), which ICU matches in ways the engine does not reproduce";

/// `XLOOKUP` search mode 2 or -2 over a range.
pub const XLOOKUP_BINARY: &str = "use XLOOKUP with a binary search (search mode 2 or -2) over a range, which LibreOffice runs through an iterator of its own that is not reproduced";

/// `F.INV(1, …)`: LibreOffice's iteration ends at the top of its search range and answers a
/// number (2.37E+81 for 5 and 8 degrees of freedom) that is an artifact of that range.
pub const F_INV_ONE: &str = "use F.INV with a probability of 1, for which LibreOffice answers the end of its search range (an artifact the engine does not reproduce)";

/// A distribution (t, F, beta) with more than 10 million degrees of freedom or shape: the
/// continued fraction of LibreOffice loses digits there (`T.DIST(2, 1E8, TRUE)` is off by 2E-9,
/// `F.DIST(2, 5, 1E9, TRUE)` by 4E-8) in ways the engine does not reproduce.
pub const BETA_LARGE: &str = "use a t, F or beta distribution with more than 10 million degrees of freedom or shape, where LibreOffice's calculation loses digits that the engine does not reproduce";

/// `SORT` of an array, or `SORTBY`, by keys that have an error: the order LibreOffice gives is not
/// a function of the keys (`SORTBY(A1:A4, M1:M4)` with `#N/A` in M2 gives the rows as 1, ?, ?, 4
/// and, reversed, as 4, ?, ?, 3). `SORT` of a range does order them (the errors last, as 0).
pub const SORT_ERROR_KEY: &str = "use SORT of an array or SORTBY with sort keys that have an error, where LibreOffice's order is not reproduced";

/// `ERRORTYPE` of `#VALUE!`: LibreOffice has several codes behind it (`Err:502` for an argument
/// outside the domain, `Err:519` for a value of the wrong type, and others) that the engine does
/// not keep apart.
pub const ERRORTYPE_VALUE: &str = "use ERRORTYPE of a #VALUE! error or of the error of a cycle, where LibreOffice answers one of several internal codes that the engine does not tell apart";

/// `DCOUNT` and `DCOUNTA` with the field 0: what LibreOffice counts depends on the criteria in ways
/// that are not a function of the records (`DCOUNT(…, 0, criteria)` is 7 over 8 matching records,
/// `DCOUNTA` of the same is `#VALUE!`).
pub const DB_FIELD_ZERO: &str = "use DCOUNT or DCOUNTA with field 0, where LibreOffice's answer is not a function of the records that match";

/// `LINEST` or `LOGEST` with statistics over a perfect fit: the standard errors, `F` and the sums
/// of squares are the rounding noise of LibreOffice's computation (`F` of 4E31 in one case, `#N/A`
/// in an equal one), which the engine's own rounding does not repeat.
pub const PERFECT_FIT: &str = "use LINEST or LOGEST with statistics over data that a line fits exactly, where LibreOffice answers the rounding noise of its computation";

/// `FREQUENCY` with equal bins: which of them gets the count depends on LibreOffice's sort (the
/// first of two bins of 9 over data that are all at most 9, the second of two bins of 3 over data
/// with three values of at most 3).
pub const FREQUENCY_EQUAL_BINS: &str = "use FREQUENCY with equal bins, where LibreOffice's choice of the bin that takes the count is not reproduced";

/// `TBILLYIELD` with a settlement and a maturity that are both the last day of February: the days
/// LibreOffice counts there (359 or 360, where the rule of the other dates gives 361) follow no
/// rule the probes could find.
pub const TBILLYIELD_FEB: &str = "use TBILLYIELD with a settlement and a maturity on the last day of February, where LibreOffice's day count is not reproduced";

/// `LEFTB`, `RIGHTB`, `MIDB` or `REPLACEB` that cut through a character above the Basic Multilingual
/// Plane (an emoji, a rare ideograph): LibreOffice answers a lone half of a surrogate pair, which
/// is not text the engine can hold.
pub const FOURIER_LITERAL: &str = "use FOURIER on an array written in the formula, grouped by rows: LibreOffice transposes that array in the formula itself each time it calculates, so the result depends on how many times the document has been calculated";
pub const ETS_RANDOM: &str = "use FORECAST.ETS.PI.ADD or FORECAST.ETS.PI.MULT with seasons: LibreOffice computes the interval from random scenarios, so its result changes with each calculation";
pub const ETS_BOUNDS: &str = "use FORECAST.ETS with a target or series for which LibreOffice reads beyond the end of its arrays, which gives no defined result";
pub const ETS_SORT: &str = "use FORECAST.ETS with a timeline that LibreOffice's sort orders by heap sort, which is not reproduced";
pub const CELL_FORMAT: &str = "use CELL with a type that depends on the cell's format, its protection (color, format, parentheses, prefix, protect, width)";
pub const CELL_SHEET_NAME: &str =
    "use CELL(\"address\") of a cell of a sheet whose name needs quotes";
pub const SURROGATE_CUT: &str = "use LEFTB, RIGHTB, MIDB or REPLACEB with a cut through a character above the Basic Multilingual Plane, where LibreOffice answers half of a surrogate pair";

thread_local! {
    static NOTES: RefCell<Vec<&'static str>> = const { RefCell::new(Vec::new()) };
}

pub fn note(what: &'static str) {
    NOTES.with(|n| {
        let mut n = n.borrow_mut();
        if !n.contains(&what) {
            n.push(what);
        }
    });
}

/// The notes since the last call.
pub fn take() -> Vec<&'static str> {
    NOTES.with(|n| std::mem::take(&mut *n.borrow_mut()))
}
