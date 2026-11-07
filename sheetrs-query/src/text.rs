//! What a cell says as text.
//!
//! Two texts: the one a person **sees** ([`TextContext::display`]: the value written with its
//! format, in a locale) and the one a search **sees** ([`TextContext::searched`]), which is not
//! the same. `LibreOffice` searches the text of the input line of a cell, not its display:
//! a number is looked for as the digits it was typed with (`1234.5` for a cell that shows
//! `1,234.50`), a percentage as `25.6%`, a date as the date the input line writes (`03/15/2023`, or
//! the ISO date when the format of the cell is an ISO one), and the formula of a formula cell in the
//! notation of `LibreOffice` (`=SUM(A1,B1)`, `=$'My Sheet'.A1`, written with the separators of the
//! locale) — never in the notation of Excel. These are the observations of 24.2.7 for the `en-US`
//! locale that [`TextContext`] reproduces.

use crate::error::{Error, Result};
use sheetrs_engine::dates::serial_to_ymd_quiet;
use sheetrs_engine::options::{DateOrder, DateSystem, Locale};
use sheetrs_formula::calc_print::{CalcSyntax, to_calc};
use sheetrs_io::reader::{Cell, CellValue};
use sheetrs_protocol::SearchIn;

/// What decides how a number is written: the locale and the date system of the workbook.
#[derive(Debug, Clone)]
pub struct TextContext {
    locale: Locale,
    sys: DateSystem,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    General,
    Percent,
    Date { iso: bool, time: bool },
    Time { total_hours: bool },
}

/// What a format code is about, without the quoted text and the bracketed parts that are not
/// elapsed time (`[Red]`, `[$EUR]`).
fn kind(format: Option<&str>) -> Kind {
    let Some(format) = format else {
        return Kind::General;
    };
    let mut letters = String::new();
    let mut percent = false;
    let mut total_hours = false;
    let mut chars = format.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                let mut inner = String::new();
                for q in chars.by_ref() {
                    if q == ']' {
                        break;
                    }
                    inner.push(q);
                }
                let inner = inner.to_ascii_lowercase();
                if inner.chars().all(|c| c == 'h') && !inner.is_empty() {
                    total_hours = true;
                    letters.push('h');
                } else if !inner.is_empty() && inner.chars().all(|c| matches!(c, 'm' | 's')) {
                    letters.push(inner.chars().next().unwrap_or('m'));
                }
            }
            '%' => percent = true,
            c if c.is_ascii_alphabetic() => letters.push(c.to_ascii_lowercase()),
            _ => {}
        }
    }
    // `E+00` is scientific, `General` and `AM/PM` carry letters of their own.
    let letters = letters.replace("general", "").replace("am/pm", "");
    let has = |c: char| letters.contains(c);
    let date = has('y') || has('d') || (has('m') && !(has('h') || has('s')));
    let time = has('h') || has('s');
    if date {
        let order: String = letters
            .chars()
            .filter(|c| matches!(c, 'y' | 'm' | 'd'))
            .collect();
        let iso = order.starts_with('y') && format.contains('-');
        Kind::Date { iso, time }
    } else if time {
        Kind::Time { total_hours }
    } else if percent {
        Kind::Percent
    } else {
        Kind::General
    }
}

fn is_general(format: Option<&str>) -> bool {
    format.is_none_or(|f| f.eq_ignore_ascii_case("general"))
}

impl TextContext {
    /// # Errors
    ///
    /// When `locale` is not one the engine has.
    pub fn new(locale: Option<&str>, date1904: bool) -> Result<Self> {
        let locale = match locale {
            None => Locale::en_us(),
            Some(tag) => Locale::by_tag(tag)
                .ok_or_else(|| Error::Request(format!("unknown locale `{tag}`")))?,
        };
        Ok(Self {
            locale,
            sys: if date1904 {
                DateSystem::NULL_1904
            } else {
                DateSystem::NULL_1899
            },
        })
    }

    /// The number as `General` writes it.
    fn general(&self, x: f64) -> String {
        sheetrs_engine::coerce::fmt_number(x, &self.locale)
    }

    fn boolean(&self, b: bool) -> String {
        if b {
            self.locale.true_text.clone()
        } else {
            self.locale.false_text.clone()
        }
    }

    /// The cell as a person sees it: the value written with the format of the cell.
    #[must_use]
    pub fn display(&self, cell: &Cell) -> String {
        let format = cell.num_fmt.as_deref();
        match &cell.value {
            CellValue::Empty => String::new(),
            CellValue::Text(t) => format
                .and_then(|f| sheetrs_engine::format::format_text(t, f, &self.locale))
                .unwrap_or_else(|| t.to_string()),
            CellValue::Boolean(b) => self.boolean(*b),
            CellValue::Error(e) => e.to_string(),
            CellValue::Number(x) => match format {
                Some(f) if !is_general(Some(f)) => {
                    sheetrs_engine::format::format_number(*x, f, &self.locale, self.sys)
                        .unwrap_or_else(|| self.general(*x))
                }
                _ => self.general(*x),
            },
        }
    }

    /// A number as the input line of the cell writes it.
    #[must_use]
    pub fn input_line(&self, x: f64, format: Option<&str>) -> String {
        match kind(format) {
            Kind::General => self.general(x),
            Kind::Percent => format!("{}%", self.general(x * 100.0)),
            Kind::Date { iso, time } => {
                let Some((y, m, d)) = serial_to_ymd_quiet(x, self.sys) else {
                    return self.general(x);
                };
                let sep = self.locale.date_sep;
                let date = if iso {
                    format!("{y:04}-{m:02}-{d:02}")
                } else {
                    match self.locale.date_order {
                        DateOrder::Mdy => format!("{m:02}{sep}{d:02}{sep}{y:04}"),
                        DateOrder::Dmy => format!("{d:02}{sep}{m:02}{sep}{y:04}"),
                        DateOrder::Ymd => format!("{y:04}{sep}{m:02}{sep}{d:02}"),
                    }
                };
                if time {
                    format!("{date} {}", clock(x, false))
                } else {
                    date
                }
            }
            Kind::Time { total_hours } => clock(x, total_hours),
        }
    }

    /// The formula in the notation of `LibreOffice`, with the `=`.
    #[must_use]
    pub fn formula_text(&self, formula: &str, sheet: &str) -> String {
        let syntax = CalcSyntax::for_decimal(self.locale.decimal);
        match sheetrs_formula::parse(formula) {
            Ok(expr) => format!("={}", to_calc(&expr, sheet, &syntax)),
            Err(_) => format!("={formula}"),
        }
    }

    /// The text of the cell that a search of `search_in` looks at, or `None` when the search does
    /// not look at the cell: an empty one, an error, and (unless `include_formulas`) the result of
    /// a formula.
    #[must_use]
    pub fn searched(
        &self,
        cell: &Cell,
        sheet: &str,
        search_in: SearchIn,
        include_formulas: bool,
    ) -> Option<String> {
        if let Some(formula) = cell.formula.as_deref() {
            return match search_in {
                SearchIn::Formula => Some(self.formula_text(formula, sheet)),
                SearchIn::Value if include_formulas => self.constant_text(cell),
                SearchIn::Value => None,
            };
        }
        self.constant_text(cell)
    }

    fn constant_text(&self, cell: &Cell) -> Option<String> {
        match &cell.value {
            CellValue::Empty | CellValue::Error(_) => None,
            CellValue::Text(t) => Some(t.to_string()),
            CellValue::Boolean(b) => Some(self.boolean(*b)),
            CellValue::Number(x) => Some(self.input_line(*x, cell.num_fmt.as_deref())),
        }
    }
}

/// `HH:MM:SS` of the fraction of a day (the hours of the whole value when it is elapsed time).
fn clock(x: f64, total_hours: bool) -> String {
    let seconds = (x.rem_euclid(1.0) * 86_400.0).round() as i64;
    let hours = if total_hours {
        (x * 24.0).floor() as i64
    } else {
        seconds / 3600
    };
    format!("{hours:02}:{:02}:{:02}", seconds / 60 % 60, seconds % 60)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx() -> TextContext {
        TextContext::new(None, false).unwrap()
    }

    // Observed in `LibreOffice` 24.2.7 (en-US): the text a search finds a cell by.
    #[test]
    fn a_number_is_searched_as_typed_and_not_as_shown() {
        let c = ctx();
        assert_eq!(c.input_line(1234.5, Some("#,##0.00")), "1234.5");
        assert_eq!(c.input_line(1234.5, Some("$#,##0.00")), "1234.5");
        assert_eq!(c.input_line(1_234_567.891, Some("0.00E+00")), "1234567.891");
        assert_eq!(c.input_line(0.1 + 0.2, None), "0.3");
        assert_eq!(
            c.input_line(-1234.5, Some("#,##0.00;[Red]-#,##0.00")),
            "-1234.5"
        );
    }

    #[test]
    fn a_percentage_is_searched_with_its_sign() {
        assert_eq!(ctx().input_line(0.256, Some("0.0%")), "25.6%");
        assert_eq!(ctx().input_line(0.5, Some("0%")), "50%");
    }

    #[test]
    fn a_date_is_searched_as_the_input_line_writes_it() {
        let c = ctx();
        assert_eq!(c.input_line(45000.0, Some("d-mmm-yy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("mmmm d, yyyy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("yyyy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("yyyy-mm-dd")), "2023-03-15");
        assert_eq!(
            c.input_line(45000.0, Some("yyyy-mm-dd hh:mm:ss")),
            "2023-03-15 00:00:00"
        );
        assert_eq!(
            c.input_line(45000.75, Some("dd/mm/yyyy hh:mm")),
            "03/15/2023 18:00:00"
        );
        assert_eq!(c.input_line(0.75, Some("[hh]:mm:ss")), "18:00:00");
    }

    #[test]
    fn a_formula_is_searched_in_the_notation_of_libreoffice() {
        let c = ctx();
        assert_eq!(c.formula_text("SUM(A1,A1)", "T"), "=SUM(A1,A1)");
        assert_eq!(c.formula_text("'U Sheet'!A1+1", "T"), "=$'U Sheet'.A1+1");
        assert_eq!(c.formula_text("A1+$A$1", "T"), "=A1+$A$1");
    }

    #[test]
    fn the_display_is_the_value_with_its_format() {
        let c = ctx();
        let cell = Cell {
            value: CellValue::Number(1234.5),
            num_fmt: Some("#,##0.00".into()),
            ..Cell::default()
        };
        assert_eq!(c.display(&cell), "1,234.50");
        let general = Cell {
            value: CellValue::Number(0.5),
            ..Cell::default()
        };
        assert_eq!(c.display(&general), "0.5");
    }

    #[test]
    fn a_search_skips_errors_and_the_results_of_formulas_unless_asked() {
        let c = ctx();
        let formula = Cell {
            value: CellValue::Number(2469.0),
            formula: Some("A1*2".into()),
            ..Cell::default()
        };
        assert_eq!(c.searched(&formula, "T", SearchIn::Value, false), None);
        assert_eq!(
            c.searched(&formula, "T", SearchIn::Value, true).as_deref(),
            Some("2469")
        );
        assert_eq!(
            c.searched(&formula, "T", SearchIn::Formula, false)
                .as_deref(),
            Some("=A1*2")
        );
    }
}
