//! A search as `LibreOffice`'s Find & Replace runs it on the text of a cell.
//!
//! Observed in `LibreOffice` 24.2:
//!
//! - Case is ignored unless `match_case`; `entire_cells` asks the whole text of the cell, not a
//!   part of it.
//! - `Wildcards`: `*` any text, `?` one character, `\` the next character as it is (not `~`, which
//!   is the escape of `COUNTIF`).
//! - `Regex`: ICU syntax; the part of it that was compared with `LibreOffice` is the subset of
//!   `sheetrs_engine::regex_subset`, and a pattern outside it is refused instead of giving another
//!   result. `.` does not match a line break.
//! - A match of no characters is not a hit.
//! - Every hit of a cell is replaced. In a regular expression the replacement reads `&` as the
//!   whole match, `$0` to `$9` as the groups, `\t` as a tab, and `\&`, `\$` and `\\` as the
//!   character they escape; in the other modes the replacement is taken as it is.

use crate::error::{Error, Result};
use regex::Regex;
use sheetrs_protocol::{Find, Mode};

/// A [`Find`] ready to test texts.
#[derive(Debug)]
pub struct Matcher {
    re: Regex,
    regex_mode: bool,
}

/// The pattern of a literal text.
fn escape(text: &str) -> String {
    regex::escape(text)
}

/// The pattern of a text with wildcards.
fn wildcards(text: &str) -> String {
    let mut out = String::new();
    let mut chars = text.chars();
    while let Some(c) = chars.next() {
        match c {
            '*' => out.push_str("(?s:.*)"),
            '?' => out.push_str("(?s:.)"),
            '\\' => match chars.next() {
                Some(next) => out.push_str(&escape(&next.to_string())),
                None => out.push_str("\\\\"),
            },
            c => out.push_str(&escape(&c.to_string())),
        }
    }
    out
}

impl Matcher {
    /// # Errors
    ///
    /// When the regular expression is outside what is reproduced, or does not compile.
    pub fn new(find: &Find) -> Result<Self> {
        let body = match find.mode {
            Mode::Literal => escape(&find.text),
            Mode::Wildcards => wildcards(&find.text),
            Mode::Regex => sheetrs_engine::regex_subset::translate_capturing(&find.text)
                .map_err(|e| Error::Unsupported(e.to_string()))?,
        };
        let body = if find.entire_cells {
            format!("\\A(?:{body})\\z")
        } else {
            body
        };
        let flags = if find.match_case { "" } else { "(?i)" };
        let re =
            Regex::new(&format!("{flags}{body}")).map_err(|e| Error::Unsupported(e.to_string()))?;
        Ok(Self {
            re,
            regex_mode: find.mode == Mode::Regex,
        })
    }

    /// Whether the text holds a hit.
    #[must_use]
    pub fn is_match(&self, text: &str) -> bool {
        self.re.find_iter(text).any(|m| !m.as_str().is_empty())
    }

    /// The text with every hit replaced by `with`, or `None` when it holds none.
    #[must_use]
    pub fn replace_all(&self, text: &str, with: &str) -> Option<String> {
        let mut out = String::with_capacity(text.len());
        let mut last = 0;
        let mut hit = false;
        for caps in self.re.captures_iter(text) {
            let whole = caps.get(0)?;
            if whole.as_str().is_empty() {
                continue;
            }
            hit = true;
            out.push_str(&text[last..whole.start()]);
            if self.regex_mode {
                expand(with, &caps, &mut out);
            } else {
                out.push_str(with);
            }
            last = whole.end();
        }
        if !hit {
            return None;
        }
        out.push_str(&text[last..]);
        Some(out)
    }
}

/// The replacement of a regular expression, with its `&`, `$n` and escapes.
fn expand(with: &str, caps: &regex::Captures<'_>, out: &mut String) {
    let mut chars = with.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '&' => out.push_str(caps.get(0).map_or("", |m| m.as_str())),
            '$' => {
                if let Some(d) = chars.peek().and_then(|d| d.to_digit(10)) {
                    chars.next();
                    out.push_str(caps.get(d as usize).map_or("", |m| m.as_str()));
                } else {
                    out.push('$');
                }
            }
            '\\' => match chars.peek().copied() {
                Some('t') => {
                    chars.next();
                    out.push('\t');
                }
                Some(e @ ('&' | '$' | '\\')) => {
                    chars.next();
                    out.push(e);
                }
                _ => out.push('\\'),
            },
            c => out.push(c),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn find(text: &str, mode: Mode) -> Find {
        Find {
            mode,
            ..Find::literal(text)
        }
    }

    fn replaced(f: &Find, text: &str, with: &str) -> Option<String> {
        Matcher::new(f).unwrap().replace_all(text, with)
    }

    // Each of these is a result of `LibreOffice` 24.2.7 (`Sheet.replaceAll` with a `Cell`'s
    // string): the comment says what it did.

    #[test]
    fn every_hit_of_a_cell_is_replaced_and_case_is_ignored_unless_asked() {
        let f = find("banana", Mode::Literal);
        assert_eq!(replaced(&f, "Banana BANANA", "x").as_deref(), Some("x x"));
        let f = Find {
            match_case: true,
            ..find("banana", Mode::Literal)
        };
        assert_eq!(
            replaced(&f, "Banana banana", "x").as_deref(),
            Some("Banana x")
        );
    }

    #[test]
    fn entire_cells_asks_the_whole_text() {
        let f = Find {
            entire_cells: true,
            ..find("banana", Mode::Literal)
        };
        assert_eq!(replaced(&f, "banana banana", "x"), None);
        assert_eq!(replaced(&f, "banana", "x").as_deref(), Some("x"));
    }

    #[test]
    fn a_literal_replacement_is_not_expanded() {
        assert_eq!(
            replaced(&find("b", Mode::Literal), "abc", "[&]").as_deref(),
            Some("a[&]c")
        );
    }

    #[test]
    fn a_regular_expression_replacement_has_groups_and_escapes() {
        let f = find("b", Mode::Regex);
        assert_eq!(replaced(&f, "abc", "[&]").as_deref(), Some("a[b]c"));
        assert_eq!(replaced(&f, "abc", "[$0]").as_deref(), Some("a[b]c"));
        assert_eq!(
            replaced(&f, "abc", "x\\ty").as_deref(),
            Some("a".to_string() + "x\ty" + "c").as_deref()
        );
        assert_eq!(replaced(&f, "abc", "x\\&y").as_deref(), Some("ax&yc"));
        assert_eq!(replaced(&f, "abc", "x\\$y").as_deref(), Some("ax$yc"));
        assert_eq!(replaced(&f, "abc", "x\\\\y").as_deref(), Some("ax\\yc"));
        assert_eq!(replaced(&f, "abc", "x\\ny").as_deref(), Some("ax\\nyc"));
        let g = find("(b)(c)", Mode::Regex);
        assert_eq!(replaced(&g, "abc", "$2$1").as_deref(), Some("acb"));
    }

    #[test]
    fn a_dot_does_not_match_a_line_break() {
        assert_eq!(replaced(&find("a.b", Mode::Regex), "a\nb", "Z"), None);
    }

    #[test]
    fn wildcards_use_the_backslash_to_escape() {
        let w = |t: &str| find(t, Mode::Wildcards);
        assert_eq!(replaced(&w("*.txt"), "abc.txt", "z").as_deref(), Some("z"));
        assert_eq!(replaced(&w("a?c"), "abcd", "Z").as_deref(), Some("Zd"));
        assert_eq!(replaced(&w("a*c"), "xabbcx", "Z").as_deref(), Some("xZx"));
        assert_eq!(replaced(&w("a\\*c"), "a*c", "Z").as_deref(), Some("Z"));
        assert_eq!(replaced(&w("a\\*c"), "abc", "Z"), None);
        assert_eq!(
            replaced(&w("a~*c"), "a*c", "Z"),
            None,
            "`~` escapes only in COUNTIF"
        );
        let whole = Find {
            entire_cells: true,
            ..w("a*c")
        };
        assert_eq!(replaced(&whole, "xabbcx", "Z"), None);
    }

    #[test]
    fn a_pattern_that_is_not_reproduced_is_refused() {
        // `LibreOffice` replaces at the end of the text with `$` but not at the start with `^`:
        // behaviours the subset does not promise.
        assert!(matches!(
            Matcher::new(&find("^", Mode::Regex)),
            Err(Error::Unsupported(_))
        ));
    }
}
