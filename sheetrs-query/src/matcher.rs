//! A search as `LibreOffice`'s Find & Replace runs it on the text of a cell.
//!
//! Observed in `LibreOffice` 26.8:
//!
//! - Case is ignored unless `match_case`; `entire_cells` asks the whole text of the cell, not a
//!   part of it: the **first** match must cover it all (`a|ab` does not match `ab`, and `a*b` does
//!   not match `abab`, because the first match of `a*b` is `ab`).
//! - `Wildcards`: `*` any text (as little as it can), `?` one character, `\` the next character as it is (not `~`, which
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
    entire: bool,
}

/// The pattern of a literal text.
fn escape(text: &str) -> String {
    regex::escape(text)
}

/// The pattern of a text with wildcards.
fn wildcards(text: &str) -> String {
    let mut out = String::new();
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            // A `*` takes as little as it can, but one that ends the pattern takes the rest of the
            // text (`LibreOffice` 26.8: `a*` finds `abab` whole, `b*b` finds `bb` in `xabbcx`).
            '*' if chars.peek().is_none() => out.push_str("(?s:.*)"),
            '*' => out.push_str("(?s:.*?)"),
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
        let flags = if find.match_case { "" } else { "(?i)" };
        let re =
            Regex::new(&format!("{flags}{body}")).map_err(|e| Error::Unsupported(e.to_string()))?;
        Ok(Self {
            re,
            regex_mode: find.mode == Mode::Regex,
            entire: find.entire_cells,
        })
    }

    /// Whether the text holds a hit.
    #[must_use]
    pub fn is_match(&self, text: &str) -> bool {
        self.hits(text).next().is_some()
    }

    /// The hits to take: every non-empty match, or, when the whole cell is asked, the first match if
    /// it covers the text.
    fn hits<'a>(&'a self, text: &'a str) -> impl Iterator<Item = regex::Captures<'a>> {
        let whole = |c: &regex::Captures<'_>| {
            c.get(0)
                .is_some_and(|m| m.start() == 0 && m.end() == text.len())
        };
        let entire = self.entire;
        self.re
            .captures_iter(text)
            .filter(|c| c.get(0).is_some_and(|m| !m.as_str().is_empty()))
            .take(if entire { 1 } else { usize::MAX })
            .filter(move |c| !entire || whole(c))
    }

    /// The text with every hit replaced by `with`, or `None` when it holds none.
    #[must_use]
    pub fn replace_all(&self, text: &str, with: &str) -> Option<String> {
        let mut out = String::with_capacity(text.len());
        let mut last = 0;
        let mut hit = false;
        for caps in self.hits(text) {
            let whole = caps.get(0)?;
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

    // Each of these is a result of `LibreOffice` 26.8 (`Sheet.replaceAll` with a `Cell`'s
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
    fn the_whole_cell_is_the_first_match_and_a_star_takes_as_little_as_it_can() {
        let whole = |t: &str, mode| Find {
            entire_cells: true,
            ..find(t, mode)
        };
        // `LibreOffice` 26.8: the first match must be the whole text.
        assert_eq!(replaced(&whole("a*b", Mode::Wildcards), "abab", "X"), None);
        // A `*` that ends the pattern takes the rest of the text.
        assert_eq!(
            replaced(
                &whole("*empty*", Mode::Wildcards),
                "Hello on an empty cell",
                "X"
            )
            .as_deref(),
            Some("X")
        );
        assert_eq!(
            replaced(&find("a*", Mode::Wildcards), "Hello on an empty cell", "X").as_deref(),
            Some("Hello on X")
        );
        assert_eq!(
            replaced(&whole("a*b", Mode::Wildcards), "axb", "X").as_deref(),
            Some("X")
        );
        assert_eq!(replaced(&whole("a|ab", Mode::Regex), "ab", "X"), None);
        assert_eq!(
            replaced(&whole("ab|a", Mode::Regex), "ab", "X").as_deref(),
            Some("X")
        );
        assert_eq!(
            replaced(&whole("a.*b", Mode::Regex), "abab", "X").as_deref(),
            Some("X")
        );
        assert_eq!(replaced(&whole("a.*?b", Mode::Regex), "abab", "X"), None);
        // Not the whole cell: `a*b` stops at the first `b`.
        assert_eq!(
            replaced(&find("a*b", Mode::Wildcards), "xabcbx", "X").as_deref(),
            Some("xXcbx")
        );
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
