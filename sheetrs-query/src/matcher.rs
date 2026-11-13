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
//! - `^` and `$` anchor the text of the cell, not its lines: `^` is its start, `$` its end or the
//!   place before a line terminator that ends it. A search ([`Matcher::for_find`]) and a
//!   replacement ([`Matcher::new`]) read them.
//! - A match of no characters is not a hit, with one exception: at the very end of the text it is
//!   (`abc` with `$` replaced by `X` gives `abcX`, with `^` or `x*` nothing changes). The hits are
//!   found as `LibreOffice` finds them: from the start of the text, each search begins where the
//!   replacement of the last hit ended (the replacement is made in the text itself, so `^a` replaced
//!   by nothing empties `aa`), or one character after an empty match that was skipped, and none
//!   begins at the end of the text (`a` with `a*` replaced by `X` gives `X`, not `XX`; `a` and a line break with
//!   `$` is not changed: the one empty match, before the line break, is skipped, and the end is
//!   not searched). A search ([`Matcher::for_find`]) refuses a pattern that can match the empty
//!   text; a replacement reads it.
//! - Every hit of a cell is replaced. In a regular expression the replacement reads `&` as the
//!   whole match, `$0` to `$9` as the groups, `\t` as a tab, and `\&`, `\$` and `\\` as the
//!   character they escape; in the other modes the replacement is taken as it is.

use crate::error::{Error, Result};
use fancy_regex::{Captures, Regex};
use sheetrs_protocol::{Find, Mode};

/// A [`Find`] ready to test texts.
#[derive(Debug)]
pub struct Matcher {
    re: Regex,
    regex_mode: bool,
    entire: bool,
    /// Whether the pattern matches the empty text in the middle of a text (`x*` does, `$` and `^`
    /// do not): see [`Matcher::next_hit`].
    empty_inside: bool,
}

/// The pattern of a literal text.
fn escape(text: &str) -> String {
    fancy_regex::escape(text).into_owned()
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
    /// A matcher to replace with: a regular expression may have the anchors `^` and `$` and may
    /// match the empty text.
    ///
    /// # Errors
    ///
    /// When the regular expression is outside what is reproduced, or does not compile.
    pub fn new(find: &Find) -> Result<Self> {
        Self::build(find, true)
    }

    /// A matcher to find the cells with a hit: a regular expression may have the anchors `^` and
    /// `$`, but not match the empty text (what `findAll` does with those is not reproduced yet).
    ///
    /// # Errors
    ///
    /// When the regular expression is outside what is reproduced, or does not compile.
    pub fn for_find(find: &Find) -> Result<Self> {
        Self::build(find, false)
    }

    fn build(find: &Find, replacing: bool) -> Result<Self> {
        let body = match find.mode {
            Mode::Literal => escape(&find.text),
            Mode::Wildcards => wildcards(&find.text),
            Mode::Regex => {
                let translated = if replacing {
                    sheetrs_engine::regex_subset::translate_replace(&find.text)
                } else {
                    sheetrs_engine::regex_subset::translate_find(&find.text)
                };
                translated.map_err(|e| Error::Unsupported(e.to_string()))?
            }
        };
        let flags = if find.match_case { "" } else { "(?i)" };
        let re =
            Regex::new(&format!("{flags}{body}")).map_err(|e| Error::Unsupported(e.to_string()))?;
        // A text of three characters that no pattern of a search names, and its middle.
        let probe = "\u{E000}\u{E000}\u{E000}";
        let empty_inside = re
            .captures_from_pos(probe, '\u{E000}'.len_utf8())
            .ok()
            .flatten()
            .and_then(|c| c.get(0).map(|m| m.start() == 3 && m.end() == 3))
            .unwrap_or(false);
        Ok(Self {
            re,
            regex_mode: find.mode == Mode::Regex,
            entire: find.entire_cells,
            empty_inside,
        })
    }

    /// Whether the text holds a hit.
    #[must_use]
    pub fn is_match(&self, text: &str) -> bool {
        self.next_hit(text, 0, false).is_some_and(|caps| {
            !self.entire
                || caps
                    .get(0)
                    .is_some_and(|m| m.start() == 0 && m.end() == text.len())
        })
    }

    /// The next hit that begins at or after `start`, as `LibreOffice` searches: an empty match is
    /// not a hit, unless it is at the end of the text, and no search begins at the end of it.
    ///
    /// `LibreOffice` counts in UTF-16 units and goes on one unit after a skipped empty match, so
    /// after one before a character outside the Basic Multilingual Plane its next search begins
    /// inside that character (`mid`), which is not the end of the text yet when the character is
    /// the last one. There a pattern that matches the empty text anywhere (`x*`) matches it again
    /// and is skipped again; one that does not (`$`) finds its match after the character, at the
    /// end of the text if the character is the last one. A pattern that matches the half of the
    /// character (`.`) gives a text that is not valid: that is not reproduced.
    fn next_hit<'a>(
        &self,
        text: &'a str,
        mut start: usize,
        mut mid: bool,
    ) -> Option<Captures<'a, str>> {
        while start < text.len() || mid {
            let caps = self.re.captures_from_pos(text, start).ok().flatten()?;
            let whole = caps.get(0)?;
            if !whole.as_str().is_empty() || whole.start() == text.len() {
                return Some(caps);
            }
            // Skipped: the next search begins one character on.
            let skipped = text[whole.start()..].chars().next()?;
            start = whole.start() + skipped.len_utf8();
            mid = skipped.len_utf16() == 2 && !self.empty_inside;
        }
        None
    }

    /// The text with every hit replaced by `with`, or `None` when it holds none.
    ///
    /// The replacement is made in the text itself and the next search begins after it, in what is
    /// left: `^a` replaced by nothing empties `aa`, since the second `a` is then at the start.
    /// With `entire_cells` only the first hit is taken, and only if it covers the text.
    #[must_use]
    pub fn replace_all(&self, text: &str, with: &str) -> Option<String> {
        let mut text = text.to_string();
        let mut start = 0;
        let mut replaced = false;
        while let Some(caps) = self.next_hit(&text, start, false) {
            let whole = caps.get(0)?;
            let range = whole.start()..whole.end();
            if self.entire && range != (0..text.len()) {
                return None;
            }
            let mut new = String::new();
            if self.regex_mode {
                expand(with, &caps, &mut new);
            } else {
                new.push_str(with);
            }
            start = range.start + new.len();
            text.replace_range(range, &new);
            replaced = true;
            if self.entire {
                break;
            }
        }
        replaced.then_some(text)
    }
}

/// The replacement of a regular expression, with its `&`, `$n` and escapes.
fn expand(with: &str, caps: &Captures<'_, str>, out: &mut String) {
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

    fn found(pattern: &str, entire_cells: bool, text: &str) -> bool {
        let f = Find {
            entire_cells,
            ..find(pattern, Mode::Regex)
        };
        Matcher::for_find(&f).unwrap().is_match(text)
    }

    #[test]
    fn a_replacement_reads_the_anchors_as_libreoffice_does() {
        let r = |p: &str, text: &str, with: &str| replaced(&find(p, Mode::Regex), text, with);
        assert_eq!(r("^a", "aa", "X").as_deref(), Some("Xa"));
        assert_eq!(r("a$", "aa", "X").as_deref(), Some("aX"));
        assert_eq!(r("^(a|b)$", "b", "X").as_deref(), Some("X"));
        // An empty match is not replaced, but one at the end of the text is.
        assert_eq!(r("^", "abc", "X"), None);
        assert_eq!(r("$", "abc", "X").as_deref(), Some("abcX"));
        assert_eq!(r("x*", "abc", "X"), None);
        // No search begins at the end of the text.
        assert_eq!(r("a*", "ab", "X").as_deref(), Some("Xb"));
        assert_eq!(r("(^|$)", "a", "X"), None);
        // The empty match before the line break that ends the text is skipped, and so is the end.
        assert_eq!(r("$", "a\n", "X"), None);
        assert_eq!(r("a$", "a\n", "X").as_deref(), Some("X\n"));
    }

    #[test]
    fn a_replacement_refuses_what_is_not_reproduced() {
        for pattern in ["", "^*", "\\b", "(a*)*"] {
            assert!(
                matches!(
                    Matcher::new(&find(pattern, Mode::Regex)),
                    Err(Error::Unsupported(_))
                ),
                "{pattern}"
            );
        }
    }

    #[test]
    fn a_search_reads_the_anchors_of_the_text_of_the_cell() {
        assert!(found("^ab", false, "abc"));
        assert!(!found("^bc", false, "abc"));
        assert!(found("bc$", false, "abc"));
        assert!(!found("ab$", false, "abc"));
        // The text, not its lines.
        assert!(!found("^b", false, "a\nb"));
        assert!(!found("a$", false, "a\nb"));
        // `$` is also the place before the line terminator that ends the text.
        for end in [
            "\n", "\r", "\r\n", "\u{b}", "\u{c}", "\u{85}", "\u{2028}", "\u{2029}",
        ] {
            assert!(found("a$", false, &format!("a{end}")), "{end:?}");
        }
        assert!(!found("a$", false, "a\n\n"));
        assert!(!found("a$", false, "a\r\nb"));
    }

    #[test]
    fn an_anchor_in_a_group_or_in_the_middle_of_a_pattern() {
        assert!(found("^(Code|Resumen Contadia)$", true, "Resumen Contadia"));
        assert!(!found("^(Code|Resumen Contadia)$", false, "Code x"));
        assert!(found("(^a|b$)", false, "xxb"));
        assert!(found("(^a|b$)", false, "axx"));
        assert!(!found("(^a|b$)", false, "xax"));
        assert!(!found("a^b", false, "ab"));
        assert!(!found("a$b", false, "ab"));
        // The first match must be the whole text with `entire_cells`, anchors or not.
        assert!(!found("^a|ab$", true, "ab"));
        assert!(found("^ab|a$", true, "ab"));
    }

    #[test]
    fn a_dot_next_to_an_anchor_leaves_out_a_line_break() {
        assert!(!found("^.b", false, "\nb"));
        assert!(!found("a.$", false, "a\n"));
        assert!(found("^a.$", false, "ab"));
    }

    #[test]
    fn a_search_refuses_a_pattern_that_can_match_the_empty_text() {
        for pattern in ["^", "$", "^$", "^a*", "a*$", "(^|$)"] {
            assert!(
                matches!(
                    Matcher::for_find(&find(pattern, Mode::Regex)),
                    Err(Error::Unsupported(_))
                ),
                "{pattern}"
            );
        }
    }
}
