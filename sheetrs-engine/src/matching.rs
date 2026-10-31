//! Text pattern matching under the document's search settings.
//!
//! Matching is always case-insensitive: LibreOffice's "case sensitive" setting
//! governs the comparison operators only, not search criteria or lookups
//! (observed: `COUNTIF(r,"APPLE")` counts `apple` with the setting on).

use crate::options::{Search, SearchMode};
use fancy_regex::Regex;

pub enum Matcher {
    /// No special characters: compare directly.
    Plain {
        /// The pattern with the case ignored as the equality of a whole cell does.
        text: String,
        /// The pattern as a search for a part of the text folds it.
        part: String,
        whole: bool,
    },
    Re(Regex),
    /// A regular expression of the document's search mode, read as ICU reads it: the first match
    /// is the one that counts, and for a criterion it must be the whole text.
    Icu(Regex, bool),
    /// The pattern did not compile (invalid regular expression): matches nothing.
    Never,
}

fn wildcard_to_regex(p: &str) -> String {
    let mut out = String::new();
    let mut it = p.chars().peekable();
    while let Some(c) = it.next() {
        match c {
            '*' => out.push_str(".*"),
            '?' => out.push('.'),
            '~' => match it.next() {
                Some(n) => out.push_str(&fancy_regex::escape(&n.to_string())),
                None => out.push('~'),
            },
            c => out.push_str(&fancy_regex::escape(&c.to_string())),
        }
    }
    out
}

use crate::casemap::fold;

/// Whether LibreOffice reads the text as a regular expression when the document's search mode is
/// one (`MayBeRegExp`): one character other than `.` is text, and `{` and `}` alone do not make
/// an expression.
pub fn may_be_regex(pattern: &str) -> bool {
    (pattern.chars().count() > 1 || pattern == ".")
        && pattern.contains(|c: char| "?*+.[]^$\\<>()|".contains(c))
}

impl Matcher {
    /// `whole`: the pattern must cover the entire text (criteria functions take
    /// it from the document; `SEARCH` passes `false`).
    pub fn new(pattern: &str, s: &Search, whole: bool) -> Self {
        let special = match s.mode {
            SearchMode::Wildcards => pattern.contains(['*', '?', '~']),
            SearchMode::Regex => may_be_regex(pattern),
            SearchMode::Literal => false,
        };
        if !special {
            return Self::Plain {
                text: fold(pattern).into_owned(),
                part: crate::casemap::fold_search(pattern).0,
                whole,
            };
        }
        if s.mode == SearchMode::Regex {
            if crate::casemap::has_multi_fold(pattern) {
                crate::notes::note(crate::notes::REGEX_FOLD);
                return Self::Never;
            }
            return match crate::regex_subset::translate(pattern) {
                Ok(body) => Regex::new(&format!("(?i){body}")).map_or_else(
                    |_| {
                        crate::notes::note(crate::notes::REGEX);
                        Self::Never
                    },
                    |re| Self::Icu(re, whole),
                ),
                Err(_) => {
                    crate::notes::note(crate::notes::REGEX);
                    Self::Never
                }
            };
        }
        let body = wildcard_to_regex(pattern);
        let anchored = if whole { format!("^(?:{body})$") } else { body };
        Regex::new(&format!("(?is){anchored}")).map_or(Self::Never, Self::Re)
    }

    /// The lowercased text when the pattern is plain and covers the whole cell: the
    /// match is then an equality that a hash index can answer.
    pub fn plain_whole(&self) -> Option<&str> {
        match self {
            Self::Plain {
                text, whole: true, ..
            } => Some(text),
            _ => None,
        }
    }

    pub fn is_match(&self, text: &str) -> bool {
        self.find(text).is_some()
    }

    /// Character index of the first match.
    pub fn find(&self, text: &str) -> Option<usize> {
        match self {
            Self::Never => None,
            // An empty match (`a*` against `xyz`) is not a hit, as in LibreOffice.
            Self::Re(re) => re
                .find_iter(text)
                .filter_map(Result::ok)
                .find(|m| !m.as_str().is_empty())
                .map(|m| text[..m.start()].chars().count()),
            Self::Icu(re, whole) => {
                if crate::casemap::has_multi_fold(text) {
                    crate::notes::note(crate::notes::REGEX_FOLD);
                    return None;
                }
                let m = re.find(text).ok().flatten()?;
                if *whole && (m.start() != 0 || m.end() != text.len()) {
                    return None;
                }
                Some(text[..m.start()].chars().count())
            }
            Self::Plain {
                text: p,
                part,
                whole,
            } => {
                if *whole {
                    // No copy of the text when it is plain: this runs once per cell of a
                    // searched column.
                    let same = if text.is_ascii() && p.is_ascii() {
                        text.eq_ignore_ascii_case(p)
                    } else {
                        fold(text) == p.as_str()
                    };
                    same.then_some(0)
                } else {
                    let (t, from) = crate::casemap::fold_search(text);
                    t.find(part.as_str()).map(|b| {
                        let at = t[..b].chars().count();
                        from.get(at).copied().unwrap_or(text.chars().count())
                    })
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(mode: SearchMode, whole: bool, cs: bool) -> Search {
        Search {
            mode,
            whole_cell: whole,
            case_sensitive: cs,
        }
    }

    #[test]
    fn wildcards() {
        let m = Matcher::new("a*n?", &s(SearchMode::Wildcards, true, false), true);
        assert!(m.is_match("anx"));
        assert!(Matcher::new("a~*", &s(SearchMode::Wildcards, true, false), true).is_match("a*"));
        assert!(!Matcher::new("a~*", &s(SearchMode::Wildcards, true, false), true).is_match("ab"));
        assert!(!Matcher::new("a.c", &s(SearchMode::Wildcards, true, false), true).is_match("abc"));
    }

    #[test]
    fn regex_is_read_as_icu_reads_it() {
        let re = s(SearchMode::Regex, true, false);
        assert!(Matcher::new("a.c", &re, true).is_match("ABC"));
        assert!(!Matcher::new("a.", &re, true).is_match("abc"));
        assert!(Matcher::new("a.", &re, false).is_match("xabc"));
        // `.` leaves out ICU's line terminators, among them the vertical tab.
        assert!(!Matcher::new("a.c", &re, true).is_match("a\u{b}c"));
        assert!(Matcher::new("a.c", &re, true).is_match("a\tc"));
        // The first match is the one that counts, and a criterion needs it to be the whole text.
        assert!(!Matcher::new("a|ab", &re, true).is_match("ab"));
        assert!(Matcher::new("ab|a", &re, true).is_match("ab"));
        // Braces alone do not make an expression: the text is searched as written.
        assert!(!Matcher::new("b{1,2}", &re, true).is_match("b"));
        assert!(Matcher::new("b{1,2}", &re, true).is_match("B{1,2}"));
        // What is outside the subset matches nothing (and the build refuses it).
        assert!(!Matcher::new("(?=a)a", &re, true).is_match("a"));
        assert!(!Matcher::new("(a", &re, true).is_match("(a"));
        // One character other than `.` is text.
        assert!(Matcher::new("(", &re, true).is_match("("));
    }

    #[test]
    fn literal_and_wildcard_modes() {
        let lit = s(SearchMode::Literal, true, true);
        assert!(Matcher::new("a*", &lit, true).is_match("a*"));
        assert!(!Matcher::new("a*", &lit, true).is_match("ab"));
        // Search is case-insensitive whatever the comparison setting says.
        assert!(Matcher::new("abc", &lit, true).is_match("ABC"));
        assert!(!Matcher::new("straße", &lit, true).is_match("STRASSE"));
        assert!(Matcher::new("straße", &lit, true).is_match("STRAẞE"));
        // A part of a text is searched with the full case folding: `s` is in `ß`, `i` is not in
        // `İ` and `I` is not `ı`; the position is that of the character.
        let part = s(SearchMode::Literal, false, false);
        assert_eq!(Matcher::new("s", &part, false).find("aßb"), Some(1));
        assert_eq!(Matcher::new("ss", &part, false).find("xß"), Some(1));
        assert!(Matcher::new("k", &part, false).is_match("\u{212a}"));
        assert!(!Matcher::new("i", &part, false).is_match("İ"));
        assert!(!Matcher::new("I", &part, false).is_match("ı"));
    }

    #[test]
    fn whole_cell_and_position() {
        let w = s(SearchMode::Wildcards, false, false);
        assert!(Matcher::new("app", &w, false).is_match("apple"));
        assert_eq!(Matcher::new("n?n", &w, false).find("banana"), Some(2));
        assert!(!Matcher::new("app", &w, true).is_match("apple"));
    }
}
