//! The regular expressions the engine runs: a subset of ICU's syntax.
//!
//! `LibreOffice` runs a regular expression with ICU's engine, whose results differ from other
//! engines' in details: what an empty match is, where `$` and `\b` match, what `.` and `\s` stand
//! for, how a pattern that is not valid is treated. This module reads the part of the syntax in
//! which the results were compared with `LibreOffice`'s (`tests/regex_subset.rs` lists the cases)
//! and writes it for the `regex` engine with the meanings ICU gives; anything else is refused, so
//! the build can stop instead of giving another result.
//!
//! The subset: literals and escaped punctuation, `.`, `\d \D \w \W \s \S`, sets (`[a-c]`, `[^x]`,
//! with ranges and `\d \w \s` inside), groups `( )` and `(?: )`, alternation, and the quantifiers
//! `* + ? {n} {n,} {n,m}` with their lazy forms. Nothing that repeats may match the empty text.
//!
//! The anchors `^` and `$` are read by the searches of the Find & Replace of `LibreOffice`
//! ([`translate_find`]) and refused by the functions ([`translate`],
//! [`translate_capturing`]: `REGEX`, `COUNTIF`...), whose semantics are another matter. They mean
//! what `LibreOffice` gives them in a cell: the start of the text, and its end or the place before
//! a line terminator that ends it; never the start or end of a line.
//!
//! A pattern that can match the empty text (`^`, `$`, `x*`, `a*$`) is refused by the functions and
//! read by [`translate_find`]: what a search or a replacement does with an empty match is the loop of
//! `LibreOffice` (an empty match is not a hit, except one at the very end of the text, and no search
//! starts at the end), which the caller of that translation runs (`sheetrs_query::Matcher`). The
//! empty pattern itself matches nothing in `LibreOffice` and never reaches a translation: the search
//! and the replacement that carry it are answered before (nothing found, nothing replaced).

use std::fmt::Write as _;

#[derive(Debug, Clone, PartialEq)]
enum Node {
    Empty,
    Lit(char),
    Dot,
    Digit(bool),
    Word(bool),
    Space(bool),
    Set(bool, Vec<SetItem>),
    Group(Box<Node>, bool),
    Concat(Vec<Node>),
    Alt(Vec<Node>),
    Repeat(Box<Node>, u32, Option<u32>, bool),
    /// `^`: the start of the text.
    Start,
    /// `$`: the end of the text, or the place before the line terminator that ends it.
    End,
}

#[derive(Debug, Clone, PartialEq)]
enum SetItem {
    Char(char),
    Range(char, char),
    Digit,
    Word,
    Space,
}

struct Parser {
    chars: Vec<char>,
    at: usize,
    /// Whether `^` and `$` are read (Find & Replace) or refused (a function).
    anchors: bool,
}

/// A construct of a pattern that is outside the subset, described as the text says it.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[error("{0}")]
pub struct Unsupported(String);

impl From<&str> for Unsupported {
    fn from(what: &str) -> Self {
        Self(what.to_string())
    }
}

impl From<String> for Unsupported {
    fn from(what: String) -> Self {
        Self(what)
    }
}

type R<T> = Result<T, Unsupported>;

const MAX_REPEAT: u32 = 1000;

impl Parser {
    fn peek(&self) -> Option<char> {
        self.chars.get(self.at).copied()
    }

    fn next(&mut self) -> Option<char> {
        let c = self.peek();
        if c.is_some() {
            self.at += 1;
        }
        c
    }

    fn eat(&mut self, c: char) -> bool {
        if self.peek() == Some(c) {
            self.at += 1;
            true
        } else {
            false
        }
    }

    fn alt(&mut self) -> R<Node> {
        let mut alts = vec![self.concat()?];
        while self.eat('|') {
            alts.push(self.concat()?);
        }
        Ok(if alts.len() == 1 {
            alts.remove(0)
        } else {
            Node::Alt(alts)
        })
    }

    fn concat(&mut self) -> R<Node> {
        let mut items = Vec::new();
        while let Some(c) = self.peek() {
            if c == '|' || c == ')' {
                break;
            }
            let atom = self.atom()?;
            items.push(self.quantified(atom)?);
        }
        Ok(match items.len() {
            0 => Node::Empty,
            1 => items.remove(0),
            _ => Node::Concat(items),
        })
    }

    fn number(&mut self) -> Option<u32> {
        let start = self.at;
        while self.peek().is_some_and(|c| c.is_ascii_digit()) {
            self.at += 1;
        }
        if self.at == start || self.at - start > 4 {
            return None;
        }
        self.chars[start..self.at]
            .iter()
            .collect::<String>()
            .parse()
            .ok()
    }

    fn quantified(&mut self, atom: Node) -> R<Node> {
        let (min, max) = match self.peek() {
            Some('*') => {
                self.at += 1;
                (0, None)
            }
            Some('+') => {
                self.at += 1;
                (1, None)
            }
            Some('?') => {
                self.at += 1;
                (0, Some(1))
            }
            Some('{') => {
                self.at += 1;
                let min = self.number().ok_or("a repetition `{` without a count")?;
                let max = if self.eat(',') {
                    if self.peek() == Some('}') {
                        None
                    } else {
                        Some(self.number().ok_or("a repetition with a bad count")?)
                    }
                } else {
                    Some(min)
                };
                if !self.eat('}') {
                    return Err("a repetition `{` that is not closed".into());
                }
                if max.is_some_and(|m| m < min || m > MAX_REPEAT) || min > MAX_REPEAT {
                    return Err("a repetition count that is not valid".into());
                }
                (min, max)
            }
            _ => return Ok(atom),
        };
        let lazy = self.eat('?');
        if matches!(self.peek(), Some('*' | '+' | '?' | '{')) {
            return Err("a repetition that follows a repetition (or is possessive)".into());
        }
        if min_len(&atom) == 0 {
            return Err("a repetition of something that can match the empty text".into());
        }
        Ok(Node::Repeat(Box::new(atom), min, max, lazy))
    }

    fn atom(&mut self) -> R<Node> {
        let c = self.next().ok_or("the pattern ends early")?;
        Ok(match c {
            '(' => {
                let plain = self.eat('?');
                if plain && !self.eat(':') {
                    return Err("a group with options, a name or a look-around".into());
                }
                let inner = self.alt()?;
                if !self.eat(')') {
                    return Err("a group that is not closed".into());
                }
                Node::Group(Box::new(inner), !plain)
            }
            '[' => self.set()?,
            '.' => Node::Dot,
            '\\' => match self.next().ok_or("a `\\` at the end")? {
                'd' => Node::Digit(false),
                'D' => Node::Digit(true),
                'w' => Node::Word(false),
                'W' => Node::Word(true),
                's' => Node::Space(false),
                'S' => Node::Space(true),
                p if p.is_ascii_punctuation() => Node::Lit(p),
                other => return Err(format!("the escape `\\{other}`").into()),
            },
            '^' | '$' if !self.anchors => return Err("an anchor (`^`, `$`)".into()),
            '^' => Node::Start,
            '$' => Node::End,
            '*' | '+' | '?' | '{' => return Err("a repetition with nothing to repeat".into()),
            ']' | '}' => return Err(format!("a lone `{c}`").into()),
            c => Node::Lit(c),
        })
    }

    fn set(&mut self) -> R<Node> {
        let negated = self.eat('^');
        let mut items: Vec<SetItem> = Vec::new();
        loop {
            let c = self.next().ok_or("a set that is not closed")?;
            let first = match c {
                ']' if !items.is_empty() => break,
                ']' => return Err("an empty set".into()),
                '[' => return Err("a set inside a set".into()),
                '&' if self.peek() == Some('&') => return Err("a set operation (`&&`)".into()),
                '-' if !items.is_empty()
                    && self.peek() != Some(']')
                    && !matches!(items.last(), Some(SetItem::Char(_))) =>
                {
                    return Err("a `-` after a range or a class inside a set".into());
                }
                '\\' => match self.next().ok_or("a `\\` at the end")? {
                    'd' => {
                        items.push(SetItem::Digit);
                        continue;
                    }
                    'w' => {
                        items.push(SetItem::Word);
                        continue;
                    }
                    's' => {
                        items.push(SetItem::Space);
                        continue;
                    }
                    p if p.is_ascii_punctuation() => p,
                    other => return Err(format!("the escape `\\{other}` in a set").into()),
                },
                c => c,
            };
            // A range: `a-z`, where the `-` is not the last character of the set. A `--` is ICU's
            // set difference, and a `-` after a range or a class is not a plain character.
            if (self.peek() == Some('-') && self.chars.get(self.at + 1) == Some(&'-'))
                || (first == '-' && self.peek() == Some('-'))
            {
                return Err("a set operation (`--`)".into());
            }
            if self.peek() == Some('-') && self.chars.get(self.at + 1).is_some_and(|&n| n != ']') {
                self.at += 1;
                let end = match self.next().ok_or("a set that is not closed")? {
                    '\\' => match self.next().ok_or("a `\\` at the end")? {
                        p if p.is_ascii_punctuation() => p,
                        other => return Err(format!("the escape `\\{other}` in a range").into()),
                    },
                    '[' => return Err("a set inside a set".into()),
                    e => e,
                };
                if end < first {
                    return Err("a range that is backwards".into());
                }
                items.push(SetItem::Range(first, end));
            } else {
                items.push(SetItem::Char(first));
            }
        }
        Ok(Node::Set(negated, items))
    }
}

fn min_len(n: &Node) -> u32 {
    match n {
        Node::Empty | Node::Start | Node::End => 0,
        Node::Lit(_)
        | Node::Dot
        | Node::Digit(_)
        | Node::Word(_)
        | Node::Space(_)
        | Node::Set(..) => 1,
        Node::Group(g, _) => min_len(g),
        Node::Concat(v) => v.iter().map(min_len).sum(),
        Node::Alt(v) => v.iter().map(min_len).min().unwrap_or(0),
        Node::Repeat(g, min, ..) => min_len(g).saturating_mul(*min),
    }
}

/// What ICU's `.` leaves out: its line terminators.
const LINE_ENDS: &str = "\\n\\x0B\\x0C\\r\\x{85}\\x{2028}\\x{2029}";

/// ICU's `$` without the multi-line flag: the end of the text, or the place before the one line
/// terminator that ends it (`\r\n` counts as one, so the place between the two is not it).
const END: &str =
    "(?:\\z|(?=\\r\\n\\z)|(?<!\\r)(?=\\n\\z)|(?=[\\x0B\\x0C\\r\\x{85}\\x{2028}\\x{2029}]\\z))";

fn lit(c: char, out: &mut String) {
    let mut b = [0u8; 4];
    out.push_str(&regex_syntax_escape(c.encode_utf8(&mut b)));
}

fn regex_syntax_escape(s: &str) -> String {
    fancy_regex::escape(s).into_owned()
}

fn set_char(c: char, out: &mut String) {
    if c.is_ascii_alphanumeric() {
        out.push(c);
    } else if c.is_ascii() || c.is_control() {
        let _ = write!(out, "\\x{{{:X}}}", u32::from(c));
    } else {
        out.push(c);
    }
}

fn write(n: &Node, out: &mut String, captures: bool) {
    match n {
        Node::Empty => {}
        Node::Start => out.push_str("\\A"),
        Node::End => out.push_str(END),
        Node::Lit(c) => lit(*c, out),
        Node::Dot => {
            out.push_str("[^");
            out.push_str(LINE_ENDS);
            out.push(']');
        }
        Node::Digit(neg) => out.push_str(if *neg { "\\D" } else { "\\d" }),
        Node::Word(neg) => out.push_str(if *neg { "\\W" } else { "\\w" }),
        Node::Space(neg) => out.push_str(if *neg { "\\S" } else { "\\s" }),
        Node::Set(neg, items) => {
            out.push_str(if *neg { "[^" } else { "[" });
            for i in items {
                match i {
                    SetItem::Char(c) => set_char(*c, out),
                    SetItem::Range(a, b) => {
                        set_char(*a, out);
                        out.push('-');
                        set_char(*b, out);
                    }
                    SetItem::Digit => out.push_str("\\d"),
                    SetItem::Word => out.push_str("\\w"),
                    SetItem::Space => out.push_str("\\s"),
                }
            }
            out.push(']');
        }
        Node::Group(g, capturing) => {
            out.push_str(if *capturing && captures { "(" } else { "(?:" });
            write(g, out, captures);
            out.push(')');
        }
        Node::Concat(v) => v.iter().for_each(|x| write(x, out, captures)),
        Node::Alt(v) => {
            for (i, x) in v.iter().enumerate() {
                if i > 0 {
                    out.push('|');
                }
                out.push_str("(?:");
                write(x, out, captures);
                out.push(')');
            }
        }
        Node::Repeat(g, min, max, lazy) => {
            out.push_str("(?:");
            write(g, out, captures);
            out.push(')');
            match (min, max) {
                (0, None) => out.push('*'),
                (1, None) => out.push('+'),
                (0, Some(1)) => out.push('?'),
                (n, None) => {
                    let _ = write!(out, "{{{n},}}");
                }
                (n, Some(m)) if n == m => {
                    let _ = write!(out, "{{{n}}}");
                }
                (n, Some(m)) => {
                    let _ = write!(out, "{{{n},{m}}}");
                }
            }
            if *lazy {
                out.push('?');
            }
        }
    }
}

/// The pattern as the `regex` engine reads it (case is ignored by the caller's flags), or why the
/// pattern is outside the subset.
///
/// # Errors
///
/// When `pattern` uses a construct outside the subset.
pub fn translate(pattern: &str) -> R<String> {
    translate_with(pattern, false, false, false)
}

/// Like [`translate`], keeping the groups that capture (for the replacements of `REGEX`).
///
/// # Errors
///
/// When `pattern` uses a construct outside the subset.
pub fn translate_capturing(pattern: &str) -> R<String> {
    translate_with(pattern, true, false, false)
}

/// Like [`translate_capturing`], and reading the anchors `^` and `$`, for the searches of Find &
/// Replace (`find` and `replace`). The pattern may match the empty text (`x*`, `^`, `$`, `a*$`),
/// which the caller handles as `LibreOffice` does (an empty match is a hit only at the end of the
/// text, and a search never starts there). The result needs an engine with look-around
/// (`fancy_regex`) when it has a `$`. The empty pattern is refused: the caller answers it before.
///
/// # Errors
///
/// When `pattern` uses a construct outside the subset.
pub fn translate_find(pattern: &str) -> R<String> {
    if pattern.is_empty() {
        return Err("an empty pattern".into());
    }
    translate_with(pattern, true, true, true)
}

fn translate_with(pattern: &str, captures: bool, anchors: bool, empty: bool) -> R<String> {
    let mut p = Parser {
        chars: pattern.chars().collect(),
        at: 0,
        anchors,
    };
    let node = p.alt()?;
    if p.at < p.chars.len() {
        return Err("a `)` that closes nothing".into());
    }
    if !empty && min_len(&node) == 0 {
        return Err("a pattern that can match the empty text".into());
    }
    let mut out = String::new();
    write(&node, &mut out, captures);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_the_subset() {
        for ok in [
            "abc",
            "a.c",
            "[a-c]+",
            "(ab|cd)+",
            "a{2,3}?",
            "\\d+\\.\\d",
            "[^\\s]x",
            "(?:a|b)c*?",
        ] {
            assert!(translate(ok).is_ok(), "{ok}");
        }
    }

    #[test]
    fn refuses_the_rest() {
        for bad in [
            "a*",
            "^a",
            "a$",
            "\\bfoo",
            "(?=a)b",
            "(?i)a",
            "(a)\\1",
            "\\p{L}",
            "[[:alpha:]]",
            "[a-z&&[^b]]",
            "a++",
            "(a*)*b",
            "a{,3}",
            "a{2",
            "(a",
            "a)",
            "\\h",
            "[z-a]",
            "?a",
            "a|",
            "",
            "[--a]",
            "[a-c-e]",
            "[\\d-x]",
        ] {
            assert!(translate(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn reads_the_anchors_of_a_search_only() {
        for ok in ["^a", "a$", "^(a|b)$", "(^a|b$)", "a^b", "(^|x)a", "^a.$"] {
            assert!(translate_find(ok).is_ok(), "{ok}");
            assert!(translate(ok).is_err(), "{ok}");
            assert!(translate_capturing(ok).is_err(), "{ok}");
        }
    }

    #[test]
    fn a_search_reads_the_anchors_and_the_patterns_that_match_the_empty_text() {
        for ok in [
            "^a", "a$", "^", "$", "^$", "a*$", "^a*", "x*", "(^|x)a", "(a|)", "a|", "(|a)$", "a?",
            "(^|$)",
        ] {
            assert!(translate_find(ok).is_ok(), "{ok}");
        }
    }

    #[test]
    fn a_search_still_refuses_what_is_not_reproduced() {
        for bad in [
            "", "^*", "(^)+a", "^+", "(a*)*", "(^|$)*", "\\b", "(?=a)", "a**",
        ] {
            assert!(translate_find(bad).is_err(), "{bad}");
        }
    }
}
