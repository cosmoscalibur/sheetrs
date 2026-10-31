//! `REGEX` of LibreOffice: the first match of an expression in a text, or the text with the
//! first (or the nth, or every) match replaced. The expressions are those of
//! [`crate::regex_subset`]; any other stops the build.

use super::{Reg, opt};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};
use fancy_regex::Regex;

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// A piece of a replacement: text, or the group of the match it takes.
enum Piece {
    Text(String),
    Group(usize),
}

/// The replacement as LibreOffice reads it: `$0` to `$9` are groups, a backslash takes the next
/// character as it is (a lone one at the end is dropped) and a `$` that is not followed by a digit
/// is an error.
fn pieces(replacement: &str) -> Res<Vec<Piece>> {
    let mut out = Vec::new();
    let mut text = String::new();
    let mut it = replacement.chars();
    while let Some(c) = it.next() {
        match c {
            '\\' => {
                if let Some(n) = it.next() {
                    text.push(n);
                }
            }
            '$' => match it.next() {
                Some(d) if d.is_ascii_digit() => {
                    if !text.is_empty() {
                        out.push(Piece::Text(std::mem::take(&mut text)));
                    }
                    out.push(Piece::Group(d as usize - '0' as usize));
                }
                _ => return Err(DOMAIN),
            },
            c => text.push(c),
        }
    }
    if !text.is_empty() {
        out.push(Piece::Text(text));
    }
    Ok(out)
}

fn regex(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let text = e.text(&a[0])?;
    let pattern = e.text(&a[1])?;
    let Ok(body) = crate::regex_subset::translate_capturing(&pattern) else {
        crate::notes::note(crate::notes::REGEX);
        return Err(ErrorKind::Unsupported);
    };
    let Ok(re) = Regex::new(&body) else {
        crate::notes::note(crate::notes::REGEX);
        return Err(ErrorKind::Unsupported);
    };
    if crate::casemap::has_multi_fold(&text) || crate::casemap::has_multi_fold(&pattern) {
        crate::notes::note(crate::notes::REGEX_FOLD);
        return Err(ErrorKind::Unsupported);
    }
    let replacement = match opt(a, 2) {
        Some(v) => Some(e.text(v)?),
        None => None,
    };
    // The fourth argument: "g" for every match, or the number of the match.
    let (mut global, mut nth) = (false, 1usize);
    if let Some(v) = opt(a, 3) {
        match e.scalar(v.clone()) {
            Value::Error(k) => return Err(k),
            Value::Text(t) if &*t == "g" => global = true,
            Value::Text(t) if t.is_empty() => {}
            Value::Text(_) => return Err(DOMAIN),
            v => {
                let n = e.num(&v)?.trunc();
                if n < 0.0 {
                    return Err(DOMAIN);
                }
                if n == 0.0 {
                    return Ok(Value::text(text));
                }
                nth = n as usize;
            }
        }
    }
    let matches: Vec<_> = re
        .captures_iter(&text)
        .filter_map(Result::ok)
        .filter(|c| c.get(0).is_some_and(|m| !m.as_str().is_empty()))
        .collect();
    let Some(replacement) = replacement else {
        return match matches.get(if global { 0 } else { nth - 1 }) {
            Some(c) => Ok(Value::text(c.get(0).map_or("", |m| m.as_str()))),
            None => Err(ErrorKind::Na),
        };
    };
    if matches.is_empty() {
        return Ok(Value::text(text));
    }
    let parts = pieces(&replacement)?;
    let groups = matches[0].len();
    if parts
        .iter()
        .any(|p| matches!(p, Piece::Group(g) if *g >= groups))
    {
        return Err(DOMAIN);
    }
    let mut out = String::new();
    let mut last = 0;
    for (i, c) in matches.iter().enumerate() {
        if !global && i + 1 != nth {
            continue;
        }
        let whole = c.get(0).map_or(0..0, |m| m.range());
        out.push_str(&text[last..whole.start]);
        for p in &parts {
            match p {
                Piece::Text(t) => out.push_str(t),
                Piece::Group(g) => out.push_str(c.get(*g).map_or("", |m| m.as_str())),
            }
        }
        last = whole.end;
    }
    out.push_str(&text[last..]);
    Ok(Value::text(out))
}

pub(super) fn register(r: &mut Reg) {
    r.eager("REGEX", 2, 4, regex);
}
