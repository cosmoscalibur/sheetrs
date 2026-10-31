//! The byte and width functions of text: `LENB`, `LEFTB`, `RIGHTB`, `MIDB`, `FINDB`, `SEARCHB`,
//! `REPLACEB`, `ASC` and `JIS`. LibreOffice counts a character as one or two bytes by a table of
//! its own (`width_data.rs`, generated from it), whatever the language of the document, and a
//! double-byte character that a cut splits becomes a space.

use super::width_data::{ASC, JIS, JIS_PAIRS, SINGLE_BYTE};
use super::{Reg, num, opt_num};
use crate::engine::{Engine, Res};
use crate::matching::Matcher;
use crate::value::{ErrorKind, Value};

/// The bytes of a character: one in the single-byte ranges of the plane, two elsewhere.
fn width(c: char) -> usize {
    let n = c as u32;
    if SINGLE_BYTE.iter().any(|&(a, b)| (a..=b).contains(&n)) {
        1
    } else {
        2
    }
}

/// The text as cells of one byte: `(character, which byte of it, its width)`.
fn cells(t: &str) -> Vec<(char, usize, usize)> {
    let mut out = Vec::new();
    for c in t.chars() {
        let w = width(c);
        for k in 0..w {
            out.push((c, k, w));
        }
    }
    out
}

/// The cells as text: a character with all its bytes is itself, one with a part of them a space
/// for each byte that is there. Half of a character above the plane is half of a surrogate pair,
/// which no text of the engine (nor an XML file) can hold: that is not reproduced.
fn render(part: &[(char, usize, usize)]) -> Res<String> {
    let mut out = String::new();
    let mut i = 0;
    while i < part.len() {
        let (c, k, w) = part[i];
        if w == 1 {
            out.push(c);
            i += 1;
        } else if k == 0 && i + 1 < part.len() && part[i + 1].0 == c && part[i + 1].1 == 1 {
            out.push(c);
            i += 2;
        } else {
            if (c as u32) >= 0x10000 {
                crate::notes::note(crate::notes::SURROGATE_CUT);
                return Err(ErrorKind::Unsupported);
            }
            out.push(' ');
            i += 1;
        }
    }
    Ok(out)
}

fn count(e: &mut Engine, a: &[Value], i: usize) -> Res<usize> {
    let n = opt_num(e, a, i, 1.0)?.trunc();
    if n < 0.0 {
        return Err(ErrorKind::Domain);
    }
    Ok(n as usize)
}

fn find_bytes(e: &mut Engine, a: &[Value], sensitive: bool) -> Res<Value> {
    let needle = e.text(&a[0])?;
    let hay = e.text(&a[1])?;
    let start = opt_num(e, a, 2, 1.0)?.trunc();
    if start < 1.0 {
        return Err(ErrorKind::Domain);
    }
    if needle.is_empty() {
        return Err(ErrorKind::Value);
    }
    let chars: Vec<char> = hay.chars().collect();
    // The byte at which each character starts (1-based), and the total.
    let mut starts = Vec::with_capacity(chars.len());
    let mut at = 1usize;
    for &c in &chars {
        starts.push(at);
        at += width(c);
    }
    // Searching starts at the first character that starts at or after the byte given.
    let first = starts.partition_point(|&p| (p as f64) < start);
    if first >= chars.len() {
        return Err(ErrorKind::Value);
    }
    let found = if sensitive {
        let nd: Vec<char> = needle.chars().collect();
        (first..chars.len()).find(|&i| chars[i..].starts_with(&nd))
    } else {
        let search = crate::options::Search {
            case_sensitive: false,
            ..e.opts.search
        };
        let m = Matcher::new(&needle, &search, false);
        let rest: String = chars[first..].iter().collect();
        m.find(&rest).map(|i| first + i)
    };
    found.map_or(Err(ErrorKind::Value), |i| num(starts[i] as f64))
}

/// The conversion of `ASC` or `JIS` over a text.
fn convert(text: &str, table: &[(u32, u32, u32)], join_pairs: bool) -> String {
    let chars: Vec<char> = text.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        let n = chars[i] as u32;
        if join_pairs
            && let Some(&next) = chars.get(i + 1)
            && let Some(&(_, _, joined)) = JIS_PAIRS
                .iter()
                .find(|&&(b, m, _)| b == n && m == next as u32)
        {
            out.extend(char::from_u32(joined));
            i += 2;
            continue;
        }
        match table.binary_search_by_key(&n, |&(from, _, _)| from) {
            Ok(k) => {
                let (_, first, second) = table[k];
                out.extend(char::from_u32(first));
                if second != 0 {
                    out.extend(char::from_u32(second));
                }
            }
            Err(_) => out.push(chars[i]),
        }
        i += 1;
    }
    out
}

pub(super) fn register(r: &mut Reg) {
    r.eager("LENB", 1, 1, |e, a| {
        num(e.text(&a[0])?.chars().map(width).sum::<usize>() as f64)
    });
    r.eager("LEFTB", 1, 2, |e, a| {
        let t = cells(&e.text(&a[0])?);
        let n = count(e, a, 1)?;
        Ok(Value::text(render(&t[..n.min(t.len())])?))
    });
    r.eager("RIGHTB", 1, 2, |e, a| {
        let t = cells(&e.text(&a[0])?);
        let n = count(e, a, 1)?;
        Ok(Value::text(render(&t[t.len().saturating_sub(n)..])?))
    });
    r.eager("MIDB", 3, 3, |e, a| {
        let t = cells(&e.text(&a[0])?);
        let (s, n) = (e.num(&a[1])?.trunc(), e.num(&a[2])?.trunc());
        if s < 1.0 || n < 0.0 {
            return Err(ErrorKind::Domain);
        }
        let from = (s as usize - 1).min(t.len());
        let to = (from + n as usize).min(t.len());
        Ok(Value::text(render(&t[from..to])?))
    });
    r.eager("FINDB", 2, 3, |e, a| find_bytes(e, a, true));
    r.eager("SEARCHB", 2, 3, |e, a| find_bytes(e, a, false));
    r.eager("REPLACEB", 4, 4, |e, a| {
        let t = cells(&e.text(&a[0])?);
        let (s, n) = (e.num(&a[1])?.trunc(), e.num(&a[2])?.trunc());
        let new = e.text(&a[3])?;
        if s < 1.0 || n < 0.0 {
            return Err(ErrorKind::Domain);
        }
        let (from, len) = (s as usize - 1, n as usize);
        if from >= t.len() || from + len > t.len() {
            return Err(ErrorKind::Domain);
        }
        let mut out = render(&t[..from])?;
        out.push_str(&new);
        out.push_str(&render(&t[from + len..])?);
        Ok(Value::text(out))
    });
    r.eager("ASC", 1, 1, |e, a| {
        Ok(Value::text(convert(&e.text(&a[0])?, ASC, false)))
    });
    r.eager("JIS", 1, 1, |e, a| {
        Ok(Value::text(convert(&e.text(&a[0])?, JIS, true)))
    });
}
