//! Text functions.

use super::{Reg, VAR, num, opt, opt_num, round_decimal};
use crate::coerce;
use crate::engine::{Engine, Res};
use crate::format;
use crate::matching::Matcher;
use crate::value::{ErrorKind, Value};

pub(super) fn register(r: &mut Reg) {
    r.eager("CONCATENATE", 1, VAR, |e, a| concat(e, a, false));
    r.eager("CONCAT", 1, VAR, |e, a| concat(e, a, true));
    r.eager("TEXTJOIN", 3, VAR, |e, a| {
        let sep = e.text(&a[0])?;
        let skip = e.boolean(&a[1])?;
        let mut parts = Vec::new();
        for v in &a[2..] {
            for x in &e.array_of(v.clone()).data {
                let t = coerce::to_text(x, &e.opts.locale)?;
                if !(skip && t.is_empty()) {
                    parts.push(t);
                }
            }
        }
        Ok(Value::text(parts.join(&sep)))
    });
    r.eager("LEN", 1, 1, |e, a| {
        num(e.text(&a[0])?.chars().count() as f64)
    });
    r.eager("UPPER", 1, 1, |e, a| {
        Ok(Value::text(crate::casemap::upper(&e.text(&a[0])?)))
    });
    r.eager("LOWER", 1, 1, |e, a| {
        Ok(Value::text(crate::casemap::lower(&e.text(&a[0])?)))
    });
    r.eager("PROPER", 1, 1, |e, a| {
        Ok(Value::text(crate::casemap::proper(&e.text(&a[0])?)))
    });
    r.eager("TRIM", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        Ok(Value::text(
            t.split(' ')
                .filter(|w| !w.is_empty())
                .collect::<Vec<_>>()
                .join(" "),
        ))
    });
    r.eager("LEFT", 1, 2, |e, a| {
        let t = e.text(&a[0])?;
        let n = count(e, a, 1)?;
        Ok(Value::text(t.chars().take(n).collect::<String>()))
    });
    r.eager("RIGHT", 1, 2, |e, a| {
        let t = e.text(&a[0])?;
        let n = count(e, a, 1)?;
        let len = t.chars().count();
        Ok(Value::text(
            t.chars().skip(len.saturating_sub(n)).collect::<String>(),
        ))
    });
    r.eager("MID", 3, 3, |e, a| {
        let t = e.text(&a[0])?;
        let (s, n) = (e.num(&a[1])?.trunc(), e.num(&a[2])?.trunc());
        if s < 1.0 || n < 0.0 {
            return Err(ErrorKind::Value);
        }
        Ok(Value::text(
            t.chars()
                .skip(s as usize - 1)
                .take(n as usize)
                .collect::<String>(),
        ))
    });
    r.eager("REPT", 2, 2, |e, a| {
        let (t, n) = (e.text(&a[0])?, e.num(&a[1])?.trunc());
        if !(0.0..=32767.0).contains(&n) {
            return Err(ErrorKind::Value);
        }
        Ok(Value::text(t.repeat(n as usize)))
    });
    r.eager("EXACT", 2, 2, |e, a| {
        Ok(Value::Bool(e.text(&a[0])? == e.text(&a[1])?))
    });
    r.eager("FIND", 2, 3, |e, a| find(e, a, true));
    r.eager("SEARCH", 2, 3, |e, a| find(e, a, false));
    r.eager("SUBSTITUTE", 3, 4, |e, a| {
        let (t, old, new) = (e.text(&a[0])?, e.text(&a[1])?, e.text(&a[2])?);
        if old.is_empty() {
            return Ok(Value::text(t));
        }
        match opt(a, 3) {
            None => Ok(Value::text(t.replace(&old, &new))),
            Some(n) => {
                let n = e.num(n)?.trunc();
                if n < 1.0 {
                    return Err(ErrorKind::Value);
                }
                let mut out = String::new();
                let mut rest = t.as_str();
                let mut seen = 0.0;
                while let Some(p) = rest.find(&old) {
                    seen += 1.0;
                    out.push_str(&rest[..p]);
                    out.push_str(if seen == n { &new } else { &old });
                    rest = &rest[p + old.len()..];
                }
                out.push_str(rest);
                Ok(Value::text(out))
            }
        }
    });
    r.eager("REPLACE", 4, 4, |e, a| {
        let t: Vec<char> = e.text(&a[0])?.chars().collect();
        let (s, n) = (e.num(&a[1])?.trunc(), e.num(&a[2])?.trunc());
        let new = e.text(&a[3])?;
        if s < 1.0 || n < 0.0 {
            return Err(ErrorKind::Value);
        }
        let s = (s as usize - 1).min(t.len());
        let end = (s + n as usize).min(t.len());
        Ok(Value::text(
            t[..s].iter().collect::<String>() + &new + &t[end..].iter().collect::<String>(),
        ))
    });
    r.eager("VALUE", 1, 1, |e, a| match e.scalar(a[0].clone()) {
        Value::Text(t) => coerce::parse_number(&t, &e.opts.locale, e.opts.date_ctx())
            .map_or(Err(ErrorKind::Value), |n| Ok(Value::Number(n))),
        v => num(coerce::to_num(&v, &e.opts.locale, e.opts.date_ctx())?),
    });
    r.eager("NUMBERVALUE", 1, 3, |e, a| {
        let t = e.text(&a[0])?;
        let dec = opt(a, 1)
            .map(|v| e.text(v))
            .transpose()?
            .and_then(|s| s.chars().next());
        let grp = opt(a, 2)
            .map(|v| e.text(v))
            .transpose()?
            .and_then(|s| s.chars().next());
        let mut loc = e.opts.locale.clone();
        loc.decimal = dec.unwrap_or(loc.decimal);
        loc.group = grp.unwrap_or(loc.group);
        coerce::parse_number(&t, &loc, e.opts.date_ctx())
            .map_or(Err(ErrorKind::Value), |n| Ok(Value::Number(n)))
    });
    r.eager("T", 1, 1, |e, a| match e.scalar(a[0].clone()) {
        Value::Text(t) => Ok(Value::Text(t)),
        Value::Error(k) => Err(k),
        _ => Ok(Value::text("")),
    });
    r.eager("TEXT", 2, 2, |e, a| {
        let v = e.scalar(a[0].clone());
        let fmt = e.text(&a[1])?;
        if fmt.is_empty() {
            return Ok(Value::text(""));
        }
        match v {
            Value::Text(t) => Ok(Value::text(
                format::format_text(&t, &fmt, &e.opts.locale).unwrap_or_else(|| t.to_string()),
            )),
            Value::Error(k) => Err(k),
            v => {
                let x = coerce::to_num(&v, &e.opts.locale, e.opts.date_ctx())?;
                format::format_number(x, &fmt, &e.opts.locale, e.opts.date_system)
                    .map(Value::text)
                    .ok_or(ErrorKind::Value)
            }
        }
    });
    r.eager("FIXED", 1, 3, |e, a| {
        let x = e.num(&a[0])?;
        let d = opt_num(e, a, 1, 2.0)?.trunc().max(0.0) as usize;
        let no_commas = match opt(a, 2) {
            Some(v) => e.boolean(v)?,
            None => false,
        };
        let mut fmt = String::from(if no_commas { "0" } else { "#,##0" });
        if d > 0 {
            fmt.push('.');
            fmt.push_str(&"0".repeat(d));
        }
        Ok(Value::text(format::format_number_en(
            x,
            &fmt,
            &e.opts.locale,
            e.opts.date_system,
        )))
    });
    r.eager("DOLLAR", 1, 2, |e, a| {
        let x = e.num(&a[0])?;
        let d = opt_num(e, a, 1, 2.0)?.trunc();
        // Rounded to the digits asked (tens for a negative count), then shown
        // with the locale's currency format.
        let x = round_decimal(x, d as i32);
        let fmt = if d > 0.0 {
            format!("#,##0.{}", "0".repeat(d as usize))
        } else {
            "#,##0".to_string()
        };
        let body = format::format_number_en(x.abs(), &fmt, &e.opts.locale, e.opts.date_system);
        let loc = &e.opts.locale;
        let negative = x < 0.0 && body.chars().any(|c| c.is_ascii_digit() && c != '0');
        let template = if negative {
            &loc.currency_neg
        } else {
            &loc.currency_pos
        };
        Ok(Value::text(
            template.replace('¤', &loc.currency).replace('#', &body),
        ))
    });
    // LibreOffice 26.8 on Linux (a UTF-8 system): the characters 128–255 do not exist in its text
    // encoding and are U+FFFD, and `CODE` is the first byte of the UTF-8 of the first character
    // (0 for a character outside the Basic Multilingual Plane and for no text).
    r.eager("CHAR", 1, 1, |e, a| {
        let n = e.num(&a[0])?.trunc();
        if !(0.0..=255.0).contains(&n) {
            return Err(ErrorKind::Value);
        }
        let n = n as u32;
        Ok(Value::text(
            if n < 128 {
                char::from_u32(n).unwrap_or('\u{fffd}')
            } else {
                '\u{fffd}'
            }
            .to_string(),
        ))
    });
    r.eager("CODE", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        let code = t.chars().next().map_or(0, |c| {
            if u32::from(c) > 0xffff {
                0
            } else {
                u32::from(c.to_string().as_bytes()[0])
            }
        });
        num(f64::from(code))
    });
}

fn count(e: &mut Engine, a: &[Value], i: usize) -> Res<usize> {
    let n = opt_num(e, a, i, 1.0)?.trunc();
    if n < 0.0 {
        Err(ErrorKind::Value)
    } else {
        Ok(n as usize)
    }
}

fn concat(e: &mut Engine, a: &[Value], flatten: bool) -> Res<Value> {
    let mut out = String::new();
    for v in a {
        if flatten && matches!(v, Value::Ref(_) | Value::Array(_)) {
            for x in &e.array_of(v.clone()).data {
                out.push_str(&coerce::to_text(x, &e.opts.locale)?);
            }
        } else {
            out.push_str(&e.text(v)?);
        }
    }
    Ok(Value::text(out))
}

fn find(e: &mut Engine, a: &[Value], case_sensitive: bool) -> Res<Value> {
    let needle = e.text(&a[0])?;
    // LibreOffice: nothing is found when searching for the empty text.
    if needle.is_empty() {
        return Err(ErrorKind::Value);
    }
    let hay: Vec<char> = e.text(&a[1])?.chars().collect();
    let start = opt_num(e, a, 2, 1.0)?.trunc();
    if start < 1.0 || start as usize > hay.len() + 1 {
        return Err(ErrorKind::Value);
    }
    let from = start as usize - 1;
    if case_sensitive {
        // FIND: literal and case-sensitive.
        let nd: Vec<char> = needle.chars().collect();
        return (from..=hay.len())
            .find(|&i| hay[i..].starts_with(&nd))
            .map_or(Err(ErrorKind::Value), |i| num((i + 1) as f64));
    }
    // SEARCH: case-insensitive; patterns follow the document's search mode.
    let search = crate::options::Search {
        case_sensitive: false,
        ..e.opts.search
    };
    let m = Matcher::new(&needle, &search, false);
    let rest: String = hay[from..].iter().collect();
    m.find(&rest)
        .map_or(Err(ErrorKind::Value), |i| num((from + i + 1) as f64))
}
