//! More text: cleaning, code points, Roman numerals and number bases.

use super::{Reg, num, opt_num};
use crate::engine::Res;
use crate::value::{ErrorKind, Value};

const ROMAN_CHARS: [char; 7] = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
const ROMAN_VALUES: [i64; 7] = [1000, 500, 100, 50, 10, 5, 1];

/// Roman numeral of `n` (1–3999) in form 0 (classic) to 4 (simplified).
fn roman(n: i64, mode: i64) -> String {
    let (mut out, mut val) = (String::new(), n);
    for i in 0..=3usize {
        let mut idx = 2 * i;
        let digit = val / ROMAN_VALUES[idx];
        if digit % 5 == 4 {
            let idx2 = if digit == 4 { idx - 1 } else { idx - 2 };
            let mut steps = 0;
            while steps < mode && idx < 6 {
                steps += 1;
                if ROMAN_VALUES[idx2] - ROMAN_VALUES[idx + 1] <= val {
                    idx += 1;
                } else {
                    steps = mode;
                }
            }
            out.push(ROMAN_CHARS[idx]);
            out.push(ROMAN_CHARS[idx2]);
            val += ROMAN_VALUES[idx];
            val -= ROMAN_VALUES[idx2];
        } else {
            if digit > 4 {
                out.push(ROMAN_CHARS[idx - 1]);
            }
            for _ in 0..digit % 5 {
                out.push(ROMAN_CHARS[idx]);
            }
            val %= ROMAN_VALUES[idx];
        }
    }
    out
}

fn arabic(s: &str) -> Option<i64> {
    let t = s.trim().to_ascii_uppercase();
    let (neg, body) = match t.strip_prefix('-') {
        Some(b) => (true, b),
        None => (false, t.as_str()),
    };
    let val = |c: char| {
        ROMAN_CHARS
            .iter()
            .position(|x| *x == c)
            .map(|i| ROMAN_VALUES[i])
    };
    let digits: Option<Vec<i64>> = body.chars().map(val).collect();
    let digits = digits?;
    let mut total = 0;
    for (i, d) in digits.iter().enumerate() {
        if digits.get(i + 1).is_some_and(|n| n > d) {
            total -= d;
        } else {
            total += d;
        }
    }
    Some(if neg { -total } else { total })
}

fn digits_of(n: u64, radix: u64) -> String {
    if n == 0 {
        return "0".into();
    }
    let (mut n, mut out) = (n, Vec::new());
    while n > 0 {
        out.push(
            char::from_digit((n % radix) as u32, radix as u32)
                .unwrap_or('0')
                .to_ascii_uppercase(),
        );
        n /= radix;
    }
    out.iter().rev().collect()
}

fn to_value(r: Res<String>) -> Res<Value> {
    r.map(Value::text)
}

pub(super) fn register(r: &mut Reg) {
    r.eager("CLEAN", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        Ok(Value::text(
            t.chars().filter(|c| *c >= ' ').collect::<String>(),
        ))
    });
    r.eager("UNICODE", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        t.chars()
            .next()
            .map_or(Err(ErrorKind::Value), |c| num(f64::from(c as u32)))
    });
    r.eager("UNICHAR", 1, 1, |e, a| {
        let n = e.num(&a[0])?.trunc();
        if n < 1.0 {
            return Err(ErrorKind::Value);
        }
        char::from_u32(n as u32).map_or(Err(ErrorKind::Value), |c| Ok(Value::text(c.to_string())))
    });
    r.eager("ROMAN", 1, 2, |e, a| {
        let n = e.num(&a[0])?.trunc() as i64;
        // The form is a number: TRUE is 1, FALSE is 0.
        let mode = match a.get(1) {
            None | Some(Value::Empty) => 0,
            Some(v) => e.num(v)?.trunc() as i64,
        };
        if !(0..=3999).contains(&n) || !(0..=4).contains(&mode) {
            return Err(ErrorKind::Value);
        }
        Ok(Value::text(roman(n, mode)))
    });
    r.eager("ARABIC", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        if t.trim().is_empty() {
            return num(0.0);
        }
        // LibreOffice has no negative numerals.
        if t.trim().starts_with('-') {
            return Err(ErrorKind::Value);
        }
        arabic(&t).map_or(Err(ErrorKind::Value), |n| num(n as f64))
    });
    r.eager("BASE", 2, 3, |e, a| {
        let (n, radix) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
        let min = opt_num(e, a, 2, 0.0)?.trunc();
        if n < 0.0
            || n >= 9_007_199_254_740_992.0
            || !(2.0..=36.0).contains(&radix)
            || !(0.0..=255.0).contains(&min)
        {
            return Err(ErrorKind::Value);
        }
        let mut s = digits_of(n as u64, radix as u64);
        while (s.len() as f64) < min {
            s.insert(0, '0');
        }
        Ok(Value::text(s))
    });
    r.eager("DECIMAL", 2, 2, |e, a| {
        let t = e.text(&a[0])?;
        let radix = e.num(&a[1])?.trunc();
        if !(2.0..=36.0).contains(&radix) {
            return Err(ErrorKind::Value);
        }
        let mut total = 0.0_f64;
        for c in t.trim().chars() {
            let d = c.to_digit(radix as u32).ok_or(ErrorKind::Value)?;
            total = total * radix + f64::from(d);
        }
        num(total)
    });

    // ----- number bases, as the Excel engineering functions -----
    for (name, from, to) in [
        ("DEC2BIN", 10u64, 2u64),
        ("DEC2OCT", 10, 8),
        ("DEC2HEX", 10, 16),
        ("BIN2DEC", 2, 10),
        ("BIN2OCT", 2, 8),
        ("BIN2HEX", 2, 16),
        ("OCT2DEC", 8, 10),
        ("OCT2BIN", 8, 2),
        ("OCT2HEX", 8, 16),
        ("HEX2DEC", 16, 10),
        ("HEX2BIN", 16, 2),
        ("HEX2OCT", 16, 8),
    ] {
        let f: super::EagerFn = match (from, to) {
            (10, 2) => |e, a| convert(e, a, 10, 2),
            (10, 8) => |e, a| convert(e, a, 10, 8),
            (10, 16) => |e, a| convert(e, a, 10, 16),
            (2, 10) => |e, a| convert(e, a, 2, 10),
            (2, 8) => |e, a| convert(e, a, 2, 8),
            (2, 16) => |e, a| convert(e, a, 2, 16),
            (8, 10) => |e, a| convert(e, a, 8, 10),
            (8, 2) => |e, a| convert(e, a, 8, 2),
            (8, 16) => |e, a| convert(e, a, 8, 16),
            (16, 10) => |e, a| convert(e, a, 16, 10),
            (16, 2) => |e, a| convert(e, a, 16, 2),
            _ => |e, a| convert(e, a, 16, 8),
        };
        r.eager(name, 1, if to == 10 { 1 } else { 2 }, f);
    }
    r.eager("DELTA", 1, 2, |e, a| {
        let (x, y) = (e.num(&a[0])?, opt_num(e, a, 1, 0.0)?);
        num(f64::from(u8::from(x == y)))
    });
    r.eager("GESTEP", 1, 2, |e, a| {
        let (x, y) = (e.num(&a[0])?, opt_num(e, a, 1, 0.0)?);
        num(f64::from(u8::from(x >= y)))
    });
}

/// Bits of a base's 10-digit two's complement.
fn bits(radix: u64) -> u32 {
    match radix {
        2 => 10,
        8 => 30,
        _ => 40,
    }
}

fn convert(e: &mut crate::engine::Engine, a: &[Value], from: u64, to: u64) -> Res<Value> {
    // Decode to a signed integer.
    let n: i64 = if from == 10 {
        let x = e.num(&a[0])?.trunc();
        if x.abs() > 2f64.powi(40) {
            return Err(ErrorKind::Domain);
        }
        x as i64
    } else {
        let t = match e.scalar(a[0].clone()) {
            Value::Number(n) if n >= 0.0 => format!("{}", n as u64),
            v => e.text(&v)?,
        };
        let t = t.trim();
        if t.is_empty() || t.len() > 10 {
            return Err(ErrorKind::Domain);
        }
        let mut v: i64 = 0;
        for c in t.chars() {
            let d = c.to_digit(from as u32).ok_or(ErrorKind::Domain)?;
            v = v * i64::try_from(from).unwrap_or(10) + i64::from(d);
        }
        let total = bits(from);
        if t.len() == 10 && v >= 1i64 << (total - 1) {
            v -= 1i64 << total;
        }
        v
    };
    if to == 10 {
        return num(n as f64);
    }
    let total = bits(to);
    let limit = 1i64 << (total - 1);
    if n < -limit || n >= limit {
        return Err(ErrorKind::Domain);
    }
    let places = match a.get(1) {
        Some(v) if !matches!(v, Value::Empty) => Some(e.num(v)?.trunc()),
        _ => None,
    };
    if n < 0 {
        return to_value(Ok(digits_of((n + (1i64 << total)) as u64, to)));
    }
    let mut s = digits_of(n as u64, to);
    if let Some(p) = places {
        if p < 1.0 || p > 10.0 || (s.len() as f64) > p {
            return Err(ErrorKind::Domain);
        }
        while (s.len() as f64) < p {
            s.insert(0, '0');
        }
    }
    Ok(Value::text(s))
}
