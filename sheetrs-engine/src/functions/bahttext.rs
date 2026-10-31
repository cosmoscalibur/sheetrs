//! `BAHTTEXT`: an amount in Thai words, with baht and satang.

use super::{Reg, round_decimal};
use crate::engine::{Engine, Res};
use crate::value::Value;

const DIGITS: [&str; 10] = [
    "ศูนย์",
    "หนึ่ง",
    "สอง",
    "สาม",
    "สี่",
    "ห้า",
    "หก",
    "เจ็ด",
    "แปด",
    "เก้า",
];
const PLACES: [&str; 6] = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];

/// A number below a million, in words (nothing for 0).
fn group(n: u32) -> String {
    let mut out = String::new();
    let tens = (n / 10) % 10;
    for p in (0..6).rev() {
        let d = (n / 10u32.pow(p)) % 10;
        if d == 0 {
            continue;
        }
        match p {
            1 => out.push_str(match d {
                1 => "สิบ",
                2 => "ยี่สิบ",
                _ => "",
            }),
            0 => {}
            _ => {}
        }
        match (p, d) {
            (1, 1 | 2) => {}
            (0, 1) if tens != 0 => out.push_str("เอ็ด"),
            _ => {
                out.push_str(DIGITS[d as usize]);
                out.push_str(PLACES[p as usize]);
            }
        }
    }
    out
}

/// The digits of a whole number as groups of a million, from the lowest.
fn thai(digits: &str) -> String {
    let trimmed = digits.trim_start_matches('0');
    if trimmed.is_empty() {
        return DIGITS[0].to_string();
    }
    let split = trimmed.len().saturating_sub(6);
    let (high, low) = trimmed.split_at(split);
    let low: u32 = low.parse().unwrap_or(0);
    if high.is_empty() {
        return group(low);
    }
    format!("{}ล้าน{}", thai(high), group(low))
}

fn bahttext(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let x = e.num(&a[0])?;
    let rounded = round_decimal(x.abs(), 2);
    let baht = rounded.trunc();
    let satang = ((rounded - baht) * 100.0).round() as u32;
    // Fifteen significant digits: what LibreOffice writes for the amount.
    let digits = if baht < 1e15 {
        format!("{baht:.0}")
    } else {
        let s = format!("{baht:.14e}");
        let (mantissa, exp) = s.split_once('e').unwrap_or((&s, "0"));
        let exp: usize = exp.parse().unwrap_or(0);
        let mut d: String = mantissa.chars().filter(char::is_ascii_digit).collect();
        while d.len() < exp + 1 {
            d.push('0');
        }
        d
    };
    let mut out = String::new();
    if x < 0.0 && (baht != 0.0 || satang != 0) {
        out.push_str("ลบ");
    }
    let has_baht = baht != 0.0;
    if has_baht || satang == 0 {
        out.push_str(&thai(&digits));
        out.push_str("บาท");
    }
    if satang == 0 {
        out.push_str("ถ้วน");
    } else {
        out.push_str(&thai(&satang.to_string()));
        out.push_str("สตางค์");
    }
    Ok(Value::text(out))
}

pub(super) fn register(r: &mut Reg) {
    r.eager("BAHTTEXT", 1, 1, bahttext);
}
