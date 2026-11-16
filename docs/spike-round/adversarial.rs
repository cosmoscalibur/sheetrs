mod old;
use old::{round_dec, Mode};
use std::collections::BTreeMap;

const P: [f64; 23] = { let mut t = [1.0; 23]; let mut i = 1; while i < 23 { t[i] = t[i-1] * 10.0; i += 1; } t };
const PI: [u64; 20] = { let mut t = [1u64; 20]; let mut i = 1; while i < 20 { t[i] = t[i-1] * 10; i += 1; } t };

/// Candidate F: the digit rounding of the text definition with f64 arithmetic only (no u128): the 15 digits
/// m = round(ax * 10^k) in f64, then integer steps on m (exact below 2^53).
fn round_f64_digits(x: f64, digits: i32, mode: Mode) -> Option<f64> {
    if x == 0.0 || !x.is_finite() { return Some(x); }
    let ax = x.abs();
    let mut e10 = ax.log10().floor() as i32;
    let m = loop {
        let k = 14 - e10;
        if !(-22..=22).contains(&k) { return None; }
        let s = if k >= 0 { ax * P[k as usize] } else { ax / P[(-k) as usize] };
        let m = s.round_ties_even();
        if m < 1e14 { e10 -= 1; continue; }
        if m >= 1e15 { e10 += 1; continue; }
        break m as u64;
    };
    let keep = e10 + 1 + digits;
    if keep >= 15 { return Some(x); }
    if keep < 0 { return Some(if mode == Mode::Away { 10f64.powi(-digits).copysign(x) } else { 0.0 }); }
    let k = keep as usize;
    let below = PI[15 - k];
    let mut d = m / below;
    let up = match mode { Mode::Nearest => (m / PI[14 - k]) % 10 >= 5, Mode::Away => m % below != 0, Mode::Down => false };
    if up { d += 1; }
    if d == 0 { return Some(0.0); }
    let scale = e10 - keep + 1;
    if scale.unsigned_abs() > 22 { return None; }
    let r = if scale >= 0 { d as f64 * P[scale as usize] } else { d as f64 / P[(-scale) as usize] };
    Some(r.copysign(x))
}

/// Candidate L: scale, then floor(v + 0.5 + a correction that depends on the magnitude), then unscale.
fn round_lo_corrected(x: f64, digits: i32) -> f64 {
    const CORR: [f64; 16] = [0.0, 9e-1, 9e-2, 9e-3, 9e-4, 9e-5, 9e-6, 9e-7, 9e-8, 9e-9, 9e-10, 9e-11, 9e-12, 9e-13, 9e-14, 9e-15];
    if x == 0.0 || !x.is_finite() { return x; }
    let neg = x < 0.0; let mut v = x.abs();
    let fac = if digits.abs() <= 22 { P[digits.unsigned_abs() as usize] } else { return f64::NAN };
    if digits > 0 { v *= fac } else if digits < 0 { v /= fac }
    let nexp = v.log10().floor() as i32;
    let idx = if nexp < 0 { 15 } else if nexp >= 14 { 0 } else { 15 - nexp } as usize;
    v = (v + 0.5 + CORR[idx]).floor();
    if digits > 0 { v /= fac } else if digits < 0 { v *= fac }
    if neg { -v } else { v }
}

/// Candidate P: plain f64, half away from zero.
fn round_plain(x: f64, digits: i32, mode: Mode) -> f64 {
    if x == 0.0 || !x.is_finite() || digits.abs() > 22 { return x; }
    let f = P[digits.unsigned_abs() as usize];
    let y = if digits >= 0 { x * f } else { x / f };
    let r = match mode { Mode::Nearest => y.round(), Mode::Away => y.abs().ceil().copysign(y), Mode::Down => y.trunc() };
    if digits >= 0 { r / f } else { r * f }
}

/// Exact digits with u64 only: Some when mant * 10^k fits a u64 (the integer part of the scaling).
fn u64_fits(x: f64) -> bool {
    if x == 0.0 || !x.is_finite() { return true; }
    let bits = x.abs().to_bits(); let be = ((bits >> 52) & 0x7ff) as i32;
    if be == 0 { return false; }
    let mant = (bits & ((1u64 << 52) - 1)) | (1u64 << 52); let e2 = be - 1075;
    let e10 = x.abs().log10().floor() as i32; let k = 14 - e10;
    if !(0..=19).contains(&k) { return k < 0 && e2 <= 0 && false; }
    let Some(num) = mant.checked_mul(PI[k as usize]) else { return false };
    if e2 >= 0 { num.leading_zeros() as i32 > e2 } else { -e2 < 64 }
}


use std::time::Instant;
fn main() {
    let mut st = 0x9E3779B97F4A7C15u64; let mut rnd = move || { st ^= st << 13; st ^= st >> 7; st ^= st << 17; st };
    let modes = [Mode::Nearest, Mode::Away, Mode::Down];
    let (mut n, mut diff, mut decl) = (0u64, 0u64, 0u64);
    let mut unit = std::collections::BTreeMap::<&str, u64>::new();
    let mut shown = 0;
    for i in 0..30_000_000u64 {
        let d = (rnd() % 17) as i32 - 6; let m = modes[(i % 3) as usize];
        let x = match rnd() % 6 { 0 => f64::from_bits(rnd()), 1 => (rnd() % 100_000_000) as f64 / 100.0, 2 => ((rnd() % 20_000_000) as f64 + 0.5) / 10f64.powi((rnd() % 7) as i32),
            3 => (rnd() % 1_000_000) as f64 / 1000.0 + 0.005, 4 => (rnd() as f64 / u64::MAX as f64) * 10f64.powi((rnd() % 40) as i32 - 20),
            _ => ((rnd() % 9_000) as f64 + 0.9999999999999) * 10f64.powi((rnd() % 30) as i32 - 10) };
        let t = round_dec(x, d, m);
        match round_f64_digits(x, d, m) {
            None => decl += 1,
            Some(f) => { n += 1; if f.to_bits() != t.to_bits() { diff += 1;
                let q = (f - t).abs(); let unit_size = 10f64.powi(-d);
                *unit.entry(if q <= unit_size * 1.0000001 { "one unit of the rounded place" } else { "more" }).or_default() += 1;
                if shown < 6 { shown += 1; println!("DIFF x={x:e} digits={d} text={t:e} f64digits={f:e}"); } } }
        }
    }
    println!("answered {n}, declined {decl}, differences vs the text definition {diff} ({:.4} ppm): {unit:?}", 1e6 * diff as f64 / n as f64);
    // speed: money-like values, ROUND(x,2)
    let xs: Vec<f64> = (0..1_000_000).map(|_| ((rnd() % 100_000_000) as f64) / 997.0).collect();
    macro_rules! t { ($name:expr, $f:expr) => {{ let mut best = f64::MAX; let mut a = 0.0; for _ in 0..7 { let t0 = Instant::now(); for &x in &xs { a += $f(x); } best = best.min(t0.elapsed().as_secs_f64()); } println!("{:34} {:.1} ns per call (chk {a:.1})", $name, best * 1e9 / xs.len() as f64); }}; }
    t!("text definition (before)", |x| round_dec(x, 2, Mode::Nearest));
    t!("f64 digits (no u128)", |x| round_f64_digits(x, 2, Mode::Nearest).unwrap());
    t!("LO-corrected port", |x| round_lo_corrected(x, 2));
    t!("plain f64", |x| round_plain(x, 2, Mode::Nearest));
}
