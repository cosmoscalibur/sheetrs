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

fn ulps(a: f64, b: f64) -> String {
    if a.is_nan() || b.is_nan() { return "nan".into(); }
    let k = |v: f64| { let b = v.to_bits() as i64; if b < 0 { i64::MIN - b } else { b } };
    let d = (k(a) as i128 - k(b) as i128).unsigned_abs();
    if d <= 4 { format!("{d}ulp") } else { "big".into() }
}

fn main() {
    let text = std::fs::read_to_string(std::env::args().nth(1).unwrap()).unwrap();
    struct Acc { n: u64, bad: BTreeMap<String, u64> }
    let mut by: BTreeMap<String, Acc> = BTreeMap::new();
    let names = ["text(current) vs LO", "f64 digits vs LO", "LO-corrected port vs LO", "plain f64 vs LO", "f64 digits vs text"];
    let mut kinds: BTreeMap<(String, &str), BTreeMap<String, u64>> = BTreeMap::new();
    let (mut u64_ok, mut u64_n) = (0u64, 0u64);
    let mut f_decl = 0u64; let mut total = 0u64;
    for line in text.lines() {
        let c: Vec<&str> = line.split('\t').collect();
        let fam = c[0].trim_end_matches('-').to_string();
        let x = f64::from_bits(u64::from_str_radix(c[1], 16).unwrap());
        let d: i32 = c[2].parse().unwrap(); let fnn: u32 = c[3].parse().unwrap();
        let lo = f64::from_bits(u64::from_str_radix(c[4], 16).unwrap());
        let mode = [Mode::Nearest, Mode::Away, Mode::Down][fnn as usize];
        total += 1;
        let t = round_dec(x, d, mode);
        let f = round_f64_digits(x, d, mode);
        if f.is_none() { f_decl += 1; }
        let f = f.unwrap_or(t);
        let l = round_lo_corrected(x, d);
        let p = round_plain(x, d, mode);
        if fnn == 0 { u64_n += 1; if u64_fits(x) { u64_ok += 1; } }
        let key = (format!("{} fn{}", fam, fnn), "");
        let acc = by.entry(key.0.clone()).or_insert(Acc { n: 0, bad: BTreeMap::new() });
        acc.n += 1;
        let mut note = |name: &str, a: f64, b: f64| {
            if !(a == b) { *acc.bad.entry(name.to_string()).or_default() += 1; *kinds.entry((name.to_string(), "")).or_default().entry(ulps(a, b)).or_default() += 1; }
        };
        note(names[0], t, lo); note(names[1], f, lo); note(names[4], f, t);
        if fnn == 0 { note(names[2], l, lo); }
        note(names[3], p, lo);
    }
    println!("{total} cases; f64-digits declines {f_decl}; exact-with-u64 answers {u64_ok} of {u64_n} ROUND cases ({:.1} %)", 100.0 * u64_ok as f64 / u64_n as f64);
    println!("\nmismatches (bit-different results), by family and function (fn0 ROUND, fn1 ROUNDUP, fn2 ROUNDDOWN):");
    print!("{:14} {:>7}", "", "cases"); for n in names { print!(" | {:>22}", n); } println!();
    for (k, a) in &by {
        print!("{:14} {:>7}", k, a.n);
        for n in names { print!(" | {:>22}", a.bad.get(n).copied().unwrap_or(0)); }
        println!();
    }
    println!("\nsize of the mismatches (ulps of the result; 'big' = more than 4 ulps):");
    for ((n, _), m) in &kinds { println!("  {n:26} {m:?}"); }
    let mut tot: BTreeMap<&str, u64> = BTreeMap::new();
    for a in by.values() { for n in names { *tot.entry(n).or_default() += a.bad.get(n).copied().unwrap_or(0); } }
    println!("\nTOTAL mismatches: {tot:?}");
}
