mod old;
use old::{round_dec, Mode};
use std::collections::BTreeMap;
fn n10(n: i32) -> f64 { 10f64.powi(n) }
fn approx_value(v: f64) -> f64 {
    const BIG: f64 = 2199023255552.0;
    if v == 0.0 || !v.is_finite() || v > BIG { return v; }
    let mut e = 14 - v.abs().log10().floor() as i32; let neg = e < 0; if neg { e = -e; }
    let f = n10(e); let mut w = if neg { v / f } else { v * f };
    if !w.is_finite() { return v; } w = w.round(); if neg { w *= f } else { w /= f } w
}
fn lo_round(x: f64, d: i32, mode: Mode) -> f64 {
    if !x.is_finite() || x == 0.0 { return x; }
    let neg = x.is_sign_negative(); let mut v = x.abs();
    if d >= 0 && v >= (1u64 << 52) as f64 { return x; }
    let mut fac = 1.0;
    if d != 0 { if d < -20 || d > 20 || v > f64::MAX / 1e20 { return x; } fac = n10(d.abs()); if d < 0 { v /= fac } else { v *= fac } }
    v = match mode { Mode::Nearest => approx_value(v + 0.5).floor(), Mode::Down => approx_value(v).floor(), Mode::Away => approx_value(v).ceil() };
    if d != 0 { if d < 0 { v *= fac } else { v /= fac } }
    if neg { -v } else { v }
}
fn main() {
    let text = std::fs::read_to_string(std::env::args().nth(1).unwrap()).unwrap();
    let kk = |r: f64, d: i32| -> Option<i64> { let k = if d >= 0 { r * 10f64.powi(d) } else { r / 10f64.powi(-d) }; if k.abs() < 1e15 { Some(k.round() as i64) } else { None } };
    // per (function, digits): cases, engine-now mismatches, new-algorithm mismatches
    let mut t: BTreeMap<(usize, i32), [u64; 3]> = BTreeMap::new();
    for l in text.lines() {
        let c: Vec<&str> = l.split('\t').collect();
        let x = f64::from_bits(u64::from_str_radix(c[1], 16).unwrap());
        let d: i32 = c[2].parse().unwrap(); let fnn: usize = c[3].parse().unwrap();
        let lo: f64 = c[4].parse().unwrap(); let mode = [Mode::Nearest, Mode::Away, Mode::Down][fnn];
        let kl = kk(lo, d); if kl.is_none() { continue; }
        let e = t.entry((fnn, d)).or_default(); e[0] += 1;
        if kk(round_dec(x, d, mode), d) != kl { e[1] += 1; }
        if kk(lo_round(x, d, mode), d) != kl { e[2] += 1; }
    }
    for (fnn, name) in [(0usize, "ROUND"), (2, "ROUNDDOWN"), (1, "ROUNDUP")] {
        println!("{name}: digits -> mismatches of [engine now | new algorithm] per cases");
        let mut line = String::new();
        let (mut a, mut b, mut n) = (0u64, 0u64, 0u64);
        for ((f, d), e) in &t { if *f != fnn { continue; } line += &format!(" {d:>3}:{}|{}", e[1], e[2]); a += e[1]; b += e[2]; n += e[0]; }
        println!("{line}\n   total {n} cases: engine now {a}, new algorithm {b}");
    }
}
