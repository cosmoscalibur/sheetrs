mod old;
use old::{round_dec, Mode};
fn n10(n: i32) -> f64 { 10f64.powi(n) }
fn approx_value(v: f64, big: f64) -> f64 {
    if v == 0.0 || !v.is_finite() || v > big { return v; }
    let mut e = 14 - v.abs().log10().floor() as i32;
    let neg = e < 0; if neg { e = -e; }
    let f = n10(e);
    let mut w = if neg { v / f } else { v * f };
    if !w.is_finite() { return v; }
    w = w.round();
    if neg { w *= f } else { w /= f }
    w
}
fn lo_round(x: f64, d: i32, mode: Mode, big: f64, shortcut: bool) -> f64 {
    if !x.is_finite() { return x; }
    if shortcut && d == 0 && mode == Mode::Nearest { return x.round(); }
    if x == 0.0 { return x; }
    let neg = x.is_sign_negative(); let mut v = x.abs();
    if d >= 0 && v >= (1u64 << 52) as f64 { return x; }
    let mut fac = 1.0;
    if d != 0 { if d < -20 || d > 20 || v > f64::MAX / 1e20 { return x; } fac = n10(d.abs()); if d < 0 { v /= fac } else { v *= fac } }
    v = match mode { Mode::Nearest => approx_value(v + 0.5, big).floor(), Mode::Down => approx_value(v, big).floor(), Mode::Away => approx_value(v, big).ceil() };
    if d != 0 { if d < 0 { v *= fac } else { v /= fac } }
    if neg { -v } else { v }
}
fn main() {
    let text = std::fs::read_to_string(std::env::args().nth(1).unwrap()).unwrap();
    let bigs: Vec<(String, f64)> = vec![("2^36".into(), 2f64.powi(36)), ("2^40".into(), 2f64.powi(40)), ("2^41".into(), 2f64.powi(41)), ("2^42".into(), 2f64.powi(42)), ("2^44".into(), 2f64.powi(44)), ("2^46".into(), 2f64.powi(46)), ("1e14".into(), 1e14), ("none".into(), f64::INFINITY)];
    let cases: Vec<(f64, i32, usize, f64)> = text.lines().map(|l| { let c: Vec<&str> = l.split('\t').collect();
        (f64::from_bits(u64::from_str_radix(c[1], 16).unwrap()), c[2].parse().unwrap(), c[3].parse().unwrap(), c[4].parse().unwrap()) }).collect();
    let kk = |r: f64, d: i32| -> Option<i64> { let k = if d >= 0 { r * 10f64.powi(d) } else { r / 10f64.powi(-d) }; if k.abs() < 1e15 { Some(k.round() as i64) } else { None } };
    println!("mismatches against LibreOffice by function, {} cases", cases.len());
    println!("{:22} {:>8} {:>8} {:>8}", "variant", "ROUND", "ROUNDUP", "ROUNDDOWN");
    let modes = [Mode::Nearest, Mode::Away, Mode::Down];
    let mut line = |name: &str, f: &dyn Fn(f64, i32, usize) -> f64| {
        let mut m = [0u64; 3];
        for &(x, d, fnn, lo) in &cases { let (a, b) = (kk(f(x, d, fnn), d), kk(lo, d)); if a.is_some() && b.is_some() && a != b { m[fnn] += 1; } }
        println!("{:22} {:>8} {:>8} {:>8}", name, m[0], m[1], m[2]);
    };
    line("engine now (text)", &|x, d, f| round_dec(x, d, modes[f]));
    for (n, b) in &bigs { line(&format!("approx, bail above {n}"), &|x, d, f| lo_round(x, d, modes[f], *b, false)); }
    line("approx 2^41, d=0 shortcut", &|x, d, f| lo_round(x, d, modes[f], 2f64.powi(41), true));
}
