mod old;
use old::{round_dec, Mode};
use std::collections::BTreeMap;
fn main() {
    let text = std::fs::read_to_string("cases3.tsv").unwrap();
    let mut t: BTreeMap<(String, i32), (u64, u64)> = BTreeMap::new();
    for line in text.lines() {
        let c: Vec<&str> = line.split('\t').collect();
        let fam = c[0].trim_end_matches('-');
        if !["money", "calc", "tie", "noise"].contains(&fam) { continue; }
        let x = f64::from_bits(u64::from_str_radix(c[1], 16).unwrap());
        let d: i32 = c[2].parse().unwrap(); let fnn: usize = c[3].parse().unwrap();
        let lo = f64::from_bits(u64::from_str_radix(c[4], 16).unwrap());
        let r = round_dec(x, d, [Mode::Nearest, Mode::Away, Mode::Down][fnn]);
        let e = t.entry((["ROUND", "ROUNDUP", "ROUNDDOWN"][fnn].to_string(), d)).or_default();
        e.0 += 1; if r != lo { e.1 += 1; }
    }
    println!("typical families (money, calc, ties, noise): engine vs LibreOffice, value mismatches / cases, by digits");
    for ((f, d), (n, b)) in &t { if *b > 0 || (*d == 2) { println!("{f:10} digits {d:>2}: {b:>5} / {n:>6}"); } }
}
