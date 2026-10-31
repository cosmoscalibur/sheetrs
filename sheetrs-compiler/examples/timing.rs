//! Phase timing and peak memory on a large workbook:
//! `cargo run --release -p sheetrs-compiler --example timing -- FILE [--in R]... [--out R]...`
use sheetrs_compiler::{BindingSpec, Spec, compile, loader};
use std::time::Instant;

fn peak_rss_mb() -> f64 {
    std::fs::read_to_string("/proc/self/status")
        .ok()
        .and_then(|s| {
            s.lines()
                .find(|l| l.starts_with("VmHWM:"))
                .and_then(|l| l.split_whitespace().nth(1)?.parse::<f64>().ok())
        })
        .map_or(0.0, |kb| kb / 1024.0)
}

fn main() {
    let mut args = std::env::args().skip(1);
    let file = args.next().expect("file");
    let (mut ins, mut outs) = (vec![], vec![]);
    while let Some(a) = args.next() {
        let v = args.next().expect("value");
        match a.as_str() {
            "--in" => ins.push(BindingSpec::parse(&v)),
            "--out" => outs.push(BindingSpec::parse(&v)),
            _ => panic!("unknown flag {a}"),
        }
    }
    let t = Instant::now();
    let wb = sheetrs_core::reader::read_workbook(&file).unwrap();
    println!(
        "read (sheetrs)      {:>8.1?}   peak {:.0} MB",
        t.elapsed(),
        peak_rss_mb()
    );
    let t = Instant::now();
    let model = loader::from_workbook_owned(wb, false);
    println!(
        "model               {:>8.1?}   peak {:.0} MB",
        t.elapsed(),
        peak_rss_mb()
    );
    let t = Instant::now();
    let (p, r) = compile(
        &model,
        &Spec {
            inputs: ins,
            outputs: outs,
            ..Spec::default()
        },
    )
    .unwrap();
    println!(
        "compile             {:>8.1?}   peak {:.0} MB",
        t.elapsed(),
        peak_rss_mb()
    );
    println!(
        "kept {} cells / {} formulas; {} formulas folded",
        r.cells_kept, r.formulas_kept, r.formulas_folded
    );
    let t = Instant::now();
    let out = p
        .execute(&serde_json::Value::Null, &Default::default())
        .unwrap();
    println!("run                 {:>8.1?}   -> {out}", t.elapsed());
}
