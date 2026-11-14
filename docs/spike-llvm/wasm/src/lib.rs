mod rd;
use std::cell::RefCell;
#[unsafe(no_mangle)] pub extern "C" fn round2(x: f64) -> f64 { rd::round_dec(x, 2, rd::Mode::Nearest) }
unsafe extern "C" {
    fn kernel_e(rate: f64, n: i64, a: *const f64, b: *const f64, c: *const f64, d: *mut f64, e: *mut f64, g: *mut f64, h: *mut f64);
    fn kernel_a(rate: f64, n: i64, a: *const f64, b: *const f64, c: *const f64, d: *mut f64, e: *mut f64, g: *mut f64, h: *mut f64);
}
struct S { a: Vec<f64>, b: Vec<f64>, c: Vec<f64>, d: Vec<f64>, e: Vec<f64>, g: Vec<f64>, h: Vec<f64> }
thread_local! { static ST: RefCell<Option<S>> = const { RefCell::new(None) }; }
#[unsafe(no_mangle)] pub extern "C" fn setup(n: u32) {
    let n = n as usize; let mut s = 12345u64; let mut rnd = || { s ^= s << 13; s ^= s >> 7; s ^= s << 17; s };
    let a = (0..n).map(|_| (rnd() % 1000 + 1) as f64).collect();
    let b = (0..n).map(|_| ((rnd() % 49900 + 100) as f64) / 100.0).collect();
    let c = (0..n).map(|_| ((rnd() % 15000 + 5000) as f64) / 10000.0).collect();
    ST.with(|x| *x.borrow_mut() = Some(S { a, b, c, d: vec![0.0; n], e: vec![0.0; n], g: vec![0.0; n], h: vec![0.0; n] }));
}
fn sum(s: &S) -> f64 { s.e[s.e.len() - 1] + s.h.iter().sum::<f64>() + s.g[0] }
#[unsafe(no_mangle)] pub extern "C" fn run_llvm(rate: f64, exact: u32) -> f64 {
    ST.with(|x| { let mut x = x.borrow_mut(); let s = x.as_mut().unwrap(); let n = s.a.len() as i64;
        unsafe { if exact == 1 { kernel_e(rate, n, s.a.as_ptr(), s.b.as_ptr(), s.c.as_ptr(), s.d.as_mut_ptr(), s.e.as_mut_ptr(), s.g.as_mut_ptr(), s.h.as_mut_ptr()) } else { kernel_a(rate, n, s.a.as_ptr(), s.b.as_ptr(), s.c.as_ptr(), s.d.as_mut_ptr(), s.e.as_mut_ptr(), s.g.as_mut_ptr(), s.h.as_mut_ptr()) } }
        sum(s) })
}
#[unsafe(no_mangle)] pub extern "C" fn run_rust(rate: f64, exact: u32) -> f64 {
    ST.with(|x| { let mut x = x.borrow_mut(); let s = x.as_mut().unwrap(); let n = s.a.len(); let mut run = 0.0;
        for i in 0..n {
            let xx = s.a[i] * s.b[i]; let a1 = if xx > 50000.0 { xx * (1.0 - rate) } else { xx }; let y = a1 * s.c[i];
            let d = if exact == 1 { round2(y) } else { (y * 100.0).round() / 100.0 };
            s.d[i] = d; run += d; s.e[i] = run; let q = d / s.a[i]; s.g[i] = if q.is_finite() { q } else { 0.0 };
            s.h[i] = d.max(s.g[i]) - d.min(s.g[i]);
        }
        sum(s) })
}
