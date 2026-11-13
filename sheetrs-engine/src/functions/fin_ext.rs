//! More finance: annuity parts, depreciation, irregular cash flows, rates.

use super::{Reg, num, opt_num};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

fn fv_of(rate: f64, n: f64, pmt: f64, pv: f64, ty: f64) -> f64 {
    if rate == 0.0 {
        return -(pv + pmt * n);
    }
    let g = (1.0 + rate).powf(n);
    -(pv * g + pmt * (1.0 + rate * ty) * (g - 1.0) / rate)
}

fn pmt_of(rate: f64, n: f64, pv: f64, fv: f64, ty: f64) -> f64 {
    if rate == 0.0 {
        return -(pv + fv) / n;
    }
    let g = (1.0 + rate).powf(n);
    -(rate * (fv + pv * g)) / ((1.0 + rate * ty) * (g - 1.0))
}

/// Interest part of payment `per` (1-based) of a loan. For payments at the start
/// of the period (`ty` 1) the first one has no interest and the rest are the
/// end-of-period ones discounted by a period.
fn ipmt_of(rate: f64, per: f64, n: f64, pv: f64, fv: f64, ty: f64) -> f64 {
    let pmt = pmt_of(rate, n, pv, fv, 0.0);
    let ordinary = fv_of(rate, per - 1.0, pmt, pv, 0.0) * rate;
    if ty == 1.0 {
        if per == 1.0 {
            0.0
        } else {
            ordinary / (1.0 + rate)
        }
    } else {
        ordinary
    }
}

fn opt_num_or(e: &mut Engine, a: &[Value], i: usize, d: f64) -> Res<f64> {
    opt_num(e, a, i, d)
}

fn flows_dates(e: &mut Engine, v: &Value, d: &Value) -> Res<(Vec<f64>, Vec<f64>)> {
    let (v, d) = (
        e.numbers(std::slice::from_ref(v))?,
        e.numbers(std::slice::from_ref(d))?,
    );
    if v.len() != d.len() || v.len() < 2 {
        return Err(ErrorKind::Domain);
    }
    let d: Vec<f64> = d.into_iter().map(f64::floor).collect();
    if d.iter().any(|x| *x < d[0]) {
        return Err(ErrorKind::Domain);
    }
    Ok((v, d))
}

fn xnpv(rate: f64, v: &[f64], d: &[f64]) -> f64 {
    v.iter()
        .zip(d)
        .map(|(c, t)| c / (1.0 + rate).powf((t - d[0]) / 365.0))
        .sum()
}

pub(super) fn register(r: &mut Reg) {
    r.eager("NPER", 3, 5, |e, a| {
        let (rate, pmt, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        let fv = opt_num_or(e, a, 3, 0.0)?;
        let ty = f64::from(u8::from(opt_num_or(e, a, 4, 0.0)? != 0.0));
        if rate == 0.0 {
            if pmt == 0.0 {
                return Err(ErrorKind::Domain);
            }
            return num(-(pv + fv) / pmt);
        }
        let p = pmt * (1.0 + rate * ty);
        let (top, bottom) = (p - fv * rate, p + pv * rate);
        if top / bottom <= 0.0 || bottom == 0.0 {
            return Err(ErrorKind::Domain);
        }
        num((top / bottom).ln() / (1.0 + rate).ln())
    });
    r.eager("RATE", 3, 6, |e, a| {
        let (n, pmt, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        let fv = opt_num_or(e, a, 3, 0.0)?;
        let ty = f64::from(u8::from(opt_num_or(e, a, 4, 0.0)? != 0.0));
        let mut x = opt_num_or(e, a, 5, 0.1)?;
        let f = |r: f64| {
            if r.abs() < 1e-12 {
                pv + pmt * n + fv
            } else {
                let g = (1.0 + r).powf(n);
                pv * g + pmt * (1.0 + r * ty) * (g - 1.0) / r + fv
            }
        };
        for _ in 0..200 {
            let h = 1e-7_f64.max(x.abs() * 1e-7);
            let (y, dy) = (f(x), (f(x + h) - f(x - h)) / (2.0 * h));
            if !y.is_finite() || dy == 0.0 || !dy.is_finite() {
                break;
            }
            let nx = x - y / dy;
            if nx <= -1.0 {
                x = (x - 1.0) / 2.0;
                continue;
            }
            if (nx - x).abs() < 1e-12 {
                return num(nx);
            }
            x = nx;
        }
        Err(ErrorKind::Domain)
    });
    r.eager("IPMT", 4, 6, |e, a| {
        let (rate, per, n, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        let fv = opt_num_or(e, a, 4, 0.0)?;
        let ty = f64::from(u8::from(opt_num_or(e, a, 5, 0.0)? != 0.0));
        if per < 1.0 || per > n {
            return Err(ErrorKind::Domain);
        }
        num(ipmt_of(rate, per, n, pv, fv, ty))
    });
    r.eager("PPMT", 4, 6, |e, a| {
        let (rate, per, n, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        let fv = opt_num_or(e, a, 4, 0.0)?;
        let ty = f64::from(u8::from(opt_num_or(e, a, 5, 0.0)? != 0.0));
        if per < 1.0 || per > n {
            return Err(ErrorKind::Domain);
        }
        num(pmt_of(rate, n, pv, fv, ty) - ipmt_of(rate, per, n, pv, fv, ty))
    });
    r.eager("ISPMT", 4, 4, |e, a| {
        let (rate, per, n, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        if n == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        num(pv * rate * (per / n - 1.0))
    });
    for (name, principal) in [("CUMIPMT", false), ("CUMPRINC", true)] {
        let f: super::EagerFn = if principal {
            |e, a| cumulative(e, a, true)
        } else {
            |e, a| cumulative(e, a, false)
        };
        r.eager(name, 6, 6, f);
    }
    r.eager("SLN", 3, 3, |e, a| {
        let (c, s, l) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if l == 0.0 {
            Err(ErrorKind::DivZero)
        } else {
            num((c - s) / l)
        }
    });
    r.eager("SYD", 4, 4, |e, a| {
        let (c, s, l, p) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        if l <= 0.0 || p <= 0.0 || p > l || s < 0.0 {
            return Err(ErrorKind::Domain);
        }
        num((c - s) * (l - p + 1.0) * 2.0 / (l * (l + 1.0)))
    });
    r.eager("DDB", 4, 5, |e, a| {
        let (cost, salvage, life, period) =
            (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        let factor = opt_num_or(e, a, 4, 2.0)?;
        if cost < 0.0
            || salvage < 0.0
            || life <= 0.0
            || period <= 0.0
            || period > life
            || factor <= 0.0
        {
            return Err(ErrorKind::Domain);
        }
        let mut total = 0.0;
        let mut dep = 0.0;
        let mut i = 1.0;
        while i <= period.ceil() {
            dep = ((cost - total) * factor / life).min((cost - total - salvage).max(0.0));
            if i < period.ceil() || period.fract() == 0.0 {
                total += dep;
            }
            i += 1.0;
        }
        num(dep)
    });
    r.eager("DB", 4, 5, |e, a| {
        let (cost, salvage, life, period) =
            (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?, e.num(&a[3])?);
        let month = opt_num_or(e, a, 4, 12.0)?.trunc();
        if cost < 0.0
            || salvage < 0.0
            || life <= 0.0
            || period <= 0.0
            || !(1.0..=12.0).contains(&month)
            || period > life + if month < 12.0 { 1.0 } else { 0.0 }
            || cost == 0.0
        {
            return Err(ErrorKind::Domain);
        }
        let rate = ((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0).round() / 1000.0;
        let mut total = 0.0;
        let mut dep = 0.0;
        let mut p = 1.0;
        while p <= period.trunc() {
            dep = if p == 1.0 {
                cost * rate * month / 12.0
            } else if p == life + 1.0 {
                (cost - total) * rate * (12.0 - month) / 12.0
            } else {
                (cost - total) * rate
            };
            total += dep;
            p += 1.0;
        }
        num(dep)
    });
    r.eager("MIRR", 3, 3, |e, a| {
        let v = e.numbers(&a[..1])?;
        let (fin, re) = (e.num(&a[1])?, e.num(&a[2])?);
        let n = v.len() as f64;
        if !v.iter().any(|x| *x > 0.0) || !v.iter().any(|x| *x < 0.0) {
            return Err(ErrorKind::DivZero);
        }
        let npv = |rate: f64, keep: fn(f64) -> bool| -> f64 {
            v.iter()
                .enumerate()
                .filter(|(_, c)| keep(**c))
                .map(|(i, c)| c / (1.0 + rate).powf(i as f64 + 1.0))
                .sum()
        };
        let pos = npv(re, |c| c > 0.0);
        let neg = npv(fin, |c| c < 0.0);
        num((-pos * (1.0 + re).powf(n) / (neg * (1.0 + fin))).powf(1.0 / (n - 1.0)) - 1.0)
    });
    r.eager("XNPV", 3, 3, |e, a| {
        let rate = e.num(&a[0])?;
        let (v, d) = flows_dates(e, &a[1], &a[2])?;
        if rate <= -1.0 {
            return Err(ErrorKind::Domain);
        }
        num(xnpv(rate, &v, &d))
    });
    r.eager("XIRR", 2, 3, |e, a| {
        let (v, d) = flows_dates(e, &a[0], &a[1])?;
        if !v.iter().any(|x| *x > 0.0) || !v.iter().any(|x| *x < 0.0) {
            return Err(ErrorKind::Domain);
        }
        let mut x = opt_num_or(e, a, 2, 0.1)?;
        for _ in 0..200 {
            let y = xnpv(x, &v, &d);
            let dy: f64 = v
                .iter()
                .zip(&d)
                .map(|(c, t)| {
                    let k = (t - d[0]) / 365.0;
                    -k * c / (1.0 + x).powf(k + 1.0)
                })
                .sum();
            if dy == 0.0 || !dy.is_finite() || !y.is_finite() {
                break;
            }
            let nx = x - y / dy;
            if nx <= -1.0 {
                x = (x - 1.0) / 2.0;
                continue;
            }
            if (nx - x).abs() < 1e-12 {
                return num(nx);
            }
            x = nx;
        }
        Err(ErrorKind::Domain)
    });
    r.eager("EFFECT", 2, 2, |e, a| {
        let (n, p) = (e.num(&a[0])?, e.num(&a[1])?.trunc());
        if n <= 0.0 || p < 1.0 {
            return Err(ErrorKind::Domain);
        }
        num((1.0 + n / p).powf(p) - 1.0)
    });
    r.eager("NOMINAL", 2, 2, |e, a| {
        let (n, p) = (e.num(&a[0])?, e.num(&a[1])?.trunc());
        if n <= 0.0 || p < 1.0 {
            return Err(ErrorKind::Domain);
        }
        num(p * ((1.0 + n).powf(1.0 / p) - 1.0))
    });
    r.eager("PDURATION", 3, 3, |e, a| {
        let (rate, pv, fv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if rate <= 0.0 || pv <= 0.0 || fv <= 0.0 {
            return Err(ErrorKind::Domain);
        }
        num((fv / pv).ln() / (1.0 + rate).ln())
    });
    r.eager("RRI", 3, 3, |e, a| {
        let (n, pv, fv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if n <= 0.0 || pv == 0.0 {
            return Err(ErrorKind::Domain);
        }
        num((fv / pv).powf(1.0 / n) - 1.0)
    });
    r.eager("FVSCHEDULE", 2, 2, |e, a| {
        let p = e.num(&a[0])?;
        let rates = e.numbers(&a[1..2])?;
        num(rates.iter().fold(p, |acc, r| acc * (1.0 + r)))
    });
    r.eager("DOLLARDE", 2, 2, |e, a| dollar(e, a, true));
    r.eager("DOLLARFR", 2, 2, |e, a| dollar(e, a, false));
}

fn cumulative(e: &mut Engine, a: &[Value], principal: bool) -> Res<Value> {
    let (rate, n, pv) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
    let (s, t) = (e.num(&a[3])?.trunc(), e.num(&a[4])?.trunc());
    let ty = e.num(&a[5])?;
    if rate <= 0.0
        || n <= 0.0
        || pv <= 0.0
        || s < 1.0
        || t < s
        || t > n
        || !(ty == 0.0 || ty == 1.0)
    {
        return Err(ErrorKind::Domain);
    }
    let mut sum = 0.0;
    let mut per = s;
    while per <= t {
        let i = ipmt_of(rate, per, n, pv, 0.0, ty);
        sum += if principal {
            pmt_of(rate, n, pv, 0.0, ty) - i
        } else {
            i
        };
        per += 1.0;
    }
    num(sum)
}

fn dollar(e: &mut Engine, a: &[Value], to_decimal: bool) -> Res<Value> {
    let (x, frac) = (e.num(&a[0])?, e.num(&a[1])?.trunc());
    if frac < 0.0 {
        return Err(ErrorKind::Domain);
    }
    if frac == 0.0 {
        return Err(ErrorKind::DivZero);
    }
    let int = x.trunc();
    let digits = (frac.log10().ceil()).max(1.0);
    let scale = 10f64.powf(digits);
    if to_decimal {
        num(int + (x - int) * scale / frac)
    } else {
        num(int + (x - int) * frac / scale)
    }
}
