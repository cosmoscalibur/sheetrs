//! More statistics: ranks and percentiles, dispersion, regression, the normal
//! distribution and the "A" variants that count text and booleans.

use super::math_ext::{combin, norm_cdf, norm_inv};
use super::stats::variance;
use super::{Reg, VAR, num, opt_num};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

fn sorted(e: &mut Engine, v: &Value) -> Res<Vec<f64>> {
    let mut x = e.numbers(std::slice::from_ref(v))?;
    x.sort_by(f64::total_cmp);
    Ok(x)
}

/// Values of two equal-size ranges where both are numbers.
fn pairs(e: &mut Engine, ys: &Value, xs: &Value) -> Res<(Vec<f64>, Vec<f64>)> {
    let (ay, ax) = (e.array_of(ys.clone()), e.array_of(xs.clone()));
    if ay.data.len() != ax.data.len() {
        return Err(ErrorKind::Na);
    }
    let (mut y, mut x) = (Vec::new(), Vec::new());
    for (a, b) in ay.data.iter().zip(&ax.data) {
        if let Value::Error(k) = a {
            return Err(*k);
        }
        if let Value::Error(k) = b {
            return Err(*k);
        }
        if let (Value::Number(p), Value::Number(q)) = (a, b) {
            y.push(*p);
            x.push(*q);
        }
    }
    Ok((y, x))
}

fn mean(v: &[f64]) -> f64 {
    v.iter().sum::<f64>() / v.len() as f64
}

/// `(sxx, syy, sxy)` of the centred sums.
fn sums(y: &[f64], x: &[f64]) -> (f64, f64, f64) {
    let (my, mx) = (mean(y), mean(x));
    let mut s = (0.0, 0.0, 0.0);
    for (a, b) in y.iter().zip(x) {
        s.0 += (b - mx) * (b - mx);
        s.1 += (a - my) * (a - my);
        s.2 += (a - my) * (b - mx);
    }
    s
}

fn percentile_inc(v: &[f64], k: f64) -> Res<f64> {
    if v.is_empty() || !(0.0..=1.0).contains(&k) {
        return Err(ErrorKind::Domain);
    }
    let pos = k * (v.len() - 1) as f64;
    let lo = pos.floor() as usize;
    let hi = (lo + 1).min(v.len() - 1);
    Ok(v[lo] + (pos - lo as f64) * (v[hi] - v[lo]))
}

fn percentile_exc(v: &[f64], k: f64) -> Res<f64> {
    let n = v.len() as f64;
    if v.is_empty() || k < 1.0 / (n + 1.0) || k > n / (n + 1.0) {
        return Err(ErrorKind::Domain);
    }
    let pos = k * (n + 1.0);
    let lo = pos.floor();
    let (i, j) = (lo as usize - 1, (lo as usize).min(v.len() - 1));
    Ok(v[i] + (pos - lo) * (v[j] - v[i]))
}

fn quartile(e: &mut Engine, a: &[Value], exclusive: bool) -> Res<Value> {
    let v = sorted(e, &a[0])?;
    let q = e.num(&a[1])?.trunc();
    let ok = if exclusive {
        (1.0..=3.0).contains(&q)
    } else {
        (0.0..=4.0).contains(&q)
    };
    if !ok {
        return Err(ErrorKind::Domain);
    }
    let k = q / 4.0;
    (if exclusive {
        percentile_exc(&v, k)
    } else {
        percentile_inc(&v, k)
    })
    .and_then(num)
}

fn rank(e: &mut Engine, a: &[Value], average: bool) -> Res<Value> {
    let x = e.num(&a[0])?;
    let v = e.numbers(&a[1..2])?;
    let asc = opt_num(e, a, 2, 0.0)? != 0.0;
    if !v.contains(&x) {
        // No numbers at all is #VALUE!.
        return Err(if v.is_empty() {
            ErrorKind::Value
        } else {
            ErrorKind::Na
        });
    }
    let better = v
        .iter()
        .filter(|y| if asc { **y < x } else { **y > x })
        .count() as f64;
    let ties = v.iter().filter(|y| **y == x).count() as f64;
    num(better + 1.0 + if average { (ties - 1.0) / 2.0 } else { 0.0 })
}

fn percent_rank(e: &mut Engine, a: &[Value], exclusive: bool) -> Res<Value> {
    let v = sorted(e, &a[0])?;
    let x = e.num(&a[1])?;
    let sig = opt_num(e, a, 2, 3.0)?.trunc();
    if sig < 1.0 {
        return Err(ErrorKind::Domain);
    }
    let n = v.len();
    if n == 0 || x < v[0] || x > v[n - 1] {
        return Err(ErrorKind::Value);
    }
    // Position of x among the sorted values, interpolated when between two.
    let below = v.iter().filter(|y| **y < x).count();
    let pos = if v.get(below) == Some(&x) {
        below as f64
    } else {
        let (lo, hi) = (v[below - 1], v[below]);
        below as f64 - 1.0 + (x - lo) / (hi - lo)
    };
    let r = if exclusive {
        (pos + 1.0) / (n as f64 + 1.0)
    } else {
        pos / (n as f64 - 1.0)
    };
    let f = 10f64.powf(sig);
    // Rounded to the significance asked.
    num((r * f).round() / f)
}

/// Numbers of the "A" functions: text counts as 0 and TRUE as 1 inside ranges.
fn numbers_a(e: &mut Engine, a: &[Value]) -> Res<Vec<f64>> {
    let mut out = Vec::new();
    for v in a {
        match v {
            Value::Ref(_) | Value::Array(_) => {
                for x in &e.array_of(v.clone()).data {
                    match x {
                        Value::Number(n) => out.push(*n),
                        Value::Bool(b) => out.push(f64::from(u8::from(*b))),
                        Value::Text(_) => out.push(0.0),
                        Value::Error(k) => return Err(*k),
                        _ => {}
                    }
                }
            }
            v => out.extend(e.numbers(std::slice::from_ref(v))?),
        }
    }
    Ok(out)
}

fn var_a(v: &[f64], sample: bool) -> Res<f64> {
    let n = v.len() as f64;
    if n < if sample { 2.0 } else { 1.0 } {
        return Err(ErrorKind::DivZero);
    }
    let m = mean(v);
    Ok(v.iter().map(|x| (x - m).powi(2)).sum::<f64>() / if sample { n - 1.0 } else { n })
}

pub(super) fn register(r: &mut Reg) {
    // ----- dispersion -----
    r.eager("STDEV.S", 1, VAR, |e, a| {
        variance(e, a, true).map(f64::sqrt).and_then(num)
    });
    for name in ["STDEVP", "STDEV.P"] {
        r.eager(name, 1, VAR, |e, a| {
            variance(e, a, false).map(f64::sqrt).and_then(num)
        });
    }
    r.eager("VAR.S", 1, VAR, |e, a| variance(e, a, true).and_then(num));
    for name in ["VARP", "VAR.P"] {
        r.eager(name, 1, VAR, |e, a| variance(e, a, false).and_then(num));
    }
    r.eager("STDEVA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        var_a(&v, true).map(f64::sqrt).and_then(num)
    });
    r.eager("STDEVPA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        var_a(&v, false).map(f64::sqrt).and_then(num)
    });
    r.eager("VARA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        var_a(&v, true).and_then(num)
    });
    r.eager("VARPA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        var_a(&v, false).and_then(num)
    });
    r.eager("AVERAGEA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        if v.is_empty() {
            Err(ErrorKind::DivZero)
        } else {
            num(mean(&v))
        }
    });
    r.eager("MAXA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        num(v.into_iter().reduce(f64::max).unwrap_or(0.0))
    });
    r.eager("MINA", 1, VAR, |e, a| {
        let v = numbers_a(e, a)?;
        num(v.into_iter().reduce(f64::min).unwrap_or(0.0))
    });
    r.eager("DEVSQ", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        if v.is_empty() {
            return Err(ErrorKind::Domain);
        }
        let m = mean(&v);
        num(v.iter().map(|x| (x - m).powi(2)).sum())
    });
    r.eager("AVEDEV", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        if v.is_empty() {
            return Err(ErrorKind::Domain);
        }
        let m = mean(&v);
        num(v.iter().map(|x| (x - m).abs()).sum::<f64>() / v.len() as f64)
    });
    r.eager("GEOMEAN", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        if v.is_empty() || v.iter().any(|x| *x <= 0.0) {
            return Err(ErrorKind::Domain);
        }
        num((v.iter().map(|x| x.ln()).sum::<f64>() / v.len() as f64).exp())
    });
    r.eager("HARMEAN", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        if v.is_empty() || v.iter().any(|x| *x <= 0.0) {
            return Err(ErrorKind::Domain);
        }
        num(v.len() as f64 / v.iter().map(|x| 1.0 / x).sum::<f64>())
    });
    r.eager("SKEW", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        let n = v.len() as f64;
        let s = variance(e, a, true)?.sqrt();
        if n < 3.0 || s == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        let m = mean(&v);
        num(n / ((n - 1.0) * (n - 2.0)) * v.iter().map(|x| ((x - m) / s).powi(3)).sum::<f64>())
    });
    for name in ["SKEW.P", "SKEWP"] {
        r.eager(name, 1, VAR, |e, a| {
            let v = e.numbers(a)?;
            let s = variance(e, a, false)?.sqrt();
            if s == 0.0 {
                return Err(ErrorKind::DivZero);
            }
            let m = mean(&v);
            num(v.iter().map(|x| ((x - m) / s).powi(3)).sum::<f64>() / v.len() as f64)
        });
    }
    r.eager("KURT", 1, VAR, |e, a| {
        let v = e.numbers(a)?;
        let n = v.len() as f64;
        if n < 4.0 {
            return Err(ErrorKind::DivZero);
        }
        let s = variance(e, a, true)?.sqrt();
        if s == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        let m = mean(&v);
        let k = n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0))
            * v.iter().map(|x| ((x - m) / s).powi(4)).sum::<f64>();
        num(k - 3.0 * (n - 1.0) * (n - 1.0) / ((n - 2.0) * (n - 3.0)))
    });
    r.eager("TRIMMEAN", 2, 2, |e, a| {
        let v = sorted(e, &a[0])?;
        let p = e.num(&a[1])?;
        if !(0.0..1.0).contains(&p) || v.is_empty() {
            return Err(ErrorKind::Domain);
        }
        let k = (v.len() as f64 * p / 2.0).floor() as usize;
        num(mean(&v[k..v.len() - k]))
    });
    for name in ["MODE", "MODE.SNGL"] {
        r.eager(name, 1, VAR, |e, a| {
            let v = e.numbers(a)?;
            let mut best: Option<(usize, f64)> = None;
            for x in &v {
                let c = v.iter().filter(|y| *y == x).count();
                if c > 1 && best.is_none_or(|(b, _)| c > b) {
                    best = Some((c, *x));
                }
            }
            // #VALUE! when nothing repeats.
            best.map_or(Err(ErrorKind::Value), |(_, x)| num(x))
        });
    }

    // ----- ranks and percentiles -----
    r.eager("RANK", 2, 3, |e, a| rank(e, a, false));
    r.eager("RANK.EQ", 2, 3, |e, a| rank(e, a, false));
    r.eager("RANK.AVG", 2, 3, |e, a| rank(e, a, true));
    for name in ["PERCENTILE", "PERCENTILE.INC"] {
        r.eager(name, 2, 2, |e, a| {
            let v = sorted(e, &a[0])?;
            percentile_inc(&v, e.num(&a[1])?).and_then(num)
        });
    }
    r.eager("PERCENTILE.EXC", 2, 2, |e, a| {
        let v = sorted(e, &a[0])?;
        percentile_exc(&v, e.num(&a[1])?).and_then(num)
    });
    for name in ["QUARTILE", "QUARTILE.INC"] {
        r.eager(name, 2, 2, |e, a| quartile(e, a, false));
    }
    r.eager("QUARTILE.EXC", 2, 2, |e, a| quartile(e, a, true));
    for name in ["PERCENTRANK", "PERCENTRANK.INC"] {
        r.eager(name, 2, 3, |e, a| percent_rank(e, a, false));
    }
    r.eager("PERCENTRANK.EXC", 2, 3, |e, a| percent_rank(e, a, true));

    // ----- regression and correlation -----
    for name in ["CORREL", "PEARSON"] {
        r.eager(name, 2, 2, |e, a| {
            let (y, x) = pairs(e, &a[0], &a[1])?;
            let (sxx, syy, sxy) = sums(&y, &x);
            if y.len() < 2 || sxx == 0.0 || syy == 0.0 {
                return Err(ErrorKind::DivZero);
            }
            num(sxy / (sxx * syy).sqrt())
        });
    }
    r.eager("RSQ", 2, 2, |e, a| {
        let (y, x) = pairs(e, &a[0], &a[1])?;
        let (sxx, syy, sxy) = sums(&y, &x);
        if y.len() < 2 || sxx == 0.0 || syy == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        num(sxy * sxy / (sxx * syy))
    });
    r.eager("SLOPE", 2, 2, |e, a| {
        let (y, x) = pairs(e, &a[0], &a[1])?;
        let (sxx, _, sxy) = sums(&y, &x);
        if y.len() < 2 || sxx == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        num(sxy / sxx)
    });
    r.eager("INTERCEPT", 2, 2, |e, a| {
        let (y, x) = pairs(e, &a[0], &a[1])?;
        let (sxx, _, sxy) = sums(&y, &x);
        if y.len() < 2 || sxx == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        num(mean(&y) - sxy / sxx * mean(&x))
    });
    for name in ["FORECAST", "FORECAST.LINEAR"] {
        r.eager(name, 3, 3, |e, a| {
            let x0 = e.num(&a[0])?;
            let (y, x) = pairs(e, &a[1], &a[2])?;
            let (sxx, _, sxy) = sums(&y, &x);
            if y.len() < 2 || sxx == 0.0 {
                return Err(ErrorKind::DivZero);
            }
            let b = sxy / sxx;
            num(mean(&y) + b * (x0 - mean(&x)))
        });
    }
    r.eager("STEYX", 2, 2, |e, a| {
        let (y, x) = pairs(e, &a[0], &a[1])?;
        let (sxx, syy, sxy) = sums(&y, &x);
        let n = y.len() as f64;
        if n < 3.0 || sxx == 0.0 {
            return Err(ErrorKind::DivZero);
        }
        num(((syy - sxy * sxy / sxx) / (n - 2.0)).max(0.0).sqrt())
    });
    for (name, sample) in [
        ("COVAR", false),
        ("COVARIANCE.P", false),
        ("COVARIANCE.S", true),
    ] {
        let f: super::EagerFn = if sample {
            |e, a| covariance(e, a, true)
        } else {
            |e, a| covariance(e, a, false)
        };
        r.eager(name, 2, 2, f);
    }

    // ----- the normal distribution -----
    r.eager("NORMSDIST", 1, 1, |e, a| num(norm_cdf(e.num(&a[0])?)));
    r.eager("NORM.S.DIST", 1, 2, |e, a| {
        let z = e.num(&a[0])?;
        if cumulative(e, a, 1)? {
            num(norm_cdf(z))
        } else {
            num(pdf(z))
        }
    });
    for name in ["NORMDIST", "NORM.DIST"] {
        r.eager(name, 4, 4, |e, a| {
            let (x, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if s <= 0.0 {
                return Err(ErrorKind::Domain);
            }
            let z = (x - m) / s;
            if cumulative(e, a, 3)? {
                num(norm_cdf(z))
            } else {
                num(pdf(z) / s)
            }
        });
    }
    for name in ["NORMSINV", "NORM.S.INV"] {
        r.eager(name, 1, 1, |e, a| {
            let p = e.num(&a[0])?;
            if p <= 0.0 || p >= 1.0 {
                Err(ErrorKind::Domain)
            } else {
                num(norm_inv(p))
            }
        });
    }
    for name in ["NORMINV", "NORM.INV"] {
        r.eager(name, 3, 3, |e, a| {
            let (p, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if p <= 0.0 || p >= 1.0 || s <= 0.0 {
                Err(ErrorKind::Domain)
            } else {
                num(m + s * norm_inv(p))
            }
        });
    }
    r.eager("STANDARDIZE", 3, 3, |e, a| {
        let (x, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if s <= 0.0 {
            // A zero deviation divides by zero.
            Err(ErrorKind::DivZero)
        } else {
            num((x - m) / s)
        }
    });
    r.eager("GAUSS", 1, 1, |e, a| num(norm_cdf(e.num(&a[0])?) - 0.5));
    r.eager("PHI", 1, 1, |e, a| num(pdf(e.num(&a[0])?)));

    // ----- discrete distributions -----
    for name in ["BINOMDIST", "BINOM.DIST"] {
        r.eager(name, 4, 4, |e, a| {
            let (k, n, p) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc(), e.num(&a[2])?);
            if n < 0.0 || k < 0.0 || k > n || !(0.0..=1.0).contains(&p) {
                return Err(ErrorKind::Domain);
            }
            let pmf = |i: f64| combin(n, i).map(|c| c * p.powf(i) * (1.0 - p).powf(n - i));
            if cumulative(e, a, 3)? {
                let mut s = 0.0;
                let mut i = 0.0;
                while i <= k {
                    s += pmf(i).ok_or(ErrorKind::Domain)?;
                    i += 1.0;
                }
                num(s)
            } else {
                pmf(k).map_or(Err(ErrorKind::Domain), num)
            }
        });
    }
    for name in ["POISSON", "POISSON.DIST"] {
        r.eager(name, 3, 3, |e, a| {
            let (k, m) = (e.num(&a[0])?.trunc(), e.num(&a[1])?);
            if k < 0.0 || m < 0.0 {
                return Err(ErrorKind::Domain);
            }
            let pmf = |i: f64| (-m + i * m.ln() - ln_fact(i)).exp();
            let pmf = |i: f64| {
                if m == 0.0 {
                    f64::from(u8::from(i == 0.0))
                } else {
                    pmf(i)
                }
            };
            if cumulative(e, a, 2)? {
                let mut s = 0.0;
                let mut i = 0.0;
                while i <= k {
                    s += pmf(i);
                    i += 1.0;
                }
                num(s)
            } else {
                num(pmf(k))
            }
        });
    }
    for name in ["EXPONDIST", "EXPON.DIST"] {
        r.eager(name, 3, 3, |e, a| {
            let (x, l) = (e.num(&a[0])?, e.num(&a[1])?);
            if l <= 0.0 {
                return Err(ErrorKind::Domain);
            }
            // LibreOffice answers 0 below zero.
            if x < 0.0 {
                return num(0.0);
            }
            if cumulative(e, a, 2)? {
                num(1.0 - (-l * x).exp())
            } else {
                num(l * (-l * x).exp())
            }
        });
    }
}

fn ln_fact(n: f64) -> f64 {
    (2..=n as u64).map(|i| (i as f64).ln()).sum()
}

fn pdf(z: f64) -> f64 {
    (-z * z / 2.0).exp() / (2.0 * std::f64::consts::PI).sqrt()
}

fn cumulative(e: &mut Engine, a: &[Value], i: usize) -> Res<bool> {
    match a.get(i) {
        None | Some(Value::Empty) => Ok(true),
        Some(v) => e.boolean(v),
    }
}

fn covariance(e: &mut Engine, a: &[Value], sample: bool) -> Res<Value> {
    let (y, x) = pairs(e, &a[0], &a[1])?;
    let n = y.len() as f64;
    if n < if sample { 2.0 } else { 1.0 } {
        return Err(ErrorKind::DivZero);
    }
    let (_, _, sxy) = sums(&y, &x);
    num(sxy / if sample { n - 1.0 } else { n })
}
