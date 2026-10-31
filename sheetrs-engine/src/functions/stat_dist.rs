//! The continuous and discrete distributions beyond the normal one: gamma, chi-square, t, F,
//! beta, log-normal, Weibull, hypergeometric, negative binomial and the tests and intervals built
//! on them. Arguments outside a function's domain are `Domain` (LibreOffice's `Err:502`), as the
//! oracle (`tests/assets/semantics.*`) has them.

use super::math_ext::{norm_cdf, norm_inv};
use super::special::{
    beta_reg, beta_reg_xy, gamma, gamma_p, gamma_q, invert_increasing, invert_on, ln_beta, ln_gamma,
};
use super::{Reg, num};
use crate::engine::{Engine, Res};
use crate::ir::Ns;
use crate::value::{ErrorKind, Value};
use std::f64::consts::PI;

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// A value of the incomplete gamma function, or `Domain` where LibreOffice does not converge.
fn converged(v: f64) -> Res<f64> {
    if v.is_nan() { Err(DOMAIN) } else { Ok(v) }
}

/// The density of the chi-square distribution as LibreOffice computes it: the logarithm form for a
/// large `df * x`, the product of the terms otherwise.
fn chisq_pdf(x: f64, df: f64) -> f64 {
    if df * x > 1_391_000.0 {
        return ((0.5 * df - 1.0) * (x * 0.5).ln() - 0.5 * x - 2f64.ln() - ln_gamma(0.5 * df))
            .exp();
    }
    let (mut value, mut count) = if df % 2.0 < 0.5 {
        (0.5, 2.0)
    } else {
        (1.0 / (x * 2.0 * std::f64::consts::PI).sqrt(), 1.0)
    };
    while count < df {
        value *= x / count;
        count += 2.0;
    }
    if x >= 1425.0 {
        (value.ln() - x / 2.0).exp()
    } else {
        value * (-x / 2.0).exp()
    }
}

/// The shape of a t, F or beta distribution: what LibreOffice computes accurately.
fn shape(x: f64) -> Res<()> {
    if x > 1e7 {
        crate::notes::note(crate::notes::BETA_LARGE);
        Err(ErrorKind::Unsupported)
    } else {
        Ok(())
    }
}

/// `invert_increasing` of a function that may not converge (`NaN`): then LibreOffice's iteration
/// fails as well.
fn invert_checked(f: impl Fn(f64) -> f64, target: f64) -> Res<f64> {
    let bad = std::cell::Cell::new(false);
    let x = invert_increasing(
        |x| {
            let v = f(x);
            if v.is_nan() {
                bad.set(true);
            }
            v
        },
        target,
    );
    if bad.get() { Err(DOMAIN) } else { Ok(x) }
}

fn cumulative(e: &mut Engine, a: &[Value], i: usize) -> Res<bool> {
    match a.get(i) {
        None | Some(Value::Empty) => Ok(true),
        Some(v) => e.boolean(v),
    }
}

/// The argument as a whole number (LibreOffice cuts the degrees of freedom and the counts).
fn whole(e: &mut Engine, a: &[Value], i: usize) -> Res<f64> {
    Ok(e.num(&a[i])?.floor())
}

fn mean(v: &[f64]) -> f64 {
    v.iter().sum::<f64>() / v.len() as f64
}

/// The sum of squared deviations.
fn ss(v: &[f64]) -> f64 {
    let m = mean(v);
    v.iter().map(|x| (x - m) * (x - m)).sum()
}

// ----- the t distribution -----

/// `P(T > t)` for `t >= 0`.
fn t_tail(t: f64, df: f64) -> f64 {
    0.5 * t_two(t, df)
}

fn t_cdf(t: f64, df: f64) -> f64 {
    if t >= 0.0 {
        1.0 - t_tail(t, df)
    } else {
        t_tail(-t, df)
    }
}

/// `P(|T| > t)` for `t >= 0`.
fn t_two(t: f64, df: f64) -> f64 {
    let d = df + t * t;
    beta_reg_xy(df / d, t * t / d, df / 2.0, 0.5)
}

fn t_pdf(t: f64, df: f64) -> f64 {
    (ln_gamma((df + 1.0) / 2.0)
        - ln_gamma(df / 2.0)
        - 0.5 * (df * PI).ln()
        - (df + 1.0) / 2.0 * (t * t / df).ln_1p())
    .exp()
}

/// The `t >= 0` with `P(T > t) = q`.
fn t_tail_inv(q: f64, df: f64) -> f64 {
    invert_increasing(|t| -t_tail(t, df), -q)
}

fn t_inv(p: f64, df: f64) -> f64 {
    if p == 0.5 {
        0.0
    } else if p > 0.5 {
        t_tail_inv(1.0 - p, df)
    } else {
        -t_tail_inv(p, df)
    }
}

/// The `t >= 0` with `P(|T| > t) = p`.
fn t_two_inv(p: f64, df: f64) -> f64 {
    if p >= 1.0 {
        0.0
    } else {
        invert_increasing(|t| -t_two(t, df), -p)
    }
}

// ----- the F distribution -----

fn f_cdf(x: f64, d1: f64, d2: f64) -> f64 {
    let d = d1 * x + d2;
    beta_reg_xy(d1 * x / d, d2 / d, d1 / 2.0, d2 / 2.0)
}

fn f_rt(x: f64, d1: f64, d2: f64) -> f64 {
    let d = d1 * x + d2;
    beta_reg_xy(d2 / d, d1 * x / d, d2 / 2.0, d1 / 2.0)
}

fn f_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x == 0.0 {
        return if d1 == 2.0 {
            1.0
        } else if d1 < 2.0 {
            f64::INFINITY
        } else {
            0.0
        };
    }
    ((d1 / 2.0) * (d1 / d2).ln() + (d1 / 2.0 - 1.0) * x.ln()
        - (d1 + d2) / 2.0 * (d1 * x / d2).ln_1p()
        - ln_beta(d1 / 2.0, d2 / 2.0))
    .exp()
}

// ----- discrete helpers -----

fn ln_choose(n: f64, k: f64) -> f64 {
    ln_gamma(n + 1.0) - ln_gamma(k + 1.0) - ln_gamma(n - k + 1.0)
}

fn binom_pmf(n: f64, k: f64, p: f64) -> f64 {
    if p == 0.0 {
        return f64::from(u8::from(k == 0.0));
    }
    if p == 1.0 {
        return f64::from(u8::from(k == n));
    }
    (ln_choose(n, k) + k * p.ln() + (n - k) * (-p).ln_1p()).exp()
}

fn hyper_pmf(s: f64, n: f64, m: f64, big_n: f64) -> f64 {
    (ln_choose(m, s) + ln_choose(big_n - m, n - s) - ln_choose(big_n, n)).exp()
}

/// The hypergeometric arguments `(s, n, M, N)`: whole numbers that can happen.
fn hyper_args(e: &mut Engine, a: &[Value]) -> Res<(f64, f64, f64, f64)> {
    let (s, n, m, big_n) = (
        whole(e, a, 0)?,
        whole(e, a, 1)?,
        whole(e, a, 2)?,
        whole(e, a, 3)?,
    );
    if s < 0.0
        || n < 0.0
        || m < 0.0
        || s > n
        || s > m
        || n > big_n
        || m > big_n
        || n - s > big_n - m
    {
        return Err(DOMAIN);
    }
    Ok((s, n, m, big_n))
}

fn negbinom_pmf(f: f64, s: f64, p: f64) -> f64 {
    // `0 * ln(0)` is zero here: no failures, or no chance of one.
    let failures = if f == 0.0 { 0.0 } else { f * (-p).ln_1p() };
    let v = (ln_gamma(f + s) - ln_gamma(s) - ln_gamma(f + 1.0) + s * p.ln() + failures).exp();
    if v.is_nan() { 0.0 } else { v }
}

fn numbers_of(e: &mut Engine, v: &Value) -> Res<Vec<f64>> {
    e.numbers(std::slice::from_ref(v))
}

/// The numbers of a range with its shape; text and empty cells are not numbers.
fn grid(e: &mut Engine, v: &Value) -> Res<(usize, usize, Vec<Value>)> {
    let a = e.array_of(v.clone());
    Ok((a.rows as usize, a.cols as usize, a.data.to_vec()))
}

pub(crate) fn register(r: &mut Reg) {
    // ----- gamma -----
    for name in ["GAMMALN", "GAMMALN.PRECISE"] {
        r.eager(name, 1, 1, |e, a| {
            let x = e.num(&a[0])?;
            if x <= 0.0 {
                return Err(DOMAIN);
            }
            num(ln_gamma(x))
        });
    }
    r.eager("GAMMA", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        match gamma(x) {
            None => Err(DOMAIN),
            Some(g) => num(g),
        }
    });
    r.eager("FISHER", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        if x <= -1.0 || x >= 1.0 {
            return Err(DOMAIN);
        }
        num(0.5 * ((1.0 + x) / (1.0 - x)).ln())
    });
    r.eager("FISHERINV", 1, 1, |e, a| num(e.num(&a[0])?.tanh()));
    for name in ["GAMMADIST", "GAMMA.DIST"] {
        r.eager(name, 4, 4, |e, a| {
            let (x, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if x < 0.0 || al <= 0.0 || be <= 0.0 {
                return Err(DOMAIN);
            }
            if cumulative(e, a, 3)? {
                return num(converged(gamma_p(al, x / be))?);
            }
            if x == 0.0 {
                return if al == 1.0 {
                    num(1.0 / be)
                } else if al < 1.0 {
                    Err(ErrorKind::DivZero)
                } else {
                    num(0.0)
                };
            }
            num(((al - 1.0) * x.ln() - x / be - ln_gamma(al) - al * be.ln()).exp())
        });
    }
    for name in ["GAMMAINV", "GAMMA.INV"] {
        r.eager(name, 3, 3, |e, a| {
            let (p, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if !(0.0..1.0).contains(&p) || al <= 0.0 || be <= 0.0 {
                return Err(DOMAIN);
            }
            if p == 0.0 {
                return num(0.0);
            }
            num(be * invert_checked(|x| gamma_p(al, x), p)?)
        });
    }

    // ----- chi-square -----
    r.eager("CHISQ.DIST", 3, 3, |e, a| {
        let (x, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if df < 1.0 || df > 1e10 || x < 0.0 {
            return Err(DOMAIN);
        }
        if cumulative(e, a, 2)? {
            return num(converged(gamma_p(df / 2.0, x / 2.0))?);
        }
        if x == 0.0 {
            return num(0.0);
        }
        num(
            ((df / 2.0 - 1.0) * (x / 2.0).ln() - x / 2.0 - ln_gamma(df / 2.0) - (2f64).ln() * 1.0)
                .exp(),
        )
    });
    for name in ["CHISQ.DIST.RT", "CHIDIST"] {
        r.eager(name, 2, 2, |e, a| {
            let (x, df) = (e.num(&a[0])?, whole(e, a, 1)?);
            if df < 1.0 || x < 0.0 {
                return Err(DOMAIN);
            }
            num(converged(gamma_q(df / 2.0, x / 2.0))?)
        });
    }
    r.eager("CHISQ.INV", 2, 2, |e, a| {
        let (p, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if !(0.0..1.0).contains(&p) || df < 1.0 || df > 1e10 {
            return Err(DOMAIN);
        }
        if p == 0.0 {
            return num(0.0);
        }
        num(2.0 * invert_checked(|x| gamma_p(df / 2.0, x), p)?)
    });
    for name in ["CHISQ.INV.RT", "CHIINV"] {
        r.eager(name, 2, 2, |e, a| {
            let (p, df) = (e.num(&a[0])?, whole(e, a, 1)?);
            if p <= 0.0 || p > 1.0 || df < 1.0 || df > 1e10 {
                return Err(DOMAIN);
            }
            if p == 1.0 {
                return num(0.0);
            }
            num(2.0 * invert_checked(|x| -gamma_q(df / 2.0, x), -p)?)
        });
    }
    // The ODF `CHISQDIST` (left tail, the type optional) and `CHISQINV`, which take any degrees of
    // freedom and give 0 for a negative `x`.
    r.eager("CHISQDIST", 2, 3, |e, a| {
        let (x, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if df < 1.0 {
            return Err(DOMAIN);
        }
        if cumulative(e, a, 2)? {
            return num(if x <= 0.0 {
                0.0
            } else {
                converged(gamma_p(df / 2.0, x / 2.0))?
            });
        }
        num(if x <= 0.0 { 0.0 } else { chisq_pdf(x, df) })
    });
    r.eager("CHISQINV", 2, 2, |e, a| {
        let (p, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if !(0.0..1.0).contains(&p) || df < 1.0 {
            return Err(DOMAIN);
        }
        if p == 0.0 {
            return num(0.0);
        }
        if df > 1e10 {
            return Err(DOMAIN);
        }
        num(2.0 * invert_checked(|x| gamma_p(df / 2.0, x), p)?)
    });
    for name in ["CHISQ.TEST", "CHITEST"] {
        r.eager(name, 2, 2, |e, a| {
            let (r1, c1, obs) = grid(e, &a[0])?;
            let (r2, c2, exp) = grid(e, &a[1])?;
            if (r1, c1) != (r2, c2) {
                return Err(DOMAIN);
            }
            let (mut chi, mut n) = (0.0, 0.0);
            for (o, x) in obs.iter().zip(&exp) {
                match (o, x) {
                    (Value::Number(o), Value::Number(x)) => {
                        if *x == 0.0 {
                            return Err(ErrorKind::DivZero);
                        }
                        chi += (o - x) * (o - x) / x;
                        n += 1.0;
                    }
                    (Value::Error(k), _) | (_, Value::Error(k)) => return Err(*k),
                    _ => {}
                }
            }
            let df = if r1 == 1 || c1 == 1 {
                n - 1.0
            } else {
                (r1 as f64 - 1.0) * (c1 as f64 - 1.0)
            };
            if df < 1.0 {
                return Err(ErrorKind::Value);
            }
            num(converged(gamma_q(df / 2.0, chi / 2.0))?)
        });
    }

    // ----- t -----
    r.eager("T.DIST", 3, 3, |e, a| {
        let (t, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if df < 1.0 {
            return Err(DOMAIN);
        }
        if cumulative(e, a, 2)? {
            shape(df)?;
            num(t_cdf(t, df))
        } else {
            num(t_pdf(t, df))
        }
    });
    r.eager("T.DIST.RT", 2, 2, |e, a| {
        let (t, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if df < 1.0 {
            return Err(DOMAIN);
        }
        shape(df)?;
        num(t_cdf(-t, df))
    });
    r.eager("T.DIST.2T", 2, 2, |e, a| {
        let (t, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if df < 1.0 || t < 0.0 {
            return Err(DOMAIN);
        }
        shape(df)?;
        num(t_two(t, df))
    });
    r.eager("TDIST", 3, 3, |e, a| {
        let (t, df, tails) = (e.num(&a[0])?, whole(e, a, 1)?, whole(e, a, 2)?);
        if df < 1.0 || t < 0.0 || !(tails == 1.0 || tails == 2.0) {
            return Err(DOMAIN);
        }
        shape(df)?;
        num(if tails == 1.0 {
            t_tail(t, df)
        } else {
            t_two(t, df)
        })
    });
    r.eager("T.INV", 2, 2, |e, a| {
        let (p, df) = (e.num(&a[0])?, whole(e, a, 1)?);
        if p <= 0.0 || p >= 1.0 || df < 1.0 {
            return Err(DOMAIN);
        }
        shape(df)?;
        num(t_inv(p, df))
    });
    for name in ["T.INV.2T", "TINV"] {
        r.eager(name, 2, 2, |e, a| {
            let (p, df) = (e.num(&a[0])?, whole(e, a, 1)?);
            if p <= 0.0 || p > 1.0 || df < 1.0 {
                return Err(DOMAIN);
            }
            shape(df)?;
            num(t_two_inv(p, df))
        });
    }
    for name in ["T.TEST", "TTEST"] {
        r.eager(name, 4, 4, |e, a| {
            let (tails, kind) = (whole(e, a, 2)?, whole(e, a, 3)?);
            if !(tails == 1.0 || tails == 2.0) || !(1.0..=3.0).contains(&kind) {
                return Err(DOMAIN);
            }
            let (t, df) = if kind == 1.0 {
                let (r1, c1, x) = grid(e, &a[0])?;
                let (r2, c2, y) = grid(e, &a[1])?;
                if (r1, c1) != (r2, c2) {
                    return Err(DOMAIN);
                }
                let d: Vec<f64> = x
                    .iter()
                    .zip(&y)
                    .filter_map(|p| match p {
                        (Value::Number(x), Value::Number(y)) => Some(x - y),
                        _ => None,
                    })
                    .collect();
                let n = d.len() as f64;
                if n < 2.0 {
                    return Err(ErrorKind::Value);
                }
                let var = ss(&d) / (n - 1.0);
                if var == 0.0 {
                    return Err(ErrorKind::DivZero);
                }
                (mean(&d) / (var / n).sqrt(), n - 1.0)
            } else {
                let x = numbers_of(e, &a[0])?;
                let y = numbers_of(e, &a[1])?;
                let (n1, n2) = (x.len() as f64, y.len() as f64);
                if n1 < 2.0 || n2 < 2.0 {
                    return Err(ErrorKind::Value);
                }
                let (s1, s2) = (ss(&x), ss(&y));
                if kind == 2.0 {
                    let pooled = (s1 + s2) / (n1 + n2 - 2.0);
                    (
                        (mean(&x) - mean(&y)) / (pooled * (1.0 / n1 + 1.0 / n2)).sqrt(),
                        n1 + n2 - 2.0,
                    )
                } else {
                    let (v1, v2) = (s1 / (n1 - 1.0) / n1, s2 / (n2 - 1.0) / n2);
                    (
                        (mean(&x) - mean(&y)) / (v1 + v2).sqrt(),
                        (v1 + v2) * (v1 + v2) / (v1 * v1 / (n1 - 1.0) + v2 * v2 / (n2 - 1.0)),
                    )
                }
            };
            if t.is_nan() {
                return Err(ErrorKind::Num);
            }
            shape(df)?;
            let t = t.abs();
            num(if tails == 1.0 {
                t_tail(t, df)
            } else {
                t_two(t, df)
            })
        });
    }

    // ----- F -----
    // `FDIST` and `FINV` are two functions: the OpenDocument ones (`F.DIST`, `F.INV` of Excel,
    // four and three arguments, left tail) are `core`; the legacy ones (right tail) are `legacy`.
    for name in ["F.DIST", "FDIST"] {
        r.eager(name, 4, 4, |e, a| {
            let (x, d1, d2) = (e.num(&a[0])?, whole(e, a, 1)?, whole(e, a, 2)?);
            if d1 < 1.0 || d2 < 1.0 || d1 >= 1e10 || d2 >= 1e10 || x < 0.0 {
                return Err(DOMAIN);
            }
            if cumulative(e, a, 3)? {
                shape(d1)?;
                shape(d2)?;
                num(if x == 0.0 { 0.0 } else { f_cdf(x, d1, d2) })
            } else {
                num(f_pdf(x, d1, d2))
            }
        });
    }
    fn f_rt_args(e: &mut Engine, a: &[Value]) -> Res<Value> {
        let (x, d1, d2) = (e.num(&a[0])?, whole(e, a, 1)?, whole(e, a, 2)?);
        if d1 < 1.0 || d2 < 1.0 || d1 >= 1e10 || d2 >= 1e10 || x < 0.0 {
            return Err(DOMAIN);
        }
        shape(d1)?;
        shape(d2)?;
        num(if x == 0.0 { 1.0 } else { f_rt(x, d1, d2) })
    }
    r.eager("F.DIST.RT", 3, 3, f_rt_args);
    r.eager_in(Ns::Legacy, "FDIST", 3, 3, f_rt_args);
    for name in ["F.INV", "FINV"] {
        r.eager(name, 3, 3, |e, a| {
            let (p, d1, d2) = (e.num(&a[0])?, whole(e, a, 1)?, whole(e, a, 2)?);
            if p <= 0.0 || p > 1.0 || d1 < 1.0 || d2 < 1.0 || d1 >= 1e10 || d2 >= 1e10 {
                return Err(DOMAIN);
            }
            if p == 1.0 {
                crate::notes::note(crate::notes::F_INV_ONE);
                return Err(ErrorKind::Unsupported);
            }
            shape(d1)?;
            shape(d2)?;
            num(invert_increasing(|x| f_cdf(x, d1, d2), p))
        });
    }
    fn f_rt_inv(e: &mut Engine, a: &[Value]) -> Res<Value> {
        let (p, d1, d2) = (e.num(&a[0])?, whole(e, a, 1)?, whole(e, a, 2)?);
        if p <= 0.0 || p > 1.0 || d1 < 1.0 || d2 < 1.0 || d1 >= 1e10 || d2 >= 1e10 {
            return Err(DOMAIN);
        }
        if p == 1.0 {
            return num(0.0);
        }
        shape(d1)?;
        shape(d2)?;
        num(invert_increasing(|x| -f_rt(x, d1, d2), -p))
    }
    r.eager("F.INV.RT", 3, 3, f_rt_inv);
    r.eager_in(Ns::Legacy, "FINV", 3, 3, f_rt_inv);
    for name in ["F.TEST", "FTEST"] {
        r.eager(name, 2, 2, |e, a| {
            let (x, y) = (numbers_of(e, &a[0])?, numbers_of(e, &a[1])?);
            if x.len() < 2 || y.len() < 2 {
                return Err(ErrorKind::Value);
            }
            let (d1, d2) = (x.len() as f64 - 1.0, y.len() as f64 - 1.0);
            let (v1, v2) = (ss(&x) / d1, ss(&y) / d2);
            if v1 == 0.0 || v2 == 0.0 {
                return Err(ErrorKind::Value);
            }
            let f = v1 / v2;
            num((2.0 * f_cdf(f, d1, d2).min(f_rt(f, d1, d2))).min(1.0))
        });
    }

    // ----- beta -----
    fn beta_scale(e: &mut Engine, a: &[Value], i: usize) -> Res<(f64, f64)> {
        let lo = super::opt_num(e, a, i, 0.0)?;
        let hi = super::opt_num(e, a, i + 1, 1.0)?;
        Ok((lo, hi))
    }
    r.eager("BETA.DIST", 4, 6, |e, a| {
        let (x, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        let (lo, hi) = beta_scale(e, a, 4)?;
        if al <= 0.0 || be <= 0.0 || x < lo || x > hi || lo >= hi {
            return Err(DOMAIN);
        }
        let u = (x - lo) / (hi - lo);
        if cumulative(e, a, 3)? {
            shape(al)?;
            shape(be)?;
            num(beta_reg(u, al, be))
        } else {
            if (u == 0.0 && al < 1.0) || (u == 1.0 && be < 1.0) {
                return Err(DOMAIN);
            }
            num(u.powf(al - 1.0) * (1.0 - u).powf(be - 1.0) / ln_beta(al, be).exp() / (hi - lo))
        }
    });
    r.eager("BETADIST", 3, 5, |e, a| {
        let (x, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        let (lo, hi) = beta_scale(e, a, 3)?;
        if al <= 0.0 || be <= 0.0 || x < lo || x > hi || lo >= hi {
            return Err(DOMAIN);
        }
        shape(al)?;
        shape(be)?;
        num(beta_reg((x - lo) / (hi - lo), al, be))
    });
    for name in ["BETA.INV", "BETAINV"] {
        r.eager(name, 3, 5, |e, a| {
            let (p, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            let (lo, hi) = beta_scale(e, a, 3)?;
            if !(0.0..=1.0).contains(&p) || al <= 0.0 || be <= 0.0 || lo >= hi {
                return Err(DOMAIN);
            }
            shape(al)?;
            shape(be)?;
            let u = if p == 0.0 {
                0.0
            } else if p == 1.0 {
                1.0
            } else {
                invert_on(|x| beta_reg(x, al, be), p, 0.0, 1.0)
            };
            num(lo + (hi - lo) * u)
        });
    }

    // ----- log-normal and Weibull -----
    r.eager("LOGNORM.DIST", 4, 4, |e, a| {
        let (x, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if s <= 0.0 {
            return Err(DOMAIN);
        }
        let cum = cumulative(e, a, 3)?;
        if x <= 0.0 {
            return if cum { num(0.0) } else { Err(DOMAIN) };
        }
        let z = (x.ln() - m) / s;
        if cum {
            num(norm_cdf(z))
        } else {
            num((-z * z / 2.0).exp() / (x * s * (2.0 * PI).sqrt()))
        }
    });
    r.eager("LOGNORMDIST", 3, 3, |e, a| {
        let (x, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
        if s <= 0.0 {
            return Err(DOMAIN);
        }
        num(if x <= 0.0 {
            0.0
        } else {
            norm_cdf((x.ln() - m) / s)
        })
    });
    for name in ["LOGNORM.INV", "LOGINV"] {
        r.eager(name, 3, 3, |e, a| {
            let (p, m, s) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if p <= 0.0 || p >= 1.0 || s <= 0.0 {
                return Err(DOMAIN);
            }
            num((m + s * norm_inv(p)).exp())
        });
    }
    for name in ["WEIBULL", "WEIBULL.DIST"] {
        r.eager(name, 4, 4, |e, a| {
            let (x, al, be) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
            if x < 0.0 || al <= 0.0 || be <= 0.0 {
                return Err(DOMAIN);
            }
            if cumulative(e, a, 3)? {
                num(1.0 - (-(x / be).powf(al)).exp())
            } else {
                num(al / be.powf(al) * x.powf(al - 1.0) * (-(x / be).powf(al)).exp())
            }
        });
    }

    // ----- discrete -----
    r.eager("HYPGEOM.DIST", 5, 5, |e, a| {
        let (s, n, m, big_n) = hyper_args(e, a)?;
        if cumulative(e, a, 4)? {
            let lo = (n + m - big_n).max(0.0);
            let mut sum = 0.0;
            let mut i = lo;
            while i <= s {
                sum += hyper_pmf(i, n, m, big_n);
                i += 1.0;
            }
            num(sum.min(1.0))
        } else {
            num(hyper_pmf(s, n, m, big_n))
        }
    });
    r.eager("HYPGEOMDIST", 4, 4, |e, a| {
        let (s, n, m, big_n) = hyper_args(e, a)?;
        num(hyper_pmf(s, n, m, big_n))
    });
    r.eager("NEGBINOM.DIST", 4, 4, |e, a| {
        let (f, s, p) = (whole(e, a, 0)?, whole(e, a, 1)?, e.num(&a[2])?);
        if f < 0.0 || s < 0.0 || !(0.0..=1.0).contains(&p) || f + s - 1.0 <= 0.0 {
            return Err(DOMAIN);
        }
        if cumulative(e, a, 3)? {
            if s < 1.0 {
                return Err(DOMAIN);
            }
            shape(s)?;
            shape(f + 1.0)?;
            num(beta_reg(p, s, f + 1.0))
        } else {
            num(negbinom_pmf(f, s, p))
        }
    });
    r.eager("NEGBINOMDIST", 3, 3, |e, a| {
        let (f, s, p) = (whole(e, a, 0)?, whole(e, a, 1)?, e.num(&a[2])?);
        if f < 0.0 || s < 0.0 || !(0.0..=1.0).contains(&p) || f + s - 1.0 <= 0.0 {
            return Err(DOMAIN);
        }
        num(negbinom_pmf(f, s, p))
    });
    for name in ["BINOM.INV", "CRITBINOM"] {
        r.eager(name, 3, 3, |e, a| {
            let (n, p, alpha) = (whole(e, a, 0)?, e.num(&a[1])?, e.num(&a[2])?);
            if n < 0.0 || !(0.0..=1.0).contains(&p) || !(0.0..=1.0).contains(&alpha) {
                return Err(DOMAIN);
            }
            if alpha == 0.0 {
                return num(0.0);
            }
            let (mut sum, mut k) = (0.0, 0.0);
            while k < n {
                sum += binom_pmf(n, k, p);
                if sum >= alpha {
                    break;
                }
                k += 1.0;
            }
            num(k)
        });
    }
    for name in ["B", "BINOM.DIST.RANGE"] {
        r.eager(name, 3, 4, |e, a| {
            let (n, p, k1) = (whole(e, a, 0)?, e.num(&a[1])?, whole(e, a, 2)?);
            let k2 = if a.len() > 3 { whole(e, a, 3)? } else { k1 };
            if n < 0.0 || !(0.0..=1.0).contains(&p) || k1 < 0.0 || k1 > k2 || k2 > n {
                return Err(DOMAIN);
            }
            let mut sum = 0.0;
            let mut k = k1;
            while k <= k2 {
                sum += binom_pmf(n, k, p);
                k += 1.0;
            }
            num(sum.min(1.0))
        });
    }

    // ----- intervals and tests -----
    for name in ["CONFIDENCE", "CONFIDENCE.NORM"] {
        r.eager(name, 3, 3, |e, a| {
            let (alpha, sd, n) = (e.num(&a[0])?, e.num(&a[1])?, whole(e, a, 2)?);
            if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || n < 1.0 {
                return Err(DOMAIN);
            }
            num(norm_inv(1.0 - alpha / 2.0) * sd / n.sqrt())
        });
    }
    r.eager("CONFIDENCE.T", 3, 3, |e, a| {
        let (alpha, sd, n) = (e.num(&a[0])?, e.num(&a[1])?, whole(e, a, 2)?);
        if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || n < 1.0 {
            return Err(DOMAIN);
        }
        if n == 1.0 {
            return Err(ErrorKind::DivZero);
        }
        shape(n - 1.0)?;
        num(t_two_inv(alpha, n - 1.0) * sd / n.sqrt())
    });
    for name in ["Z.TEST", "ZTEST"] {
        r.eager(name, 2, 3, |e, a| {
            let x = numbers_of(e, &a[0])?;
            let mu = e.num(&a[1])?;
            let n = x.len() as f64;
            if n < 2.0 {
                return Err(ErrorKind::DivZero);
            }
            let sigma = if a.len() > 2 {
                let s = e.num(&a[2])?;
                if s <= 0.0 {
                    return Err(DOMAIN);
                }
                s
            } else {
                let sd = (ss(&x) / (n - 1.0)).sqrt();
                if sd == 0.0 {
                    return Err(ErrorKind::DivZero);
                }
                sd
            };
            num(norm_cdf(-(mean(&x) - mu) * n.sqrt() / sigma))
        });
    }
    r.eager("PROB", 3, 4, |e, a| {
        let (r1, c1, xs) = grid(e, &a[0])?;
        let (r2, c2, ps) = grid(e, &a[1])?;
        if (r1, c1) != (r2, c2) {
            return Err(ErrorKind::Na);
        }
        let lo = e.num(&a[2])?;
        let hi = if a.len() > 3 { e.num(&a[3])? } else { lo };
        let (lo, hi) = if lo > hi { (hi, lo) } else { (lo, hi) };
        let (mut total, mut sum) = (0.0, 0.0);
        for (x, p) in xs.iter().zip(&ps) {
            let (x, p) = match (x, p) {
                (Value::Number(x), Value::Number(p)) => (*x, *p),
                (Value::Error(k), _) | (_, Value::Error(k)) => return Err(*k),
                _ => return Err(ErrorKind::Value),
            };
            if !(0.0..=1.0).contains(&p) {
                return Err(ErrorKind::Value);
            }
            total += p;
            if x >= lo && x <= hi {
                sum += p;
            }
        }
        if (total - 1.0).abs() > 1e-9 {
            return Err(ErrorKind::Value);
        }
        num(sum)
    });
}
