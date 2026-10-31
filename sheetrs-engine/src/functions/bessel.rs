//! `BESSELI`, `BESSELJ`, `BESSELK` and `BESSELY`, with the algorithms of LibreOffice's analysis
//! add-in (`scaddins/source/analysis/bessel.cxx`): series and adjoint summation for the first
//! kinds, polynomial approximations for the second, so that the digits are LibreOffice's.
//! The result is checked against the oracle (`tests/assets/semantics.*`).

use std::f64::consts::{EULER_GAMMA as EGAMMA, FRAC_2_PI, FRAC_PI_2, FRAC_PI_4, PI};

use super::{Reg, num};
use crate::engine::Res;
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;
/// The iteration did not converge (`Err:523`, which the files store as `#VALUE!`).
const NO_CONVERGENCE: ErrorKind = ErrorKind::Value;
const TWO_64: f64 = 18446744073709551616.0;

fn bessel_j(x: f64, n: i32) -> Res<f64> {
    if n < 0 {
        return Err(DOMAIN);
    }
    if x == 0.0 {
        return Ok(if n == 0 { 1.0 } else { 0.0 });
    }
    // An odd order makes the function odd in `x`.
    let sign = if n % 2 == 1 && x < 0.0 { -1.0 } else { 1.0 };
    let fx = x.abs();
    const MAX_ITERATION: f64 = 9_000_000.0;
    let estimate = fx * 1.5 + f64::from(n);
    if estimate > MAX_ITERATION {
        if fx.powf(0.4) > f64::from(n) {
            return Ok(sign
                * (FRAC_2_PI / fx).sqrt()
                * (fx - f64::from(n) * FRAC_PI_2 - FRAC_PI_4).cos());
        }
        return Err(NO_CONVERGENCE);
    }
    let epsilon = 1.0e-15;
    let mut u;
    let mut g = 0.0;
    let mut delta_u = 0.0;
    let mut f_bar = -1.0;
    let mut k;
    if n == 0 {
        u = 1.0;
        let g_bar_delta_u = 0.0;
        let g_bar = -2.0 / fx;
        delta_u = g_bar_delta_u / g_bar;
        u += delta_u;
        g = -1.0 / g_bar;
        f_bar *= g;
        k = 2.0;
    } else {
        u = 0.0;
        k = 1.0;
        while k <= f64::from(n - 1) {
            let m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
            let g_bar_delta_u = -g * delta_u - m_bar * u;
            let g_bar = m_bar - 2.0 * k / fx + g;
            delta_u = g_bar_delta_u / g_bar;
            u += delta_u;
            g = -1.0 / g_bar;
            f_bar *= g;
            k += 1.0;
        }
        let m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
        let g_bar_delta_u = f_bar - g * delta_u - m_bar * u;
        let g_bar = m_bar - 2.0 * k / fx + g;
        delta_u = g_bar_delta_u / g_bar;
        u += delta_u;
        g = -1.0 / g_bar;
        f_bar *= g;
        k += 1.0;
    }
    loop {
        let m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
        let g_bar_delta_u = -g * delta_u - m_bar * u;
        let g_bar = m_bar - 2.0 * k / fx + g;
        delta_u = g_bar_delta_u / g_bar;
        u += delta_u;
        g = -1.0 / g_bar;
        f_bar *= g;
        let found = delta_u.abs() <= u.abs() * epsilon;
        k += 1.0;
        if found {
            return Ok(u * sign);
        }
        if k > MAX_ITERATION {
            return Err(NO_CONVERGENCE);
        }
    }
}

fn bessel_i(x: f64, n: i32) -> Res<f64> {
    const MAX_ITERATION: i32 = 2000;
    let xhalf = x / 2.0;
    if n < 0 {
        return Err(DOMAIN);
    }
    let mut term = 1.0;
    for k in 1..=n {
        term = term / f64::from(k) * xhalf;
        if term == 0.0 {
            break;
        }
    }
    let mut result = term;
    if term != 0.0 {
        let mut k = 1;
        loop {
            term = term * xhalf / f64::from(k) * xhalf / (f64::from(k) + f64::from(n));
            result += term;
            k += 1;
            if !(term.abs() > result.abs() * 1.0e-15 && k < MAX_ITERATION) {
                break;
            }
        }
    }
    Ok(result)
}

/// The polynomial keeps the add-in's coefficients, Euler's constant truncated to 8 decimals
/// included.
fn bessel_k0(x: f64) -> Res<f64> {
    Ok(if x <= 2.0 {
        let x2 = x * 0.5;
        let y = x2 * x2;
        let euler = (std::f64::consts::EULER_GAMMA * 1e8).trunc() / 1e8;
        -x2.ln() * bessel_i(x, 0)?
            + (-euler
                + y * (0.42278420
                    + y * (0.23069756
                        + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))))
    } else {
        let y = 2.0 / x;
        (-x).exp() / x.sqrt()
            * (1.25331414
                + y * (-0.7832358e-1
                    + y * (0.2189568e-1
                        + y * (-0.1062446e-1
                            + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))))
    })
}

fn bessel_k1(x: f64) -> Res<f64> {
    Ok(if x <= 2.0 {
        let x2 = x * 0.5;
        let y = x2 * x2;
        x2.ln() * bessel_i(x, 1)?
            + (1.0
                + y * (0.15443144
                    + y * (-0.67278579
                        + y * (-0.18156897
                            + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4))))))
                / x
    } else {
        let y = 2.0 / x;
        (-x).exp() / x.sqrt()
            * (1.25331414
                + y * (0.23498619
                    + y * (-0.3655620e-1
                        + y * (0.1504268e-1
                            + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))))
    })
}

fn bessel_k(x: f64, order: i32) -> Res<f64> {
    match order {
        0 => bessel_k0(x),
        1 => bessel_k1(x),
        _ => {
            let tox = 2.0 / x;
            let mut bkm = bessel_k0(x)?;
            let mut bk = bessel_k1(x)?;
            for n in 1..order {
                let bkp = bkm + f64::from(n) * tox * bk;
                bkm = bk;
                bk = bkp;
            }
            Ok(bk)
        }
    }
}

fn bessel_y0(x: f64) -> Res<f64> {
    if x <= 0.0 || x.abs() >= TWO_64 {
        return Err(DOMAIN);
    }
    const MAX_ITERATION: f64 = 9_000_000.0;
    if x > 5.0e+6 {
        return Ok((1.0 / PI / x).sqrt() * (x.sin() - x.cos()));
    }
    let epsilon = 1.0e-15;
    let mut alpha = (x / 2.0).ln() + EGAMMA;
    let mut u = alpha;
    let mut k = 1.0;
    let mut g_bar_delta_u = 0.0;
    let mut g_bar = -2.0 / x;
    let mut delta_u = g_bar_delta_u / g_bar;
    let mut g = -1.0 / g_bar;
    let mut f_bar = -g;
    let mut sign_alpha = 1.0;
    k += 1.0;
    loop {
        let km1mod2 = (k - 1.0) % 2.0;
        let m_bar = (2.0 * km1mod2) * f_bar;
        if km1mod2 == 0.0 {
            alpha = 0.0;
        } else {
            alpha = sign_alpha * (4.0 / k);
            sign_alpha = -sign_alpha;
        }
        g_bar_delta_u = f_bar * alpha - g * delta_u - m_bar * u;
        g_bar = m_bar - (2.0 * k) / x + g;
        delta_u = g_bar_delta_u / g_bar;
        u += delta_u;
        g = -1.0 / g_bar;
        f_bar *= g;
        let found = delta_u.abs() <= u.abs() * epsilon;
        k += 1.0;
        if found {
            return Ok(u * FRAC_2_PI);
        }
        if k >= MAX_ITERATION {
            return Err(NO_CONVERGENCE);
        }
    }
}

fn bessel_y1(x: f64) -> Res<f64> {
    if x <= 0.0 || x.abs() >= TWO_64 {
        return Err(DOMAIN);
    }
    const MAX_ITERATION: f64 = 9_000_000.0;
    if x > 5.0e+6 {
        return Ok(-(1.0 / PI / x).sqrt() * (x.sin() + x.cos()));
    }
    let epsilon = 1.0e-15;
    let mut alpha = 1.0 / x;
    let mut f_bar = -1.0;
    let mut u = alpha;
    let mut k = 1.0;
    alpha = 1.0 - EGAMMA - (x / 2.0).ln();
    let mut g_bar_delta_u = -alpha;
    let mut g_bar = -2.0 / x;
    let mut delta_u = g_bar_delta_u / g_bar;
    u += delta_u;
    let mut g = -1.0 / g_bar;
    f_bar *= g;
    let mut sign_alpha = -1.0;
    k += 1.0;
    loop {
        let km1mod2 = (k - 1.0) % 2.0;
        let m_bar = (2.0 * km1mod2) * f_bar;
        let q = (k - 1.0) / 2.0;
        if km1mod2 == 0.0 {
            alpha = sign_alpha * (1.0 / q + 1.0 / (q + 1.0));
            sign_alpha = -sign_alpha;
        } else {
            alpha = 0.0;
        }
        g_bar_delta_u = f_bar * alpha - g * delta_u - m_bar * u;
        g_bar = m_bar - (2.0 * k) / x + g;
        delta_u = g_bar_delta_u / g_bar;
        u += delta_u;
        g = -1.0 / g_bar;
        f_bar *= g;
        let found = delta_u.abs() <= u.abs() * epsilon;
        k += 1.0;
        if found {
            return Ok(-u * 2.0 / PI);
        }
        if k >= MAX_ITERATION {
            return Err(NO_CONVERGENCE);
        }
    }
}

fn bessel_y(x: f64, order: i32) -> Res<f64> {
    match order {
        0 => bessel_y0(x),
        1 => bessel_y1(x),
        _ => {
            let tox = 2.0 / x;
            let mut bym = bessel_y0(x)?;
            let mut by = bessel_y1(x)?;
            for n in 1..order {
                let byp = f64::from(n) * tox * by - bym;
                bym = by;
                by = byp;
            }
            Ok(by)
        }
    }
}

/// A result the add-in accepts: finite.
fn finite(x: f64) -> Res<Value> {
    if x.is_finite() { num(x) } else { Err(DOMAIN) }
}

pub(super) fn register(r: &mut Reg) {
    macro_rules! bessel {
        ($name:literal, $f:expr, $positive:expr) => {
            r.eager($name, 2, 2, |e, a| {
                let x = e.num(&a[0])?;
                let n = e.num(&a[1])?.trunc();
                if n.abs() > f64::from(i32::MAX) {
                    return Err(DOMAIN);
                }
                let n = n as i32;
                let f: fn(f64, i32) -> Res<f64> = $f;
                if $positive && (n < 0 || x <= 0.0) {
                    return Err(DOMAIN);
                }
                finite(f(x, n)?)
            });
        };
    }
    bessel!("BESSELI", bessel_i, false);
    bessel!("BESSELJ", bessel_j, false);
    bessel!("BESSELK", bessel_k, true);
    bessel!("BESSELY", bessel_y, true);
}
