//! Special functions the distributions are made of: the gamma function, the regularized
//! incomplete gamma and beta functions, and the inverse of an increasing function.
//!
//! They are accurate to about 1e-14, which is what the oracle (1e-9) asks of them; the iterations
//! stop on the size of the last step, not on a fixed count.

use std::f64::consts::PI;

/// `ln Γ(x)` for `x > 0` (Lanczos, g = 7, nine terms).
pub(crate) fn ln_gamma(x: f64) -> f64 {
    const C: [f64; 9] = [
        0.999_999_999_999_809_9,
        676.520_368_121_885_1,
        -1_259.139_216_722_402_8,
        771.323_428_777_653_1,
        -176.615_029_162_140_6,
        12.507_343_278_686_905,
        -0.138_571_095_265_720_12,
        9.984_369_578_019_572e-6,
        1.505_632_735_149_311_6e-7,
    ];
    if x < 0.5 {
        // Reflection: Γ(x) Γ(1 - x) = π / sin(πx).
        return (PI / (PI * x).sin().abs()).ln() - ln_gamma(1.0 - x);
    }
    let x = x - 1.0;
    let t = x + 7.5;
    let mut a = C[0];
    for (i, c) in C.iter().enumerate().skip(1) {
        a += c / (x + i as f64);
    }
    0.5 * (2.0 * PI).ln() + (x + 0.5) * t.ln() - t + a.ln()
}

/// `Γ(x)`, or `None` at its poles (zero and the negative integers). Overflows to infinity.
pub(crate) fn gamma(x: f64) -> Option<f64> {
    if x == x.floor() {
        if x <= 0.0 {
            return None;
        }
        if x <= 171.0 {
            return Some((2..x as u64).map(|k| k as f64).product());
        }
        return Some(f64::INFINITY);
    }
    if x < 0.0 {
        // Γ(x) = π / (sin(πx) Γ(1 - x))
        let g = gamma(1.0 - x)?;
        return Some(PI / ((PI * x).sin() * g));
    }
    Some(ln_gamma(x).exp())
}

/// `ln B(a, b)`. For large arguments the difference of three huge logarithms would lose the
/// digits the incomplete beta function needs, so it is written as the Stirling series of the
/// combination, which has no large terms to cancel.
pub(crate) fn ln_beta(a: f64, b: f64) -> f64 {
    if a.min(b) < 10.0 {
        return ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b);
    }
    let s = a + b;
    let corr = |z: f64| {
        let z2 = z * z;
        1.0 / (12.0 * z) - 1.0 / (360.0 * z * z2) + 1.0 / (1260.0 * z * z2 * z2)
            - 1.0 / (1680.0 * z * z2 * z2 * z2)
    };
    0.5 * (2.0 * PI).ln() + a * (-b / s).ln_1p() + b * (-a / s).ln_1p() - 0.5 * (a * b / s).ln()
        + corr(a)
        + corr(b)
        - corr(s)
}

const EPS: f64 = 1.110_223_024_625_156_5e-16;
const MAX_ITER: u32 = 1_000_000;
/// LibreOffice gives up on the series and the continued fraction of the incomplete gamma function
/// after this many terms (`Err:523`, which the files store as `#VALUE!`).
const GAMMA_LIMIT: u32 = 10_000;
const TINY: f64 = 1e-300;

/// The regularized lower incomplete gamma function `P(a, x)`; `NaN` where LibreOffice does not
/// converge (`a` large and `x` near it).
pub(crate) fn gamma_p(a: f64, x: f64) -> f64 {
    if x <= 0.0 {
        0.0
    } else if x < a + 1.0 {
        gamma_series(a, x)
    } else {
        1.0 - gamma_fraction(a, x)
    }
}

/// The regularized upper incomplete gamma function `Q(a, x) = 1 - P(a, x)`; `NaN` as `gamma_p`.
pub(crate) fn gamma_q(a: f64, x: f64) -> f64 {
    if x <= 0.0 {
        1.0
    } else if x < a + 1.0 {
        1.0 - gamma_series(a, x)
    } else {
        gamma_fraction(a, x)
    }
}

fn gamma_series(a: f64, x: f64) -> f64 {
    let mut ap = a;
    let mut term = 1.0 / a;
    let mut sum = term;
    let mut converged = false;
    for _ in 0..GAMMA_LIMIT {
        ap += 1.0;
        term *= x / ap;
        sum += term;
        if term.abs() <= sum.abs() * EPS {
            converged = true;
            break;
        }
    }
    if !converged {
        return f64::NAN;
    }
    sum * (-x + a * x.ln() - ln_gamma(a)).exp()
}

fn gamma_fraction(a: f64, x: f64) -> f64 {
    let mut b = x + 1.0 - a;
    let mut c = 1.0 / TINY;
    let mut d = 1.0 / b;
    let mut h = d;
    let mut converged = false;
    for i in 1..GAMMA_LIMIT {
        let an = -f64::from(i) * (f64::from(i) - a);
        b += 2.0;
        d = an * d + b;
        if d.abs() < TINY {
            d = TINY;
        }
        c = b + an / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() <= EPS {
            converged = true;
            break;
        }
    }
    if !converged {
        return f64::NAN;
    }
    h * (-x + a * x.ln() - ln_gamma(a)).exp()
}

/// The regularized incomplete beta function `I_x(a, b)`.
pub(crate) fn beta_reg(x: f64, a: f64, b: f64) -> f64 {
    beta_reg_xy(x, 1.0 - x, a, b)
}

/// `I_x(a, b)` for `x` and `y = 1 - x` that the caller has computed without the cancellation of
/// subtracting (`x` near 1 is where the degrees of freedom are large).
pub(crate) fn beta_reg_xy(x: f64, y: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if y <= 0.0 {
        return 1.0;
    }
    let front = (a * x.ln() + b * y.ln() - ln_beta(a, b)).exp();
    if x < (a + 1.0) / (a + b + 2.0) {
        front * beta_fraction(x, a, b) / a
    } else {
        1.0 - front * beta_fraction(y, b, a) / b
    }
}

fn beta_fraction(x: f64, a: f64, b: f64) -> f64 {
    let (qab, qap, qam) = (a + b, a + 1.0, a - 1.0);
    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < TINY {
        d = TINY;
    }
    d = 1.0 / d;
    let mut h = d;
    for m in 1..MAX_ITER {
        let m = f64::from(m);
        let m2 = 2.0 * m;
        let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < TINY {
            d = TINY;
        }
        c = 1.0 + aa / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        h *= d * c;
        let aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < TINY {
            d = TINY;
        }
        c = 1.0 + aa / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() < EPS {
            break;
        }
    }
    h
}

/// The `x >= 0` where the increasing `f` reaches `target` (`f` is `0` at zero and tends to
/// something above `target`).
pub(crate) fn invert_increasing(f: impl Fn(f64) -> f64, target: f64) -> f64 {
    let mut hi = 1.0;
    let mut doublings = 0;
    while f(hi) < target {
        hi *= 2.0;
        doublings += 1;
        if doublings > 1100 {
            return hi;
        }
    }
    let mut lo = if hi > 1.0 { hi / 2.0 } else { 0.0 };
    for _ in 0..3000 {
        let mid = 0.5 * (lo + hi);
        if mid == lo || mid == hi {
            break;
        }
        if f(mid) < target {
            lo = mid;
        } else {
            hi = mid;
        }
        if hi - lo <= 1e-16 * hi {
            break;
        }
    }
    0.5 * (lo + hi)
}

/// The `x` in `[lo, hi]` where the increasing `f` reaches `target`.
pub(crate) fn invert_on(f: impl Fn(f64) -> f64, target: f64, mut lo: f64, mut hi: f64) -> f64 {
    for _ in 0..200 {
        let mid = 0.5 * (lo + hi);
        if f(mid) < target {
            lo = mid;
        } else {
            hi = mid;
        }
        if hi - lo <= 1e-16 {
            break;
        }
    }
    0.5 * (lo + hi)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn close(a: f64, b: f64) -> bool {
        (a - b).abs() <= 1e-12 * a.abs().max(b.abs()).max(1.0)
    }

    #[test]
    fn gamma_and_its_logarithm() {
        assert!(close(ln_gamma(5.0), 24f64.ln()));
        assert!(close(ln_gamma(0.5), PI.sqrt().ln()));
        assert_eq!(gamma(5.0), Some(24.0));
        assert!(close(gamma(-1.5).unwrap(), 4.0 * PI.sqrt() / 3.0));
        assert_eq!(gamma(0.0), None);
        assert_eq!(gamma(-2.0), None);
        assert!(gamma(172.0).unwrap().is_infinite());
    }

    #[test]
    fn incomplete_functions() {
        // Chi-square with 4 degrees of freedom: P(2, 1.5) = 1 - e^-1.5 (1 + 1.5).
        assert!(close(gamma_p(2.0, 1.5), 1.0 - (-1.5f64).exp() * 2.5));
        assert!(close(gamma_p(3.0, 2.0) + gamma_q(3.0, 2.0), 1.0));
        // I_x(2, 3) = 0.5248 at 0.4 (a closed form).
        assert!(close(beta_reg(0.4, 2.0, 3.0), 0.5248));
        assert!(close(beta_reg(0.5, 1.5, 1.5), 0.5));
    }

    #[test]
    fn beta_function_of_large_arguments() {
        // B(a, b) by the gamma function where it does not cancel, and by the series where it does.
        let direct = ln_gamma(12.0) + ln_gamma(15.0) - ln_gamma(27.0);
        assert!(close(ln_beta(12.0, 15.0), direct));
        // The symmetric half: I_0.5(a, a) = 0.5.
        assert!((beta_reg(0.5, 1e6, 1e6) - 0.5).abs() < 1e-11);
    }

    #[test]
    fn inverse() {
        let x = invert_increasing(|x| gamma_p(3.0, x), 0.5);
        assert!(close(gamma_p(3.0, x), 0.5));
        let y = invert_on(|x| beta_reg(x, 2.0, 3.0), 0.5248, 0.0, 1.0);
        assert!(close(y, 0.4));
    }
}
