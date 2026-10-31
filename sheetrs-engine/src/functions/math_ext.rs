//! More arithmetic: number theory, multiples, inverse and hyperbolic
//! trigonometry, combinatorics, bit operations and the error function.

use super::{Reg, VAR, num, opt_num};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

/// A quotient that is an integer up to the noise of the division (`0.9/0.3`).
fn snap(q: f64) -> f64 {
    let r = q.round();
    if (q - r).abs() <= 1e-12 * q.abs().max(1.0) {
        r
    } else {
        q
    }
}

fn one(e: &mut Engine, a: &[Value], f: fn(f64) -> Option<f64>) -> Res<Value> {
    let x = e.num(&a[0])?;
    f(x).map_or(Err(ErrorKind::Domain), num)
}

fn trunc_all(e: &mut Engine, a: &[Value]) -> Res<Vec<f64>> {
    let v = e.numbers(a)?;
    if v.iter().any(|x| *x < 0.0) {
        return Err(ErrorKind::Domain);
    }
    Ok(v.into_iter().map(f64::trunc).collect())
}

fn gcd(mut a: f64, mut b: f64) -> f64 {
    while b != 0.0 {
        (a, b) = (b, a % b);
    }
    a
}

/// `n` choose `k` for non-negative integers, exact while it fits a double.
pub(crate) fn combin(n: f64, k: f64) -> Option<f64> {
    if n < 0.0 || k < 0.0 || k > n {
        return None;
    }
    let k = k.min(n - k);
    let mut r = 1.0_f64;
    let mut i = 1.0;
    while i <= k {
        r = (r * (n - k + i) / i).round();
        i += 1.0;
    }
    r.is_finite().then_some(r)
}

fn fact(n: f64) -> Option<f64> {
    (0.0..=170.0)
        .contains(&n)
        .then(|| (1..=n as u32).map(f64::from).product())
}

/// `ROUNDSIG`: `x` rounded to `n` significant digits.
fn round_significant(x: f64, n: f64) -> Res<Value> {
    if n < 1.0 {
        return Err(ErrorKind::Domain);
    }
    if x == 0.0 || !x.is_finite() {
        return num(x);
    }
    let exp: i32 = format!("{:.14e}", x.abs())
        .split_once('e')
        .and_then(|(_, e)| e.parse().ok())
        .unwrap_or(0);
    let digits = n.min(1e6) as i32 - 1 - exp;
    num(super::math::round_dec(
        x,
        digits,
        super::math::Mode::Nearest,
    ))
}

/// `SERIESSUM`: the sum of `coef * x^(n + i*m)` over the coefficients in the order they are read.
fn series_sum(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let (x, n, m) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
    let coefs: Vec<Value> = match &a[3] {
        Value::Ref(_) | Value::Array(_) => e.array_of(a[3].clone()).data.clone(),
        v => vec![Value::Number(e.num(v)?)],
    };
    let mut terms = Vec::with_capacity(coefs.len());
    for c in &coefs {
        terms.push(match c {
            Value::Number(v) => *v,
            Value::Bool(b) => f64::from(u8::from(*b)),
            Value::Error(_) => return Err(ErrorKind::Domain),
            _ => 0.0,
        });
    }
    // Nothing of the series is left at zero, except its first power.
    if x == 0.0 {
        return if n == 0.0 {
            Err(ErrorKind::Value)
        } else {
            num(0.0)
        };
    }
    let mut sum = 0.0;
    for (i, c) in terms.iter().enumerate() {
        let p = x.powf(n + i as f64 * m);
        if p.is_nan() {
            return Err(ErrorKind::Domain);
        }
        sum += c * p;
    }
    num(sum)
}

pub(super) fn register(r: &mut Reg) {
    r.eager("ROUNDSIG", 2, 2, |e, a| {
        let (x, n) = (e.num(&a[0])?, e.num(&a[1])?);
        round_significant(x, n.floor())
    });
    r.eager("SERIESSUM", 4, 4, series_sum).array_args();
    // A subtraction without the snap to zero of the operator: `0.3 - 0.1 - 0.2` is not 0.
    r.eager("RAWSUBTRACT", 2, VAR, |e, a| {
        let mut acc = e.num(&a[0])?;
        for v in &a[1..] {
            acc -= e.num(v)?;
        }
        num(acc)
    });
    r.eager("ASIN", 1, 1, |e, a| {
        one(e, a, |x| (-1.0..=1.0).contains(&x).then(|| x.asin()))
    });
    r.eager("ACOS", 1, 1, |e, a| {
        one(e, a, |x| (-1.0..=1.0).contains(&x).then(|| x.acos()))
    });
    r.eager("SINH", 1, 1, |e, a| one(e, a, |x| Some(x.sinh())));
    r.eager("COSH", 1, 1, |e, a| one(e, a, |x| Some(x.cosh())));
    r.eager("TANH", 1, 1, |e, a| one(e, a, |x| Some(x.tanh())));
    r.eager("ASINH", 1, 1, |e, a| one(e, a, |x| Some(x.asinh())));
    r.eager("ACOSH", 1, 1, |e, a| {
        one(e, a, |x| (x >= 1.0).then(|| x.acosh()))
    });
    r.eager("ATANH", 1, 1, |e, a| {
        one(e, a, |x| (x.abs() < 1.0).then(|| x.atanh()))
    });
    r.eager("COT", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        if x == 0.0 {
            Err(ErrorKind::DivZero)
        } else {
            num(1.0 / x.tan())
        }
    });
    r.eager("COTH", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        if x == 0.0 {
            Err(ErrorKind::DivZero)
        } else {
            num(1.0 / x.tanh())
        }
    });
    r.eager("ACOT", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        num(std::f64::consts::FRAC_PI_2 - x.atan())
    });
    r.eager("ACOTH", 1, 1, |e, a| {
        one(e, a, |x| (x.abs() > 1.0).then(|| (1.0 / x).atanh()))
    });
    r.eager("SEC", 1, 1, |e, a| one(e, a, |x| Some(1.0 / x.cos())));
    r.eager("CSC", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        if x == 0.0 {
            Err(ErrorKind::DivZero)
        } else {
            num(1.0 / x.sin())
        }
    });
    r.eager("SECH", 1, 1, |e, a| one(e, a, |x| Some(1.0 / x.cosh())));
    r.eager("CSCH", 1, 1, |e, a| {
        let x = e.num(&a[0])?;
        if x == 0.0 {
            Err(ErrorKind::DivZero)
        } else {
            num(1.0 / x.sinh())
        }
    });
    r.eager("SQRTPI", 1, 1, |e, a| {
        one(e, a, |x| {
            (x >= 0.0).then(|| (x * std::f64::consts::PI).sqrt())
        })
    });

    r.eager("GCD", 1, VAR, |e, a| {
        let v = trunc_all(e, a)?;
        let g = v.iter().fold(0.0, |g, x| gcd(g, *x));
        if g >= 9_007_199_254_740_992.0 {
            Err(ErrorKind::Domain)
        } else {
            num(g)
        }
    });
    r.eager("LCM", 1, VAR, |e, a| {
        let v = trunc_all(e, a)?;
        let mut l = 1.0_f64;
        for x in v {
            if x == 0.0 {
                return num(0.0);
            }
            l = l / gcd(l, x) * x;
        }
        if l >= 9_007_199_254_740_992.0 {
            Err(ErrorKind::Domain)
        } else {
            num(l)
        }
    });
    r.eager("QUOTIENT", 2, 2, |e, a| {
        let (x, y) = (e.num(&a[0])?, e.num(&a[1])?);
        if y == 0.0 {
            // Err:502 (invalid argument), not a division error.
            Err(ErrorKind::Domain)
        } else {
            num((x / y).trunc())
        }
    });
    r.eager("MROUND", 2, 2, |e, a| {
        let (x, m) = (e.num(&a[0])?, e.num(&a[1])?);
        if m == 0.0 || x == 0.0 {
            return num(0.0);
        }
        // Mixed signs round the quotient anyway.
        num(snap(x / m).round() * m)
    });

    // Multiples: Excel's 2010 family. CEILING.MATH / FLOOR.MATH round toward
    // +/- infinity (or away from zero with a non-zero mode for negatives);
    // the PRECISE ones ignore the sign of the multiple.
    r.eager("CEILING.MATH", 1, 3, |e, a| {
        let x = e.num(&a[0])?;
        let s = opt_num(e, a, 1, 1.0)?.abs();
        let mode = opt_num(e, a, 2, 0.0)?;
        if s == 0.0 {
            return num(0.0);
        }
        let q = snap(x / s);
        num(if x < 0.0 && mode != 0.0 {
            q.floor()
        } else {
            q.ceil()
        } * s)
    });
    r.eager("FLOOR.MATH", 1, 3, |e, a| {
        let x = e.num(&a[0])?;
        let s = opt_num(e, a, 1, 1.0)?.abs();
        let mode = opt_num(e, a, 2, 0.0)?;
        if s == 0.0 {
            return num(0.0);
        }
        let q = snap(x / s);
        num(if x < 0.0 && mode != 0.0 {
            q.ceil()
        } else {
            q.floor()
        } * s)
    });
    for name in ["CEILING.PRECISE", "ISO.CEILING"] {
        r.eager(name, 1, 2, |e, a| {
            let x = e.num(&a[0])?;
            let s = opt_num(e, a, 1, 1.0)?.abs();
            if s == 0.0 {
                num(0.0)
            } else {
                num(snap(x / s).ceil() * s)
            }
        });
    }
    r.eager("FLOOR.PRECISE", 1, 2, |e, a| {
        let x = e.num(&a[0])?;
        let s = opt_num(e, a, 1, 1.0)?.abs();
        if s == 0.0 {
            num(0.0)
        } else {
            num(snap(x / s).floor() * s)
        }
    });

    r.eager("COMBIN", 2, 2, |e, a| {
        let (n, k) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
        combin(n, k).map_or(Err(ErrorKind::Domain), num)
    });
    r.eager("COMBINA", 2, 2, |e, a| {
        let (n, k) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
        if n < 0.0 || k < 0.0 {
            return Err(ErrorKind::Domain);
        }
        if n == 0.0 {
            return if k == 0.0 {
                num(1.0)
            } else {
                Err(ErrorKind::Domain)
            };
        }
        combin(n + k - 1.0, k).map_or(Err(ErrorKind::Domain), num)
    });
    r.eager("PERMUT", 2, 2, |e, a| {
        let (n, k) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
        if n <= 0.0 && !(n == 0.0 && k == 0.0) || k < 0.0 || k > n {
            return Err(ErrorKind::Domain);
        }
        let mut p = 1.0_f64;
        let mut i = 0.0;
        while i < k {
            p *= n - i;
            i += 1.0;
        }
        num(p)
    });
    r.eager("PERMUTATIONA", 2, 2, |e, a| {
        let (n, k) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
        if n < 0.0 || k < 0.0 {
            return Err(ErrorKind::Domain);
        }
        num(n.powf(k))
    });
    r.eager("FACTDOUBLE", 1, 1, |e, a| {
        let n = e.num(&a[0])?.trunc();
        if n < -1.0 {
            return Err(ErrorKind::Domain);
        }
        let mut r = 1.0_f64;
        let mut i = n;
        while i > 1.0 {
            r *= i;
            i -= 2.0;
        }
        num(r)
    });
    r.eager("MULTINOMIAL", 1, VAR, |e, a| {
        let v = trunc_all(e, a)?;
        let total: f64 = v.iter().sum();
        let mut r = fact(total).ok_or(ErrorKind::Domain)?;
        for x in v {
            r /= fact(x).ok_or(ErrorKind::Domain)?;
        }
        num(r.round())
    });

    for (name, op) in [("BITAND", 0), ("BITOR", 1), ("BITXOR", 2)] {
        let f: super::EagerFn = match op {
            0 => |e, a| bits(e, a, |x, y| x & y),
            1 => |e, a| bits(e, a, |x, y| x | y),
            _ => |e, a| bits(e, a, |x, y| x ^ y),
        };
        r.eager(name, 2, 2, f);
    }
    r.eager("BITLSHIFT", 2, 2, |e, a| shift(e, a, true));
    r.eager("BITRSHIFT", 2, 2, |e, a| shift(e, a, false));

    r.eager("ERF", 1, 2, |e, a| {
        let lo = e.num(&a[0])?;
        match a.get(1) {
            Some(v) if !matches!(v, Value::Empty) => {
                let hi = e.num(v)?;
                num(erf(hi) - erf(lo))
            }
            _ => num(erf(lo)),
        }
    });
    r.eager("ERF.PRECISE", 1, 1, |e, a| num(erf(e.num(&a[0])?)));
    r.eager("ERFC", 1, 1, |e, a| num(erfc(e.num(&a[0])?)));
    r.eager("ERFC.PRECISE", 1, 1, |e, a| num(erfc(e.num(&a[0])?)));
}

fn bits(e: &mut Engine, a: &[Value], f: fn(u64, u64) -> u64) -> Res<Value> {
    let (x, y) = (e.num(&a[0])?, e.num(&a[1])?);
    let max = 281_474_976_710_655.0; // 2^48 - 1
    if x < 0.0 || y < 0.0 || x.fract() != 0.0 || y.fract() != 0.0 || x > max || y > max {
        return Err(ErrorKind::Domain);
    }
    num(f(x as u64, y as u64) as f64)
}

fn shift(e: &mut Engine, a: &[Value], left: bool) -> Res<Value> {
    let (x, s) = (e.num(&a[0])?, e.num(&a[1])?.trunc());
    if x < 0.0 || x.fract() != 0.0 || x > 281_474_976_710_655.0 {
        return Err(ErrorKind::Domain);
    }
    let left = left == (s >= 0.0);
    let s = s.abs();
    // A multiplication or division by a power of two, without limits on the shift or on the
    // result.
    let p = 2f64.powf(s);
    num(if left { x * p } else { (x / p).floor() })
}

// The error function of the C library (Sun's algorithm as `s_erf.c` of glibc has it, with the
// polynomials split in the way it does), so that the digits are those LibreOffice gets from
// `std::erf` and `std::erfc`: the same bits as glibc 2.39 gives.
//
// ====================================================
// Copyright (C) 1993 by Sun Microsystems, Inc. All rights reserved.
//
// Developed at SunPro, a Sun Microsystems, Inc. business.
// Permission to use, copy, modify, and distribute this
// software is freely granted, provided that this notice
// is preserved.
// ====================================================
mod sun_erf {
    const ERX: f64 = 8.450_629_115_104_675e-1;
    /*
     * Coefficients for approximation to  erf on [0,0.84375]
     */
    const EFX: f64 = 1.283_791_670_955_126e-1;
    const TINY: f64 = 1e-300;
    const PP0: f64 = 1.283_791_670_955_125_6e-1;
    const PP1: f64 = -3.250_421_072_470_015e-1;
    const PP2: f64 = -2.848_174_957_559_851e-2;
    const PP3: f64 = -5.770_270_296_489_442e-3;
    const PP4: f64 = -2.376_301_665_665_016_3e-5;
    const QQ1: f64 = 3.979_172_239_591_553_5e-1;
    const QQ2: f64 = 6.502_224_998_876_73e-2;
    const QQ3: f64 = 5.081_306_281_875_766e-3;
    const QQ4: f64 = 1.324_947_380_043_216_4e-4;
    const QQ5: f64 = -3.960_228_278_775_368e-6;
    /*
     * Coefficients for approximation to  erf  in [0.84375,1.25]
     */
    const PA0: f64 = -2.362_118_560_752_659_4e-3;
    const PA1: f64 = 4.148_561_186_837_483_3e-1;
    const PA2: f64 = -3.722_078_760_357_013e-1;
    const PA3: f64 = 3.183_466_199_011_617_5e-1;
    const PA4: f64 = -1.108_946_942_823_966_8e-1;
    const PA5: f64 = 3.547_830_432_561_823_6e-2;
    const PA6: f64 = -2.166_375_594_868_791e-3;
    const QA1: f64 = 1.064_208_804_008_442_3e-1;
    const QA2: f64 = 5.403_979_177_021_71e-1;
    const QA3: f64 = 7.182_865_441_419_627e-2;
    const QA4: f64 = 1.261_712_198_087_616_4e-1;
    const QA5: f64 = 1.363_708_391_202_905e-2;
    const QA6: f64 = 1.198_449_984_679_910_7e-2;
    /*
     * Coefficients for approximation to  erfc in [1.25,1/0.35]
     */
    const RA0: f64 = -9.864_944_034_847_148e-3;
    const RA1: f64 = -6.938_585_727_071_818e-1;
    const RA2: f64 = -1.055_862_622_532_329_1e1;
    const RA3: f64 = -6.237_533_245_032_600_6e1;
    const RA4: f64 = -1.623_966_694_625_734_7e2;
    const RA5: f64 = -1.846_050_929_067_110_4e2;
    const RA6: f64 = -8.128_743_550_630_66e1;
    const RA7: f64 = -9.814_329_344_169_145;
    const SA1: f64 = 1.965_127_166_743_925_7e1;
    const SA2: f64 = 1.376_577_541_435_190_4e2;
    const SA3: f64 = 4.345_658_774_752_292_3e2;
    const SA4: f64 = 6.453_872_717_332_679e2;
    const SA5: f64 = 4.290_081_400_275_678_3e2;
    const SA6: f64 = 1.086_350_055_417_794_4e2;
    const SA7: f64 = 6.570_249_770_319_282;
    const SA8: f64 = -6.042_441_521_485_81e-2;
    /*
     * Coefficients for approximation to  erfc in [1/.35,28]
     */
    const RB0: f64 = -9.864_942_924_700_1e-3;
    const RB1: f64 = -7.992_832_376_805_23e-1;
    const RB2: f64 = -1.775_795_491_775_475_2e1;
    const RB3: f64 = -1.606_363_848_558_219_2e2;
    const RB4: f64 = -6.375_664_433_683_896e2;
    const RB5: f64 = -1.025_095_131_611_077_2e3;
    const RB6: f64 = -4.835_191_916_086_514e2;
    const SB1: f64 = 3.033_806_074_348_246e1;
    const SB2: f64 = 3.257_925_129_965_739e2;
    const SB3: f64 = 1.536_729_586_084_437e3;
    const SB4: f64 = 3.199_858_219_508_595_5e3;
    const SB5: f64 = 2.553_050_406_433_164_4e3;
    const SB6: f64 = 4.745_285_412_069_553_7e2;
    const SB7: f64 = -2.244_095_244_658_582e1;

    /// `a * b + c`, rounded twice as the C library does (not fused).
    fn mad(a: f64, b: f64, c: f64) -> f64 {
        a * b + c
    }

    fn high_word(x: f64) -> u32 {
        (x.to_bits() >> 32) as u32
    }

    /// The rational function of the small arguments: `(r, s)` of `x + x r / s`.
    fn small(x: f64) -> f64 {
        let z = x * x;
        let r1 = mad(z, PP1, PP0);
        let z2 = z * z;
        let r2 = mad(z, PP3, PP2);
        let z4 = z2 * z2;
        let s1 = mad(z, QQ1, 1.0);
        let s2 = mad(z, QQ3, QQ2);
        let s3 = mad(z, QQ5, QQ4);
        let r = mad(z4, PP4, mad(z2, r2, r1));
        let s = mad(z4, s3, mad(z2, s2, s1));
        r / s
    }

    /// `P / Q` of the arguments between 0.84375 and 1.25, with `s = |x| - 1`.
    fn middle(s: f64) -> f64 {
        let p1 = mad(s, PA1, PA0);
        let s2 = s * s;
        let q1 = mad(s, QA1, 1.0);
        let s4 = s2 * s2;
        let p2 = mad(s, PA3, PA2);
        let s6 = s4 * s2;
        let q2 = mad(s, QA3, QA2);
        let p3 = mad(s, PA5, PA4);
        let q3 = mad(s, QA5, QA4);
        let p = mad(s6, PA6, mad(s4, p3, mad(s2, p2, p1)));
        let q = mad(s6, QA6, mad(s4, q3, mad(s2, q2, q1)));
        p / q
    }

    /// `exp(-x x - 0.5625 + R / S) / x` for the arguments from 1.25 to 28; `None` below -6.
    fn tail(ix: u32, hx: i32, x: f64, erfc_side: bool) -> Option<f64> {
        let x = x.abs();
        let s = 1.0 / (x * x);
        let (r, big_s);
        if ix < if erfc_side { 0x4006_DB6D } else { 0x4006_DB6E } {
            let r1 = mad(s, RA1, RA0);
            let s2 = s * s;
            let s1 = mad(s, SA1, 1.0);
            let s4 = s2 * s2;
            let r2 = mad(s, RA3, RA2);
            let s6 = s4 * s2;
            let q2 = mad(s, SA3, SA2);
            let s8 = s4 * s4;
            let r3 = mad(s, RA5, RA4);
            let q3 = mad(s, SA5, SA4);
            let r4 = mad(s, RA7, RA6);
            let q4 = mad(s, SA7, SA6);
            r = mad(s6, r4, mad(s4, r3, mad(s2, r2, r1)));
            big_s = mad(s8, SA8, mad(s6, q4, mad(s4, q3, mad(s2, q2, s1))));
        } else {
            if erfc_side && hx < 0 && ix >= 0x4018_0000 {
                return None;
            }
            let r1 = mad(s, RB1, RB0);
            let s2 = s * s;
            let s1 = mad(s, SB1, 1.0);
            let s4 = s2 * s2;
            let r2 = mad(s, RB3, RB2);
            let s6 = s4 * s2;
            let q2 = mad(s, SB3, SB2);
            let r3 = mad(s, RB5, RB4);
            let q3 = mad(s, SB5, SB4);
            let q4 = mad(s, SB7, SB6);
            r = mad(s6, RB6, mad(s4, r3, mad(s2, r2, r1)));
            big_s = mad(s6, q4, mad(s4, q3, mad(s2, q2, s1)));
        }
        let z = f64::from_bits(x.to_bits() & 0xffff_ffff_0000_0000);
        let e = mad(-z, z, -0.5625).exp() * mad(z - x, z + x, r / big_s).exp();
        Some(e / x)
    }

    pub fn erf(x: f64) -> f64 {
        let hx = high_word(x).cast_signed();
        let ix = (hx & 0x7fff_ffff).cast_unsigned();
        if ix >= 0x7ff0_0000 {
            let i = (hx.cast_unsigned() >> 31) << 1;
            return f64::from(1 - i.cast_signed()) + 1.0 / x;
        }
        if ix < 0x3feb_0000 {
            if ix < 0x3e30_0000 {
                if ix < 0x0080_0000 {
                    return 0.0625 * (16.0 * x + (16.0 * EFX) * x);
                }
                return mad(EFX, x, x);
            }
            return mad(x, small(x), x);
        }
        if ix < 0x3ff4_0000 {
            let pq = middle(x.abs() - 1.0);
            return if hx >= 0 { ERX + pq } else { -ERX - pq };
        }
        if ix >= 0x4018_0000 {
            return if hx >= 0 { 1.0 - TINY } else { TINY - 1.0 };
        }
        let r = tail(ix, hx, x, false).unwrap_or(0.0);
        if hx >= 0 { 1.0 - r } else { r - 1.0 }
    }

    pub fn erfc(x: f64) -> f64 {
        let hx = high_word(x).cast_signed();
        let ix = (hx & 0x7fff_ffff).cast_unsigned();
        if ix >= 0x7ff0_0000 {
            return f64::from((hx.cast_unsigned() >> 31) << 1) + 1.0 / x;
        }
        if ix < 0x3feb_0000 {
            if ix < 0x3c70_0000 {
                return 1.0 - x;
            }
            let y = small(x);
            if hx < 0x3fd0_0000 {
                return 1.0 - mad(x, y, x);
            }
            let r = mad(x, y, x - 0.5);
            return 0.5 - r;
        }
        if ix < 0x3ff4_0000 {
            let pq = middle(x.abs() - 1.0);
            return if hx >= 0 {
                (1.0 - ERX) - pq
            } else {
                1.0 + (ERX + pq)
            };
        }
        if ix < 0x403c_0000 {
            return match tail(ix, hx, x, true) {
                None => 2.0 - TINY,
                Some(v) if hx > 0 => v,
                Some(v) => 2.0 - v,
            };
        }
        if hx > 0 { TINY * TINY } else { 2.0 - TINY }
    }
}

pub(crate) use sun_erf::{erf, erfc};

/// Standard normal distribution function.
pub(crate) fn norm_cdf(z: f64) -> f64 {
    0.5 * erfc(-z / std::f64::consts::SQRT_2)
}

/// Inverse of the standard normal distribution function (Acklam's rational
/// approximation refined with Halley steps to full double precision).
pub(crate) fn norm_inv(p: f64) -> f64 {
    const A: [f64; 6] = [
        -3.969_683_028_665_376e1,
        2.209_460_984_245_205e2,
        -2.759_285_104_469_687e2,
        1.383_577_518_672_69e2,
        -3.066_479_806_614_716e1,
        2.506_628_277_459_239,
    ];
    const B: [f64; 5] = [
        -5.447_609_879_822_406e1,
        1.615_858_368_580_409e2,
        -1.556_989_798_598_866e2,
        6.680_131_188_771_972e1,
        -1.328_068_155_288_572e1,
    ];
    const C: [f64; 6] = [
        -7.784_894_002_430_293e-3,
        -3.223_964_580_411_365e-1,
        -2.400_758_277_161_838,
        -2.549_732_539_343_734,
        4.374_664_141_464_968,
        2.938_163_982_698_783,
    ];
    const D: [f64; 4] = [
        7.784_695_709_041_462e-3,
        3.224_671_290_700_398e-1,
        2.445_134_137_142_996,
        3.754_408_661_907_416,
    ];
    let low = 0.02425;
    let mut x = if p < low {
        let q = (-2.0 * p.ln()).sqrt();
        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    } else if p <= 1.0 - low {
        let q = p - 0.5;
        let r = q * q;
        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r + A[5]) * q
            / (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r + 1.0)
    } else {
        let q = (-2.0 * (1.0 - p).ln()).sqrt();
        -(((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    };
    for _ in 0..3 {
        let e = norm_cdf(x) - p;
        let u = e * (2.0 * std::f64::consts::PI).sqrt() * (x * x / 2.0).exp();
        x -= u / (1.0 + x * u / 2.0);
    }
    x
}

#[cfg(test)]
mod tests {
    /// `erf` and `erfc` of glibc 2.39, bit by bit.
    #[test]
    fn erf_and_erfc_are_the_c_library_s() {
        let cases: [(u64, u64, u64); 50] = [
            (0x0, 0x0, 0x3ff0000000000000),
            (0x3bc79ca10c924223, 0x3bcaa4a230244ae0, 0x3ff0000000000000),
            (0x3e35798ee2308c3a, 0x3e383b547fb79514, 0x3feffffffcf89570),
            (0x3fd3333333333333, 0x3fd50838881dea0f, 0x3fe57be3bbf10af8),
            (0x3fe3333333333333, 0x3fe352ca0235d4f7, 0x3fd95a6bfb945612),
            (0x3feae147ae147ae1, 0x3fe87c0c914f1a59, 0x3fce0fcdbac3969e),
            (0x3feccccccccccccd, 0x3fe98045a6c8a2e6, 0x3fc9fee964dd7469),
            (0x3ff199999999999a, 0x3fec2aa3d27302c0, 0x3fbeaae16c67ea00),
            (0x3ff3d70a3d70a3d7, 0x3fed74c749cb9bf0, 0x3fb459c5b1a32084),
            (0x3ff4cccccccccccd, 0x3fede364a1372d3a, 0x3fb0e4daf646962e),
            (0x4000000000000000, 0x3fefd9ae142795e3, 0x3f7328f5ec350e66),
            (0x4006666666666666, 0x3fefff62af990752, 0x3f13aa0cdf15cedb),
            (0x4007333333333333, 0x3fefffa9cfbea4e3, 0x3f058c1056c73872),
            (0x4010000000000000, 0x3feffffff7b91176, 0x3e508ddd13bd35e7),
            (0x401799999999999a, 0x3fefffffffffffff, 0x3c94b99601a343be),
            (0x401a000000000000, 0x3ff0000000000000, 0x3be6ae172414ceba),
            (0x4028000000000000, 0x3ff0000000000000, 0x32ac90f21d2d475f),
            (0x403be66666666666, 0x3ff0000000000000, 0x0),
            (0x403c800000000000, 0x3ff0000000000000, 0x0),
            (0xbbc79ca10c924223, 0xbbcaa4a230244ae0, 0x3ff0000000000000),
            (0xbfd3333333333333, 0xbfd50838881dea0f, 0x3ff5420e22077a84),
            (0xbfeccccccccccccd, 0xbfe98045a6c8a2e6, 0x3ffcc022d3645173),
            (0xbff199999999999a, 0xbfec2aa3d27302c0, 0x3ffe1551e9398160),
            (0xc004000000000000, 0xbfeffcaa8f4c9bea, 0x3ffffe5547a64df5),
            (0xc014000000000000, 0xbfefffffffffc9e8, 0x3fffffffffffe4f4),
            (0xc01a000000000000, 0xbff0000000000000, 0x4000000000000000),
            (0xc03e000000000000, 0xbff0000000000000, 0x4000000000000000),
            (0xbfe5557de0d66f0d, 0xbfe4ef7f14857a8b, 0x3ffa77bf8a42bd46),
            (0x3feac72c0e7bc3c6, 0x3fe86d766d9bef37, 0x3fce4a2649904324),
            (0x4017c17b95a29414, 0x3ff0000000000000, 0x3c89f88ee73ad2e2),
            (0xbfdec70c996b7671, 0xbfe01d1f586a71fd, 0x3ff80e8fac3538fe),
            (0x3fbc1a6937d1fe65, 0x3fbf9587c80a20d2, 0x3fec0d4f06febbe6),
            (0x3ff392ff4ba51a00, 0x3fed53180474059f, 0x3fb5673fdc5fd30b),
            (0xc011a8b588e368f1, 0xbfefffffffc52813, 0x3fffffffffe29409),
            (0x3fc5571d1d4738a4, 0x3fc7dbd23e9b2697, 0x3fea090b7059365a),
            (0x3ffd17fe8ee6b830, 0x3fefad0f4cb7d21d, 0x3f84bc2cd20b78b0),
            (0x4010681dc0db2703, 0x3feffffffc74140e, 0x3e3c5f5f93d8f7e7),
            (0xc016baa582dbe7f3, 0xbfeffffffffffff8, 0x3ffffffffffffffc),
            (0xc00604e0daa0cae6, 0xbfefff2fe3df5f4a, 0x3fffff97f1efafa5),
            (0xc016ec258d5842b7, 0xbfeffffffffffffb, 0x3ffffffffffffffe),
            (0x4011571044f1a198, 0x3fefffffff87b60c, 0x3e0e127d04a92c1d),
            (0x4005aa44c7b02d5a, 0x3feffef323fc5015, 0x3f20cdc03afeb652),
            (0xc019a79a6b50b0f2, 0xbff0000000000000, 0x4000000000000000),
            (0x401b00b9991361dd, 0x3ff0000000000000, 0x3b993b8833ad0210),
            (0x401a06c4827b6fe3, 0x3ff0000000000000, 0x3be4ca8dc507f9f7),
            (0x40013d441355475a, 0x3fefed18e8bc9114, 0x3f62e717436eec74),
            (0x3ff9e2d40aaeafab, 0x3fef4aa94ec12e80, 0x3f96aad627da2ff4),
            (0xc0132e2a3cea6c1a, 0xbfeffffffffe5ce0, 0x3fffffffffff2e70),
            (0xc01b28f32378ab0d, 0xbff0000000000000, 0x4000000000000000),
            (0x3fd96dfc5cdd50a9, 0x3fdb40cc8472dca2, 0x3fe25f99bdc691af),
        ];
        for (x, erf, erfc) in cases {
            let x = f64::from_bits(x);
            assert_eq!(super::erf(x).to_bits(), erf, "erf({x:e})");
            assert_eq!(super::erfc(x).to_bits(), erfc, "erfc({x:e})");
        }
    }
}
