//! Complex numbers: `COMPLEX`, `IMREAL`, `IMAGINARY`, `IMABS`, `IMARGUMENT`, `IMCONJUGATE`, the
//! arithmetic (`IMSUM`, `IMSUB`, `IMPRODUCT`, `IMDIV`, `IMPOWER`), the roots, exponentials and
//! logarithms, and the trigonometric and hyperbolic functions with their reciprocals. A complex
//! number is a text (`3+4i`, with `i` or `j`), which is how LibreOffice takes it and gives it back;
//! its numbers are written with 15 significant digits. The results are what the oracle
//! (`tests/assets/semantics.*`) has from LibreOffice 26.8.

use super::{Reg, VAR, num};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

#[derive(Clone, Copy, Debug)]
struct Z {
    re: f64,
    im: f64,
    /// The suffix the number was written with, if it had one.
    suffix: Option<char>,
}

impl Z {
    fn new(re: f64, im: f64) -> Self {
        Self {
            re,
            im,
            suffix: None,
        }
    }

    fn abs(self) -> f64 {
        self.re.hypot(self.im)
    }

    /// The argument, from `atan2` (the small angles keep their digits).
    fn arg(self) -> Res<f64> {
        if self.re == 0.0 && self.im == 0.0 {
            return Err(DOMAIN);
        }
        Ok(self.im.atan2(self.re))
    }

    fn times(self, o: Z) -> Z {
        Z {
            re: self.re * o.re - self.im * o.im,
            im: self.re * o.im + self.im * o.re,
            suffix: self.suffix.or(o.suffix),
        }
    }

    fn reciprocal(self) -> Res<Z> {
        Z {
            re: 1.0,
            im: 0.0,
            suffix: self.suffix,
        }
        .div(self)
    }

    fn div(self, d: Z) -> Res<Z> {
        if d.re == 0.0 && d.im == 0.0 {
            return Err(DOMAIN);
        }
        let (mut a, mut b, mut c, mut e) = (self.re, self.im, d.re, d.im);
        // The division of the C compiler runtime: Smith's algorithm, and where it gives NaN the
        // limits with an infinite operand.
        let (mut re, mut im) = if c.abs() < e.abs() {
            let ratio = c / e;
            let den = c * ratio + e;
            ((a * ratio + b) / den, (b * ratio - a) / den)
        } else {
            let ratio = e / c;
            let den = e * ratio + c;
            ((b * ratio + a) / den, (b - a * ratio) / den)
        };
        if re.is_nan() && im.is_nan() {
            if (a.is_infinite() || b.is_infinite()) && c.is_finite() && e.is_finite() {
                a = if a.is_infinite() {
                    1.0f64.copysign(a)
                } else {
                    0.0f64.copysign(a)
                };
                b = if b.is_infinite() {
                    1.0f64.copysign(b)
                } else {
                    0.0f64.copysign(b)
                };
                re = f64::INFINITY * (a * c + b * e);
                im = f64::INFINITY * (b * c - a * e);
            } else if (c.is_infinite() || e.is_infinite()) && a.is_finite() && b.is_finite() {
                c = if c.is_infinite() {
                    1.0f64.copysign(c)
                } else {
                    0.0f64.copysign(c)
                };
                e = if e.is_infinite() {
                    1.0f64.copysign(e)
                } else {
                    0.0f64.copysign(e)
                };
                re = 0.0 * (a * c + b * e);
                im = 0.0 * (b * c - a * e);
            }
        }
        Ok(Z {
            re,
            im,
            suffix: self.suffix,
        })
    }
}

/// `%.15g` of a number, as LibreOffice writes the parts of a complex number.
fn g15(x: f64) -> String {
    if x == 0.0 {
        return if x.is_sign_negative() {
            "-0".into()
        } else {
            "0".into()
        };
    }
    let s = format!("{:.14e}", x.abs());
    let (mantissa, exp) = s.split_once('e').unwrap_or((&s, "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let digits: String = mantissa.chars().filter(char::is_ascii_digit).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let sign = if x < 0.0 { "-" } else { "" };
    if (-4..15).contains(&exp) {
        if exp >= 0 {
            let int_len = exp as usize + 1;
            let mut int = digits.chars().take(int_len).collect::<String>();
            while int.len() < int_len {
                int.push('0');
            }
            let frac: String = digits.chars().skip(int_len).collect();
            if frac.is_empty() {
                format!("{sign}{int}")
            } else {
                format!("{sign}{int}.{frac}")
            }
        } else {
            let zeros = "0".repeat((-exp - 1) as usize);
            format!("{sign}0.{zeros}{digits}")
        }
    } else {
        let (first, rest) = digits.split_at(1);
        let m = if rest.is_empty() {
            first.to_string()
        } else {
            format!("{first}.{rest}")
        };
        format!(
            "{sign}{m}e{}{:02}",
            if exp < 0 { '-' } else { '+' },
            exp.abs()
        )
    }
}

fn write_z(z: Z) -> String {
    let suffix = z.suffix.unwrap_or('i');
    let mut out = String::new();
    if z.im == 0.0 {
        return g15(z.re);
    }
    if z.re != 0.0 {
        out.push_str(&g15(z.re));
    }
    if z.im == 1.0 || z.im == -1.0 {
        if z.im < 0.0 {
            out.push('-');
        } else if !out.is_empty() {
            out.push('+');
        }
    } else {
        if z.im > 0.0 && !out.is_empty() {
            out.push('+');
        }
        out.push_str(&g15(z.im));
    }
    out.push(suffix);
    out
}

/// A number as LibreOffice reads it inside a complex number: digits with a `.` or a `,`, and an
/// exponent. The fraction is added digit by digit with a multiplier divided by ten each time, so
/// `0.0000001` is 1.0000000000000005e-7, which the results of the functions show.
fn number(s: &[char], at: &mut usize) -> Option<f64> {
    let start = *at;
    let (mut int, mut frac, mut mult) = (0.0f64, 0.0f64, 0.1f64);
    let mut digits = 0;
    while *at < s.len() && s[*at].is_ascii_digit() {
        int = int * 10.0 + f64::from(s[*at] as u8 - b'0');
        *at += 1;
        digits += 1;
    }
    if *at < s.len() && (s[*at] == '.' || s[*at] == ',') {
        *at += 1;
        while *at < s.len() && s[*at].is_ascii_digit() {
            frac += f64::from(s[*at] as u8 - b'0') * mult;
            mult *= 0.1;
            *at += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        *at = start;
        return None;
    }
    let mut value = int + frac;
    if *at < s.len() && (s[*at] == 'e' || s[*at] == 'E') {
        let mut k = *at + 1;
        let negative = k < s.len() && s[k] == '-';
        if k < s.len() && (s[k] == '+' || s[k] == '-') {
            k += 1;
        }
        let from = k;
        let mut exp = 0.0f64;
        while k < s.len() && s[k].is_ascii_digit() {
            exp = exp * 10.0 + f64::from(s[k] as u8 - b'0');
            k += 1;
        }
        if k > from {
            value *= 10f64.powf(if negative { -exp } else { exp });
            *at = k;
        }
    }
    Some(value)
}

fn is_unit(c: char) -> bool {
    c == 'i' || c == 'j'
}

/// The complex number a text writes, or `None` when it is not one.
fn parse(text: &str) -> Option<Z> {
    let s: Vec<char> = text.chars().collect();
    if s.is_empty() {
        return None;
    }
    let mut at = 0;
    let mut sign = 1.0;
    if s[at] == '+' || s[at] == '-' {
        if s[at] == '-' {
            sign = -1.0;
        }
        at += 1;
    }
    let first_unit_only = at < s.len() && is_unit(s[at]);
    let (first, suffix) = if first_unit_only {
        let c = s[at];
        at += 1;
        (1.0, Some(c))
    } else {
        let n = number(&s, &mut at)?;
        if at < s.len() && is_unit(s[at]) {
            let c = s[at];
            at += 1;
            (n, Some(c))
        } else {
            (n, None)
        }
    };
    let first = sign * first;
    if at == s.len() {
        return Some(match suffix {
            Some(c) => Z {
                re: 0.0,
                im: first,
                suffix: Some(c),
            },
            None => Z {
                re: first,
                im: 0.0,
                suffix: None,
            },
        });
    }
    // A real part followed by a signed imaginary one.
    if suffix.is_some() || (s[at] != '+' && s[at] != '-') {
        return None;
    }
    let sign2 = if s[at] == '-' { -1.0 } else { 1.0 };
    at += 1;
    let (second, unit) = if at < s.len() && is_unit(s[at]) {
        let c = s[at];
        at += 1;
        (1.0, c)
    } else {
        let n = number(&s, &mut at)?;
        if at < s.len() && is_unit(s[at]) {
            let c = s[at];
            at += 1;
            (n, c)
        } else {
            return None;
        }
    };
    if at != s.len() {
        return None;
    }
    Some(Z {
        re: first,
        im: sign2 * second,
        suffix: Some(unit),
    })
}

/// A complex argument: text, a number (through its 15 digits) or a boolean; a blank or text that
/// is not a complex number is outside the domain.
fn arg(e: &mut Engine, v: &Value) -> Res<Z> {
    match e.scalar(v.clone()) {
        Value::Error(k) => Err(k),
        Value::Text(t) => parse(&t).ok_or(DOMAIN),
        Value::Number(n) => parse(&g15(n)).ok_or(DOMAIN),
        Value::Bool(b) => Ok(Z::new(f64::from(u8::from(b)), 0.0)),
        _ => Err(DOMAIN),
    }
}

fn out(z: Z) -> Res<Value> {
    if !z.re.is_finite() || !z.im.is_finite() {
        return Err(DOMAIN);
    }
    Ok(Value::text(write_z(z)))
}

/// The arguments of `IMSUM` and `IMPRODUCT`: the cells of ranges and arrays, a blank of a range
/// skipped, an error in a range outside the domain.
fn list(e: &mut Engine, a: &[Value]) -> Res<Vec<Z>> {
    let mut zs = Vec::new();
    for v in a {
        match v {
            // A single blank cell is 0; the blanks of a range are not there.
            Value::Ref(r) if r.is_single() && matches!(e.scalar(v.clone()), Value::Empty) => {
                zs.push(Z::new(0.0, 0.0));
            }
            Value::Ref(_) | Value::Array(_) => {
                let m = e.array_of(v.clone());
                for x in &m.data {
                    match x {
                        Value::Empty => {}
                        Value::Error(_) => return Err(DOMAIN),
                        x => zs.push(arg(e, x)?),
                    }
                }
            }
            Value::Text(t) if t.is_empty() => {}
            v => zs.push(arg(e, v)?),
        }
    }
    Ok(zs)
}

fn first_suffix(zs: &[Z]) -> Option<char> {
    zs.iter().find_map(|z| z.suffix)
}

fn finish(mut z: Z, suffix: Option<char>) -> Res<Value> {
    z.suffix = suffix;
    out(z)
}

/// Whether `x` is an argument that sine and cosine of LibreOffice accept.
fn valid_arc(x: f64, limit: f64) -> Res<()> {
    if x.abs() < limit { Ok(()) } else { Err(DOMAIN) }
}

const TWO_64: f64 = 18446744073709551616.0;
const TWO_63: f64 = 9223372036854775808.0;

fn sin(z: Z) -> Res<Z> {
    valid_arc(z.re, TWO_64)?;
    if z.im != 0.0 {
        Ok(Z {
            re: z.re.sin() * z.im.cosh(),
            im: z.re.cos() * z.im.sinh(),
            ..z
        })
    } else {
        Ok(Z {
            re: z.re.sin(),
            ..z
        })
    }
}

fn cos(z: Z) -> Res<Z> {
    valid_arc(z.re, TWO_64)?;
    if z.im != 0.0 {
        Ok(Z {
            re: z.re.cos() * z.im.cosh(),
            im: -(z.re.sin() * z.im.sinh()),
            ..z
        })
    } else {
        Ok(Z {
            re: z.re.cos(),
            ..z
        })
    }
}

/// The tangent as the C library computes it for a complex number: the limit `±i` where the
/// hyperbolic functions overflow.
fn tan(z: Z) -> Res<Z> {
    valid_arc(z.re, TWO_63)?;
    if z.im == 0.0 {
        return Ok(Z {
            re: z.re.tan(),
            ..z
        });
    }
    const T: f64 = 354.0;
    let (sinrx, cosrx) = if z.re.abs() > f64::MIN_POSITIVE {
        (z.re.sin(), z.re.cos())
    } else {
        (z.re, 1.0)
    };
    if z.im.abs() > T {
        let exp_2t = (2.0 * T).exp();
        let mut re = 4.0 * sinrx * cosrx / exp_2t;
        let y = z.im.abs() - T;
        re /= if y > T { exp_2t } else { (2.0 * y).exp() };
        return Ok(Z {
            re,
            im: 1.0f64.copysign(z.im),
            ..z
        });
    }
    let (sinhix, coshix) = if z.im.abs() > f64::MIN_POSITIVE {
        (z.im.sinh(), z.im.cosh())
    } else {
        (z.im, 1.0)
    };
    let den = if sinhix.abs() > cosrx.abs() * f64::EPSILON {
        cosrx * cosrx + sinhix * sinhix
    } else {
        cosrx * cosrx
    };
    Ok(Z {
        re: sinrx * cosrx / den,
        im: sinhix * coshix / den,
        ..z
    })
}

fn sinh(z: Z) -> Res<Z> {
    valid_arc(z.im, TWO_64)?;
    if z.im != 0.0 {
        Ok(Z {
            re: z.re.sinh() * z.im.cos(),
            im: z.re.cosh() * z.im.sin(),
            ..z
        })
    } else {
        Ok(Z {
            re: z.re.sinh(),
            ..z
        })
    }
}

fn cosh(z: Z) -> Res<Z> {
    valid_arc(z.im, TWO_64)?;
    if z.im != 0.0 {
        Ok(Z {
            re: z.re.cosh() * z.im.cos(),
            im: z.re.sinh() * z.im.sin(),
            ..z
        })
    } else {
        Ok(Z {
            re: z.re.cosh(),
            ..z
        })
    }
}

fn ln(z: Z) -> Res<Z> {
    if z.re == 0.0 && z.im == 0.0 {
        return Err(DOMAIN);
    }
    Ok(Z {
        re: clog_real(z.re, z.im),
        im: z.im.atan2(z.re),
        ..z
    })
}

/// The root, with the part that would cancel taken from the other (`im = i / (2 re)`); a purely
/// imaginary number has equal parts.
/// The logarithm to a base, from the natural one (its natural logarithm `base_ln`).
fn log_base(z: Z, base_ln: f64) -> Res<Z> {
    let l = ln(z)?;
    Ok(Z {
        re: l.re / base_ln,
        im: l.im / base_ln,
        ..l
    })
}

fn sqrt(z: Z) -> Z {
    let p = z.abs();
    let neg = z.im < 0.0;
    let (re, im) = if z.re == 0.0 {
        let m = (z.im.abs() / 2.0).sqrt();
        (m, if neg { -m } else { m })
    } else if z.re > 0.0 {
        let re = ((p + z.re) / 2.0).sqrt();
        (re, z.im / (2.0 * re))
    } else {
        let m = ((p - z.re) / 2.0).sqrt();
        let im = if neg { -m } else { m };
        (z.im / (2.0 * im), im)
    };
    Z { re, im, ..z }
}

/// The real part of the complex logarithm as the C library computes it (`log1p` near the unit
/// circle, where `log(hypot)` loses digits).
fn clog_real(x: f64, y: f64) -> f64 {
    let (mut ax, mut ay) = (x.abs(), y.abs());
    if ax < ay {
        std::mem::swap(&mut ax, &mut ay);
    }
    if ax == 1.0 {
        (ay * ay).ln_1p() / 2.0
    } else if ax > 1.0 && ax < 2.0 && ay < 1.0 {
        let mut d = (ax - 1.0) * (ax + 1.0);
        if ay >= f64::EPSILON {
            d += ay * ay;
        }
        d.ln_1p() / 2.0
    } else if (0.5..1.0).contains(&ax) && ay < f64::EPSILON / 2.0 {
        ((ax - 1.0) * (ax + 1.0)).ln_1p() / 2.0
    } else if (0.5..1.0).contains(&ax) && ax * ax + ay * ay >= 0.5 {
        ((ax - 1.0) * (ax + 1.0) + ay * ay).ln_1p() / 2.0
    } else {
        ax.hypot(ay).ln()
    }
}

/// `z^n` as the C++ library does it for a complex base and a real exponent: `exp(n log z)`.
fn power(z: Z, n: f64) -> Res<Z> {
    if z.re == 0.0 && z.im == 0.0 {
        return if n > 0.0 { Ok(z) } else { Err(DOMAIN) };
    }
    if z.im == 0.0 && z.re > 0.0 {
        return Ok(Z {
            re: z.re.powf(n),
            ..z
        });
    }
    let p = (n * clog_real(z.re, z.im)).exp();
    let t = n * z.im.atan2(z.re);
    Ok(Z {
        re: p * t.cos(),
        im: p * t.sin(),
        ..z
    })
}

pub(super) fn register(r: &mut Reg) {
    macro_rules! complex_function {
        ($name:literal, $f:expr) => {
            r.eager($name, 1, 1, |e, a| {
                let z = arg(e, &a[0])?;
                let f: fn(Z) -> Res<Z> = $f;
                let w = f(z)?;
                out(Z {
                    suffix: z.suffix,
                    ..w
                })
            });
        };
    }
    r.eager("COMPLEX", 2, 3, |e, a| {
        let (re, im) = (e.num(&a[0])?, e.num(&a[1])?);
        let suffix = match a.get(2) {
            None => 'i',
            Some(v) => match e.scalar(v.clone()) {
                Value::Error(k) => return Err(k),
                Value::Text(t) if t.is_empty() => 'i',
                Value::Text(t) if &*t == "i" => 'i',
                Value::Text(t) if &*t == "j" => 'j',
                _ => return Err(DOMAIN),
            },
        };
        out(Z {
            re,
            im,
            suffix: Some(suffix),
        })
    });
    r.eager("IMREAL", 1, 1, |e, a| num(arg(e, &a[0])?.re));
    r.eager("IMAGINARY", 1, 1, |e, a| num(arg(e, &a[0])?.im));
    r.eager("IMABS", 1, 1, |e, a| num(arg(e, &a[0])?.abs()));
    r.eager("IMARGUMENT", 1, 1, |e, a| num(arg(e, &a[0])?.arg()?));
    complex_function!("IMCONJUGATE", |z| Ok(Z { im: -z.im, ..z }));
    complex_function!("IMSQRT", |z| Ok(sqrt(z)));
    complex_function!("IMEXP", |z| {
        let f = z.re.exp();
        Ok(Z {
            re: f * z.im.cos(),
            im: f * z.im.sin(),
            ..z
        })
    });
    complex_function!("IMLN", ln);
    complex_function!("IMLOG10", |z| log_base(z, std::f64::consts::LN_10));
    complex_function!("IMLOG2", |z| {
        let l = ln(z)?;
        Ok(Z {
            re: l.re * std::f64::consts::LOG2_E,
            im: l.im * std::f64::consts::LOG2_E,
            ..l
        })
    });
    complex_function!("IMSIN", sin);
    complex_function!("IMCOS", cos);
    complex_function!("IMTAN", tan);
    complex_function!("IMSEC", |z| cos(z)?.reciprocal());
    complex_function!("IMCSC", |z| sin(z)?.reciprocal());
    complex_function!("IMCOT", |z| tan(z)?.reciprocal());
    complex_function!("IMSINH", sinh);
    complex_function!("IMCOSH", cosh);
    complex_function!("IMSECH", |z| cosh(z)?.reciprocal());
    complex_function!("IMCSCH", |z| sinh(z)?.reciprocal());
    r.eager("IMPOWER", 2, 2, |e, a| {
        // An error in either argument comes before a text that is not a complex number.
        for v in a {
            if let Value::Error(k) = e.scalar(v.clone()) {
                return Err(k);
            }
        }
        let z = arg(e, &a[0])?;
        let n = e.num(&a[1])?;
        out(power(z, n)?)
    });
    r.eager("IMSUM", 1, VAR, |e, a| {
        let zs = list(e, a)?;
        let mut acc = Z::new(0.0, 0.0);
        for z in &zs {
            acc.re += z.re;
            acc.im += z.im;
        }
        finish(acc, first_suffix(&zs))
    });
    r.eager("IMPRODUCT", 1, VAR, |e, a| {
        let zs = list(e, a)?;
        if zs.is_empty() {
            return Ok(Value::text("0"));
        }
        let mut acc = Z::new(1.0, 0.0);
        for z in &zs {
            acc = acc.times(*z);
        }
        finish(acc, first_suffix(&zs))
    });
    r.eager("IMSUB", 2, 2, |e, a| {
        let (x, y) = (arg(e, &a[0])?, arg(e, &a[1])?);
        out(Z {
            re: x.re - y.re,
            im: x.im - y.im,
            suffix: x.suffix.or(y.suffix),
        })
    });
    r.eager("IMDIV", 2, 2, |e, a| {
        let (x, y) = (arg(e, &a[0])?, arg(e, &a[1])?);
        out(x.div(y)?)
    });
}
