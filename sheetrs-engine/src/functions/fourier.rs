//! `FOURIER`, with the transforms of LibreOffice (`sc/source/core/tool/interpr3.cxx`): the
//! radix-2 transform for lengths that are a power of 2, Bluestein's for the other lengths and a
//! packed transform for real input of even length, each with the twiddle factors it computes, so
//! that the digits are LibreOffice's. The results are checked against the oracle
//! (`tests/assets/semantics.*`).

use std::f64::consts::PI;
use std::sync::Arc;

use super::Reg;
use crate::engine::{Engine, Res};
use crate::ir::Expr;
use crate::value::{Array, ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// The least power of two that is not below `n`, and its logarithm.
fn round_up_pow2(n: usize) -> (usize, usize) {
    let bits = (usize::BITS - n.leading_zeros()) as usize;
    let pow2 = if bits == 0 { 0 } else { 1usize << (bits - 1) };
    if pow2 != n {
        (if pow2 != 0 { pow2 << 1 } else { 1 }, bits)
    } else {
        (n, bits - 1)
    }
}

fn bit_reverse(n: usize, bound: usize) -> usize {
    let mut out = 0;
    let mut mask = 1;
    while mask < bound {
        out <<= 1;
        if n & mask != 0 {
            out |= 1;
        }
        mask <<= 1;
    }
    out
}

struct Twiddle {
    re: Vec<f64>,
    im: Vec<f64>,
    inverse: bool,
}

impl Twiddle {
    fn new(n: usize, inverse: bool) -> Self {
        let mut t = Self {
            re: vec![0.0; n],
            im: vec![0.0; n],
            inverse,
        };
        t.compute();
        t
    }

    fn conjugate(&mut self) {
        self.inverse = !self.inverse;
        for x in &mut self.im {
            *x = -*x;
        }
    }

    fn compute(&mut self) {
        let n = self.re.len();
        let inv = self.inverse;
        let w = if inv { 2.0 } else { -2.0 } * PI / n as f64;
        let (re, im) = (&mut self.re, &mut self.im);
        if n == 1 {
            re[0] = 1.0;
            im[0] = 0.0;
        } else if n == 2 {
            re[0] = 1.0;
            im[0] = 0.0;
            re[1] = -1.0;
            im[1] = 0.0;
        } else if n == 4 {
            re[0] = 1.0;
            im[0] = 0.0;
            re[1] = 0.0;
            im[1] = if inv { 1.0 } else { -1.0 };
            re[2] = -1.0;
            im[2] = 0.0;
            re[3] = 0.0;
            im[3] = if inv { -1.0 } else { 1.0 };
        } else if n.is_multiple_of(4) {
            let q = n >> 2;
            for (i, x) in re[..=q].iter_mut().enumerate() {
                *x = (w * i as f64).cos();
            }
            if inv {
                let q1 = q;
                for i in 0..=q1 {
                    im[i] = re[q1 - i];
                }
                let q2 = q1 << 1;
                for i in q1 + 1..=q2 {
                    re[i] = -re[q2 - i];
                    im[i] = im[q2 - i];
                }
                let q3 = q2 + q1;
                for i in q2 + 1..=q3 {
                    re[i] = -re[i - q2];
                    im[i] = -im[i - q2];
                }
                for i in q3 + 1..n {
                    re[i] = re[n - i];
                    im[i] = -im[n - i];
                }
            } else {
                let q4 = q;
                let q3 = q << 1;
                let q2 = q3 + q;
                for i in 0..=q4 {
                    im[i] = -re[q4 - i];
                }
                for i in q4 + 1..=q3 {
                    re[i] = -re[q3 - i];
                    im[i] = im[q3 - i];
                }
                for i in q3 + 1..=q2 {
                    re[i] = -re[i - q3];
                    im[i] = -im[i - q3];
                }
                for i in q2 + 1..n {
                    re[i] = re[n - i];
                    im[i] = -im[n - i];
                }
            }
        } else {
            for i in 0..n {
                let angle = w * i as f64;
                re[i] = angle.cos();
                im[i] = angle.sin();
            }
        }
    }
}

fn normalize(a: &mut [f64], only_real: bool) {
    let points = a.len() / 2;
    let scale = 1.0 / points as f64;
    for x in &mut a[..points] {
        *x *= scale;
    }
    if !only_real {
        for x in &mut a[points..] {
            *x *= scale;
        }
    }
}

fn to_polar(a: &mut [f64], min_mag: f64) {
    let points = a.len() / 2;
    for i in 0..points {
        let (r, im) = (a[i], a[points + i]);
        let mag = r.hypot(im);
        let (mag, phase) = if mag < min_mag {
            (0.0, 0.0)
        } else {
            (mag, im.atan2(r))
        };
        a[i] = mag;
        a[points + i] = phase;
    }
}

struct Options {
    inverse: bool,
    polar: bool,
    min_mag: f64,
}

/// The radix-2 transform of a complex array (the real parts first, then the imaginary parts) of
/// a power of two points.
fn fft2(a: &mut [f64], tf: &Twiddle, o: &Options, subsample: bool, disable_normalize: bool) {
    let points = a.len() / 2;
    let (_, stages) = round_up_pow2(points);
    for i in 0..points {
        let rev = bit_reverse(i, points);
        if i < rev {
            a.swap(i, rev);
            a.swap(points + i, points + rev);
        }
    }
    let flies_in_stage = points / 2;
    for stage in 0..stages {
        let scale_bits = stages - stage - 1;
        let flies_in_group = 1usize << stage;
        let groups = flies_in_stage / flies_in_group;
        let width = flies_in_group;
        let mut top = 0;
        for _ in 0..groups {
            for _ in 0..flies_in_group {
                let bottom = top + width;
                let mut w1 = (top << scale_bits) & (points - 1);
                let mut w2 = (bottom << scale_bits) & (points - 1);
                if subsample {
                    w1 <<= 1;
                    w2 <<= 1;
                }
                let (x1r, x2r) = (a[top], a[bottom]);
                let (w1r, w1i) = (tf.re[w1], tf.im[w1]);
                let (w2r, w2i) = (tf.re[w2], tf.im[w2]);
                let (x1i, x2i) = (a[points + top], a[points + bottom]);
                a[top] = x1r + x2r * w1r - x2i * w1i;
                a[points + top] = x1i + x2i * w1r + x2r * w1i;
                a[bottom] = x1r + x2r * w2r - x2i * w2i;
                a[points + bottom] = x1i + x2i * w2r + x2r * w2i;
                top += 1;
            }
            top += width;
        }
    }
    if o.polar {
        to_polar(a, o.min_mag);
    }
    if o.inverse && !disable_normalize {
        normalize(a, o.polar);
    }
}

/// Bluestein's transform, for a complex array of any number of points.
fn bluestein(a: &mut [f64], real: bool, o: &Options, disable_normalize: bool) {
    let points = a.len() / 2;
    let w = if o.inverse { 2.0 } else { -2.0 } * PI / points as f64;
    let mut real_scalars = vec![0.0; points];
    let mut imag_scalars = vec![0.0; points];
    for i in 0..points {
        let angle = 0.5 * w * (i * i) as f64;
        real_scalars[i] = angle.cos();
        imag_scalars[i] = angle.sin();
    }
    let (extended, _) = round_up_pow2(points * 2 - 1);
    let mut a_signal = vec![0.0; extended * 2];
    let mut b_signal = vec![0.0; extended * 2];
    for i in 0..points {
        a_signal[i] = a[i] * real_scalars[i]
            + if real {
                0.0
            } else {
                -a[points + i] * imag_scalars[i]
            };
        a_signal[extended + i] = a[i] * imag_scalars[i]
            + if real {
                0.0
            } else {
                a[points + i] * real_scalars[i]
            };
        let (re, im) = (real_scalars[i], -imag_scalars[i]);
        b_signal[i] = re;
        b_signal[extended + i] = im;
        if i != 0 {
            b_signal[extended - i] = re;
            b_signal[(extended << 1) - i] = im;
        }
    }
    let plain = Options {
        inverse: false,
        polar: false,
        min_mag: 0.0,
    };
    let mut tf = Twiddle::new(extended, false);
    fft2(&mut a_signal, &tf, &plain, false, true);
    fft2(&mut b_signal, &tf, &plain, false, true);
    for i in 0..extended {
        let (ar, ai) = (a_signal[i], a_signal[extended + i]);
        let (br, bi) = (b_signal[i], b_signal[extended + i]);
        a_signal[i] = ar * br - ai * bi;
        a_signal[extended + i] = ar * bi + ai * br;
    }
    tf.conjugate();
    let inverse = Options {
        inverse: true,
        polar: false,
        min_mag: 0.0,
    };
    fft2(&mut a_signal, &tf, &inverse, false, false);
    for i in 0..points {
        let (re, im) = (a_signal[i], a_signal[extended + i]);
        a[i] = re * real_scalars[i] - im * imag_scalars[i];
        a[points + i] = re * imag_scalars[i] + im * real_scalars[i];
    }
    if o.polar {
        to_polar(a, o.min_mag);
    }
    if o.inverse && !disable_normalize {
        normalize(a, o.polar);
    }
}

/// The transform of real input of even length, through a complex one of half the points.
fn real_fft(input: &[f64], o: &Options) -> Vec<f64> {
    let n = input.len() / 2;
    let mut out = vec![0.0; input.len() * 2];
    let mut work = vec![0.0; n * 2];
    for i in 0..n {
        work[i] = input[2 * i];
        work[n + i] = input[2 * i + 1];
    }
    let tfs = Twiddle::new(n * 2, o.inverse);
    let (next_pow2, _) = round_up_pow2(n);
    let plain = Options {
        inverse: o.inverse,
        polar: false,
        min_mag: 0.0,
    };
    if next_pow2 == n {
        fft2(&mut work, &tfs, &plain, true, true);
    } else {
        bluestein(&mut work, false, &plain, true);
    }
    let (two_n, three_n) = (2 * n, 3 * n);
    for i in 0..n {
        let rev = if i != 0 { n - i } else { 0 };
        let (y1r, y2r) = (work[i], work[rev]);
        let (y1i, y2i) = (work[n + i], work[n + rev]);
        let (wr, wi) = (tfs.re[i], tfs.im[i]);
        let res_r = 0.5 * (y1r + y2r + wr * (y1i + y2i) + wi * (y1r - y2r));
        out[i] = res_r;
        let res_i = 0.5 * (y1i - y2i + wi * (y1i + y2i) - wr * (y1r - y2r));
        out[two_n + i] = res_i;
        if i != 0 {
            out[n + rev] = res_r;
            out[three_n + rev] = -res_i;
        } else {
            out[n] = y1r - y1i;
            out[three_n] = 0.0;
        }
    }
    if o.polar {
        to_polar(&mut out, o.min_mag);
    }
    if o.inverse {
        normalize(&mut out, o.polar);
    }
    out
}

/// The transform of `data` (the real parts, then the imaginary parts if the input is complex);
/// the result has the same layout, for `rows` points of 2 columns.
fn transform(mut data: Vec<f64>, real: bool, o: &Options) -> Vec<f64> {
    let points = if real { data.len() } else { data.len() / 2 };
    if points == 1 {
        let mut out = vec![data[0], if real { 0.0 } else { data[1] }];
        if o.polar {
            to_polar(&mut out, o.min_mag);
        }
        return out;
    }
    if real && points % 2 == 0 {
        return real_fft(&data, o);
    }
    let (next_pow2, _) = round_up_pow2(points);
    if next_pow2 == points && !real {
        let tf = Twiddle::new(points, o.inverse);
        fft2(&mut data, &tf, o, false, false);
        return data;
    }
    if real {
        data.resize(points * 2, 0.0);
    }
    bluestein(&mut data, real, o, false);
    data
}

/// A text that LibreOffice reads as a flag or a number: `TRUE` and `FALSE`, or the number it holds.
fn text_flag(e: &Engine, t: &str) -> Option<f64> {
    if t.eq_ignore_ascii_case("TRUE") {
        Some(1.0)
    } else if t.eq_ignore_ascii_case("FALSE") {
        Some(0.0)
    } else {
        crate::coerce::to_num(&Value::text(t), &e.opts.locale, e.opts.date_ctx()).ok()
    }
}

/// The value of an argument that is a flag or a number: a text that is neither is 0 for the
/// optional arguments (which LibreOffice reads without complaint) and an error for the grouping.
fn flag(e: &mut Engine, a: &[Value], i: usize, optional: bool) -> Res<Option<f64>> {
    match a.get(i) {
        None | Some(Value::Empty) => Ok(None),
        Some(v) => match e.scalar(v.clone()) {
            Value::Error(k) => Err(k),
            Value::Text(t) => match text_flag(e, &t) {
                Some(x) => Ok(Some(x)),
                None if optional => Ok(Some(0.0)),
                None => Err(ErrorKind::Value),
            },
            s => Ok(Some(e.num(&s)?)),
        },
    }
}

/// The first of the arguments that is an error is the result, before anything else is looked at.
fn first_error(e: &mut Engine, a: &[Value]) -> Res<()> {
    for v in a {
        let v = match v {
            Value::Ref(r) if r.is_single() => e.scalar(v.clone()),
            v => v.clone(),
        };
        if let Value::Error(k) = v {
            return Err(k);
        }
    }
    Ok(())
}

/// `literal`: the series is an array written in the formula.
fn fourier(e: &mut Engine, a: &[Value], literal: bool) -> Res<Value> {
    {
        first_error(e, a)?;
        let min_mag = flag(e, a, 4, true)?.unwrap_or(0.0);
        let polar = flag(e, a, 3, true)?.is_some_and(|x| x != 0.0);
        let inverse = flag(e, a, 2, true)?.is_some_and(|x| x != 0.0);
        let by_column = flag(e, a, 1, false)?.is_some_and(|x| x != 0.0);
        // LibreOffice transposes a series grouped by rows in the matrix itself, which for an array
        // written in the formula is the formula's own: each calculation flips it.
        if literal && !by_column {
            crate::notes::note(crate::notes::FOURIER_LITERAL);
            return Err(ErrorKind::Unsupported);
        }
        let m = e.array_of(a[0].clone());
        let (rows, cols) = (m.rows as usize, m.cols as usize);
        if (by_column && cols > 2) || (!by_column && rows > 2) {
            return Err(DOMAIN);
        }
        let mut cells = Vec::with_capacity(rows * cols);
        for v in &m.data {
            match v {
                Value::Number(x) => cells.push(*x),
                Value::Bool(b) => cells.push(f64::from(u8::from(*b))),
                _ => return Err(DOMAIN),
            }
        }
        // Column by column; a series by rows is transposed.
        let (data, real, points) = if by_column {
            let mut d = Vec::with_capacity(cells.len());
            for c in 0..cols {
                for r in 0..rows {
                    d.push(cells[r * cols + c]);
                }
            }
            (d, cols == 1, rows)
        } else {
            let mut d = Vec::with_capacity(cells.len());
            for r in 0..rows {
                for c in 0..cols {
                    d.push(cells[r * cols + c]);
                }
            }
            (d, rows == 1, cols)
        };
        let out = transform(
            data,
            real,
            &Options {
                inverse,
                polar,
                min_mag,
            },
        );
        let mut values = Vec::with_capacity(points * 2);
        for i in 0..points {
            values.push(Value::Number(out[i]));
            values.push(Value::Number(out[points + i]));
        }
        Ok(Value::Array(Arc::new(Array::new(points as u32, 2, values))))
    }
}

pub(super) fn register(r: &mut Reg) {
    r.lazy("FOURIER", 2, 5, |e, a| {
        let vals = e.eval_array_args(a);
        fourier(e, &vals, matches!(a[0], Expr::Const(Value::Array(_))))
    });
}
