//! `FORECAST.ETS.ADD`, `FORECAST.ETS.MULT`, `FORECAST.ETS.PI.ADD`, `FORECAST.ETS.PI.MULT`,
//! `FORECAST.ETS.STAT.ADD`, `FORECAST.ETS.STAT.MULT` and `FORECAST.ETS.SEASONALITY`: the
//! Holt-Winters exponential smoothing of LibreOffice (`ScETSForecastCalculation` of
//! `sc/source/core/tool/interpr8.cxx`), with its search of the smoothing constants, its handling
//! of gaps, repeated timeline values and month steps, and the Kahan sums it computes with. The
//! results are checked against the oracle (`tests/assets/semantics.*`).
//!
//! Not reproduced: the prediction interval of a series with seasons, which LibreOffice computes
//! from random scenarios (the result changes with each calculation): the build stops.

use std::sync::Arc;

use super::Reg;
use super::math_ext::norm_inv;
use crate::dates;
use crate::engine::{Engine, Res};
use crate::ir::Ns;
use crate::value::{Array, ErrorKind, Value};

/// The errors of LibreOffice's calculation, by the error value a file stores for them.
const ILLEGAL_FP: ErrorKind = ErrorKind::Num;
const NO_VALUE: ErrorKind = ErrorKind::Value;
const DIV_ZERO: ErrorKind = ErrorKind::DivZero;
const ILLEGAL_ARG: ErrorKind = ErrorKind::Domain;

/// The smallest change of the smoothing constants that is searched for.
const MIN_ABC_RESOLUTION: f64 = 0.001;

#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Add,
    Mult,
    PiAdd,
    PiMult,
    StatAdd,
    StatMult,
    Season,
}

impl Kind {
    fn pi(self) -> bool {
        matches!(self, Self::PiAdd | Self::PiMult)
    }

    fn stat(self) -> bool {
        matches!(self, Self::StatAdd | Self::StatMult)
    }

    fn additive(self) -> bool {
        matches!(self, Self::Add | Self::PiAdd | Self::StatAdd)
    }
}

/// `KahanSum` of LibreOffice: Neumaier's sum with the last term kept apart.
#[derive(Clone, Copy, Default)]
struct Kahan {
    sum: f64,
    err: f64,
    mem: f64,
}

impl Kahan {
    fn new(x: f64) -> Self {
        Self {
            sum: x,
            err: 0.0,
            mem: 0.0,
        }
    }

    fn neumaier(sum: &mut f64, err: &mut f64, value: f64) {
        let t = *sum + value;
        if sum.abs() >= value.abs() {
            *err += (*sum - t) + value;
        } else {
            *err += (value - t) + *sum;
        }
        *sum = t;
    }

    fn add(&mut self, x: f64) {
        if x == 0.0 {
            return;
        }
        if self.mem == 0.0 {
            self.mem = x;
            return;
        }
        Self::neumaier(&mut self.sum, &mut self.err, self.mem);
        self.mem = x;
    }

    fn get(&mut self) -> f64 {
        let total = self.sum + self.err;
        if self.mem == 0.0 {
            return total;
        }
        if ((self.mem < 0.0 && total > 0.0) || (total < 0.0 && self.mem > 0.0))
            && approx_equal(self.mem, -total)
        {
            return 0.0;
        }
        let m = self.mem;
        self.add(m);
        self.mem = 0.0;
        self.sum + self.err
    }

    fn div(&mut self, d: f64) {
        if self.mem != 0.0 {
            self.sum = self.get() / d;
            self.mem = 0.0;
        } else {
            self.sum = (self.sum + self.err) / d;
        }
        self.err = 0.0;
    }
}

/// `rtl::math::approxEqual`.
fn approx_equal(a: f64, b: f64) -> bool {
    const E48: f64 = 1.0 / 281_474_976_710_656.0;
    const HALF_15TH: f64 = 5e-15;
    if a == b {
        return true;
    }
    if a == 0.0 || b == 0.0 || a.is_sign_negative() != b.is_sign_negative() {
        return false;
    }
    let d = (a - b).abs();
    if !d.is_finite() {
        return false;
    }
    let (a, b) = (a.abs(), b.abs());
    let min_ab = a.min(b);
    let threshold1 = min_ab * E48;
    let threshold2 = format!("1e{}", min_ab.log10().floor() as i32)
        .parse::<f64>()
        .unwrap_or(0.0)
        * HALF_15TH;
    if d >= threshold1.max(threshold2) {
        return false;
    }
    let representable = |x: f64| x < 9_007_199_254_740_992.0 && x.trunc() == x;
    !(representable(a) && representable(b))
}

/// `std::sort` of libstdc++ (introsort) on the data points by timeline value, whose order of
/// equal values is the one LibreOffice's calculation depends on. `false` where it would fall back
/// to heap sort (not reproduced).
fn std_sort(v: &mut [(f64, f64)]) -> bool {
    let n = v.len();
    if n > 1 {
        let depth = 2 * (usize::BITS - 1 - n.leading_zeros());
        if !introsort_loop(v, 0, n, depth) {
            return false;
        }
        final_insertion_sort(v);
    }
    true
}

fn less(a: &(f64, f64), b: &(f64, f64)) -> bool {
    a.0 < b.0
}

fn introsort_loop(v: &mut [(f64, f64)], first: usize, mut last: usize, mut depth: u32) -> bool {
    while last - first > 16 {
        if depth == 0 {
            return false;
        }
        depth -= 1;
        // The median of the first, the middle and the last is moved to the front, as the pivot.
        let mid = first + (last - first) / 2;
        let (a, b, c) = (first + 1, mid, last - 1);
        let pick = if less(&v[a], &v[b]) {
            if less(&v[b], &v[c]) {
                b
            } else if less(&v[a], &v[c]) {
                c
            } else {
                a
            }
        } else if less(&v[a], &v[c]) {
            a
        } else if less(&v[b], &v[c]) {
            c
        } else {
            b
        };
        v.swap(first, pick);
        let mut lo = first + 1;
        let mut hi = last;
        let cut = loop {
            while less(&v[lo], &v[first]) {
                lo += 1;
            }
            hi -= 1;
            while less(&v[first], &v[hi]) {
                hi -= 1;
            }
            if lo >= hi {
                break lo;
            }
            v.swap(lo, hi);
            lo += 1;
        };
        if !introsort_loop(v, cut, last, depth) {
            return false;
        }
        last = cut;
    }
    true
}

fn unguarded_linear_insert(v: &mut [(f64, f64)], at: usize) {
    let val = v[at];
    let mut last = at;
    while last > 0 && less(&val, &v[last - 1]) {
        v[last] = v[last - 1];
        last -= 1;
    }
    v[last] = val;
}

fn insertion_sort(v: &mut [(f64, f64)], first: usize, last: usize) {
    for i in first + 1..last {
        if less(&v[i], &v[first]) {
            let val = v[i];
            v.copy_within(first..i, first + 1);
            v[first] = val;
        } else {
            unguarded_linear_insert(v, i);
        }
    }
}

fn final_insertion_sort(v: &mut [(f64, f64)]) {
    let n = v.len();
    if n > 16 {
        insertion_sort(v, 0, 16);
        for i in 16..n {
            unguarded_linear_insert(v, i);
        }
    } else {
        insertion_sort(v, 0, n);
    }
}

struct Ets {
    sys: crate::options::DateSystem,
    range: Vec<(f64, f64)>,
    base: Vec<f64>,
    trend: Vec<f64>,
    per_idx: Vec<f64>,
    forecast: Vec<f64>,
    smpl_in_prd: usize,
    step: f64,
    alpha: f64,
    beta: f64,
    gamma: f64,
    count: usize,
    initialised: bool,
    month_day: i64,
    mae: f64,
    mase: f64,
    mse: f64,
    rmse: f64,
    smape: f64,
    additive: bool,
    eds: bool,
}

type Chk = Result<(), ErrorKind>;

/// The arguments of the calculation: the seasonality, the data completion, the aggregation of
/// repeated timeline values, the first target and the function.
#[derive(Clone, Copy)]
struct Options {
    smpl: i64,
    completion: bool,
    aggregation: i64,
    first_target: Option<f64>,
    kind: Kind,
}

impl Ets {
    fn date(&self, x: f64) -> Res<(i64, i64, i64)> {
        dates::serial_to_ymd(x.trunc(), self.sys).ok_or(ILLEGAL_ARG)
    }

    fn preprocess(&mut self, xs: &[f64], ys: &[f64], o: &Options) -> Chk {
        let Options {
            smpl,
            completion,
            aggregation,
            first_target,
            kind,
        } = *o;
        self.eds = smpl == 0;
        self.additive = kind.additive();
        self.range = xs.iter().zip(ys).map(|(x, y)| (*x, *y)).collect();
        if !std_sort(&mut self.range) {
            crate::notes::note(crate::notes::ETS_SORT);
            return Err(ErrorKind::Unsupported);
        }
        if let Some(t) = first_target {
            let limit = if kind.pi() {
                self.range[self.count - 1].0
            } else {
                self.range[0].0
            };
            if t < limit {
                return Err(ILLEGAL_FP);
            }
        }
        // A step of months has no exact size: detect it and take the months as the timeline.
        let first = self.date(self.range[0].0)?;
        self.month_day = first.2;
        let first_days = self.range[0].0.trunc();
        let mut i = 1;
        while i < self.count && self.month_day != 0 {
            let x = self.range[i].0.trunc();
            if x != first_days {
                let d = self.date(self.range[i].0)?;
                if d.2 != self.month_day {
                    self.month_day = 0;
                }
            }
            i += 1;
        }
        self.step = f64::MAX;
        if self.month_day != 0 {
            for i in 0..self.count {
                let d = self.date(self.range[i].0)?;
                self.range[i].0 = (d.0 * 12 + d.1) as f64;
            }
        }
        let mut i = 1;
        while i < self.count {
            let mut step = self.range[i].0 - self.range[i - 1].0;
            if step == 0.0 {
                if aggregation == 0 {
                    return Err(NO_VALUE);
                }
                let mut tmp = self.range[i - 1].1;
                let mut counter = 1usize;
                match aggregation {
                    1 => {
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            self.range.remove(i);
                            self.count -= 1;
                        }
                    }
                    7 => {
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            tmp += self.range[i].1;
                            self.range.remove(i);
                            self.count -= 1;
                        }
                        self.range[i - 1].1 = tmp;
                    }
                    2 | 3 => {
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            counter += 1;
                            self.range.remove(i);
                            self.count -= 1;
                        }
                        self.range[i - 1].1 = counter as f64;
                    }
                    4 => {
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            if self.range[i].1 > tmp {
                                tmp = self.range[i].1;
                            }
                            self.range.remove(i);
                            self.count -= 1;
                        }
                        self.range[i - 1].1 = tmp;
                    }
                    5 => {
                        let mut values = vec![self.range[i - 1].1];
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            values.push(self.range[i].1);
                            counter += 1;
                            self.range.remove(i);
                            self.count -= 1;
                        }
                        values.sort_by(f64::total_cmp);
                        self.range[i - 1].1 = if counter % 2 == 1 {
                            values[counter / 2]
                        } else {
                            (values[counter / 2] + values[counter / 2 - 1]) / 2.0
                        };
                    }
                    6 => {
                        while i < self.count && self.range[i].0 == self.range[i - 1].0 {
                            if self.range[i].1 < tmp {
                                tmp = self.range[i].1;
                            }
                            self.range.remove(i);
                            self.count -= 1;
                        }
                        self.range[i - 1].1 = tmp;
                    }
                    _ => {}
                }
                step = if i + 1 < self.count {
                    self.range[i].0 - self.range[i - 1].0
                } else {
                    self.step
                };
            }
            if step > 0.0 && step < self.step {
                self.step = step;
            }
            i += 1;
        }
        // The step is constant, or a gap is a multiple of it.
        let mut has_gap = false;
        let mut i = 1;
        while i < self.count && !has_gap {
            let step = self.range[i].0 - self.range[i - 1].0;
            if step != self.step {
                if step % self.step != 0.0 {
                    return Err(NO_VALUE);
                }
                has_gap = true;
            }
            i += 1;
        }
        if has_gap {
            let mut missing = 0usize;
            let original = self.count as f64;
            let mut date = if self.month_day != 0 {
                Some(self.date(self.range[0].0)?)
            } else {
                None
            };
            let mut i = 1;
            while i < self.count {
                let dist;
                if self.month_day != 0 {
                    let d1 = self.date(self.range[i].0)?;
                    let d0 = date.unwrap_or(d1);
                    dist = (12 * (d1.0 - d0.0) + (d1.1 - d0.1)) as f64;
                    date = Some(d1);
                } else {
                    dist = self.range[i].0 - self.range[i - 1].0;
                }
                if dist > self.step {
                    let gap_y = (self.range[i].1 + self.range[i - 1].1) / 2.0;
                    let mut x_gap = Kahan::new(self.range[i - 1].0 + self.step);
                    while x_gap.get() < self.range[i].0 {
                        let x = x_gap.get();
                        self.range
                            .insert(i, (x, if completion { gap_y } else { 0.0 }));
                        i += 1;
                        self.count += 1;
                        missing += 1;
                        if missing as f64 / original > 0.3 {
                            return Err(NO_VALUE);
                        }
                        x_gap.add(self.step);
                    }
                }
                i += 1;
            }
        }
        if smpl != 1 {
            self.smpl_in_prd = smpl as usize;
        } else {
            self.smpl_in_prd = self.calc_period_len();
            if self.smpl_in_prd == 1 {
                self.eds = true;
            }
        }
        self.init_data()
    }

    fn init_data(&mut self) -> Chk {
        let n = self.count;
        self.base = vec![0.0; n];
        self.trend = vec![0.0; n];
        if !self.eds {
            self.per_idx = vec![0.0; n];
        }
        self.forecast = vec![0.0; n];
        self.forecast[0] = self.range[0].1;
        self.prefill_trend()?;
        self.prefill_per_idx()?;
        self.base[0] = if self.eds {
            self.range[0].1
        } else {
            self.range[0].1 / self.per_idx[0]
        };
        Ok(())
    }

    fn prefill_trend(&mut self) -> Chk {
        if self.eds {
            self.trend[0] =
                (self.range[self.count - 1].1 - self.range[0].1) / (self.count - 1) as f64;
        } else {
            // At least two periods are needed.
            if self.count < 2 * self.smpl_in_prd {
                return Err(NO_VALUE);
            }
            let mut sum = Kahan::default();
            for i in 0..self.smpl_in_prd {
                sum.add(self.range[i + self.smpl_in_prd].1);
                sum.add(-self.range[i].1);
            }
            self.trend[0] = sum.get() / (self.smpl_in_prd * self.smpl_in_prd) as f64;
        }
        Ok(())
    }

    fn prefill_per_idx(&mut self) -> Chk {
        if self.eds {
            return Ok(());
        }
        if self.smpl_in_prd == 0 {
            return Err(NO_VALUE);
        }
        let s = self.smpl_in_prd;
        let periods = self.count / s;
        let mut average = vec![Kahan::default(); periods];
        for (i, avg) in average.iter_mut().enumerate() {
            for j in 0..s {
                avg.add(self.range[i * s + j].1);
            }
            avg.div(s as f64);
            if avg.get() == 0.0 {
                return Err(DIV_ZERO);
            }
        }
        for j in 0..s {
            let mut sum = Kahan::default();
            for (i, avg) in average.iter_mut().enumerate() {
                let adj = avg.get() + (j as f64 - 0.5 * (s as f64 - 1.0)) * self.trend[0];
                if self.additive {
                    sum.add(self.range[i * s + j].1 - adj);
                } else {
                    sum.add(self.range[i * s + j].1 / adj);
                }
            }
            self.per_idx[j] = sum.get() / periods as f64;
        }
        if s < self.count {
            self.per_idx[s] = 0.0;
        }
        Ok(())
    }

    fn init_calc(&mut self) {
        if !self.initialised {
            self.calc_alpha_beta_gamma();
            self.initialised = true;
            self.calc_accuracy();
        }
    }

    fn calc_accuracy(&mut self) {
        let (mut abs_err, mut divisor, mut err_sq, mut abs_perc) = (
            Kahan::default(),
            Kahan::default(),
            Kahan::default(),
            Kahan::default(),
        );
        for i in 1..self.count {
            let e = self.forecast[i] - self.range[i].1;
            abs_err.add(e.abs());
            err_sq.add(e * e);
            abs_perc.add(e.abs() / (self.forecast[i].abs() + self.range[i].1.abs()));
        }
        for i in 2..self.count {
            divisor.add((self.range[i].1 - self.range[i - 1].1).abs());
        }
        let n = (self.count - 1) as f64;
        let abs_err = abs_err.get();
        self.mae = abs_err / n;
        self.mase = abs_err / (n * divisor.get() / (n - 1.0));
        self.mse = err_sq.get() / n;
        self.rmse = self.mse.sqrt();
        self.smape = abs_perc.get() * 2.0 / n;
    }

    /// The most likely length of a period.
    fn calc_period_len(&self) -> usize {
        let count = self.count;
        let mut best = count;
        let mut best_me = f64::MAX;
        let mut len = count / 2;
        while len >= 1 {
            let mut mean_error = Kahan::default();
            let periods = count / len;
            let start = count - periods * len + 1;
            let mut i = start;
            while i < count - len {
                mean_error.add(
                    ((self.range[i].1 - self.range[i - 1].1)
                        - (self.range[len + i].1 - self.range[len + i - 1].1))
                        .abs(),
                );
                i += 1;
            }
            let denominator = ((periods - 1).wrapping_mul(len)).wrapping_sub(1);
            let me = mean_error.get() / denominator as f64;
            if me <= best_me || me == 0.0 {
                best = len;
                best_me = me;
            }
            len -= 1;
        }
        best
    }

    fn fit(&mut self, which: u8) {
        // One search of the constant `which` (0 alpha, 1 beta, 2 gamma) on 0…1, bisecting on the
        // mean squared error, with the inner constants searched at each step.
        let set = |s: &mut Self, v: f64| match which {
            0 => s.alpha = v,
            1 => s.beta = v,
            _ => s.gamma = v,
        };
        let inner = |s: &mut Self| match which {
            0 => {
                if s.eds {
                    s.calc_gamma();
                } else {
                    s.calc_beta_gamma();
                }
            }
            1 => s.calc_gamma(),
            _ => {}
        };
        let mut f0 = 0.0;
        set(self, f0);
        if which == 0 && self.eds {
            self.beta = 0.0;
        }
        inner(self);
        self.refill();
        let mut e0 = self.mse;
        let mut f2 = 1.0;
        set(self, f2);
        inner(self);
        self.refill();
        let mut e2 = self.mse;
        let mut f1 = 0.5;
        set(self, f1);
        inner(self);
        self.refill();
        if e0 == self.mse && self.mse == e2 {
            set(self, 0.0);
            inner(self);
            self.refill();
            return;
        }
        while (f2 - f1) > MIN_ABC_RESOLUTION {
            if e2 > e0 {
                f2 = f1;
                e2 = self.mse;
                f1 = (f0 + f1) / 2.0;
            } else {
                f0 = f1;
                e0 = self.mse;
                f1 = (f1 + f2) / 2.0;
            }
            set(self, f1);
            inner(self);
            self.refill();
        }
        if e2 > e0 {
            if e0 < self.mse {
                set(self, f0);
                inner(self);
                self.refill();
            }
        } else if e2 < self.mse {
            set(self, f2);
            inner(self);
            self.refill();
        }
        if which == 0 {
            self.calc_accuracy();
        }
    }

    fn calc_alpha_beta_gamma(&mut self) {
        self.fit(0);
    }

    fn calc_beta_gamma(&mut self) {
        self.fit(1);
    }

    fn calc_gamma(&mut self) {
        self.fit(2);
    }

    /// Fill the base, trend, seasonal values and forecasts with the smoothing constants.
    fn refill(&mut self) {
        for i in 1..self.count {
            let y = self.range[i].1;
            if self.eds {
                self.base[i] =
                    self.alpha * y + (1.0 - self.alpha) * (self.base[i - 1] + self.trend[i - 1]);
                self.trend[i] = self.gamma * (self.base[i] - self.base[i - 1])
                    + (1.0 - self.gamma) * self.trend[i - 1];
                self.forecast[i] = self.base[i - 1] + self.trend[i - 1];
            } else {
                let s = self.smpl_in_prd;
                let idx;
                if self.additive {
                    idx = if i > s { i - s } else { i };
                    self.base[i] = self.alpha * (y - self.per_idx[idx])
                        + (1.0 - self.alpha) * (self.base[i - 1] + self.trend[i - 1]);
                    self.per_idx[i] =
                        self.beta * (y - self.base[i]) + (1.0 - self.beta) * self.per_idx[idx];
                } else {
                    idx = if i >= s { i - s } else { i };
                    self.base[i] = self.alpha * (y / self.per_idx[idx])
                        + (1.0 - self.alpha) * (self.base[i - 1] + self.trend[i - 1]);
                    self.per_idx[i] =
                        self.beta * (y / self.base[i]) + (1.0 - self.beta) * self.per_idx[idx];
                }
                self.trend[i] = self.gamma * (self.base[i] - self.base[i - 1])
                    + (1.0 - self.gamma) * self.trend[i - 1];
                self.forecast[i] = if self.additive {
                    self.base[i - 1] + self.trend[i - 1] + self.per_idx[idx]
                } else {
                    (self.base[i - 1] + self.trend[i - 1]) * self.per_idx[idx]
                };
            }
        }
        self.calc_accuracy();
    }

    fn to_months(&self, x: f64) -> Res<f64> {
        let (y, m, d) = self.date(x)?;
        let length = match m {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31.0,
            2 => {
                if dates::is_leap(y) {
                    29.0
                } else {
                    28.0
                }
            }
            _ => 30.0,
        };
        Ok(12.0 * y as f64 + m as f64 + (d - self.month_day) as f64 / length)
    }

    fn target(&self, t: f64) -> Res<f64> {
        if self.month_day != 0 {
            self.to_months(t)
        } else {
            Ok(t)
        }
    }

    /// What LibreOffice does when it reads beyond an array is not defined.
    fn beyond<T>() -> Res<T> {
        crate::notes::note(crate::notes::ETS_BOUNDS);
        Err(ErrorKind::Unsupported)
    }

    fn get_forecast(&mut self, target: f64) -> Res<f64> {
        self.init_calc();
        let last = self.count - 1;
        if target <= self.range[last].0 {
            let n = ((target - self.range[0].0) / self.step) as usize;
            let interpolate = (target - self.range[0].0) % self.step;
            if n >= self.count || (interpolate >= MIN_ABC_RESOLUTION && n + 1 >= self.count) {
                return Self::beyond();
            }
            let mut f = self.range[n].1;
            if interpolate >= MIN_ABC_RESOLUTION {
                let factor = interpolate / self.step;
                let fc1 = self.forecast[n + 1];
                f += factor * (fc1 - f);
            }
            Ok(f)
        } else {
            let n = ((target - self.range[last].0) / self.step) as usize;
            let interpolate = (target - self.range[last].0) % self.step;
            let s = self.smpl_in_prd;
            if !self.eds && (s == 0 || last < s) {
                return Self::beyond();
            }
            let at = |k: usize| last.wrapping_sub(s).wrapping_add(k % s.max(1));
            let mut f = if self.eds {
                self.base[last] + n as f64 * self.trend[last]
            } else if self.additive {
                self.base[last] + n as f64 * self.trend[last] + self.per_idx[at(n)]
            } else {
                (self.base[last] + n as f64 * self.trend[last]) * self.per_idx[at(n)]
            };
            if interpolate >= MIN_ABC_RESOLUTION {
                let factor = interpolate / self.step;
                let fc1 = if self.eds {
                    self.base[last] + (n + 1) as f64 * self.trend[last]
                } else if self.additive {
                    self.base[last] + (n + 1) as f64 * self.trend[last] + self.per_idx[at(n + 1)]
                } else {
                    (self.base[last] + (n + 1) as f64 * self.trend[last]) * self.per_idx[at(n + 1)]
                };
                f += factor * (fc1 - f);
            }
            Ok(f)
        }
    }

    fn statistic(&mut self, kind: i64) -> f64 {
        self.init_calc();
        match kind {
            1 => self.alpha,
            2 => self.gamma,
            3 => self.beta,
            4 => self.mase,
            5 => self.smape,
            6 => self.mae,
            7 => self.rmse,
            8 => self.step,
            _ => self.smpl_in_prd as f64,
        }
    }

    /// The prediction interval of a series without seasons (the double exponential smoothing).
    fn eds_interval(&mut self, targets: &[f64], level: f64) -> Res<Vec<f64>> {
        self.init_calc();
        let mut max = targets[0];
        for &t in targets {
            if max < t {
                max = t;
            }
        }
        let last = self.range[self.count - 1].0;
        max = self.target(max)? - last;
        let mut size = (max / self.step) as usize;
        if max % self.step != 0.0 {
            size += 1;
        }
        if size == 0 {
            return Err(ILLEGAL_ARG);
        }
        let z = norm_inv((1.0 + level) / 2.0);
        let o = 1.0 - level;
        let c: Vec<f64> = (0..size)
            .map(|i| {
                let i = i as f64;
                (1.0 + (level / (1.0 + o).powf(3.0))
                    * ((1.0 + 4.0 * o + 5.0 * o * o)
                        + 2.0 * i * level * (1.0 + 3.0 * o)
                        + 2.0 * (i * i) * level * level))
                    .sqrt()
            })
            .collect();
        let mut out = Vec::with_capacity(targets.len());
        for &t in targets {
            let t = self.target(t)? - last;
            let steps = ((t / self.step) - 1.0) as usize;
            let factor = t % self.step;
            if steps >= c.len() || (factor != 0.0 && steps + 1 >= c.len()) {
                return Self::beyond();
            }
            let mut pi = z * self.rmse * c[steps] / c[0];
            if factor != 0.0 {
                let pi1 = z * self.rmse * c[steps + 1] / c[0];
                pi += factor * (pi1 - pi);
            }
            out.push(pi);
        }
        Ok(out)
    }
}

/// The numbers of a matrix argument, column by column, with its dimensions; `None` where a cell
/// is not a number.
fn matrix(e: &mut Engine, v: &Value) -> (Vec<Option<f64>>, usize, usize) {
    let m = e.array_of(v.clone());
    let (rows, cols) = (m.rows as usize, m.cols as usize);
    let mut out = Vec::with_capacity(rows * cols);
    for c in 0..cols {
        for r in 0..rows {
            out.push(match &m.data[r * cols + c] {
                Value::Number(x) => Some(*x),
                Value::Bool(b) => Some(f64::from(u8::from(*b))),
                _ => None,
            });
        }
    }
    (out, rows, cols)
}

/// An optional numeric argument, read as `GetDoubleWithDefault`.
fn with_default(e: &mut Engine, a: &[Value], i: usize, default: f64) -> Res<f64> {
    match a.get(i) {
        None | Some(Value::Empty) => Ok(default),
        Some(v) => e.num(v),
    }
}

/// The first of the arguments that is an error is the result.
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

fn ets(e: &mut Engine, a: &[Value], kind: Kind) -> Res<Value> {
    first_error(e, a)?;
    let n = a.len();
    let pi = kind.pi();
    let season = kind == Kind::Season;
    let mut at = n;
    let aggregation = if (n == 6 && !pi) || (n == 4 && season) || n == 7 {
        at -= 1;
        with_default(e, a, at, 1.0)?.trunc() as i64
    } else {
        1
    };
    if !(1..=7).contains(&aggregation) {
        return Err(ILLEGAL_ARG);
    }
    let completion = if (n >= 5 && !pi) || (n >= 3 && season) || (n >= 6 && pi) {
        at -= 1;
        match with_default(e, a, at, 1.0)?.trunc() as i64 {
            0 => false,
            1 => true,
            _ => return Err(ILLEGAL_ARG),
        }
    } else {
        true
    };
    let smpl = if ((n >= 4 && !pi) || (n >= 5 && pi)) && !season {
        at -= 1;
        let v = with_default(e, a, at, 1.0)?;
        if v % 1.0 != 0.0 || v < 0.0 {
            return Err(ILLEGAL_FP);
        }
        v as i64
    } else {
        1
    };
    let mut level = 0.0;
    if pi {
        level = if n < 4 {
            0.95
        } else {
            at -= 1;
            with_default(e, a, at, 0.95)?
        };
        if !(0.0..=1.0).contains(&level) {
            return Err(ILLEGAL_ARG);
        }
    }
    // The arguments that are left are, from the last: the statistic types, the timeline, the
    // values and the targets.
    let mut types = None;
    if kind.stat() {
        at -= 1;
        let (t, rows, cols) = matrix(e, &a[at]);
        for x in &t {
            let k = x.unwrap_or(0.0) as i64;
            if !(1..=9).contains(&k) {
                return Err(ILLEGAL_ARG);
            }
        }
        types = Some((t, rows, cols));
    }
    at -= 1;
    let (xs, rx, cx) = matrix(e, &a[at]);
    at -= 1;
    let (ys, ry, cy) = matrix(e, &a[at]);
    if rx != ry || cx != cy || xs.iter().chain(&ys).any(Option::is_none) {
        return Err(ILLEGAL_ARG);
    }
    let (xs, ys): (Vec<f64>, Vec<f64>) = (
        xs.into_iter().flatten().collect(),
        ys.into_iter().flatten().collect(),
    );
    let mut targets = None;
    if !kind.stat() && !season {
        at -= 1;
        targets = Some(matrix(e, &a[at]));
    }
    let first_target = targets
        .as_ref()
        .and_then(|t| t.0.first().copied().flatten());
    let mut calc = Ets {
        sys: e.opts.date_system,
        range: Vec::new(),
        base: Vec::new(),
        trend: Vec::new(),
        per_idx: Vec::new(),
        forecast: Vec::new(),
        smpl_in_prd: 0,
        step: 0.0,
        alpha: 0.0,
        beta: 0.0,
        gamma: 0.0,
        count: xs.len(),
        initialised: false,
        month_day: 0,
        mae: 0.0,
        mase: 0.0,
        mse: 0.0,
        rmse: 0.0,
        smape: 0.0,
        additive: false,
        eds: false,
    };
    let first = if kind.stat() { None } else { first_target };
    calc.preprocess(
        &xs,
        &ys,
        &Options {
            smpl,
            completion,
            aggregation,
            first_target: first,
            kind,
        },
    )?;
    let array = |rows: usize, cols: usize, data: Vec<f64>| {
        Value::Array(Arc::new(Array::new(
            rows as u32,
            cols as u32,
            data.into_iter()
                .map(|x| {
                    if x.is_finite() {
                        Value::Number(x)
                    } else {
                        Value::Error(ILLEGAL_FP)
                    }
                })
                .collect(),
        )))
    };
    match kind {
        Kind::Add | Kind::Mult => {
            let (t, rows, cols) = targets.unwrap_or_default();
            let mut out = vec![0.0; rows * cols];
            for r in 0..rows {
                for c in 0..cols {
                    let target = t[c * rows + r].unwrap_or(f64::NAN);
                    let target = calc.target(target)?;
                    out[r * cols + c] = calc.get_forecast(target)?;
                }
            }
            Ok(array(rows, cols, out))
        }
        Kind::PiAdd | Kind::PiMult => {
            if smpl != 0 {
                crate::notes::note(crate::notes::ETS_RANDOM);
                return Err(ErrorKind::Unsupported);
            }
            let (t, rows, cols) = targets.unwrap_or_default();
            let flat: Vec<f64> = (0..rows)
                .flat_map(|r| (0..cols).map(move |c| (r, c)))
                .map(|(r, c)| t[c * rows + r].unwrap_or(f64::NAN))
                .collect();
            let out = calc.eds_interval(&flat, level)?;
            Ok(array(rows, cols, out))
        }
        Kind::StatAdd | Kind::StatMult => {
            let (t, rows, cols) = types.unwrap_or_default();
            let mut out = vec![0.0; rows * cols];
            for r in 0..rows {
                for c in 0..cols {
                    out[r * cols + c] = calc.statistic(t[c * rows + r].unwrap_or(0.0) as i64);
                }
            }
            Ok(array(rows, cols, out))
        }
        Kind::Season => Ok(Value::Number(calc.smpl_in_prd as f64)),
    }
}

pub(super) fn register(r: &mut Reg) {
    // A file names Excel's three by one name and an ODS by LibreOffice's (`FORECAST.ETS.ADD`).
    macro_rules! forecast {
        ($kind:expr, $min:expr, $max:expr, $($name:literal),+) => {
            $(r.eager_in(Ns::Microsoft, $name, $min, $max, |e, a| ets(e, a, $kind)).array_args();)+
        };
    }
    forecast!(Kind::Add, 3, 6, "FORECAST.ETS", "FORECAST.ETS.ADD");
    forecast!(
        Kind::PiAdd,
        3,
        7,
        "FORECAST.ETS.CONFINT",
        "FORECAST.ETS.PI.ADD"
    );
    forecast!(
        Kind::StatAdd,
        3,
        6,
        "FORECAST.ETS.STAT",
        "FORECAST.ETS.STAT.ADD"
    );
    forecast!(Kind::Season, 2, 4, "FORECAST.ETS.SEASONALITY");
    macro_rules! libreoffice {
        ($kind:expr, $min:expr, $max:expr, $name:literal) => {
            r.eager_in(Ns::LibreOffice, $name, $min, $max, |e, a| ets(e, a, $kind))
                .array_args();
        };
    }
    libreoffice!(Kind::Mult, 3, 6, "FORECAST.ETS.MULT");
    libreoffice!(Kind::PiMult, 3, 7, "FORECAST.ETS.PI.MULT");
    libreoffice!(Kind::StatMult, 3, 6, "FORECAST.ETS.STAT.MULT");
}
