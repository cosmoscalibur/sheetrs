//! Securities: the coupon functions and what is built on them (prices, yields, durations,
//! accrued interest and the discount instruments). What LibreOffice 26.8 answers for the
//! arguments outside the plain case is what `tests/assets/semantics.*` holds.

use super::{Reg, num};
use crate::dates;
use crate::engine::{Engine, Res};
use crate::options::DateSystem;
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// A date argument: the day, cut to a whole number.
fn date_arg(e: &mut Engine, v: &Value) -> Res<f64> {
    Ok(e.num(v)?.floor())
}

fn ymd(sys: DateSystem, serial: f64) -> Res<(i64, i64, i64)> {
    dates::serial_to_ymd(serial, sys).ok_or(DOMAIN)
}

/// 30/360 day count between two dates: the US rule, or the European one when `european`. In the
/// US rule the end day 31 is cut to 30 only when the start day, as given, is 30 or 31: a start
/// on the last day of February (28 or 29) leaves it at 31 (`DAYS360` does not do this).
fn days360(from: (i64, i64, i64), to: (i64, i64, i64), european: bool) -> f64 {
    let ((y1, m1, mut d1), (y2, m2, mut d2)) = (from, to);
    let last_feb = |y: i64, m: i64, d: i64| m == 2 && d == dates::days_in_month(y, 2);
    if european {
        d1 = d1.min(30);
        d2 = d2.min(30);
    } else {
        let both_feb = last_feb(y1, m1, d1) && last_feb(y2, m2, d2);
        let start_feb = last_feb(y1, m1, d1);
        if both_feb {
            d2 = 30;
        }
        if d2 == 31 && d1 >= 30 {
            d2 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
        if start_feb {
            d1 = 30;
        }
    }
    ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64
}

/// Days from `a` to `b` by a day-count basis (0 and 4: 30/360 US and European; else the days).
fn day_diff(sys: DateSystem, a: f64, b: f64, basis: i64) -> Res<f64> {
    if basis == 0 || basis == 4 {
        Ok(days360(ymd(sys, a)?, ymd(sys, b)?, basis == 4))
    } else {
        Ok(b - a)
    }
}

/// The coupon dates around a settlement and what the coupon functions answer from them.
struct Coupon {
    pcd: f64,
    ncd: f64,
    num: f64,
    /// Days from the previous coupon to the settlement.
    days_bs: f64,
    /// Days of the coupon period.
    days: f64,
    /// Days from the settlement to the next coupon.
    days_nc: f64,
}

fn frequency(e: &mut Engine, v: &Value) -> Res<i64> {
    let f = e.num(v)?.trunc() as i64;
    if matches!(f, 1 | 2 | 4) {
        Ok(f)
    } else {
        Err(DOMAIN)
    }
}

fn basis_of(e: &mut Engine, a: &[Value], i: usize) -> Res<i64> {
    let b = match super::opt(a, i) {
        Some(v) => e.num(v)?.trunc() as i64,
        None => 0,
    };
    if (0..=4).contains(&b) {
        Ok(b)
    } else {
        Err(DOMAIN)
    }
}

/// The coupon date `k` periods before `maturity`: the same day of the month, or the last day of
/// it when the maturity is one.
fn coupon_before(sys: DateSystem, maturity: (i64, i64, i64), k: i64, freq: i64) -> Res<f64> {
    let (y, m, d) = maturity;
    let eom = d == dates::days_in_month(y, m);
    let months = y * 12 + (m - 1) - k * (12 / freq);
    let (ny, nm) = (months.div_euclid(12), months.rem_euclid(12) + 1);
    let dim = dates::days_in_month(ny, nm);
    let nd = if eom { dim } else { d.min(dim) };
    dates::ymd_to_serial_quiet(ny, nm, nd, sys).ok_or(DOMAIN)
}

fn coupon(sys: DateSystem, settle: f64, maturity: f64, freq: i64, basis: i64) -> Res<Coupon> {
    if settle >= maturity {
        return Err(DOMAIN);
    }
    let mat = ymd(sys, maturity)?;
    let mut n = 0;
    let mut pcd = maturity;
    let mut ncd = maturity;
    while pcd > settle {
        n += 1;
        ncd = pcd;
        pcd = coupon_before(sys, mat, n, freq)?;
    }
    let days_bs = day_diff(sys, pcd, settle, basis)?;
    let days = match basis {
        1 => ncd - pcd,
        3 => 365.0 / freq as f64,
        _ => 360.0 / freq as f64,
    };
    let days_nc = if basis == 0 || basis == 4 {
        days - days_bs
    } else {
        day_diff(sys, settle, ncd, basis)?
    };
    Ok(Coupon {
        pcd,
        ncd,
        num: n as f64,
        days_bs,
        days,
        days_nc,
    })
}

fn coupon_of(e: &mut Engine, a: &[Value]) -> Res<Coupon> {
    let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
    let freq = frequency(e, &a[2])?;
    let basis = basis_of(e, a, 3)?;
    coupon(e.opts.date_system, s, m, freq, basis)
}

/// The years between two dates as LibreOffice's securities count them: the 30/360 bases with the
/// rule of `days360`, the others as `YEARFRAC`.
fn year_frac(e: &Engine, a: f64, b: f64, basis: i64) -> Res<f64> {
    let (a, b) = if a <= b { (a, b) } else { (b, a) };
    if basis == 0 || basis == 4 {
        let sys = e.opts.date_system;
        return Ok(days360(ymd(sys, a)?, ymd(sys, b)?, basis == 4) / 360.0);
    }
    super::date_ext::year_frac(e, a, b, basis)
}

/// The years between two dates for the instruments without coupons: the days over the days of a
/// year, which on the actual basis is the year of the first date (not the average of the years
/// the dates span).
fn year_diff(e: &Engine, a: f64, b: f64, basis: i64) -> Res<f64> {
    let sys = e.opts.date_system;
    Ok(match basis {
        1 => {
            let (y, _, _) = ymd(sys, a)?;
            (b - a) / if dates::is_leap(y) { 366.0 } else { 365.0 }
        }
        2 => (b - a) / 360.0,
        3 => (b - a) / 365.0,
        _ => {
            // The 30/360 bases without any adjustment of the days, except that the US one
            // counts February by its days when the dates are in the same year and the first
            // is in it.
            let ((y1, m1, d1), (y2, m2, d2)) = (ymd(sys, a)?, ymd(sys, b)?);
            let mut days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
            if basis == 0 && m1 == 2 && y1 == y2 && m2 > 2 {
                days -= 30 - dates::days_in_month(y1, 2);
            }
            days as f64 / 360.0
        }
    })
}

/// A positive argument (a rate, a price, a redemption): zero and below are outside the domain.
fn positive(e: &mut Engine, v: &Value) -> Res<f64> {
    let x = e.num(v)?;
    if x > 0.0 { Ok(x) } else { Err(DOMAIN) }
}

fn non_negative(e: &mut Engine, v: &Value) -> Res<f64> {
    let x = e.num(v)?;
    if x >= 0.0 { Ok(x) } else { Err(DOMAIN) }
}

/// The price of a security with coupons (per 100) for a yield.
fn price_of(c: &Coupon, rate: f64, yld: f64, redemption: f64, freq: f64) -> f64 {
    let e = c.days;
    let dsc_e = c.days_nc / e;
    let n = c.num;
    let mut price = redemption / (1.0 + yld / freq).powf(n - 1.0 + dsc_e);
    price -= 100.0 * rate / freq * c.days_bs / e;
    let t1 = 100.0 * rate / freq;
    let t2 = 1.0 + yld / freq;
    let mut k = 0.0;
    while k < n {
        price += t1 / t2.powf(k + dsc_e);
        k += 1.0;
    }
    price
}

/// The yield for a price: the search of LibreOffice (a regula falsi from yields 0 and 1, the upper
/// one doubled while the price is above that of it).
fn yield_of(c: &Coupon, rate: f64, price: f64, redemption: f64, freq: f64) -> Res<f64> {
    let at = |y: f64| price_of(c, rate, y, redemption, freq);
    let (mut y1, mut y2) = (0.0, 1.0);
    let (mut p1, mut p2) = (at(y1), at(y2));
    let mut yn = (y2 - y1) * 0.5;
    let mut pn = 0.0;
    let mut iter = 0;
    while iter < 100 && pn != price {
        pn = at(yn);
        if price == p1 {
            return Ok(y1);
        } else if price == p2 {
            return Ok(y2);
        } else if price == pn {
            return Ok(yn);
        } else if price < p2 {
            y2 *= 2.0;
            p2 = at(y2);
            yn = (y2 - y1) * 0.5;
        } else {
            if price < pn {
                y1 = yn;
                p1 = pn;
            } else {
                y2 = yn;
                p2 = pn;
            }
            yn = y2 - (y2 - y1) * ((price - p2) / (p1 - p2));
        }
        iter += 1;
    }
    if !yn.is_finite() || (price - pn).abs() > price / 100.0 {
        return Err(DOMAIN);
    }
    Ok(yn)
}

/// The Macaulay duration of a security with coupons, in years.
fn duration_of(yearfrac: f64, num: f64, coupon_rate: f64, yld: f64, freq: f64) -> f64 {
    let coup = coupon_rate * 100.0 / freq;
    let y = yld / freq + 1.0;
    let diff = yearfrac * freq - num;
    let mut dur = 0.0;
    let mut t = 1.0;
    while t < num {
        dur += (t + diff) * coup / y.powf(t + diff);
        t += 1.0;
    }
    dur += (num + diff) * (coup + 100.0) / y.powf(num + diff);
    let mut p = 0.0;
    let mut t = 1.0;
    while t < num {
        p += coup / y.powf(t + diff);
        t += 1.0;
    }
    p += (coup + 100.0) / y.powf(num + diff);
    dur / p / freq
}

fn settlement_maturity(e: &mut Engine, a: &[Value]) -> Res<(f64, f64)> {
    let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
    if s >= m {
        return Err(DOMAIN);
    }
    Ok((s, m))
}

/// The common arguments of the discount instruments: settlement and maturity, and a basis.
fn instrument(e: &mut Engine, a: &[Value], basis_at: usize) -> Res<(f64, f64, i64)> {
    let (s, m) = settlement_maturity(e, a)?;
    Ok((s, m, basis_of(e, a, basis_at)?))
}

// ----- depreciation -----

/// The depreciation of period `period` (1-based) by the declining balance of a factor, as the
/// loop of `VDB` takes it.
fn declining(cost: f64, salvage: f64, life: f64, period: f64, factor: f64) -> f64 {
    let mut rate = factor / life;
    let old_value = if rate >= 1.0 {
        rate = 1.0;
        if period == 1.0 { cost } else { 0.0 }
    } else {
        cost * (1.0 - rate).powf(period - 1.0)
    };
    let new_value = cost * (1.0 - rate).powf(period);
    let d = if new_value < salvage {
        old_value - salvage
    } else {
        old_value - new_value
    };
    d.max(0.0)
}

fn vdb(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let (cost, salvage, life) = (e.num(&a[0])?, e.num(&a[1])?, e.num(&a[2])?);
    let (start, end) = (e.num(&a[3])?, e.num(&a[4])?);
    // A factor given as an empty argument is 0, not the default.
    let factor = match a.get(5) {
        None => 2.0,
        Some(v) => e.num(v)?,
    };
    let no_switch = match a.get(6) {
        None => false,
        Some(v) => e.boolean(v)?,
    };
    if start < 0.0 || end < start || end > life || cost < 0.0 || salvage > cost || factor <= 0.0 {
        return Err(DOMAIN);
    }
    // The depreciation of each whole period (the declining balance, moving to the straight line
    // when that is larger unless `no_switch`), taken in proportion for the part of it that lies
    // between the start and the end.
    let mut total = 0.0;
    let mut rest = cost - salvage;
    let mut straight = 0.0;
    let mut switched = false;
    let last = end.ceil();
    let mut i = 1.0;
    while i <= last {
        let g = declining(cost, salvage, life, i, factor);
        let term = if no_switch {
            g
        } else if switched {
            straight
        } else {
            straight = rest / (life - (i - 1.0));
            if straight > g {
                switched = true;
                straight
            } else {
                rest -= g;
                g
            }
        };
        let overlap = (end.min(i) - start.max(i - 1.0)).max(0.0);
        total += term * overlap;
        i += 1.0;
    }
    num(total)
}

fn amor(e: &mut Engine, a: &[Value], degressive: bool) -> Res<Value> {
    let cost = e.num(&a[0])?;
    let purchased = date_arg(e, &a[1])?;
    let first = date_arg(e, &a[2])?;
    let salvage = e.num(&a[3])?;
    let period = e.num(&a[4])?.trunc();
    let rate = e.num(&a[5])?;
    let basis = basis_of(e, a, 6)?;
    if purchased > first
        || cost <= 0.0
        || salvage < 0.0
        || salvage > cost
        || period < 0.0
        || rate <= 0.0
    {
        return Err(DOMAIN);
    }
    let frac = year_frac(e, purchased, first, basis)?;
    let per = period as u64;
    if !degressive {
        let one_rate = cost * rate;
        let cost_delta = cost - salvage;
        let first_rate = frac * rate * cost;
        let full = ((cost - salvage - first_rate) / one_rate).floor();
        return num(if per == 0 {
            first_rate
        } else if (per as f64) <= full {
            one_rate
        } else if per as f64 == full + 1.0 {
            cost_delta - one_rate * full - first_rate
        } else {
            0.0
        });
    }
    let life = 1.0 / rate;
    let coeff = if life < 3.0 {
        1.0
    } else if life < 5.0 {
        1.5
    } else if life <= 6.0 {
        2.0
    } else {
        2.5
    };
    let rate = rate * coeff;
    let mut cost = cost;
    let mut paid = (frac * rate * cost + 0.5).floor();
    cost -= paid;
    let mut rest = cost - salvage;
    for n in 0..per {
        paid = (rate * cost + 0.5).floor();
        rest -= paid;
        if rest < 0.0 {
            return num(if per - n <= 1 {
                (cost * 0.5 + 0.5).floor()
            } else {
                0.0
            });
        }
        cost -= paid;
    }
    num(paid)
}

/// `ODDLPRICE` and `ODDLYIELD`: the arguments up to the redemption, and the three fractions of a
/// period (the last interest to the maturity, the settlement to the maturity, the last interest to
/// the settlement).
fn odd_last(e: &mut Engine, a: &[Value], price_known: bool) -> Res<Value> {
    let (settle, mat, last) = (
        date_arg(e, &a[0])?,
        date_arg(e, &a[1])?,
        date_arg(e, &a[2])?,
    );
    let rate = e.num(&a[3])?;
    let second = e.num(&a[4])?;
    let redemption = e.num(&a[5])?;
    let freq = frequency(e, &a[6])?;
    let basis = basis_of(e, a, 7)?;
    if rate <= 0.0
        || second < 0.0
        || (!price_known && second == 0.0)
        || redemption <= 0.0
        || mat <= settle
        || settle <= last
    {
        return Err(DOMAIN);
    }
    let f = freq as f64;
    let dci = year_frac(e, last, mat, basis)? * f;
    let dsci = year_frac(e, settle, mat, basis)? * f;
    let ai = year_frac(e, last, settle, basis)? * f;
    if price_known {
        // `second` is the yield.
        let mut price = redemption + dci * 100.0 * rate / f;
        price /= dsci * second / f + 1.0;
        price -= ai * 100.0 * rate / f;
        num(price)
    } else {
        // `second` is the price.
        let mut y = redemption + dci * 100.0 * rate / f;
        y /= second + ai * 100.0 * rate / f;
        y -= 1.0;
        num(y * f / dsci)
    }
}

/// LibreOffice has no calculation behind `ODDFPRICE` and `ODDFYIELD`: they answer `#VALUE!` for
/// every set of arguments that is a call of them (and the errors of the arguments).
fn not_implemented_in_libreoffice(e: &mut Engine, a: &[Value]) -> Res<Value> {
    for v in a {
        e.num(v)?;
    }
    Err(ErrorKind::Value)
}

pub(super) fn register(r: &mut Reg) {
    r.eager("ODDLPRICE", 7, 8, |e, a| odd_last(e, a, true));
    r.eager("ODDLYIELD", 7, 8, |e, a| odd_last(e, a, false));
    r.eager("ODDFPRICE", 8, 9, not_implemented_in_libreoffice);
    r.eager("ODDFYIELD", 8, 9, not_implemented_in_libreoffice);
    r.eager("VDB", 5, 7, vdb);
    r.eager("AMORLINC", 6, 7, |e, a| amor(e, a, false));
    r.eager("AMORDEGRC", 6, 7, |e, a| amor(e, a, true));
    r.eager("COUPPCD", 3, 4, |e, a| num(coupon_of(e, a)?.pcd));
    r.eager("COUPNCD", 3, 4, |e, a| num(coupon_of(e, a)?.ncd));
    r.eager("COUPNUM", 3, 4, |e, a| num(coupon_of(e, a)?.num));
    r.eager("COUPDAYBS", 3, 4, |e, a| num(coupon_of(e, a)?.days_bs));
    r.eager("COUPDAYS", 3, 4, |e, a| num(coupon_of(e, a)?.days));
    r.eager("COUPDAYSNC", 3, 4, |e, a| num(coupon_of(e, a)?.days_nc));

    r.eager("PRICE", 6, 7, |e, a| {
        let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
        let rate = non_negative(e, &a[2])?;
        let yld = non_negative(e, &a[3])?;
        let redemption = positive(e, &a[4])?;
        let freq = frequency(e, &a[5])?;
        let basis = basis_of(e, a, 6)?;
        let c = coupon(e.opts.date_system, s, m, freq, basis)?;
        num(price_of(&c, rate, yld, redemption, freq as f64))
    });
    r.eager("YIELD", 6, 7, |e, a| {
        let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
        let rate = non_negative(e, &a[2])?;
        let price = positive(e, &a[3])?;
        let redemption = positive(e, &a[4])?;
        let freq = frequency(e, &a[5])?;
        let basis = basis_of(e, a, 6)?;
        let c = coupon(e.opts.date_system, s, m, freq, basis)?;
        num(yield_of(&c, rate, price, redemption, freq as f64)?)
    });
    for (name, modified) in [("DURATION", false), ("MDURATION", true)] {
        let f: fn(&mut Engine, &[Value]) -> Res<Value> = if modified {
            |e, a| duration(e, a, true)
        } else {
            |e, a| duration(e, a, false)
        };
        r.eager(name, 5, 6, f);
    }
    r.eager("DISC", 4, 5, |e, a| {
        let (s, m, basis) = instrument(e, a, 4)?;
        let (price, redemption) = (positive(e, &a[2])?, positive(e, &a[3])?);
        let frac = year_frac(e, s, m, basis)?;
        num((1.0 - price / redemption) / frac)
    });
    r.eager("PRICEDISC", 4, 5, |e, a| {
        let (s, m, basis) = instrument(e, a, 4)?;
        let (disc, redemption) = (positive(e, &a[2])?, positive(e, &a[3])?);
        let frac = year_diff(e, s, m, basis)?;
        num(redemption * (1.0 - disc * frac))
    });
    r.eager("YIELDDISC", 4, 5, |e, a| {
        let (s, m, basis) = instrument(e, a, 4)?;
        let (price, redemption) = (positive(e, &a[2])?, positive(e, &a[3])?);
        let frac = year_frac(e, s, m, basis)?;
        num((redemption / price - 1.0) / frac)
    });
    r.eager("RECEIVED", 4, 5, |e, a| {
        let (s, m, basis) = instrument(e, a, 4)?;
        let (investment, disc) = (positive(e, &a[2])?, positive(e, &a[3])?);
        let frac = year_diff(e, s, m, basis)?;
        let denom = 1.0 - disc * frac;
        if denom == 0.0 {
            return Err(DOMAIN);
        }
        num(investment / denom)
    });
    r.eager("INTRATE", 4, 5, |e, a| {
        let (s, m, basis) = instrument(e, a, 4)?;
        let (investment, redemption) = (positive(e, &a[2])?, positive(e, &a[3])?);
        let frac = year_diff(e, s, m, basis)?;
        num((redemption / investment - 1.0) / frac)
    });
    r.eager("PRICEMAT", 5, 6, |e, a| {
        let (s, m, basis) = instrument(e, a, 5)?;
        let issue = date_arg(e, &a[2])?;
        let (rate, yld) = (non_negative(e, &a[3])?, non_negative(e, &a[4])?);
        let iss_mat = year_frac(e, issue, m, basis)?;
        let iss_set = year_frac(e, issue, s, basis)?;
        let set_mat = year_frac(e, s, m, basis)?;
        let mut price = 1.0 + iss_mat * rate;
        price /= 1.0 + set_mat * yld;
        price -= iss_set * rate;
        num(price * 100.0)
    });
    r.eager("YIELDMAT", 5, 6, |e, a| {
        let (s, m, basis) = instrument(e, a, 5)?;
        let issue = date_arg(e, &a[2])?;
        let rate = non_negative(e, &a[3])?;
        let price = positive(e, &a[4])?;
        if issue > s {
            return Err(DOMAIN);
        }
        let iss_mat = year_frac(e, issue, m, basis)?;
        let iss_set = year_frac(e, issue, s, basis)?;
        let set_mat = year_frac(e, s, m, basis)?;
        let mut y = 1.0 + iss_mat * rate;
        y /= price / 100.0 + iss_set * rate;
        y -= 1.0;
        num(y / set_mat)
    });
    r.eager("ACCRINTM", 3, 5, |e, a| {
        let issue = date_arg(e, &a[0])?;
        let settle = date_arg(e, &a[1])?;
        let rate = positive(e, &a[2])?;
        let par = match super::opt(a, 3) {
            Some(v) => positive(e, v)?,
            None => 1000.0,
        };
        let basis = basis_of(e, a, 4)?;
        if issue >= settle {
            return Err(DOMAIN);
        }
        num(par * rate * year_diff(e, issue, settle, basis)?)
    });
    // The first interest date is read but does not change the answer, as in LibreOffice.
    r.eager("ACCRINT", 6, 7, |e, a| {
        let issue = date_arg(e, &a[0])?;
        let _first = date_arg(e, &a[1])?;
        let settle = date_arg(e, &a[2])?;
        let rate = positive(e, &a[3])?;
        let par = positive(e, &a[4])?;
        let _freq = frequency(e, &a[5])?;
        let basis = basis_of(e, a, 6)?;
        if issue >= settle {
            return Err(DOMAIN);
        }
        let years = year_diff(e, issue, settle, basis)?;
        num(par * rate * years)
    });
    r.eager("TBILLPRICE", 3, 3, |e, a| {
        let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
        let disc = positive(e, &a[2])?;
        if s > m {
            return Err(DOMAIN);
        }
        let sys = e.opts.date_system;
        let days = days360(ymd(sys, s)?, ymd(sys, m + 1.0)?, false);
        let frac = days / 360.0;
        if frac - frac.floor() == 0.0 {
            return Err(DOMAIN);
        }
        num(100.0 * (1.0 - disc * frac))
    });
    r.eager("TBILLYIELD", 3, 3, |e, a| {
        let (s, m) = settlement_maturity(e, a)?;
        let price = positive(e, &a[2])?;
        let sys = e.opts.date_system;
        let (from, to) = (ymd(sys, s)?, ymd(sys, m)?);
        let last_feb = |(y, mo, d): (i64, i64, i64)| mo == 2 && d == dates::days_in_month(y, 2);
        if last_feb(from) && last_feb(to) && m - s > 366.0 {
            return Err(DOMAIN);
        }
        if last_feb(from) && last_feb(to) {
            crate::notes::note(crate::notes::TBILLYIELD_FEB);
            return Err(ErrorKind::Unsupported);
        }
        let days = days360_t(from, to) + 1.0;
        if days > 360.0 {
            return Err(DOMAIN);
        }
        num((100.0 / price - 1.0) / (days / 360.0))
    });
    r.eager("TBILLEQ", 3, 3, |e, a| {
        let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
        let disc = positive(e, &a[2])?;
        if s > m {
            return Err(DOMAIN);
        }
        let sys = e.opts.date_system;
        let days = days360_t(ymd(sys, s)?, ymd(sys, m + 1.0)?);
        if days > 360.0 {
            return Err(DOMAIN);
        }
        num(365.0 * disc / (360.0 - disc * days))
    });
}

/// The US 30/360 count of `TBILLEQ` and `TBILLYIELD`: the days are adjusted before the end day
/// is cut (a start on the last day of February counts as 30 for it).
fn days360_t(from: (i64, i64, i64), to: (i64, i64, i64)) -> f64 {
    let ((y1, m1, mut d1), (y2, m2, mut d2)) = (from, to);
    let last_feb = |y: i64, m: i64, d: i64| m == 2 && d == dates::days_in_month(y, 2);
    let start_feb = last_feb(y1, m1, d1);
    if start_feb && last_feb(y2, m2, d2) {
        d2 = 30;
    }
    if d1 == 31 || start_feb {
        d1 = 30;
    }
    if d2 == 31 && d1 >= 30 {
        d2 = 30;
    }
    ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64
}

fn duration(e: &mut Engine, a: &[Value], modified: bool) -> Res<Value> {
    let (s, m) = (date_arg(e, &a[0])?, date_arg(e, &a[1])?);
    let rate = non_negative(e, &a[2])?;
    let yld = non_negative(e, &a[3])?;
    let freq = frequency(e, &a[4])?;
    let basis = basis_of(e, a, 5)?;
    let c = coupon(e.opts.date_system, s, m, freq, basis)?;
    let frac = year_frac(e, s, m, basis)?;
    let d = duration_of(frac, c.num, rate, yld, freq as f64);
    num(if modified {
        d / (1.0 + yld / freq as f64)
    } else {
        d
    })
}
