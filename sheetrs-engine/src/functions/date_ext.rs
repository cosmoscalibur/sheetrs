//! More date functions: working days, day counts, week numbers, year fractions.

use super::{Reg, VAR, num, opt_num};
use crate::dates;
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

fn ymd(e: &Engine, serial: f64) -> Res<(i64, i64, i64)> {
    dates::serial_to_ymd(serial, e.opts.date_system).ok_or(ErrorKind::Domain)
}

fn last_of_feb(y: i64, m: i64, d: i64) -> bool {
    m == 2 && d == dates::days_in_month(y, 2)
}

/// Days between two dates on a 360-day year. `european` is the 30E/360 rule;
/// otherwise the US (NASD) rule.
fn days360(a: (i64, i64, i64), b: (i64, i64, i64), european: bool) -> i64 {
    let ((y1, m1, mut d1), (y2, m2, mut d2)) = (a, b);
    if european {
        d1 = d1.min(30);
        d2 = d2.min(30);
    } else {
        if last_of_feb(y1, m1, d1) {
            if last_of_feb(y2, m2, d2) {
                d2 = 30;
            }
            d1 = 30;
        }
        if d2 == 31 && d1 >= 30 {
            d2 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
    }
    (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
}

/// Non-working weekdays (index 0 = Sunday) from a `WORKDAY.INTL` weekend code.
fn weekend(e: &mut Engine, v: Option<&Value>) -> Res<[bool; 7]> {
    let mut w = [false; 7];
    let Some(v) = v.filter(|v| !matches!(v, Value::Empty)) else {
        w[0] = true;
        w[6] = true;
        return Ok(w);
    };
    let s = e.scalar(v.clone());
    if let Value::Text(t) = &s
        && t.len() == 7
        && t.bytes().all(|b| b == b'0' || b == b'1')
    {
        // Monday first.
        for (i, b) in t.bytes().enumerate() {
            w[(i + 1) % 7] = b == b'1';
        }
        if w.iter().all(|x| *x) {
            return Err(ErrorKind::Domain);
        }
        return Ok(w);
    }
    let code = e.num(&s)?.trunc() as i64;
    match code {
        1..=7 => {
            // 1 = Sat+Sun, 2 = Sun+Mon, … 7 = Fri+Sat
            let first = (code + 5) % 7; // Saturday for 1
            w[first as usize] = true;
            w[((first + 1) % 7) as usize] = true;
        }
        11..=17 => w[((code - 11 + 1) % 7) as usize] = true,
        _ => return Err(ErrorKind::Domain),
    }
    Ok(w)
}

fn holidays(e: &mut Engine, v: Option<&Value>) -> Res<Vec<f64>> {
    match v {
        None => Ok(Vec::new()),
        Some(v) => Ok(e
            .numbers(std::slice::from_ref(v))?
            .into_iter()
            .map(f64::floor)
            .collect()),
    }
}

fn working(e: &Engine, d: f64, w: &[bool; 7], hol: &[f64]) -> bool {
    !w[dates::weekday_sun0(d, e.opts.date_system) as usize] && !hol.contains(&d)
}

fn workday(e: &mut Engine, a: &[Value], intl: bool) -> Res<Value> {
    let start = e.num(&a[0])?.floor();
    let days = e.num(&a[1])?.trunc();
    let (w, hol) = if intl {
        (weekend(e, a.get(2))?, holidays(e, a.get(3))?)
    } else {
        (weekend(e, None)?, holidays(e, a.get(2))?)
    };
    if start < 0.0 {
        return Err(ErrorKind::Domain);
    }
    let step = if days < 0.0 { -1.0 } else { 1.0 };
    let (mut d, mut left) = (start, days.abs());
    while left > 0.0 {
        d += step;
        if working(e, d, &w, &hol) {
            left -= 1.0;
        }
        if !(0.0..2_958_466.0).contains(&d) {
            return Err(ErrorKind::Domain);
        }
    }
    num(d)
}

fn networkdays(e: &mut Engine, a: &[Value], intl: bool) -> Res<Value> {
    let (s, t) = (e.num(&a[0])?.floor(), e.num(&a[1])?.floor());
    let (w, hol) = if intl {
        (weekend(e, a.get(2))?, holidays(e, a.get(3))?)
    } else {
        (weekend(e, None)?, holidays(e, a.get(2))?)
    };
    let (lo, hi) = if s <= t { (s, t) } else { (t, s) };
    let mut n = 0.0;
    let mut d = lo;
    while d <= hi {
        if working(e, d, &w, &hol) {
            n += 1.0;
        }
        d += 1.0;
    }
    num(if s <= t { n } else { -n })
}

/// The number of ISO weeks of a year: 53 when it starts on a Thursday, or ends on one.
fn iso_weeks_in(year: i64) -> i64 {
    let p = |yy: i64| (yy + yy / 4 - yy / 100 + yy / 400) % 7;
    if p(year) == 4 || p(year - 1) == 3 {
        53
    } else {
        52
    }
}

/// The calendar distance `WEEKS`, `MONTHS` and `YEARS` of OpenOffice measure: `mode` 0 counts the
/// complete intervals between two dates and 1 the boundaries (Monday, first of the month, first of
/// the year) crossed.
fn date_distance(e: &mut Engine, a: &[Value], unit: char) -> Res<Value> {
    let (s, t) = (e.num(&a[0])?.floor(), e.num(&a[1])?.floor());
    let mode = e.num(&a[2])?.trunc();
    if mode != 0.0 && mode != 1.0 {
        return Err(ErrorKind::Domain);
    }
    let by_boundary = mode == 1.0;
    let sys = e.opts.date_system;
    let (lo, hi, sign) = if s <= t { (s, t, 1.0) } else { (t, s, -1.0) };
    let n = match unit {
        'w' if by_boundary => {
            let monday = |d: f64| d - ((dates::weekday_sun0(d, sys) + 6) % 7) as f64;
            ((monday(hi) - monday(lo)) / 7.0).round()
        }
        'w' => ((hi - lo) / 7.0).trunc(),
        _ => {
            let (y1, m1, d1) = ymd(e, lo)?;
            let (y2, m2, d2) = ymd(e, hi)?;
            if unit == 'm' {
                let months = (y2 - y1) * 12 + m2 - m1;
                (months - i64::from(!by_boundary && d2 < d1)) as f64
            } else {
                let years = y2 - y1;
                (years - i64::from(!by_boundary && (m2, d2) < (m1, d1))) as f64
            }
        }
    };
    num(sign * n + 0.0)
}

/// A year, month or day of the date `a[0]` for the functions that describe a calendar period.
fn calendar_of(e: &mut Engine, a: &[Value]) -> Res<(i64, i64, i64)> {
    let d = e.num(&a[0])?.floor();
    ymd(e, d)
}

fn iso_week(y: i64, serial: f64, sys: crate::options::DateSystem) -> i64 {
    let jan1 = dates::ymd_to_serial(y, 1, 1, sys).unwrap_or(serial);
    let ordinal = (serial - jan1) as i64 + 1;
    let iso_wd = (dates::weekday_sun0(serial, sys) + 6) % 7 + 1;
    let w = (ordinal - iso_wd + 10) / 7;
    if w < 1 {
        iso_weeks_in(y - 1)
    } else if w > iso_weeks_in(y) {
        1
    } else {
        w
    }
}

pub(super) fn year_frac(e: &Engine, s: f64, t: f64, basis: i64) -> Res<f64> {
    let (s, t) = if s <= t { (s, t) } else { (t, s) };
    let (a, b) = (ymd(e, s)?, ymd(e, t)?);
    let diff = t - s;
    Ok(match basis {
        0 => days360(a, b, false) as f64 / 360.0,
        4 => days360(a, b, true) as f64 / 360.0,
        2 => diff / 360.0,
        3 => diff / 365.0,
        1 => {
            let ((y1, m1, d1), (y2, m2, d2)) = (a, b);
            let year_days = if y1 == y2 {
                if dates::is_leap(y1) { 366.0 } else { 365.0 }
            } else if y1 + 1 == y2 && (m1 > m2 || (m1 == m2 && d1 >= d2)) {
                if (dates::is_leap(y1) && m1 < 3)
                    || (dates::is_leap(y2) && m2 * 31 + d2 >= 2 * 31 + 29)
                {
                    366.0
                } else {
                    365.0
                }
            } else {
                let total: i64 = (y1..=y2)
                    .map(|y| if dates::is_leap(y) { 366 } else { 365 })
                    .sum();
                total as f64 / (y2 - y1 + 1) as f64
            };
            diff / year_days
        }
        _ => return Err(ErrorKind::Domain),
    })
}

pub(super) fn register(r: &mut Reg) {
    // Gregorian Easter (Meeus/Jones/Butcher).
    r.eager("EASTERSUNDAY", 1, 1, |e, a| {
        let y = e.num(&a[0])?.trunc() as i64;
        if !(1583..=9999).contains(&y) {
            return Err(ErrorKind::Domain);
        }
        let (a_, b, c) = (y % 19, y / 100, y % 100);
        let (d, e_) = (b / 4, b % 4);
        let f = (b + 8) / 25;
        let g = (b - f + 1) / 3;
        let h = (19 * a_ + b - d - g + 15) % 30;
        let (i, k) = (c / 4, c % 4);
        let l = (32 + 2 * e_ + 2 * i - h - k) % 7;
        let m = (a_ + 11 * h + 22 * l) / 451;
        let month = (h + l - 7 * m + 114) / 31;
        let day = (h + l - 7 * m + 114) % 31 + 1;
        dates::ymd_to_serial(y, month, day, e.opts.date_system).map_or(Err(ErrorKind::Domain), num)
    });

    r.eager("WEEKS", 3, 3, |e, a| date_distance(e, a, 'w'));
    r.eager("MONTHS", 3, 3, |e, a| date_distance(e, a, 'm'));
    r.eager("YEARS", 3, 3, |e, a| date_distance(e, a, 'y'));
    r.eager("ISLEAPYEAR", 1, 1, |e, a| {
        let (y, _, _) = calendar_of(e, a)?;
        num(f64::from(u8::from(dates::is_leap(y))))
    });
    r.eager("DAYSINMONTH", 1, 1, |e, a| {
        let (y, m, _) = calendar_of(e, a)?;
        num(dates::days_in_month(y, m) as f64)
    });
    r.eager("DAYSINYEAR", 1, 1, |e, a| {
        let (y, _, _) = calendar_of(e, a)?;
        num(if dates::is_leap(y) { 366.0 } else { 365.0 })
    });
    r.eager("WEEKSINYEAR", 1, 1, |e, a| {
        let (y, _, _) = calendar_of(e, a)?;
        num(iso_weeks_in(y) as f64)
    });
    r.eager("DAYS360", 2, 3, |e, a| {
        let (s, t) = (e.num(&a[0])?, e.num(&a[1])?);
        let european = match a.get(2) {
            Some(v) if !matches!(v, Value::Empty) => e.boolean(v)?,
            _ => false,
        };
        // The start date is the one adjusted first, whichever is earlier.
        let (x, y) = (ymd(e, s)?, ymd(e, t)?);
        num(days360(x, y, european) as f64)
    });
    r.eager("WORKDAY", 2, 3, |e, a| workday(e, a, false));
    r.eager("WORKDAY.INTL", 2, 4, |e, a| workday(e, a, true));
    r.eager("NETWORKDAYS", 2, 3, |e, a| networkdays(e, a, false));
    r.eager("NETWORKDAYS.INTL", 2, 4, |e, a| networkdays(e, a, true));
    r.eager("YEARFRAC", 2, 3, |e, a| {
        let (s, t) = (e.num(&a[0])?.floor(), e.num(&a[1])?.floor());
        let basis = opt_num(e, a, 2, 0.0)?.trunc() as i64;
        year_frac(e, s, t, basis).and_then(num)
    });
    r.eager("WEEKNUM", 1, 2, weeknum);
    r.eager("WEEKNUM_OOO", 2, 2, |e, a| {
        let s = e.num(&a[0])?.floor();
        let ty = e.num(&a[1])?.trunc();
        // Weeks start on Sunday for type 1 and on Monday for any other.
        let start = if ty == 1.0 { 6 } else { 0 };
        week_of_year(e, s, start).map(|w| Value::Number(w as f64))
    });
    r.eager("ISOWEEKNUM", 1, 1, |e, a| {
        let s = e.num(&a[0])?.floor();
        let (y, _, _) = ymd(e, s)?;
        num(iso_week(y, s, e.opts.date_system) as f64)
    });
    r.eager("TIMEVALUE", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        crate::coerce::parse_date_text(&t, &e.opts.locale, e.opts.date_ctx())
            .map_or(Err(ErrorKind::Value), |n| num(n - n.floor()))
    });
    let _ = VAR;
}

fn weeknum(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let s = e.num(&a[0])?.floor();
    let ty = opt_num(e, a, 1, 1.0)?.trunc() as i64;
    let sys = e.opts.date_system;
    let (y, _, _) = ymd(e, s)?;
    if ty == 21 {
        return num(iso_week(y, s, sys) as f64);
    }
    let start_day = match ty {
        1 | 17 => 0,
        2 => 1,
        11..=16 => ty - 10,
        _ => return Err(ErrorKind::Domain),
    };
    let jan1 = dates::ymd_to_serial(y, 1, 1, sys).ok_or(ErrorKind::Domain)?;
    let offset = (dates::weekday_sun0(jan1, sys) - start_day).rem_euclid(7);
    let week = ((s - jan1) as i64 + offset) / 7 + 1;
    // The days of a week that holds January 1st of the next year are that year's first
    // week.
    if let Some(next) = dates::ymd_to_serial(y + 1, 1, 1, sys) {
        let week_start = s - (dates::weekday_sun0(s, sys) - start_day).rem_euclid(7) as f64;
        if week_start + 6.0 >= next {
            return num(1.0);
        }
    }
    num(week as f64)
}

/// `Date::GetWeekOfYear` of LibreOffice's `tools` with a week of at least 4 days, the week the
/// older `WEEKNUM_OOO` gives: `start` is the first day of the week, Monday being 0.
fn week_of_year(e: &Engine, serial: f64, start: i64) -> Res<i64> {
    let sys = e.opts.date_system;
    let (y, _, _) = ymd(e, serial)?;
    let jan1 = dates::ymd_to_serial(y, 1, 1, sys).ok_or(ErrorKind::Domain)?;
    let monday0 = |s: f64| (dates::weekday_sun0(s, sys) + 6) % 7;
    let day_of_year = (serial - jan1) as i64;
    let first = (monday0(jan1) + 7 - start) % 7;
    let mut week = if first < 4 {
        1
    } else if first == 4 {
        53
    } else if first == 5 {
        let prev_leap = (y - 1) % 4 == 0 && ((y - 1) % 100 != 0 || (y - 1) % 400 == 0);
        if prev_leap { 53 } else { 52 }
    } else {
        52
    };
    if week == 1 || day_of_year + first > 6 {
        week = if week == 1 {
            week + (day_of_year + first) / 7
        } else {
            (day_of_year + first) / 7
        };
        if week == 53 {
            // The next first day of the week is the first of the new year: the same week.
            let next = serial + (6 - (monday0(serial) + 7 - start) % 7) as f64;
            week = week_of_year(e, next, start)?;
        }
    }
    Ok(week)
}
