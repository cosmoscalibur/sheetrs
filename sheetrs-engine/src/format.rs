//! Number and date formatting for `TEXT` and friends (the common subset of the
//! format-code language, rendered through a [`Locale`]).

use crate::coerce::fmt_number;
use crate::dates;
use crate::functions::round_decimal;
use crate::options::{DateSystem, Locale};
use std::fmt::Write as _;

/// Split a format code at top-level `;`.
fn sections(fmt: &str) -> Vec<&str> {
    let (mut out, mut start, mut quoted) = (Vec::new(), 0, false);
    for (i, c) in fmt.char_indices() {
        match c {
            '"' => quoted = !quoted,
            ';' if !quoted => {
                out.push(&fmt[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push(&fmt[start..]);
    out
}

#[derive(Debug, PartialEq)]
enum Tok {
    Lit(String),
    Y(usize),
    M(usize),
    D(usize),
    H(usize),
    S(usize),
    /// `AM/PM` (`true`) or `A/P`.
    Ampm(bool),
    /// `[h]`, `[m]`, `[s]`: elapsed time since the null date, not a clock field.
    Elapsed(char),
    /// `.0`, `.00`, `.000` after the seconds.
    Frac(usize),
    /// `/` in a date format: the locale's date separator.
    DateSep,
    /// A `[$-407]` tag of a language the engine has no names for: LibreOffice writes that
    /// language's names, so the code cannot be reproduced and TEXT is an error.
    Foreign,
    /// A `[$-409]` tag: the names of months and days, and AM/PM, are those of this locale's language.
    Lang(u16),
}

fn tokenize_dt(fmt: &str) -> Option<Vec<Tok>> {
    let cs: Vec<char> = fmt.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    let mut dt = false;
    while i < cs.len() {
        let c = cs[i];
        let lc = c.to_ascii_lowercase();
        let run = cs[i..]
            .iter()
            .take_while(|x| x.to_ascii_lowercase() == lc)
            .count();
        match lc {
            '"' => {
                let end = cs[i + 1..]
                    .iter()
                    .position(|&x| x == '"')
                    .map_or(cs.len(), |p| i + 1 + p);
                out.push(Tok::Lit(cs[i + 1..end].iter().collect()));
                i = end + 1;
            }
            '\\' => {
                if let Some(&n) = cs.get(i + 1) {
                    out.push(Tok::Lit(n.to_string()));
                }
                i += 2;
            }
            // Colors, conditions and language tags are not part of the text; `[h]`
            // and friends are elapsed time.
            '[' => {
                let end = cs[i..]
                    .iter()
                    .position(|&x| x == ']')
                    .map_or(cs.len(), |p| i + p);
                let inner: String = cs[i + 1..end.min(cs.len())].iter().collect();
                let l = inner.to_ascii_lowercase();
                if let Some(k) = l.chars().next()
                    && matches!(k, 'h' | 'm' | 's')
                    && l.chars().all(|c| c == k)
                {
                    dt = true;
                    out.push(Tok::Elapsed(k));
                } else if let Some(cur) = inner.strip_prefix('$') {
                    let (sym, lang) = cur.split_once('-').unwrap_or((cur, ""));
                    // A language id (`409`) or a tag (`en-US`).
                    let row = u16::from_str_radix(lang, 16)
                        .ok()
                        .and_then(Locale::row_by_lcid)
                        .or_else(|| Locale::row(lang).filter(|_| lang.contains('-')));
                    match (row, u16::from_str_radix(lang, 16)) {
                        (Some(r), _) => out.push(Tok::Lang(r.lcid)),
                        // No language, or `40A` (Spanish, traditional sort), which LibreOffice
                        // has no locale for: the names stay the locale's.
                        (None, Ok(0x40A) | Err(_)) => {}
                        (None, _) if lang.is_empty() || sym_only_symbol(sym) => {}
                        (None, _) => out.push(Tok::Foreign),
                    }
                    out.push(Tok::Lit(sym.to_string()));
                }
                i = end + 1;
            }
            '.' if matches!(out.last(), Some(Tok::S(_))) && cs.get(i + 1) == Some(&'0') => {
                let n = cs[i + 1..].iter().take_while(|&&x| x == '0').count();
                out.push(Tok::Frac(n));
                i += 1 + n;
            }
            'y' => {
                dt = true;
                out.push(Tok::Y(run));
                i += run;
            }
            'm' => {
                dt = true;
                out.push(Tok::M(run));
                i += run;
            }
            'd' => {
                dt = true;
                out.push(Tok::D(run));
                i += run;
            }
            // LibreOffice: `NN` the short and `NNN` the long name of the day.
            'n' => {
                dt = true;
                out.push(Tok::D(if run >= 3 { 4 } else { 3 }));
                i += run;
            }
            'h' => {
                dt = true;
                out.push(Tok::H(run));
                i += run;
            }
            's' => {
                dt = true;
                out.push(Tok::S(run));
                i += run;
            }
            '/' => {
                out.push(Tok::DateSep);
                i += 1;
            }
            'a' | 'A'
                if cs[i..]
                    .iter()
                    .take(5)
                    .collect::<String>()
                    .eq_ignore_ascii_case("AM/PM") =>
            {
                dt = true;
                out.push(Tok::Ampm(true));
                i += 5;
            }
            'a' | 'A'
                if cs[i..]
                    .iter()
                    .take(3)
                    .collect::<String>()
                    .eq_ignore_ascii_case("A/P") =>
            {
                dt = true;
                out.push(Tok::Ampm(false));
                i += 3;
            }
            _ => {
                out.push(Tok::Lit(c.to_string()));
                i += 1;
            }
        }
    }
    dt.then_some(out)
}

/// A tag with a currency symbol (`[$€-407]`) names a currency, and its language is not read.
fn sym_only_symbol(sym: &str) -> bool {
    !sym.is_empty()
}

/// Whether a date or time code carries a language tag the engine has no names for (`[$-407]`).
pub fn has_foreign_language_tag(fmt: &str) -> bool {
    tokenize_dt(fmt).is_some_and(|t| t.contains(&Tok::Foreign))
}

fn render_dt(toks: &[Tok], serial: f64, loc: &Locale, sys: DateSystem) -> Option<String> {
    if toks.contains(&Tok::Foreign) {
        crate::notes::note(crate::notes::TEXT_LANGUAGE);
        return None;
    }
    let tagged = toks.iter().find_map(|t| match t {
        Tok::Lang(id) => Locale::row_by_lcid(*id).map(Locale::from_row),
        _ => None,
    });
    let names = tagged.as_ref().unwrap_or(loc);
    let elapsed = toks.iter().any(|t| matches!(t, Tok::Elapsed(_)));
    let negative = elapsed && serial < 0.0;
    let serial = if negative { -serial } else { serial };
    let digits = toks
        .iter()
        .filter_map(|t| if let Tok::Frac(n) = t { Some(*n) } else { None })
        .max()
        .unwrap_or(0);
    let scale = 10f64.powi(i32::try_from(digits).unwrap_or(0));
    // Seconds since midnight of day 0, in units of 1/scale seconds. A clock shows
    // what it has reached (it does not round up to the next second); elapsed time
    // rounds to the smallest unit the format shows.
    let total = (serial * 86_400.0 * 1e7).round() / 1e7;
    let units = if elapsed || digits > 0 {
        (total * scale).round()
    } else {
        total.floor() * scale
    };
    let per_day = 86_400.0 * scale;
    let mut day = (units / per_day).floor();
    let in_day = (units - day * per_day) as i64;
    let secs = in_day / scale as i64;
    let sub = in_day % scale as i64;
    let (hh, mi, ss) = (secs / 3600 % 24, secs / 60 % 60, secs % 60);
    let elapsed_secs = units / scale;
    let _ = &mut day;
    let twelve = toks.iter().any(|t| matches!(t, Tok::Ampm(_)));
    let needs_date = toks
        .iter()
        .any(|t| matches!(t, Tok::Y(_) | Tok::M(_) | Tok::D(_)));
    let (y, mo, d) = if needs_date {
        dates::serial_to_ymd(day, sys)?
    } else {
        (0, 1, 1)
    };
    let dow = if needs_date {
        dates::weekday_sun0(day, sys) as usize
    } else {
        0
    };
    let mut out = String::new();
    if negative {
        out.push('-');
    }
    let is_hour = |t: &Tok| matches!(t, Tok::H(_) | Tok::Elapsed('h'));
    for (k, t) in toks.iter().enumerate() {
        match t {
            Tok::Lit(s) => out.push_str(s),
            Tok::Lang(_) | Tok::Foreign => {}
            Tok::DateSep => out.push(loc.date_sep),
            Tok::Y(n) => {
                if *n <= 2 {
                    let _ = write!(out, "{:02}", y.rem_euclid(100));
                } else {
                    let _ = write!(out, "{y:04}");
                }
            }
            Tok::M(n) => {
                // `m` right after hours or right before seconds means minutes.
                let after_h = toks[..k]
                    .iter()
                    .rev()
                    .find(|t| !matches!(t, Tok::Lit(_)))
                    .is_some_and(is_hour);
                let before_s = toks[k + 1..]
                    .iter()
                    .find(|t| !matches!(t, Tok::Lit(_)))
                    .is_some_and(|t| matches!(t, Tok::S(_)));
                if after_h || before_s {
                    if *n >= 2 {
                        let _ = write!(out, "{mi:02}");
                    } else {
                        let _ = write!(out, "{mi}");
                    }
                } else {
                    let name = &names.months[(mo - 1) as usize];
                    match n {
                        1 => out.push_str(&mo.to_string()),
                        2 => {
                            let _ = write!(out, "{mo:02}");
                        }
                        3 => out.push_str(&names.months_abbr[(mo - 1) as usize]),
                        4 => out.push_str(name),
                        _ => out.extend(name.chars().take(1)),
                    }
                }
            }
            Tok::D(n) => match n {
                1 => out.push_str(&d.to_string()),
                2 => {
                    let _ = write!(out, "{d:02}");
                }
                3 => out.push_str(&names.days_abbr[dow]),
                _ => out.push_str(&names.days[dow]),
            },
            Tok::H(n) => {
                let h = if twelve { (hh + 11) % 12 + 1 } else { hh };
                if *n >= 2 {
                    let _ = write!(out, "{h:02}");
                } else {
                    let _ = write!(out, "{h}");
                }
            }
            Tok::S(n) => {
                if *n >= 2 {
                    let _ = write!(out, "{ss:02}");
                } else {
                    let _ = write!(out, "{ss}");
                }
            }
            Tok::Ampm(full) => {
                let text = if hh < 12 { &names.am } else { &names.pm };
                if *full {
                    out.push_str(text);
                } else {
                    out.extend(text.chars().take(1).flat_map(char::to_lowercase));
                }
            }
            Tok::Elapsed(k) => {
                let unit = match k {
                    'h' => 3600.0,
                    'm' => 60.0,
                    _ => 1.0,
                };
                let _ = write!(out, "{:.0}", (elapsed_secs / unit).floor());
            }
            Tok::Frac(n) => {
                if *n > 0 {
                    out.push(loc.decimal);
                    let _ = write!(out, "{:0width$}", sub, width = *n);
                }
            }
        }
    }
    Some(out)
}

// ---------------------------------------------------------------------------
// Number sections
// ---------------------------------------------------------------------------

/// One element of a number section.
#[derive(Debug, Clone, PartialEq)]
enum Nt {
    Lit(String),
    /// `0`, `#` or `?`.
    Digit(char),
    Dec,
    Group,
    Pct,
    /// `E+`/`E-` of a scientific format; `true` for `E+`.
    Exp(bool),
    Slash,
    /// `@`: the text of a text value.
    Text,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Cmp {
    Lt,
    Le,
    Gt,
    Ge,
    Eq,
    Ne,
}

#[derive(Debug, Clone, PartialEq)]
struct Section {
    cond: Option<(Cmp, f64)>,
    toks: Vec<Nt>,
}

fn parse_cond(s: &str) -> Option<(Cmp, f64)> {
    const OPS: [(&str, Cmp); 6] = [
        ("<=", Cmp::Le),
        (">=", Cmp::Ge),
        ("<>", Cmp::Ne),
        ("<", Cmp::Lt),
        (">", Cmp::Gt),
        ("=", Cmp::Eq),
    ];
    let (op, rest) = OPS
        .iter()
        .find_map(|(sym, op)| s.strip_prefix(sym).map(|r| (*op, r)))?;
    rest.trim().parse().ok().map(|v| (op, v))
}

/// The decimal and group characters a format code is written with. LibreOffice
/// reads the code of `TEXT` in the user's locale (`#.##0,00` in Spanish); the codes
/// it was observed to take in English form whatever the locale are listed.
fn code_separators(fmt: &str, loc: &Locale) -> (char, char) {
    const ENGLISH_FORM: [&str; 13] = [
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "0%",
        "0.00%",
        "0.00E+00",
        "0.00E+000",
        "# ?/?",
        "# ??/??",
        "#,##0_);(#,##0)",
        "#,##0_);[Red](#,##0)",
        "#,##0.00_);(#,##0.00)",
    ];
    if ENGLISH_FORM.contains(&fmt) {
        ('.', ',')
    } else {
        (loc.decimal, loc.group)
    }
}

fn parse_section(src: &str, dec: char, grp: char) -> Section {
    let cs: Vec<char> = src.chars().collect();
    let mut toks: Vec<Nt> = Vec::new();
    let mut cond = None;
    let mut i = 0;
    let mut seen_dec = false;
    let push_lit = |toks: &mut Vec<Nt>, s: &str| match toks.last_mut() {
        Some(Nt::Lit(l)) => l.push_str(s),
        _ => toks.push(Nt::Lit(s.to_string())),
    };
    let has_digit_before = |toks: &[Nt]| toks.iter().any(|t| matches!(t, Nt::Digit(_)));
    let has_dec = src.contains(dec);
    while i < cs.len() {
        let c = cs[i];
        match c {
            '"' => {
                let end = cs[i + 1..]
                    .iter()
                    .position(|&x| x == '"')
                    .map_or(cs.len(), |p| i + 1 + p);
                let lit: String = cs[i + 1..end.min(cs.len())].iter().collect();
                push_lit(&mut toks, &lit);
                i = end + 1;
                continue;
            }
            '\\' => {
                if let Some(&n) = cs.get(i + 1) {
                    push_lit(&mut toks, &n.to_string());
                }
                i += 2;
                continue;
            }
            '[' => {
                let end = cs[i..]
                    .iter()
                    .position(|&x| x == ']')
                    .map_or(cs.len(), |p| i + p);
                let inner: String = cs[i + 1..end.min(cs.len())].iter().collect();
                if let Some(cur) = inner.strip_prefix('$') {
                    // `[$€-407]`: the symbol, then the language it belongs to.
                    let sym = cur.split_once('-').map_or(cur, |(s, _)| s);
                    push_lit(&mut toks, sym);
                } else if let Some(c) = parse_cond(&inner) {
                    cond = Some(c);
                }
                // Colors, `[h]` (elapsed time, handled by the date path) and
                // language tags do not change a number's digits.
                i = end + 1;
                continue;
            }
            '0' | '#' | '?' => toks.push(Nt::Digit(c)),
            '%' => toks.push(Nt::Pct),
            '@' => toks.push(Nt::Text),
            '_' => {
                push_lit(&mut toks, " ");
                i += 2;
                continue;
            }
            '*' => {
                i += 2;
                continue;
            }
            '/' if has_digit_before(&toks)
                && matches!(cs.get(i + 1), Some('0' | '#' | '?' | '1'..='9')) =>
            {
                toks.push(Nt::Slash);
            }
            'E' | 'e' if matches!(cs.get(i + 1), Some('+' | '-')) && has_digit_before(&toks) => {
                toks.push(Nt::Exp(cs[i + 1] == '+'));
                i += 2;
                continue;
            }
            c if c == dec && !seen_dec => {
                seen_dec = true;
                toks.push(Nt::Dec);
            }
            c if c == grp => toks.push(Nt::Group),
            // A `.` is the decimal separator where the code has none of the locale's (`0.0%` in
            // French; `#.##0,00` there keeps it as a literal), and a `,` a group separator.
            '.' if !seen_dec && !has_dec && dec != '.' && grp != '.' => {
                seen_dec = true;
                toks.push(Nt::Dec);
            }
            ',' if dec != ',' && grp != ',' => toks.push(Nt::Group),
            c => push_lit(&mut toks, &c.to_string()),
        }
        i += 1;
    }
    Section { cond, toks }
}

/// The best fraction `n/d` of `f` (0 ≤ f < 1) with `d` up to `max_den`, or the
/// fixed denominator `fixed`.
fn fraction(f: f64, max_den: u64, fixed: Option<u64>) -> (u64, u64) {
    if let Some(d) = fixed {
        return ((f * d as f64).round() as u64, d);
    }
    let (mut best_n, mut best_d, mut best_err) = (0u64, 1u64, f64::MAX);
    for d in 1..=max_den {
        let n = (f * d as f64).round();
        let err = (f - n / d as f64).abs();
        if err < best_err - 1e-15 {
            (best_n, best_d, best_err) = (n as u64, d, err);
        }
    }
    (best_n, best_d)
}

fn group(int: &str, sep: char) -> String {
    let mut out = String::new();
    for (i, c) in int.chars().enumerate() {
        if i > 0 && (int.len() - i).is_multiple_of(3) {
            out.push(sep);
        }
        out.push(c);
    }
    out
}

/// The digit placeholders of `ts`, in order.
fn placeholders(ts: &[Nt]) -> Vec<char> {
    ts.iter()
        .filter_map(|t| if let Nt::Digit(c) = t { Some(*c) } else { None })
        .collect()
}

/// Literal text of `ts` (digits and separators left out).
fn literals(ts: &[Nt]) -> String {
    ts.iter()
        .map(|t| match t {
            Nt::Lit(s) => s.as_str(),
            Nt::Pct => "%",
            _ => "",
        })
        .collect()
}

/// Render `x` (non-negative) with a number section. `negative`: the value is
/// negative and the section does not show it by itself, so `-` is written first
/// unless the number rounds to zero.
fn render_section(sec: &Section, x: f64, negative: bool, loc: &Locale) -> String {
    let toks = &sec.toks;
    let Some(first) = toks.iter().position(|t| matches!(t, Nt::Digit(_))) else {
        // No digit placeholder: literals only.
        return format!("{}{}", if negative { "-" } else { "" }, literals(toks));
    };
    let last = toks
        .iter()
        .rposition(|t| matches!(t, Nt::Digit(_)))
        .unwrap_or(first);
    let pct = toks.iter().filter(|t| matches!(t, Nt::Pct)).count();
    let mut x = x * 100f64.powi(i32::try_from(pct).unwrap_or(0));
    // Separators right after the last digit placeholder scale by 1000 each.
    let mut end = last + 1;
    while matches!(toks.get(end), Some(Nt::Group)) {
        x /= 1000.0;
        end += 1;
    }
    let head = literals(&toks[..first]);
    let tail = literals(&toks[end..]);
    // A separator anywhere between placeholders groups the integer part.
    let grouped = toks[first..=last].iter().any(|t| matches!(t, Nt::Group));
    let mut shown_nonzero = true;
    let body = if let Some(ep) = toks.iter().position(|t| matches!(t, Nt::Exp(_))) {
        scientific(toks, first, ep, x, grouped, loc, &mut shown_nonzero)
    } else if let Some(sl) = toks.iter().position(|t| matches!(t, Nt::Slash)) {
        fractional(toks, first, sl, x, &mut shown_nonzero)
    } else {
        plain(toks, first, last, x, grouped, loc, &mut shown_nonzero)
    };
    let mut out = String::new();
    if negative && shown_nonzero {
        out.push('-');
    }
    out.push_str(&head);
    out.push_str(&body);
    out.push_str(&tail);
    out
}

fn group_digits_in(s: &str, sep: char) -> String {
    let digits: String = s.chars().filter(char::is_ascii_digit).collect();
    let lead: String = s.chars().take_while(|c| !c.is_ascii_digit()).collect();
    format!("{lead}{}", group(&digits, sep))
}

fn plain(
    toks: &[Nt],
    first: usize,
    last: usize,
    x: f64,
    grouped: bool,
    loc: &Locale,
    nonzero: &mut bool,
) -> String {
    let dp = toks[first..=last]
        .iter()
        .position(|t| matches!(t, Nt::Dec))
        .map(|p| p + first);
    let int_end = dp.unwrap_or(last + 1);
    let dec_toks: &[Nt] = dp.map_or(&[], |p| &toks[p + 1..=last]);
    let dec_ph = placeholders(dec_toks);
    let nd = dec_ph.len();
    let rounded = round_decimal(x, i32::try_from(nd).unwrap_or(0));
    *nonzero = rounded != 0.0;
    let s = format!("{rounded:.nd$}");
    let (int_s, dec_s) = s.split_once('.').unwrap_or((&s, ""));
    let int_digits = if int_s == "0" { "" } else { int_s };

    // Integer part: digits go to the placeholders from the right.
    let mut rest: Vec<char> = int_digits.chars().collect();
    let mut right_to_left: Vec<String> = Vec::new();
    for t in toks[first..int_end].iter().rev() {
        match t {
            Nt::Digit(ph) => match (rest.pop(), ph) {
                (Some(d), _) => right_to_left.push(d.to_string()),
                (None, '0') => right_to_left.push("0".into()),
                (None, '?') => right_to_left.push(" ".into()),
                _ => {}
            },
            Nt::Lit(l) => right_to_left.push(l.clone()),
            Nt::Pct => right_to_left.push("%".into()),
            _ => {}
        }
    }
    right_to_left.reverse();
    let mut int_out: String = rest.iter().collect::<String>() + &right_to_left.concat();
    if grouped {
        int_out = group_digits_in(&int_out, loc.group);
    }

    // Decimals: trailing `#` zeros go, trailing `?` zeros become blanks; text
    // between placeholders stays.
    let digits: Vec<char> = (0..nd)
        .map(|i| dec_s.chars().nth(i).unwrap_or('0'))
        .collect();
    let mut keep = nd;
    while keep > 0 && digits[keep - 1] == '0' && dec_ph[keep - 1] != '0' {
        keep -= 1;
    }
    let mut dec_out = String::new();
    let mut k = 0;
    for t in dec_toks {
        match t {
            Nt::Digit(ph) => {
                if k < keep {
                    dec_out.push(digits[k]);
                } else if *ph == '?' {
                    dec_out.push(' ');
                }
                k += 1;
            }
            Nt::Lit(l) => dec_out.push_str(l),
            Nt::Dec | Nt::Group => {
                if let Nt::Dec = t {
                    dec_out.push(loc.decimal);
                }
            }
            _ => {}
        }
    }
    let mut out = int_out;
    if keep > 0 || dec_ph.contains(&'0') || dec_ph.contains(&'?') {
        out.push(loc.decimal);
        out.push_str(&dec_out);
    } else if dec_out.chars().any(|c| c != ' ') {
        out.push_str(&dec_out);
    }
    out
}

fn scientific(
    toks: &[Nt],
    first: usize,
    ep: usize,
    x: f64,
    grouped: bool,
    loc: &Locale,
    nonzero: &mut bool,
) -> String {
    let plus = matches!(toks[ep], Nt::Exp(true));
    let mant = &toks[first..ep];
    let dpos = mant.iter().position(|t| matches!(t, Nt::Dec));
    let int_ph = placeholders(&mant[..dpos.unwrap_or(mant.len())]);
    let nd = dpos.map_or(0, |p| placeholders(&mant[p..]).len());
    let exp_ph = placeholders(&toks[ep + 1..]).len();
    let k = i32::try_from(int_ph.len().max(1)).unwrap_or(1);
    let mut exp = if x == 0.0 {
        0
    } else {
        x.log10().floor() as i32
    };
    if k > 1 {
        exp = exp.div_euclid(k) * k;
    }
    let mut m = if x == 0.0 { 0.0 } else { x / 10f64.powi(exp) };
    m = round_decimal(m, i32::try_from(nd).unwrap_or(0));
    if m >= 10f64.powi(k) {
        m /= 10f64.powi(k);
        exp += k;
    }
    *nonzero = m != 0.0;
    let s = format!("{m:.nd$}");
    let (mi, md) = s.split_once('.').unwrap_or((&s, ""));
    let mut int = mi.to_string();
    let min_int = int_ph.iter().filter(|&&c| c == '0').count();
    while int.len() < min_int {
        int.insert(0, '0');
    }
    if grouped {
        int = group_digits_in(&int, loc.group);
    }
    let mut out = int;
    if nd > 0 {
        out.push(loc.decimal);
        out.push_str(md);
    }
    out.push('E');
    if exp < 0 {
        out.push('-');
    } else if plus {
        out.push('+');
    }
    let e = exp.unsigned_abs().to_string();
    out.push_str(&"0".repeat(exp_ph.saturating_sub(e.len())));
    out.push_str(&e);
    out
}

/// `# ?/?`, `?/?`, `# ??/8`.
fn fractional(toks: &[Nt], first: usize, sl: usize, x: f64, nonzero: &mut bool) -> String {
    let numer_ph = (first..sl)
        .rev()
        .take_while(|&i| matches!(toks[i], Nt::Digit(_)))
        .count();
    let numer_start = sl - numer_ph;
    let den_ph = toks[sl + 1..]
        .iter()
        .take_while(|t| matches!(t, Nt::Digit(_)))
        .count();
    let fixed: Option<u64> = match (den_ph, toks.get(sl + 1)) {
        (0, Some(Nt::Lit(s))) => s
            .chars()
            .take_while(char::is_ascii_digit)
            .collect::<String>()
            .parse()
            .ok(),
        _ => None,
    };
    let fixed_len = fixed.map_or(0, |f| f.to_string().len());
    let int_ph = placeholders(&toks[first..numer_start]);
    let whole = x.floor();
    let (mut n, d) = fraction(
        x - whole,
        10u64.pow(den_ph as u32).saturating_sub(1).max(1),
        fixed,
    );
    let mut whole = whole;
    if n == d {
        whole += 1.0;
        n = 0;
    }
    if int_ph.is_empty() {
        // Improper fraction.
        n += whole as u64 * d;
        whole = 0.0;
    }
    *nonzero = whole != 0.0 || n != 0;
    let whole_s = if whole == 0.0 && n == 0 {
        "0".to_string()
    } else if whole == 0.0 && !int_ph.contains(&'0') {
        String::new()
    } else {
        format!("{whole:.0}")
    };
    let mid = literals(&toks[first..numer_start]);
    let mut out = whole_s + &mid;
    // LibreOffice pads the `?` places of a fraction with figure spaces (U+2007).
    const FIGURE_SPACE: char = '\u{2007}';
    let pad = |n: usize| FIGURE_SPACE.to_string().repeat(n);
    if n == 0 && !int_ph.is_empty() {
        // No fraction: blanks as wide as `?/?` would be.
        out.push_str(&pad(numer_ph));
        out.push(' ');
        out.push_str(&pad(den_ph.max(fixed_len)));
    } else {
        let (ns, ds) = (n.to_string(), d.to_string());
        out.push_str(&pad(numer_ph.saturating_sub(ns.len())));
        out.push_str(&ns);
        out.push('/');
        out.push_str(&ds);
        if fixed.is_none() {
            out.push_str(&pad(den_ph.saturating_sub(ds.len())));
        }
    }
    out
}

fn cond_holds(c: (Cmp, f64), x: f64) -> bool {
    match c.0 {
        Cmp::Lt => x < c.1,
        Cmp::Le => x <= c.1,
        Cmp::Gt => x > c.1,
        Cmp::Ge => x >= c.1,
        Cmp::Eq => x == c.1,
        Cmp::Ne => x != c.1,
    }
}

/// Whether a color or condition bracket comes after other content of its
/// section: LibreOffice refuses such a code (`yyyy [Blue]`).
fn misplaced_bracket(section: &str) -> bool {
    let cs: Vec<char> = section.chars().collect();
    let mut seen = false;
    let mut i = 0;
    while i < cs.len() {
        match cs[i] {
            '"' => {
                seen = true;
                i = cs[i + 1..]
                    .iter()
                    .position(|&x| x == '"')
                    .map_or(cs.len(), |p| i + 1 + p);
            }
            '\\' | '_' | '*' => {
                seen = true;
                i += 1;
            }
            '[' => {
                let end = cs[i..]
                    .iter()
                    .position(|&x| x == ']')
                    .map_or(cs.len(), |p| i + p);
                let inner: String = cs[i + 1..end.min(cs.len())].iter().collect();
                let l = inner.to_ascii_lowercase();
                let color = matches!(
                    l.as_str(),
                    "black" | "blue" | "cyan" | "green" | "magenta" | "red" | "white" | "yellow"
                ) || l.starts_with("color");
                if (color || parse_cond(&inner).is_some()) && seen {
                    return true;
                }
                i = end;
            }
            _ => seen = true,
        }
        i += 1;
    }
    false
}

/// Format a number with a format code written in the locale's notation, as
/// `TEXT` takes it. `None`: the code is not valid (`#VALUE!`). `General` (or an
/// empty code) uses the general number format.
/// A `[$-…]` tag without a currency symbol whose language is not a hexadecimal id (`[$-en-US]`,
/// `[$-ZZZ]`): LibreOffice answers `Err:502`. With a symbol (`[$€-en-US]`) the language is not read.
fn bad_language_tag(fmt: &str) -> bool {
    let mut rest = fmt;
    while let Some(at) = rest.find("[$-") {
        let tail = &rest[at + 3..];
        let end = tail.find(']').unwrap_or(tail.len());
        let lang = &tail[..end];
        if !lang.is_empty() && u32::from_str_radix(lang, 16).is_err() {
            return true;
        }
        rest = &tail[end..];
    }
    false
}

pub fn format_number(x: f64, fmt: &str, loc: &Locale, sys: DateSystem) -> Option<String> {
    let (dec, grp) = code_separators(fmt, loc);
    format_with(x, fmt, loc, sys, dec, grp)
}

/// [`format_number`] for a code the engine writes itself, always in English
/// notation (`FIXED`).
pub fn format_number_en(x: f64, fmt: &str, loc: &Locale, sys: DateSystem) -> String {
    format_with(x, fmt, loc, sys, '.', ',').unwrap_or_default()
}

fn format_with(
    x: f64,
    fmt: &str,
    loc: &Locale,
    sys: DateSystem,
    dec: char,
    grp: char,
) -> Option<String> {
    if fmt.trim().eq_ignore_ascii_case("general") {
        return Some(fmt_number(x, loc));
    }
    let raw = sections(fmt);
    if raw.iter().any(|s| misplaced_bracket(s)) || bad_language_tag(fmt) {
        return None;
    }
    // A last section with only `@` is the text section, not a number's.
    let is_text_section =
        |s: &str| s.contains('@') && !s.chars().any(|c| matches!(c, '0' | '#' | '?'));
    let numeric = if raw.len() > 1 && is_text_section(raw[raw.len() - 1]) {
        raw.len() - 1
    } else {
        raw.len().min(3)
    };
    let raw = &raw[..numeric.max(1)];
    let secs: Vec<Section> = raw.iter().map(|s| parse_section(s, dec, grp)).collect();
    let n = secs.len();
    // Which section, and whether the sign is written.
    let any_cond = secs.iter().any(|s| s.cond.is_some());
    let (idx, minus) = if n == 1 {
        (0, x < 0.0)
    } else if !any_cond {
        match n {
            2 => (usize::from(x < 0.0), false),
            _ => (
                if x < 0.0 {
                    1
                } else if x == 0.0 {
                    2
                } else {
                    0
                },
                false,
            ),
        }
    } else {
        let defaults = if n == 2 {
            [(Cmp::Ge, 0.0), (Cmp::Lt, 0.0)]
        } else {
            [(Cmp::Gt, 0.0), (Cmp::Lt, 0.0)]
        };
        let pick = (0..n)
            .find(|&i| i == n - 1 || cond_holds(secs[i].cond.unwrap_or(defaults[i.min(1)]), x))
            .unwrap_or(0);
        (pick, false)
    };
    let section_src = raw[idx];
    if let Some(toks) = tokenize_dt(section_src) {
        return Some(render_dt(&toks, x, loc, sys).unwrap_or_else(|| fmt_number(x, loc)));
    }
    let sec = &secs[idx];
    if sec
        .toks
        .iter()
        .any(|t| matches!(t, Nt::Lit(l) if l.eq_ignore_ascii_case("general")))
    {
        return Some(fmt_number(x, loc));
    }
    if sec.toks.iter().any(|t| matches!(t, Nt::Text)) {
        // `@` over a number: the number as General shows it.
        let g = fmt_number(x, loc);
        return Some(
            sec.toks
                .iter()
                .map(|t| match t {
                    Nt::Lit(l) => l.clone(),
                    Nt::Text => g.clone(),
                    _ => String::new(),
                })
                .collect(),
        );
    }
    Some(render_section(sec, x.abs(), minus, loc))
}

/// Format a text value with the text section (`@`) of a format code, if it has one.
pub fn format_text(text: &str, fmt: &str, loc: &Locale) -> Option<String> {
    let (dec, grp) = code_separators(fmt, loc);
    let raw = sections(fmt);
    let pick = if raw.len() >= 4 {
        Some(raw[3])
    } else {
        raw.iter().copied().find(|s| s.contains('@'))
    }?;
    let sec = parse_section(pick, dec, grp);
    Some(
        sec.toks
            .iter()
            .map(|t| match t {
                Nt::Lit(l) => l.clone(),
                Nt::Text => text.to_string(),
                _ => String::new(),
            })
            .collect(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn f(x: f64, fmt: &str) -> String {
        format_number(x, fmt, &Locale::en_us(), DateSystem::NULL_1899).unwrap()
    }

    #[test]
    fn numbers() {
        assert_eq!(f(1234.567, "0.00"), "1234.57");
        assert_eq!(f(1234.567, "#,##0.00"), "1,234.57");
        assert_eq!(f(0.256, "0%"), "26%");
        assert_eq!(f(5.0, "000"), "005");
        assert_eq!(f(0.5, "#.##"), ".5");
        assert_eq!(f(-3.5, "0.0"), "-3.5");
        assert_eq!(f(-3.5, "0.0;(0.0)"), "(3.5)");
        assert_eq!(f(12.0, "\"EUR \"0"), "EUR 12");
        assert_eq!(f(1.5, "General"), "1.5");
        assert_eq!(
            format_number(1234.5, "#,##0.00", &Locale::es_es(), DateSystem::NULL_1899).unwrap(),
            "1.234,50"
        );
    }

    #[test]
    fn a_language_tag_without_a_symbol_must_be_a_hexadecimal_id() {
        let en = Locale::en_us();
        let f = |c: &str| format_number(45367.0, c, &en, DateSystem::NULL_1899);
        for bad in ["[$-en-US]dddd", "[$-en-US]#,##0", "[$-ZZZ]dddd", "[$-x]0"] {
            assert_eq!(f(bad), None, "{bad}");
        }
        assert_eq!(f("[$-409]dddd").as_deref(), Some("Saturday"));
        assert_eq!(f("[$-]dddd").as_deref(), Some("Saturday"));
        assert_eq!(f("[$$-en-US]#,##0").as_deref(), Some("$45,367"));
        assert_eq!(f("[$€-ZZZ]0").as_deref(), Some("€45367"));
    }

    #[test]
    fn libreoffice_number_codes() {
        // Values LibreOffice 26.8 gave (tests/semantics/oracle.rs holds the full matrix).
        assert_eq!(f(1234.5678, "# ?/?"), "1234 4/7");
        assert_eq!(f(0.5, "# ??/??"), " \u{2007}1/2\u{2007}");
        assert_eq!(f(0.0, "# ?/?"), "0 \u{2007} \u{2007}");
        assert_eq!(f(1234.5678, "0.00E+00"), "1.23E+03");
        assert_eq!(f(12345678.9, "##0.0E+0"), "12.3E+6");
        assert_eq!(f(1234.5678, "000-00"), "012-35");
        assert_eq!(f(1000000.0, "0,"), "1000");
        assert_eq!(f(-1234.5678, "[>100]0.0;0"), "1235");
        assert_eq!(f(-1234.5678, "#,##0_);(#,##0)"), "(1,235)");
        assert_eq!(f(1234.5678, "[$$-en-US]#,##0"), "$1,235");
        assert_eq!(f(-1234.5678, "[$€-407]#,##0.00"), "-€1,234.57");
        assert_eq!(f(0.0, "#.##"), "");
        assert_eq!(f(0.5, "#.##"), ".5");
        assert_eq!(f(1234.5678, "@"), "1234.5678");
        assert_eq!(f(1234.5678, "#.###.##0"), "1234.567.800");
    }

    #[test]
    fn libreoffice_codes_follow_the_locale() {
        let es = Locale::es_es();
        let g = |x: f64, code: &str| format_number(x, code, &es, DateSystem::NULL_1899).unwrap();
        // `#,##0` is read in English form; `#,##0;(#,##0)` is Spanish: the comma
        // is the decimal separator.
        assert_eq!(g(1234.5678, "#,##0"), "1.235");
        assert_eq!(g(1234.5678, "#,##0;(#,##0)"), "1234,568");
        assert_eq!(g(1234.5678, "#.##0,00"), "1.234,57");
        assert_eq!(g(1234.5678, "0.0"), "1.235");
    }

    #[test]
    fn libreoffice_time_codes() {
        let g = |x: f64, code: &str| {
            format_number(x, code, &Locale::en_us(), DateSystem::NULL_1899).unwrap()
        };
        // A clock shows what it has reached; elapsed time rounds to the second.
        assert_eq!(g(0.999999, "h:mm"), "23:59");
        assert_eq!(g(0.999999, "[h]:mm"), "24:00");
        assert_eq!(g(45658.75123, "[h]:mm"), "1095810:01");
        assert_eq!(g(45658.75123, "hh:mm:ss.0"), "18:01:46.3");
        assert_eq!(g(45658.75123, "h A/P"), "6 p");
        assert_eq!(g(-1.5, "yyyy-mm-dd"), "1899-12-28");
        assert_eq!(g(-1.5, "mm/dd/yy;@"), "12/28/99");
        assert_eq!(g(45292.0, "NN"), "Mon");
        // A color after the content is not a valid code.
        assert!(
            format_number(1.0, "yyyy [Blue]", &Locale::en_us(), DateSystem::NULL_1899).is_none()
        );
    }

    #[test]
    fn dates() {
        assert_eq!(f(45292.0, "yyyy-mm-dd"), "2024-01-01");
        assert_eq!(f(45292.5, "yyyy-mm-dd hh:mm:ss"), "2024-01-01 12:00:00");
        assert_eq!(f(45292.0, "dddd d mmmm yyyy"), "Monday 1 January 2024");
        assert_eq!(f(45292.75, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(
            format_number(
                45292.0,
                "dddd d mmmm",
                &Locale::es_es(),
                DateSystem::NULL_1899
            )
            .unwrap(),
            "lunes 1 enero"
        );
    }
}
