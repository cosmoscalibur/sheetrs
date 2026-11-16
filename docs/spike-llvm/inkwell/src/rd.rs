#[derive(Clone,Copy,PartialEq)] pub enum Mode{Nearest,Away,Down}
pub fn round_dec(x: f64, digits: i32, mode: Mode) -> f64 {
    if x == 0.0 || !x.is_finite() {
        return x;
    }
    let sci = format!("{:.14e}", x.abs());
    let (mant, exp) = sci.split_once('e').unwrap_or((&sci, "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let mut d: Vec<u8> = mant
        .bytes()
        .filter(u8::is_ascii_digit)
        .map(|b| b - b'0')
        .collect();
    // Digits kept: those at or above the 10^-digits place.
    let keep = exp + 1 + digits;
    if keep >= i32::try_from(d.len()).unwrap_or(i32::MAX) {
        return x;
    }
    if keep < 0 {
        // Everything is below the rounding place.
        return if mode == Mode::Away {
            10f64.powi(-digits).copysign(x)
        } else {
            0.0
        };
    }
    let k = keep as usize;
    let up = match mode {
        Mode::Nearest => d[k] >= 5,
        Mode::Away => d[k..].iter().any(|&t| t != 0),
        Mode::Down => false,
    };
    d.truncate(k);
    let mut exp = exp;
    if up {
        let mut i = d.len();
        loop {
            if i == 0 {
                d.insert(0, 1);
                exp += 1;
                break;
            }
            i -= 1;
            if d[i] == 9 {
                d[i] = 0;
            } else {
                d[i] += 1;
                break;
            }
        }
    }
    if d.is_empty() {
        return 0.0;
    }
    let s: String = d.iter().map(|x| char::from(b'0' + x)).collect();
    let scale = exp - (i32::try_from(d.len()).unwrap_or(i32::MAX) - 1);
    format!("{s}e{scale}")
        .parse::<f64>()
        .map_or(x, |r| r.copysign(x))
}