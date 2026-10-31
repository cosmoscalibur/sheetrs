//! `COLOR`, `CURRENT`, `STYLE` and `EUROCONVERT`, the functions of LibreOffice that need no spreadsheet
//! semantics. The results are checked against the oracle (`tests/assets/semantics.*`).

use super::{Reg, num, round_decimal};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// Currency, units per euro, decimals (the table of LibreOffice's `lclConvertMoney`).
const CURRENCIES: [(&str, f64, i32); 22] = [
    ("EUR", 1.0, 2),
    ("ATS", 13.7603, 2),
    ("BEF", 40.3399, 0),
    ("DEM", 1.95583, 2),
    ("ESP", 166.386, 0),
    ("FIM", 5.94573, 2),
    ("FRF", 6.55957, 2),
    ("IEP", 0.787564, 2),
    ("ITL", 1936.27, 0),
    ("LUF", 40.3399, 0),
    ("NLG", 2.20371, 2),
    ("PTE", 200.482, 2),
    ("GRD", 340.750, 2),
    ("SIT", 239.640, 2),
    ("MTL", 0.429300, 2),
    ("CYP", 0.585274, 2),
    ("SKK", 30.1260, 2),
    ("EEK", 15.6466, 2),
    ("LVL", 0.702804, 2),
    ("LTL", 3.45280, 2),
    ("HRK", 7.53450, 2),
    ("BGN", 1.95583, 2),
];

fn currency(name: &str) -> Option<(f64, i32)> {
    CURRENCIES
        .iter()
        .find(|c| c.0.eq_ignore_ascii_case(name))
        .map(|c| (c.1, c.2))
}

/// An error argument comes before a limit that is not met; the first of them is the one.
fn first_error(e: &mut Engine, a: &[Value]) -> Res<()> {
    for v in a {
        if let Value::Error(k) = e.scalar(v.clone()) {
            return Err(k);
        }
    }
    Ok(())
}

pub(super) fn register(r: &mut Reg) {
    // The arguments are read from the last to the first, so that the error of the last one wins.
    r.eager("COLOR", 3, 4, |e, a| {
        first_error(e, a)?;
        let mut parts = [0.0; 4];
        for i in (0..a.len()).rev() {
            let v = e.num(&a[i])?.floor();
            if !(0.0..=255.0).contains(&v) {
                return Err(DOMAIN);
            }
            parts[i] = v;
        }
        let alpha = if a.len() == 4 { parts[3] } else { 0.0 };
        num(256.0 * 256.0 * 256.0 * alpha + 256.0 * 256.0 * parts[0] + 256.0 * parts[1] + parts[2])
    });
    // `CURRENT` needs the value of the cell that a conditional format is computing, which a
    // formula of a cell does not have.
    r.eager("CURRENT", 0, 0, |_, _| Err(DOMAIN));
    // `STYLE` applies a cell style (for a time, then another) and answers 0: a program has no
    // styles to apply, and the value is the same whether the style exists or not.
    r.eager("STYLE", 1, 3, |e, a| {
        first_error(e, a)?;
        e.text(&a[0])?;
        if let Some(time) = a.get(1) {
            e.num(time)?;
        }
        if let Some(style) = a.get(2) {
            e.text(style)?;
        }
        num(0.0)
    });
    r.eager("EUROCONVERT", 3, 5, |e, a| {
        first_error(e, a)?;
        let mut precision = 0.0;
        if a.len() == 5 {
            precision = e.num(&a[4])?.floor();
            if precision < 3.0 {
                return Err(DOMAIN);
            }
        }
        let full = a.len() >= 4 && e.boolean(&a[3])?;
        let to = e.text(&a[2])?;
        let from = e.text(&a[1])?;
        let value = e.num(&a[0])?;
        let (Some((from_rate, _)), Some((to_rate, to_decimals))) = (currency(&from), currency(&to))
        else {
            return Err(DOMAIN);
        };
        let result = if from.eq_ignore_ascii_case(&to) {
            value
        } else {
            let mut result = if from.eq_ignore_ascii_case("EUR") {
                value * to_rate
            } else {
                let mut middle = value / from_rate;
                if precision != 0.0 {
                    middle = round_decimal(middle, precision as i32);
                }
                middle * to_rate
            };
            if !full {
                result = round_decimal(result, to_decimals);
            }
            result
        };
        num(result)
    });
}
