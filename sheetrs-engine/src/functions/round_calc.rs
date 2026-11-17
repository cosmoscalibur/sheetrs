//! `ROUND`, `ROUNDUP` and `ROUNDDOWN` as `LibreOffice` calculates them.
//!
//! `LibreOffice` does not round on the 15 significant digits of the value (that is `round_dec`, which the
//! number formats use). Two steps decide the result:
//!
//! 1. The count of digits is cut to a whole number with a 15-digit tolerance (`ROUND(x; 2.9999999999999996)`
//!    has 3 digits) and clamped to 16 bits. `ROUNDUP` and `ROUNDDOWN` of a value with a fractional part and
//!    fewer than 12 digits first bring the value to 12 significant digits, so that `ROUNDUP(8.06 - 8; 2)` is
//!    `0.06` and not `0.07` (the noise of `8.06 - 8` is `5E-16`).
//! 2. The number is scaled by `10^digits`, the product is brought to 15 significant digits (a product that is a
//!    hair under an integer is that integer), the floor (plus one half for `ROUND`) or the ceiling is taken,
//!    and it is scaled back. The digits are cut to the fractional bits the number has, a whole number is left
//!    alone, and a negative number that rounds to nothing is `-0`.
//!
//! It is verified against what `LibreOffice` 26.8.2 answers: 1.3 million cases (`-14` to `14` digits, six
//! families of values, both signs, inputs built exactly inside it) and a probe of the forms of the call, with
//! no difference (`tests/assets/round_oracle.tsv`, `tests/assets/round_forms.tsv`).

use super::math::Mode;

/// 2^41: above it the approximation to 15 digits is not made.
const BIG: f64 = 2_199_023_255_552.0;
/// 2^52: from it a number has no fractional bits.
const NO_FRACTION: f64 = 4_503_599_627_370_496.0;
/// 2^53: the integers below are all representable.
const ALL_INTEGERS: f64 = 9_007_199_254_740_992.0;
/// The significant digits `ROUNDUP` and `ROUNDDOWN` bring a value to before rounding it.
const PRE_DIGITS: f64 = 12.0;

/// `10^n` as a decimal literal gives it (correctly rounded); 0 and infinity beyond the range of a `f64`.
fn pow10(n: i32) -> f64 {
    const EXACT: [f64; 23] = [
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16,
        1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
    ];
    match n {
        0..=22 => EXACT[n as usize],
        23..=308 => format!("1e{n}").parse().unwrap_or(f64::INFINITY),
        309.. => f64::INFINITY,
        -323..=-1 => format!("1e{n}").parse().unwrap_or(0.0),
        _ => 0.0,
    }
}

/// Whether the absolute value `a` is a whole number a `f64` holds.
fn is_whole(a: f64) -> bool {
    a < ALL_INTEGERS && a.trunc() == a
}

/// The bits a number has after the point (0 for a whole number).
fn fraction_bits(a: f64) -> i32 {
    if a == 0.0 {
        return 0;
    }
    let bits = a.to_bits();
    let exponent = ((bits >> 52) & 0x7ff) as i32 - 1023;
    if exponent >= 52 {
        return 0;
    }
    let fraction = bits & ((1u64 << 52) - 1);
    let lowest = if fraction == 0 {
        53
    } else {
        fraction.trailing_zeros() as i32 + 1
    };
    (53 - lowest - exponent).max(0)
}

/// `v` brought to 15 significant digits, in `f64` arithmetic. A number that is whole, that has few
/// fractional bits (11 or less), or that is above 2^41, is as it is.
fn to_15_digits(v: f64) -> f64 {
    if v == 0.0 || !v.is_finite() || v > BIG {
        return v;
    }
    let negative = v.is_sign_negative();
    let a = v.abs();
    if is_whole(a) || fraction_bits(a) <= 11 {
        return v;
    }
    let shift = 14 - a.log10().floor() as i32;
    let factor = pow10(shift.abs());
    let scaled = if shift < 0 { a / factor } else { a * factor };
    if !scaled.is_finite() {
        return v;
    }
    let rounded = scaled.round();
    let back = if shift < 0 {
        rounded * factor
    } else {
        rounded / factor
    };
    if !back.is_finite() {
        return v;
    }
    if negative { -back } else { back }
}

/// The rounding of the library: `x` at `digits` places in `mode`.
fn round_places(x: f64, mut digits: i32, mode: Mode) -> f64 {
    if !x.is_finite() || x == 0.0 {
        return x;
    }
    let negative = x.is_sign_negative();
    let mut v = x.abs();
    if digits >= 0 && (v >= NO_FRACTION || is_whole(v)) {
        return x;
    }
    let mut factor = 1.0;
    if digits != 0 {
        if digits > 0 {
            // No more places than the fractional bits of the number have.
            let places = 52 - (((v.to_bits() >> 52) & 0x7ff) as i32 - 1023);
            if places <= 0 {
                return x;
            }
            digits = digits.min(places);
        }
        factor = pow10(digits.abs());
        if factor == 0.0 || (digits < 0 && !factor.is_finite()) {
            return 0.0;
        }
        if !factor.is_finite() {
            return x;
        }
        v = if digits < 0 { v / factor } else { v * factor };
        if !v.is_finite() {
            return x;
        }
    }
    if v < NO_FRACTION {
        v = match mode {
            Mode::Nearest => to_15_digits(v + 0.5).floor(),
            Mode::Down => to_15_digits(v).floor(),
            Mode::Away => to_15_digits(v).ceil(),
        };
    }
    if digits != 0 {
        v = if digits < 0 { v * factor } else { v / factor };
    }
    if !v.is_finite() {
        return x;
    }
    if negative { -v } else { v }
}

/// `ROUND`, `ROUNDUP` or `ROUNDDOWN` of `x` with `digits` (`None`: the form with one argument).
pub(super) fn round_calc(x: f64, digits: Option<f64>, mode: Mode) -> f64 {
    let Some(digits) = digits else {
        return round_places(x, 0, mode);
    };
    let digits = if digits > f64::from(i16::MAX) {
        f64::from(i16::MAX)
    } else if digits < f64::from(i16::MIN) {
        f64::from(i16::MIN)
    } else if digits < 0.0 {
        to_15_digits(digits).ceil()
    } else {
        to_15_digits(digits).floor()
    };
    let places = digits as i32;
    if mode != Mode::Nearest && digits < PRE_DIGITS && x % 1.0 != 0.0 {
        // Bring the value to 12 significant digits first, so that the noise of a decimal that was
        // calculated does not decide the rounding.
        let shift = x.abs().log10().floor() + 1.0 - PRE_DIGITS;
        let mut value = if shift < 0.0 {
            x * 10f64.powf(-shift)
        } else {
            x / 10f64.powf(shift)
        };
        if value.is_finite() {
            if mode == Mode::Away {
                value = to_15_digits(value).floor();
            }
            let rounded = round_places(value, (digits + shift) as i32, mode);
            return if shift < 0.0 {
                rounded / 10f64.powf(-shift)
            } else {
                rounded * 10f64.powf(shift)
            };
        }
    }
    round_places(x, places, mode)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn round(x: f64, d: f64) -> f64 {
        round_calc(x, Some(d), Mode::Nearest)
    }

    fn up(x: f64, d: f64) -> f64 {
        round_calc(x, Some(d), Mode::Away)
    }

    fn down(x: f64, d: f64) -> f64 {
        round_calc(x, Some(d), Mode::Down)
    }

    #[test]
    fn a_half_goes_away_from_zero_and_a_decimal_that_is_a_hair_under_goes_with_it() {
        assert_eq!(round(2.5, 0.0), 3.0);
        assert_eq!(round(-2.5, 0.0), -3.0);
        assert_eq!(round(1.005, 2.0), 1.01);
        assert_eq!(round(2.675, 2.0), 2.68);
        assert_eq!(round(20_803.499_999_999_996, 0.0), 20_804.0);
        assert_eq!(round(1234.5678, -2.0), 1200.0);
    }

    #[test]
    fn a_negative_number_that_rounds_to_nothing_is_negative_zero() {
        for v in [round(-0.4, 0.0), round(-0.0001, 2.0), down(-0.4, 0.0)] {
            assert_eq!(v, 0.0);
            assert!(v.is_sign_negative());
        }
        assert!(round(0.4, 0.0).is_sign_positive());
    }

    #[test]
    fn the_product_is_brought_to_15_digits_below_2_pow_41_and_left_as_it_is_above() {
        // 0.29 · 100 is 28.999999999999996: the 15 digits bring it to 29.
        assert_eq!(down(0.29, 2.0), 0.29);
        // 36243.59 · 10^8 is 3.6e12, above 2^41: the binary value is truncated.
        assert_eq!(down(36_243.59, 8.0), 36_243.589_999_99);
    }

    #[test]
    fn roundup_and_rounddown_bring_the_value_to_12_digits_first() {
        // 8.06 - 8 is 0.0600000000000005.
        assert_eq!(up(8.06 - 8.0, 2.0), 0.06);
        // 573.9951249000158 has a tail below the 12 digits.
        assert_eq!(up(573.995_124_900_015_8, 8.0), 573.995_124_9);
        // With 12 digits or more it is not done.
        assert_eq!(up(0.123_456_789_012_34, 13.0), 0.123_456_789_012_4);
    }

    #[test]
    fn the_digits_are_cut_with_the_tolerance_of_15_digits_and_the_one_argument_form_has_none_of_it()
    {
        assert_eq!(round(1.234_56, 2.999_999_999_999_999_6), 1.235);
        assert_eq!(round(1234.5, -1.999_999_999_999_999_8), 1200.0);
        assert_eq!(round_calc(2.5, None, Mode::Nearest), 3.0);
        assert_eq!(round_calc(2.4, None, Mode::Away), 3.0);
        assert_eq!(round_calc(-2.4, None, Mode::Down), -2.0);
        // Clamped to the 16 bits of Calc: no change, and no hang.
        assert_eq!(round(1.5, 40_000.0), 1.5);
    }

    #[test]
    fn a_whole_number_and_a_number_without_fractional_bits_are_left_alone() {
        assert_eq!(round(123_456_789_012.0, 3.0), 123_456_789_012.0);
        assert_eq!(up(4_503_599_627_370_497.0, 2.0), 4_503_599_627_370_497.0);
        assert!(round(f64::NAN, 2.0).is_nan());
        assert_eq!(round(f64::INFINITY, 2.0), f64::INFINITY);
    }

    /// The decimal quantity a result stands for: `result · 10^digits` as an integer, which the 15 digits that
    /// the ODS of `LibreOffice` carries hold when the result has fewer than 15.
    fn quantity(result: f64, digits: i32) -> Option<i64> {
        let k = if digits >= 0 {
            result * pow10(digits)
        } else {
            result / pow10(-digits)
        };
        (k.abs() < 1e15).then(|| k.round() as i64)
    }

    /// Lines of `tests/assets/round_oracle.tsv`: the input as the hex of the `f64`, the digits, the function
    /// (0 `ROUND`, 1 `ROUNDUP`, 2 `ROUNDDOWN`) and the result of `LibreOffice` 26.8.2.
    fn oracle(text: &str) -> impl Iterator<Item = (f64, i32, Mode, f64)> + '_ {
        text.lines()
            .filter(|l| !l.starts_with('#') && !l.is_empty())
            .map(|l| {
                let c: Vec<&str> = l.split('\t').collect();
                (
                    f64::from_bits(u64::from_str_radix(c[1], 16).unwrap()),
                    c[2].parse().unwrap(),
                    [Mode::Nearest, Mode::Away, Mode::Down][c[3].parse::<usize>().unwrap()],
                    c[4].parse().unwrap(),
                )
            })
    }

    #[test]
    fn the_functions_give_what_libreoffice_gave_for_the_cases_of_its_oracle() {
        let text = include_str!("../../tests/assets/round_oracle.tsv");
        let mut n = 0;
        for (x, d, mode, lo) in oracle(text) {
            let got = round_calc(x, Some(f64::from(d)), mode);
            assert_eq!(
                quantity(got, d),
                quantity(lo, d),
                "x = {x:e}, digits = {d}, LibreOffice {lo}, here {got}"
            );
            n += 1;
        }
        assert!(n > 9_000, "{n} cases");
    }

    /// The number as an ODS holds it: its shortest decimal text cut to 15 significant digits, half up.
    fn as_exported(x: f64) -> f64 {
        let t = format!("{x:e}");
        let (mant, exp) = t.split_once('e').unwrap();
        let neg = mant.starts_with('-');
        let digits: Vec<u8> = mant
            .bytes()
            .filter(u8::is_ascii_digit)
            .map(|b| b - b'0')
            .collect();
        if digits.len() <= 15 {
            return x;
        }
        let mut d: Vec<u8> = digits[..15].to_vec();
        let mut exp: i32 = exp.parse().unwrap();
        if digits[15] >= 5 {
            let mut i = 14;
            loop {
                if d[i] == 9 {
                    d[i] = 0;
                    if i == 0 {
                        d.insert(0, 1);
                        d.pop();
                        exp += 1;
                        break;
                    }
                    i -= 1;
                } else {
                    d[i] += 1;
                    break;
                }
            }
        }
        let text: String = d.iter().map(|v| char::from(b'0' + v)).collect();
        let v: f64 = format!("{}.{}e{exp}", &text[..1], &text[1..])
            .parse()
            .unwrap();
        if neg { -v } else { v }
    }

    /// The forms of the call: no digits, digits that are not whole, 16-bit limits, and the sign of a zero.
    /// Lines: function, `x` (hex), digits (empty for the one-argument form), result, sign of a zero result.
    #[test]
    fn the_forms_of_the_call_give_what_libreoffice_gave() {
        let text = include_str!("../../tests/assets/round_forms.tsv");
        let mut n = 0;
        for l in text
            .lines()
            .filter(|l| !l.starts_with('#') && !l.is_empty())
        {
            let c: Vec<&str> = l.split('\t').collect();
            let x = f64::from_bits(u64::from_str_radix(c[1], 16).unwrap());
            let mode = match c[0] {
                "ROUND" => Mode::Nearest,
                "ROUNDUP" => Mode::Away,
                _ => Mode::Down,
            };
            let digits = (!c[2].is_empty()).then(|| c[2].parse::<f64>().unwrap());
            let got = round_calc(x, digits, mode);
            let lo: f64 = c[3]
                .replace("E+0", "e")
                .replace("E+", "e")
                .replace("E-0", "e-")
                .replace("E-", "e-")
                .parse()
                .unwrap();
            // The oracle (an ODS) carries 15 significant digits: the result must agree to the last of them.
            let to_15 = as_exported(got);
            assert_eq!(
                to_15, lo,
                "{} x = {x:e}, digits = {digits:?}: LibreOffice {lo}, here {got}",
                c[0]
            );
            if !c[4].is_empty() {
                assert_eq!(
                    got.is_sign_negative(),
                    c[4] == "neg",
                    "sign of the zero of {} x = {x:e}",
                    c[0]
                );
            }
            n += 1;
        }
        assert!(n > 900, "{n} cases");
    }
}
