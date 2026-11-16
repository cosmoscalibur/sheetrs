//! The rounding of spreadsheets without going through text.
//!
//! `round_dec` rounds on the 15 significant decimal digits of a value: `ROUND(1.005, 2)` is 1.01 although
//! the double nearest to 1.005 is a little below it. The definition (kept as `round_dec_text`) formats the
//! value with `{:.14e}`, rounds the digit string and parses it back. That is exact and slow (about 280 ns).
//!
//! This module computes the same `f64`, bit for bit, in integer arithmetic, and says `None` where it
//! cannot be sure, so that the caller falls back to the text definition:
//!
//! 1. The 15 digits are an integer `m` in `[10^14, 10^15)`: `round_half_even(mant · 2^e2 · 10^k)` with
//!    `k = 14 - e10`, computed in `u128`. That is what `{:.14e}` prints, ties to even included.
//! 2. The rounding at the place asked for is done on `m`: `D = m / 10^(15 - keep)`, increased by one
//!    according to the mode and the digits dropped.
//! 3. The result is `D · 10^scale`. With `D < 2^53` and `|scale| <= 22` one multiplication or division by an
//!    exact power of ten is correctly rounded, which is what parsing the text gives (Clinger's fast path).
//!
//! The tests compare the two functions over millions of values, bit by bit.

use super::math::Mode;

/// `10^i` for `i` in `0..=38` (the largest that fits a `u128`).
const P10: [u128; 39] = {
    let mut t = [1u128; 39];
    let mut i = 1;
    while i < 39 {
        t[i] = t[i - 1] * 10;
        i += 1;
    }
    t
};

/// `10^i` for `i` in `0..=19`, for the steps on the 15 digits (they fit a `u64`, and its division is a
/// machine instruction where the one of a `u128` is a call).
const P10_64: [u64; 20] = {
    let mut t = [1u64; 20];
    let mut i = 1;
    while i < 20 {
        t[i] = t[i - 1] * 10;
        i += 1;
    }
    t
};

/// `10^i` for `i` in `0..=22`, the powers of ten a `f64` holds exactly.
const F10: [f64; 23] = {
    let mut t = [1.0f64; 23];
    let mut i = 1;
    while i < 23 {
        t[i] = t[i - 1] * 10.0;
        i += 1;
    }
    t
};

/// `round_half_even(mant · 2^e2 · 10^k)`, exactly; `None` where a `u128` cannot hold the intermediate
/// value.
fn scaled(mant: u64, e2: i32, k: i32) -> Option<u128> {
    let mut num = u128::from(mant);
    let mut den10 = 1u128;
    if k >= 0 {
        if k > 22 {
            return None;
        }
        num *= P10[k as usize];
    } else {
        if -k > 38 {
            return None;
        }
        den10 = P10[(-k) as usize];
    }
    let mut shift = 0u32;
    if e2 >= 0 {
        if e2 > 74 || num.leading_zeros() < e2 as u32 + 1 {
            return None;
        }
        num <<= e2;
    } else {
        shift = e2.unsigned_abs();
        if shift > 126 {
            return None;
        }
    }
    // The value is num / (2^shift · den10); only one of the two divisors is not 1.
    if shift == 0 {
        let (q, r) = (num / den10, num % den10);
        return Some(if 2 * r > den10 || (2 * r == den10 && q & 1 == 1) {
            q + 1
        } else {
            q
        });
    }
    if den10 != 1 {
        return None;
    }
    let (q, r) = (num >> shift, num & ((1u128 << shift) - 1));
    let half = 1u128 << (shift - 1);
    Some(if r > half || (r == half && q & 1 == 1) {
        q + 1
    } else {
        q
    })
}

/// The 15 significant digits of `ax` (finite, positive) as `(m, e10)`: `m` in `[10^14, 10^15)` and the
/// decimal exponent of its first digit, as `{:.14e}` gives them.
fn digits15(ax: f64) -> Option<(u64, i32)> {
    let bits = ax.to_bits();
    let biased = ((bits >> 52) & 0x7ff) as i32;
    if biased == 0 {
        return None; // subnormal
    }
    let mant = (bits & ((1u64 << 52) - 1)) | (1u64 << 52);
    let e2 = biased - 1075;
    // The decimal exponent from the binary one: `log10(2) · 4096` is 1233 (to the integer), and the +600
    // centres the error. The estimate is within one of the true exponent; the loop looks at `m` and
    // corrects it, which also covers a value just below a power of ten.
    let mut e10 = ((biased - 1023) * 1233 + 600) >> 12;
    for _ in 0..3 {
        let k = 14 - e10;
        let m = scaled(mant, e2, k)?;
        if m < P10[14] {
            e10 -= 1;
            continue;
        }
        if m >= P10[15] {
            e10 += 1;
            continue;
        }
        if m == P10[14] {
            // A carry that made 1.00000000000000, or an exponent one too high for a value just below a
            // power of ten (9.99999999999999e11): the lower exponent says which.
            let lo = scaled(mant, e2, k + 1)?;
            if lo < P10[15] {
                return Some((lo as u64, e10 - 1));
            }
        }
        return Some((m as u64, e10));
    }
    None
}

/// The rounding of `x` at `digits` decimal places in `mode` on its 15 significant digits, or `None` when
/// this function cannot give the answer exactly (the caller then uses the text definition).
pub(super) fn round_exact(x: f64, digits: i32, mode: Mode) -> Option<f64> {
    if x == 0.0 || !x.is_finite() {
        return Some(x);
    }
    let (m, e10) = digits15(x.abs())?;
    let keep = e10 + 1 + digits;
    if keep >= 15 {
        return Some(x);
    }
    if keep < 0 {
        return Some(if mode == Mode::Away {
            10f64.powi(-digits).copysign(x)
        } else {
            0.0
        });
    }
    let k = keep as usize;
    let below = P10_64[15 - k];
    let mut d = m / below;
    let up = match mode {
        Mode::Nearest => (m / P10_64[14 - k]) % 10 >= 5,
        Mode::Away => m % below != 0,
        Mode::Down => false,
    };
    if up {
        d += 1;
    }
    if d == 0 {
        return Some(0.0);
    }
    let scale = e10 - keep + 1;
    if scale.unsigned_abs() > 22 || d >= (1u64 << 53) {
        return None;
    }
    let r = if scale >= 0 {
        d as f64 * F10[scale as usize]
    } else {
        d as f64 / F10[(-scale) as usize]
    };
    Some(r.copysign(x))
}

#[cfg(test)]
mod tests {
    use super::super::math::{Mode, round_dec_text};
    use super::round_exact;

    /// A generator that a failure can be replayed from.
    struct Rng(u64);

    impl Rng {
        fn next(&mut self) -> u64 {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            self.0
        }
    }

    const MODES: [Mode; 3] = [Mode::Nearest, Mode::Away, Mode::Down];

    fn same(x: f64, digits: i32, mode: Mode) -> Option<bool> {
        let fast = round_exact(x, digits, mode)?;
        let text = round_dec_text(x, digits, mode);
        assert_eq!(
            fast.to_bits(),
            text.to_bits(),
            "x = {x:e} ({:#x}), digits = {digits}, mode = {}: exact {fast:e}, text {text:e}",
            x.to_bits(),
            match mode {
                Mode::Nearest => "Nearest",
                Mode::Away => "Away",
                Mode::Down => "Down",
            }
        );
        Some(true)
    }

    /// A value of one of the families that matter: any bit pattern, money, ties, `x.xx5`, a wide range
    /// of magnitudes, and just below a power of ten.
    fn sample(r: &mut Rng) -> f64 {
        match r.next() % 6 {
            0 => f64::from_bits(r.next()),
            1 => {
                (r.next() % 100_000_000) as f64 / 100.0
                    * if r.next().is_multiple_of(2) {
                        1.0
                    } else {
                        -1.0
                    }
            }
            2 => ((r.next() % 20_000_000) as f64 + 0.5) / 10f64.powi((r.next() % 7) as i32),
            3 => (r.next() % 1_000_000) as f64 / 1000.0 + 0.005,
            4 => (r.next() as f64 / u64::MAX as f64) * 10f64.powi((r.next() % 40) as i32 - 20),
            _ => {
                ((r.next() % 9_000) as f64 + 0.999_999_999_999_9)
                    * 10f64.powi((r.next() % 30) as i32 - 10)
            }
        }
    }

    fn compare(cases: u64) -> (u64, u64) {
        let mut r = Rng(0x9E37_79B9_7F4A_7C15);
        let (mut answered, mut declined) = (0, 0);
        for i in 0..cases {
            let digits = (r.next() % 17) as i32 - 6;
            let x = sample(&mut r);
            match same(x, digits, MODES[(i % 3) as usize]) {
                Some(_) => answered += 1,
                None => declined += 1,
            }
        }
        (answered, declined)
    }

    #[test]
    fn exact_rounding_gives_the_bits_of_the_text_rounding() {
        let (answered, declined) = compare(3_000_000);
        // The cases that are declined are mostly bit patterns out of range; the rest must be answered.
        assert!(
            answered > declined,
            "{answered} answered, {declined} declined"
        );
    }

    /// Ten to the eighth cases: by hand, before a release (`cargo test --release -- --ignored`).
    #[test]
    #[ignore = "long: 10^8 cases"]
    fn exact_rounding_gives_the_bits_of_the_text_rounding_long() {
        compare(100_000_000);
    }

    #[test]
    fn known_cases_round_as_the_text_does() {
        for x in [
            1.005,
            2.675,
            0.5,
            1.5,
            2.5,
            -2.5,
            1e15,
            9.999_999_999_999_999e2,
            9.999_999_999_999_99e11,
            9.999_999_999_999_99e-9,
            0.000_123_456_789_012_345,
            0.1 + 0.2,
            99.995,
            1234.5678,
            5e-324,
            f64::MAX,
            f64::MIN_POSITIVE,
            -0.0,
        ] {
            for digits in -8..12 {
                for mode in MODES {
                    same(x, digits, mode);
                }
            }
        }
    }

    #[test]
    fn a_carry_into_a_new_digit_is_the_text_rounding() {
        // 9.99999999999999e11 has 15 nines: rounding at any place carries to 1e12.
        for digits in -6..8 {
            for mode in MODES {
                same(9.999_999_999_999_99e11, digits, mode);
                same(-9.999_999_999_999_99e11, digits, mode);
            }
        }
    }

    #[test]
    fn money_values_are_answered_by_the_exact_rounding() {
        let mut r = Rng(7);
        for _ in 0..100_000 {
            let x = (r.next() % 100_000_000) as f64 / 997.0;
            assert!(
                round_exact(x, 2, Mode::Nearest).is_some(),
                "{x:e} was declined"
            );
        }
    }
}
