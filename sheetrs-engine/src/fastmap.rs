//! A fast non-cryptographic hasher for the integer-keyed maps that dominate
//! large workbooks (millions of cell addresses). The default SipHash is
//! DoS-resistant, which a compiler reading its own input does not need.

use std::collections::HashMap;
use std::hash::{BuildHasherDefault, Hasher};

#[derive(Default, Clone, Copy)]
pub struct FxHasher(u64);

const SEED: u64 = 0x517c_c1b7_2722_0a95;

impl FxHasher {
    #[inline]
    fn add(&mut self, x: u64) {
        self.0 = (self.0.rotate_left(5) ^ x).wrapping_mul(SEED);
    }
}

impl Hasher for FxHasher {
    #[inline]
    fn write(&mut self, bytes: &[u8]) {
        for chunk in bytes.chunks(8) {
            let mut b = [0u8; 8];
            b[..chunk.len()].copy_from_slice(chunk);
            self.add(u64::from_le_bytes(b));
        }
    }

    #[inline]
    fn write_u8(&mut self, i: u8) {
        self.add(u64::from(i));
    }

    #[inline]
    fn write_u16(&mut self, i: u16) {
        self.add(u64::from(i));
    }

    #[inline]
    fn write_u32(&mut self, i: u32) {
        self.add(u64::from(i));
    }

    #[inline]
    fn write_u64(&mut self, i: u64) {
        self.add(i);
    }

    #[inline]
    fn write_usize(&mut self, i: usize) {
        self.add(i as u64);
    }

    #[inline]
    fn finish(&self) -> u64 {
        // The multiply leaves the low bits of a key that has none (the bits of a whole
        // number as `f64` end in 32 zeros) at zero, and the table indexes by the low
        // bits: rotate the well-mixed high bits down.
        self.0.rotate_left(26)
    }
}

pub type FastMap<K, V> = HashMap<K, V, BuildHasherDefault<FxHasher>>;
pub type FastSet<K> = std::collections::HashSet<K, BuildHasherDefault<FxHasher>>;

#[cfg(test)]
mod tests {
    use super::*;
    use std::hash::BuildHasher;

    /// Whole numbers as `f64` bits have no low bits set; the table indexes by the low bits.
    #[test]
    fn whole_numbers_as_float_bits_spread_over_the_low_bits() {
        let build = BuildHasherDefault::<FxHasher>::default();
        let slots: std::collections::HashSet<u64> = (1..=65_536u32)
            .map(|n| build.hash_one(f64::from(n).to_bits()) & 0xFFFF)
            .collect();
        // A perfect hash fills 63% of the slots of a random assignment; before the
        // rotation in `finish` every key landed in one slot.
        assert!(slots.len() > 35_000, "{} slots", slots.len());
    }
}
