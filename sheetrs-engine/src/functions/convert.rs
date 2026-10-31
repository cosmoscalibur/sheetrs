//! `CONVERT`: the units of LibreOffice's analysis add-in (`scaddins/source/analysis/
//! analysishelper.cxx`, `ConvertDataList`), with its table in its order, since the lookup of a
//! unit with a prefix takes the last unit that matches and an exact name wins over all of them.
//! The result is checked against the oracle (`tests/assets/semantics.*`).

use super::{Reg, num};
use crate::value::ErrorKind;

const DOMAIN: ErrorKind = ErrorKind::Domain;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Class {
    Mass,
    Length,
    Time,
    Pressure,
    Force,
    Energy,
    Power,
    Magnetism,
    Temperature,
    Volume,
    Area,
    Speed,
    Information,
}
use Class::*;

/// Name, constant (units per base unit), class, whether it takes prefixes, offset (temperatures),
/// whether it is a linear (temperature) unit.
type Unit = (&'static str, f64, Class, bool, f64, bool);

const UNITS: [Unit; 146] = [
    ("g", 1.0, Mass, true, 0.0, false),
    ("sg", 6.852_205_000_534_78e-5, Mass, false, 0.0, false),
    ("lbm", 2.204_622_914_691_34e-3, Mass, false, 0.0, false),
    ("u", 6.022_137e23, Mass, true, 0.0, false),
    ("ozm", 3.527_397_180_036_27e-2, Mass, false, 0.0, false),
    ("stone", 1.574730e-04, Mass, false, 0.0, false),
    ("ton", 1.102311e-06, Mass, false, 0.0, false),
    ("grain", 1.543236e01, Mass, false, 0.0, false),
    ("pweight", 7.054792e-01, Mass, false, 0.0, false),
    ("hweight", 1.968413e-05, Mass, false, 0.0, false),
    ("shweight", 2.204623e-05, Mass, false, 0.0, false),
    ("brton", 9.842065e-07, Mass, false, 0.0, false),
    ("cwt", 2.2046226218487758e-05, Mass, false, 0.0, false),
    ("shweight", 2.2046226218487758e-05, Mass, false, 0.0, false),
    ("uk_cwt", 1.9684130552221213e-05, Mass, false, 0.0, false),
    ("lcwt", 1.9684130552221213e-05, Mass, false, 0.0, false),
    ("hweight", 1.9684130552221213e-05, Mass, false, 0.0, false),
    ("uk_ton", 9.842_065_276_110_607e-7, Mass, false, 0.0, false),
    ("LTON", 9.842_065_276_110_607e-7, Mass, false, 0.0, false),
    ("m", 1.0, Length, true, 0.0, false),
    ("mi", 6.213_711_922_373_339e-4, Length, false, 0.0, false),
    ("Nmi", 5.399_568_034_557_236e-4, Length, false, 0.0, false),
    ("in", 3.937_007_874_015_748e1, Length, false, 0.0, false),
    ("ft", 3.2808398950131234e00, Length, false, 0.0, false),
    ("yd", 1.0936132983377078e00, Length, false, 0.0, false),
    ("ang", 1e10, Length, true, 0.0, false),
    ("Pica", 2.8346456692913386e03, Length, false, 0.0, false),
    ("picapt", 2.8346456692913386e03, Length, false, 0.0, false),
    ("pica", 2.36220472441e02, Length, false, 0.0, false),
    ("ell", 8.748906e-01, Length, false, 0.0, false),
    ("parsec", 3.240779e-17, Length, true, 0.0, false),
    ("pc", 3.240779e-17, Length, true, 0.0, false),
    (
        "lightyear",
        1.057_023_455_773_293e-16,
        Length,
        true,
        0.0,
        false,
    ),
    ("ly", 1.057_023_455_773_293e-16, Length, true, 0.0, false),
    (
        "survey_mi",
        6.213_699_494_949_494e-4,
        Length,
        false,
        0.0,
        false,
    ),
    ("yr", 3.168_808_781_402_895e-8, Time, false, 0.0, false),
    ("day", 1.1574074074074074e-05, Time, false, 0.0, false),
    ("d", 1.1574074074074074e-05, Time, false, 0.0, false),
    ("hr", 2.777_777_777_777_778e-4, Time, false, 0.0, false),
    ("mn", 1.6666666666666667e-02, Time, false, 0.0, false),
    ("min", 1.6666666666666667e-02, Time, false, 0.0, false),
    ("sec", 1.0, Time, true, 0.0, false),
    ("s", 1.0, Time, true, 0.0, false),
    ("Pa", 1.0, Pressure, true, 0.0, false),
    ("atm", 9.869_232_999_981_93e-6, Pressure, true, 0.0, false),
    ("at", 9.869_232_999_981_93e-6, Pressure, true, 0.0, false),
    ("mmHg", 7.500_617_079_986_27e-3, Pressure, true, 0.0, false),
    ("Torr", 7.500_638e-3, Pressure, false, 0.0, false),
    ("psi", 1.450_377e-4, Pressure, false, 0.0, false),
    ("N", 1.0, Force, true, 0.0, false),
    ("dyn", 1e5, Force, true, 0.0, false),
    ("dy", 1e5, Force, true, 0.0, false),
    ("lbf", 2.24808923655339e-01, Force, false, 0.0, false),
    ("pond", 1.019716e02, Force, true, 0.0, false),
    ("J", 1.0, Energy, true, 0.0, false),
    ("e", 1e7, Energy, true, 0.0, false),
    ("c", 2.390_062_494_734_67e-1, Energy, true, 0.0, false),
    ("cal", 2.388_461_906_420_17e-1, Energy, true, 0.0, false),
    ("eV", 6.241_457e18, Energy, true, 0.0, false),
    ("ev", 6.241_457e18, Energy, true, 0.0, false),
    ("HPh", 3.725_061_111_111_111e-7, Energy, false, 0.0, false),
    ("hh", 3.725_061_111_111_111e-7, Energy, false, 0.0, false),
    ("Wh", 2.777_777_777_777_778e-4, Energy, true, 0.0, false),
    ("wh", 2.777_777_777_777_778e-4, Energy, true, 0.0, false),
    ("flb", 2.37304222192651e01, Energy, false, 0.0, false),
    ("BTU", 9.478_150_673_490_15e-4, Energy, false, 0.0, false),
    ("btu", 9.478_150_673_490_15e-4, Energy, false, 0.0, false),
    ("W", 1.0, Power, true, 0.0, false),
    ("w", 1.0, Power, true, 0.0, false),
    ("HP", 1.341022e-03, Power, false, 0.0, false),
    ("h", 1.341022e-03, Power, false, 0.0, false),
    ("PS", 1.359622e-03, Power, false, 0.0, false),
    ("T", 1.0, Magnetism, true, 0.0, false),
    ("ga", 1e4, Magnetism, true, 0.0, false),
    ("C", 1.0, Temperature, false, -2.731_5e2, true),
    ("cel", 1.0, Temperature, false, -2.731_5e2, true),
    ("F", 1.8, Temperature, false, -2.553_722_222_222_222e2, true),
    (
        "fah",
        1.8,
        Temperature,
        false,
        -2.553_722_222_222_222e2,
        true,
    ),
    ("K", 1.0, Temperature, true, 0.0000000000000000e00, true),
    ("kel", 1.0, Temperature, true, 0.0000000000000000e00, true),
    ("Reau", 8e-1, Temperature, false, -2.731_5e2, true),
    ("Rank", 1.8, Temperature, false, 0.0000000000000000e00, true),
    ("tsp", 2.0288413621105798e02, Volume, false, 0.0, false),
    ("tbs", 6.762_804_540_368_599e1, Volume, false, 0.0, false),
    ("oz", 3.3814022701842997e01, Volume, false, 0.0, false),
    ("cup", 4.226_752_837_730_374, Volume, false, 0.0, false),
    ("pt", 2.113_376_418_865_187, Volume, false, 0.0, false),
    ("us_pt", 2.113_376_418_865_187, Volume, false, 0.0, false),
    ("uk_pt", 1.7597539863927023e00, Volume, false, 0.0, false),
    ("qt", 1.0566882094325937e00, Volume, false, 0.0, false),
    ("gal", 2.6417205235814842e-01, Volume, false, 0.0, false),
    ("l", 1.0, Volume, true, 0.0, false),
    ("L", 1.0, Volume, true, 0.0, false),
    ("lt", 1.0, Volume, true, 0.0, false),
    ("m3", 1e-3, Volume, true, 0.0, false),
    ("mi3", 2.3991275857892772e-13, Volume, false, 0.0, false),
    ("Nmi3", 1.574_262_146_858_115e-13, Volume, false, 0.0, false),
    ("in3", 6.102_374_409_473_229e1, Volume, false, 0.0, false),
    ("ft3", 3.531_466_672_148_859e-2, Volume, false, 0.0, false),
    ("yd3", 1.3079506193143922e-03, Volume, false, 0.0, false),
    ("ang3", 1e27, Volume, true, 0.0, false),
    ("Pica3", 2.2776990435870636e07, Volume, false, 0.0, false),
    ("picapt3", 2.2776990435870636e07, Volume, false, 0.0, false),
    ("pica3", 1.31811287245e04, Volume, false, 0.0, false),
    (
        "barrel",
        6.289_810_770_432_105e-3,
        Volume,
        false,
        0.0,
        false,
    ),
    ("bushel", 2.837759e-02, Volume, false, 0.0, false),
    ("regton", 3.531467e-04, Volume, false, 0.0, false),
    ("GRT", 3.531467e-04, Volume, false, 0.0, false),
    ("Schooner", 2.3529411764705882e00, Volume, false, 0.0, false),
    ("Middy", 3.508_771_929_824_561, Volume, false, 0.0, false),
    ("Glass", 5.0, Volume, false, 0.0, false),
    ("Sixpack", 0.5, Volume, false, 0.0, false),
    ("Humpen", 2.0, Volume, false, 0.0, false),
    ("ly3", 1.181_010_812_562_38e-51, Volume, false, 0.0, false),
    ("MTON", 1.4125866688595436e00, Volume, false, 0.0, false),
    ("tspm", 2e2, Volume, false, 0.0, false),
    (
        "uk_gal",
        2.199_692_482_990_878e-1,
        Volume,
        false,
        0.0,
        false,
    ),
    ("uk_qt", 8.798_769_931_963_512e-1, Volume, false, 0.0, false),
    ("m2", 1.0, Area, true, 0.0, false),
    ("mi2", 3.8610215854244585e-07, Area, false, 0.0, false),
    ("Nmi2", 2.9155334959812286e-07, Area, false, 0.0, false),
    ("in2", 1.550_003_100_006_2e3, Area, false, 0.0, false),
    ("ft2", 1.0763910416709722e01, Area, false, 0.0, false),
    ("yd2", 1.1959900463010803e00, Area, false, 0.0, false),
    ("ang2", 1e20, Area, true, 0.0, false),
    ("Pica2", 8.0352160704321409e06, Area, false, 0.0, false),
    ("picapt2", 8.0352160704321409e06, Area, false, 0.0, false),
    ("pica2", 5.58001116002232e04, Area, false, 0.0, false),
    ("Morgen", 4e-4, Area, false, 0.0, false),
    ("ar", 1.000000e-02, Area, true, 0.0, false),
    ("acre", 2.471053815e-04, Area, false, 0.0, false),
    ("uk_acre", 2.471_053_814_671_653e-4, Area, false, 0.0, false),
    ("us_acre", 2.471_043_930_466_279e-4, Area, false, 0.0, false),
    ("ly2", 1.1172985860549147e-32, Area, false, 0.0, false),
    ("ha", 1.000000e-04, Area, false, 0.0, false),
    ("m/s", 1.0, Speed, true, 0.0, false),
    ("m/sec", 1.0, Speed, true, 0.0, false),
    ("m/h", 3.6e3, Speed, true, 0.0, false),
    ("m/hr", 3.6e3, Speed, true, 0.0, false),
    ("mph", 2.2369362920544023e00, Speed, false, 0.0, false),
    ("kn", 1.9438444924406048e00, Speed, false, 0.0, false),
    ("admkn", 1.9438446603753486e00, Speed, false, 0.0, false),
    (
        "ludicrous speed",
        2.0494886343432328e-14,
        Speed,
        false,
        0.0,
        false,
    ),
    (
        "ridiculous speed",
        4.015_695_847_142_429e-6,
        Speed,
        false,
        0.0,
        false,
    ),
    ("bit", 1.00e00, Information, true, 0.0, false),
    ("byte", 1.25e-01, Information, true, 0.0, false),
];

/// `10^n` as the add-in's table of powers has it (the exact decimal literal).
fn pow10_exp(f: f64, n: i32) -> f64 {
    f * format!("1e{n}").parse::<f64>().unwrap_or(f64::NAN)
}

/// The prefix level of `name` for `unit`: 0 for the unit itself, the power of ten (or of two, for
/// the binary prefixes of the information units) with a prefix, `None` if it does not match.
fn matching_level(unit: &Unit, name: &str) -> Option<i32> {
    let chars: Vec<char> = name.chars().collect();
    let mut s: String = name.to_string();
    // `m^2` is `m2`.
    if let Some(at) = chars.iter().rposition(|&c| c == '^')
        && at > 0
        && at == chars.len() - 2
    {
        s = chars[..at].iter().chain(&chars[at + 1..]).collect();
    }
    if unit.0 == s {
        return Some(0);
    }
    if unit.3
        && let Some(prefix) = s.strip_suffix(unit.0)
    {
        {
            let p: Vec<char> = prefix.chars().collect();
            if p.len() == 1 || prefix == "da" {
                let mut n = match p[0] {
                    'y' => -24,
                    'z' => -21,
                    'a' => -18,
                    'f' => -15,
                    'p' => -12,
                    'n' => -9,
                    'u' => -6,
                    'm' => -3,
                    'c' => -2,
                    'd' if p.len() == 1 => -1,
                    'd' | 'e' => 1,
                    'h' => 2,
                    'k' => 3,
                    'M' => 6,
                    'G' => 9,
                    'T' => 12,
                    'P' => 15,
                    'E' => 18,
                    'Z' => 21,
                    'Y' => 24,
                    _ => return None,
                };
                // `cm3` is 10^-6 m^3, not 10^-2.
                if s.ends_with('2') {
                    n *= 2;
                } else if s.ends_with('3') {
                    n *= 3;
                }
                return Some(n);
            } else if p.len() == 2 && p[1] == 'i' && unit.2 == Information {
                return match p[0] {
                    'k' => Some(10),
                    'M' => Some(20),
                    'G' => Some(30),
                    'T' => Some(40),
                    'P' => Some(50),
                    'E' => Some(60),
                    'Z' => Some(70),
                    'Y' => Some(80),
                    _ => None,
                };
            }
        }
    }
    None
}

fn convert(f: f64, from: &str, to: &str) -> Option<f64> {
    let (mut pf, mut pt) = (None, None);
    let (mut lf, mut lt) = (0, 0);
    let (mut search_from, mut search_to) = (true, true);
    for unit in &UNITS {
        if search_from && let Some(n) = matching_level(unit, from) {
            pf = Some(unit);
            lf = n;
            search_from = n != 0;
        }
        if search_to && let Some(n) = matching_level(unit, to) {
            pt = Some(unit);
            lt = n;
            search_to = n != 0;
        }
        if !search_from && !search_to {
            break;
        }
    }
    let (pf, pt) = (pf?, pt?);
    if pf.2 != pt.2 {
        return None;
    }
    if pf.5 {
        // Linear units: through the base unit.
        let mut base = f;
        if lf != 0 {
            base = pow10_exp(base, lf);
        }
        base /= pf.1;
        base -= pf.4;
        let mut r = base + pt.4;
        r *= pt.1;
        if lt != 0 {
            r = pow10_exp(r, -lt);
        }
        return Some(r);
    }
    let mut f = f;
    f *= pt.1 / pf.1;
    if pf.2 == Information {
        let bin_from = lf > 0 && lf % 10 == 0;
        let bin_to = lt > 0 && lt % 10 == 0;
        if bin_from || bin_to {
            if bin_from && bin_to {
                let d = lf - lt;
                if d != 0 {
                    f *= 2f64.powf(f64::from(d));
                }
            } else if bin_from {
                f *= 2f64.powf(f64::from(lf)) / 10f64.powf(f64::from(lt));
            } else {
                f *= 10f64.powf(f64::from(lf)) / 2f64.powf(f64::from(lt));
            }
            return Some(f);
        }
    }
    let d = lf - lt;
    if d != 0 {
        f = pow10_exp(f, d);
    }
    Some(f)
}

pub(super) fn register(r: &mut Reg) {
    r.eager("CONVERT", 3, 3, |e, a| {
        let f = e.num(&a[0])?;
        let from = e.text(&a[1])?;
        let to = e.text(&a[2])?;
        match convert(f, &from, &to) {
            Some(x) if x.is_finite() => num(x),
            _ => Err(DOMAIN),
        }
    });
}
