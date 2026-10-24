//! Text collation: the order of `<`, `>`, sorted lookups and the equality of
//! text.
//!
//! Spreadsheet text compares in three levels. Letters decide first (case and
//! accents ignored), then accents, then — only where a setting asks for case
//! sensitivity — case, lowercase first. That is the CLDR root collation, the
//! one LibreOffice uses; `data.rs` is derived from it for Latin text,
//! punctuation, symbols and currency (see `scripts/gen_collation.rs`), so a
//! comparison needs no tables beyond that file and no dependency.
//!
//! Characters outside the table order by their lowercase code point, after the
//! Latin letters: right for Greek and Cyrillic alphabets, an approximation
//! elsewhere. A language whose alphabet puts a letter of its own among the
//! Latin ones (`ñ` after `n` in Spanish) says so in its locale (`tailoring`).
//! Text is compared as stored: decomposed accents (`n` + U+0303) are not
//! recomposed.

mod data;

use data::{COMBINING, EXPANSIONS, IGNORABLE, TABLE};
use std::cmp::Ordering;

/// Primary weight of every ASCII byte; 0 for what the quick path does not
/// handle: ignorable controls and every byte of a multi-byte character.
const ASCII: [u16; 256] = ascii_primaries();

/// Table weights leave room for tailored letters between two classes.
const STEP: u32 = 4;
/// Where characters outside the table start: after every table class.
const OTHER: u32 = 0x1_0000 * STEP;

const fn ascii_primaries() -> [u16; 256] {
    let mut out = [0u16; 256];
    let mut i = 0;
    while i < TABLE.len() {
        let (c, p, _, _) = TABLE[i];
        if (c as u32) < 128 {
            out[c as usize] = p + 1;
        }
        i += 1;
    }
    out
}

#[derive(Clone, Copy)]
struct Element {
    primary: u32,
    secondary: u16,
    tertiary: u8,
}

#[derive(Debug, Clone)]
pub struct Collator {
    case_sensitive: bool,
    /// `(letter, after)`, both lowercase.
    tailoring: Vec<(char, char)>,
}

/// `(primary, secondary, tertiary)`; primaries start at 1 (0 is ignorable).
fn lookup(c: char) -> Option<(u32, u16, u8)> {
    TABLE
        .binary_search_by_key(&c, |e| e.0)
        .ok()
        .map(|i| (u32::from(TABLE[i].1) + 1, u16::from(TABLE[i].2), TABLE[i].3))
}

impl Collator {
    /// `tailoring`: letters that sort as their own letter right after another
    /// (`('ñ', 'n')`); either case may be given.
    pub fn new(tailoring: &[(char, char)], case_sensitive: bool) -> Self {
        let lower = |c: char| c.to_lowercase().next().unwrap_or(c);
        Self {
            case_sensitive,
            tailoring: tailoring
                .iter()
                .map(|&(l, a)| (lower(l), lower(a)))
                .collect(),
        }
    }

    fn push(&self, c: char, out: &mut Vec<Element>) {
        // The micro sign collates as the Greek letter mu.
        let c = if c == '\u{b5}' { '\u{3bc}' } else { c };
        let lower = c.to_lowercase().next().unwrap_or(c);
        let upper = u8::from(lower != c);
        if let Some(&(_, after)) = self.tailoring.iter().find(|(l, _)| *l == lower)
            && let Some((p, _, _)) = lookup(after)
        {
            out.push(Element {
                primary: p * STEP + 2,
                secondary: 0,
                tertiary: upper,
            });
        } else if let Some((p, s, t)) = lookup(c) {
            out.push(Element {
                primary: p * STEP,
                secondary: s,
                tertiary: t,
            });
        } else if let Some(&(_, e, sec, t)) = EXPANSIONS.iter().find(|(x, ..)| *x == c) {
            for (i, x) in e.chars().enumerate() {
                let (p, _, _) = lookup(x).unwrap_or((0, 0, 0));
                out.push(Element {
                    primary: p * STEP,
                    secondary: if i == 0 { sec } else { 0 },
                    tertiary: if i == 0 { t } else { 0 },
                });
            }
        } else if let Ok(i) = COMBINING.binary_search_by_key(&c, |e| e.0) {
            out.push(Element {
                primary: 0,
                secondary: u16::from(COMBINING[i].1) + 1,
                tertiary: 0,
            });
        } else if c < ' ' || ('\u{7f}'..='\u{9f}').contains(&c) || IGNORABLE.contains(&c) {
            // completely ignorable
        } else {
            out.push(Element {
                primary: OTHER + lower as u32,
                secondary: 0,
                tertiary: upper,
            });
        }
    }

    fn elements(&self, s: &str) -> Vec<Element> {
        let mut out = Vec::with_capacity(s.len());
        for c in s.chars() {
            self.push(c, &mut out);
        }
        out
    }

    /// Order of two texts at this collator's strength.
    pub fn compare(&self, a: &str, b: &str) -> Ordering {
        if let Some(o) = self.compare_quick(a.as_bytes(), b.as_bytes()) {
            return o;
        }
        let (x, y) = (self.elements(a), self.elements(b));
        let primaries = |v: &[Element]| {
            v.iter()
                .map(|e| e.primary)
                .filter(|&p| p != 0)
                .collect::<Vec<_>>()
        };
        primaries(&x)
            .cmp(&primaries(&y))
            .then_with(|| {
                let sec = |v: &[Element]| v.iter().map(|e| e.secondary).collect::<Vec<_>>();
                sec(&x).cmp(&sec(&y))
            })
            .then_with(|| {
                if self.case_sensitive {
                    let ter = |v: &[Element]| v.iter().map(|e| e.tertiary).collect::<Vec<_>>();
                    ter(&x).cmp(&ter(&y))
                } else {
                    Ordering::Equal
                }
            })
    }

    /// Plain ASCII text, without allocating. The first difference between the
    /// letters decides, so only the part up to it has to be plain; `None` when
    /// something else shows up before the decision.
    fn compare_quick(&self, a: &[u8], b: &[u8]) -> Option<Ordering> {
        let mut case = Ordering::Equal;
        for (&x, &y) in a.iter().zip(b) {
            let (px, py) = (ASCII[x as usize], ASCII[y as usize]);
            if px == 0 || py == 0 {
                return None;
            }
            if px != py {
                return Some(px.cmp(&py));
            }
            if case == Ordering::Equal {
                case = x.is_ascii_uppercase().cmp(&y.is_ascii_uppercase());
            }
        }
        // One is a prefix of the other: the rest decides, unless it is
        // ignorable.
        let n = a.len().min(b.len());
        if a[n..]
            .iter()
            .chain(&b[n..])
            .any(|&c| ASCII[c as usize] == 0)
        {
            return None;
        }
        Some(a.len().cmp(&b.len()).then(if self.case_sensitive {
            case
        } else {
            Ordering::Equal
        }))
    }

    pub fn equal(&self, a: &str, b: &str) -> bool {
        self.compare(a, b) == Ordering::Equal
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORDS: &[&str] = &[
        "a", "A", "á", "b", "B", "ñ", "Ñ", "n", "nz", "o", "z", "Z", "1", "10", "2", "_a", "-a",
        " a", "a b", "ab", "a-b", "a_b", "Straße", "strasse", "Strasse", "Æble", "aeble", "œuvre",
        "oeuvre", "Ørsted", "€5", "$5", "résumé", "resume", "Resume", "café", "cafe", "Café",
        "日本", "Ωmega", "жук", "Жук", "ÿ", "", "x\ty", "x\ny", "…", "...", "µg", "αg", "½", "1/2",
        "Ça", "ca", "cb", "Ünï", "uni",
    ];

    // Orders produced by ICU (CLDR root; `scripts/gen_collation.rs` explains how
    // the table is derived). A stable sort keeps equal texts in input order.
    const EN_CI: &[&str] = &[
        "", " a", "_a", "-a", "…", "...", "$5", "€5", "1", "1/2", "½", "10", "2", "a", "A", "á",
        "a b", "a_b", "a-b", "ab", "aeble", "Æble", "b", "B", "ca", "Ça", "cafe", "café", "Café",
        "cb", "n", "ñ", "Ñ", "nz", "o", "oeuvre", "œuvre", "Ørsted", "resume", "Resume", "résumé",
        "strasse", "Strasse", "Straße", "uni", "Ünï", "x\ty", "x\ny", "ÿ", "z", "Z", "αg", "µg",
        "Ωmega", "жук", "Жук", "日本",
    ];
    const EN_CS: &[&str] = &[
        "", " a", "_a", "-a", "...", "…", "$5", "€5", "1", "1/2", "½", "10", "2", "a", "A", "á",
        "a b", "a_b", "a-b", "ab", "aeble", "Æble", "b", "B", "ca", "Ça", "cafe", "café", "Café",
        "cb", "n", "ñ", "Ñ", "nz", "o", "oeuvre", "œuvre", "Ørsted", "resume", "Resume", "résumé",
        "strasse", "Strasse", "Straße", "uni", "Ünï", "x\ty", "x\ny", "ÿ", "z", "Z", "αg", "µg",
        "Ωmega", "жук", "Жук", "日本",
    ];
    const ES_CI: &[&str] = &[
        "", " a", "_a", "-a", "…", "...", "$5", "€5", "1", "1/2", "½", "10", "2", "a", "A", "á",
        "a b", "a_b", "a-b", "ab", "aeble", "Æble", "b", "B", "ca", "Ça", "cafe", "café", "Café",
        "cb", "n", "nz", "ñ", "Ñ", "o", "oeuvre", "œuvre", "Ørsted", "resume", "Resume", "résumé",
        "strasse", "Strasse", "Straße", "uni", "Ünï", "x\ty", "x\ny", "ÿ", "z", "Z", "αg", "µg",
        "Ωmega", "жук", "Жук", "日本",
    ];

    fn sorted(c: &Collator) -> Vec<&'static str> {
        let mut v = WORDS.to_vec();
        v.sort_by(|a, b| c.compare(a, b));
        v
    }

    #[test]
    fn matches_the_cldr_order() {
        assert_eq!(sorted(&Collator::new(&[], false)), EN_CI);
        assert_eq!(sorted(&Collator::new(&[], true)), EN_CS);
        assert_eq!(sorted(&Collator::new(&[('ñ', 'n')], false)), ES_CI);
    }

    #[test]
    fn levels_and_tailoring() {
        let ci = Collator::new(&[], false);
        let cs = Collator::new(&[], true);
        assert!(ci.equal("Resume", "resume") && !cs.equal("Resume", "resume"));
        assert!(!ci.equal("résumé", "resume"), "accents count");
        assert_eq!(cs.compare("a", "A"), Ordering::Less, "lowercase first");
        assert_eq!(ci.compare("ña", "nb"), Ordering::Less);
        let es = Collator::new(&[('ñ', 'n')], false);
        assert_eq!(es.compare("ña", "nz"), Ordering::Greater);
        assert_eq!(es.compare("ña", "oa"), Ordering::Less);
        assert!(es.equal("Ñ", "ñ"));
    }

    #[test]
    fn ignorable_controls_and_prefixes() {
        let ci = Collator::new(&[], false);
        assert!(ci.equal("a\u{1}b", "ab"));
        assert_eq!(ci.compare("ab", "ab\u{1}"), Ordering::Equal);
        assert_eq!(ci.compare("ab", "abc"), Ordering::Less);
        assert_eq!(ci.compare("", "a"), Ordering::Less);
    }
}
