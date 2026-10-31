//! Text collation: the order of `<`, `>` and sorted lookups.
//!
//! LibreOffice orders text with ICU's collator for the language of the document: three levels,
//! letters first (case and accents ignored), then accents, then, where a setting asks for case
//! sensitivity, case. This is the same collator (ICU4X, with the data compiled in), so the order is
//! LibreOffice's for every script, emoji included, and for the alphabets the language tailors (`ñ`
//! after `n` in Spanish). The equality of text is not by collation; see [`crate::casemap`].

use icu_collator::options::{CollatorOptions, Strength};
use icu_collator::{CollatorBorrowed, CollatorPreferences};
use icu_locale_core::Locale;
use std::cmp::Ordering;

pub struct Collator(CollatorBorrowed<'static>);

impl std::fmt::Debug for Collator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("Collator")
    }
}

impl Collator {
    /// The collator of the language of `tag`; `case_sensitive` lets case break ties, lowercase
    /// first.
    pub fn new(tag: &str, case_sensitive: bool) -> Self {
        let locale: Locale = tag.replace('_', "-").parse().unwrap_or(Locale::UNKNOWN);
        let mut options = CollatorOptions::default();
        options.strength = Some(if case_sensitive {
            Strength::Tertiary
        } else {
            Strength::Secondary
        });
        let prefs = CollatorPreferences::from(&locale);
        Self(
            CollatorBorrowed::try_new(prefs, options)
                .or_else(|_| CollatorBorrowed::try_new(CollatorPreferences::default(), options))
                .expect("compiled collation data"),
        )
    }

    /// Order of two texts at this collator's strength.
    pub fn compare(&self, a: &str, b: &str) -> Ordering {
        self.0.compare(a, b)
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

    // Orders produced by ICU. A stable sort keeps equal texts in input order.
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
        assert_eq!(sorted(&Collator::new("en", false)), EN_CI);
        assert_eq!(sorted(&Collator::new("en", true)), EN_CS);
        assert_eq!(sorted(&Collator::new("es", false)), ES_CI);
    }

    #[test]
    fn levels_and_tailoring() {
        let ci = Collator::new("en", false);
        let cs = Collator::new("en", true);
        let same = |c: &Collator, a: &str, b: &str| c.compare(a, b) == Ordering::Equal;
        assert!(same(&ci, "Resume", "resume") && !same(&cs, "Resume", "resume"));
        assert!(!same(&ci, "résumé", "resume"), "accents count");
        assert_eq!(cs.compare("a", "A"), Ordering::Less, "lowercase first");
        assert_eq!(ci.compare("ña", "nb"), Ordering::Less);
        let es = Collator::new("es", false);
        assert_eq!(es.compare("ña", "nz"), Ordering::Greater);
        assert_eq!(es.compare("ña", "oa"), Ordering::Less);
        assert!(same(&es, "Ñ", "ñ"));
    }

    #[test]
    fn ignorable_controls_and_prefixes() {
        let ci = Collator::new("en", false);
        assert_eq!(ci.compare("a\u{1}b", "ab"), Ordering::Equal);
        assert_eq!(ci.compare("ab", "ab\u{1}"), Ordering::Equal);
        assert_eq!(ci.compare("ab", "abc"), Ordering::Less);
        assert_eq!(ci.compare("", "a"), Ordering::Less);
    }
}
