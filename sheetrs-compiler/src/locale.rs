//! Locales: the user's regional settings, resolved at build time.
//!
//! `en` (United States conventions) and `es` (Spain) are built in; their
//! separators, names and date order are the ones verified against Excel and
//! LibreOffice. Any other region or language is stated in the `[locale]` table
//! of the build configuration rather than guessed: what Excel and LibreOffice
//! take from the operating system's regional settings cannot be derived from a
//! tag with certainty, and a wrong decimal separator changes every number read
//! from text.

use crate::options::{DateOrder, Locale};
use serde::Deserialize;

/// The `[locale]` table. Everything but `tag` overrides what the tag gives; for
/// a language that is not built in, it provides it.
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct LocaleOverride {
    /// Locale tag (`es-CO`); default `en`.
    pub tag: Option<String>,
    pub decimal: Option<char>,
    pub group: Option<char>,
    pub true_text: Option<String>,
    pub false_text: Option<String>,
    pub currency: Option<String>,
    /// First year of the window for two-digit years (default 1930).
    pub two_digit_year_start: Option<u16>,
    /// `mdy`, `dmy` or `ymd`: the order of numeric text dates.
    pub date_order: Option<String>,
    /// Twelve month names, January first.
    pub months: Option<Vec<String>>,
    /// Seven day names, Sunday first.
    pub days: Option<Vec<String>>,
    /// Letters with a place of their own in the alphabet, as `"ñ>n"`: sorts
    /// right after `n`. `es` has `ñ>n` built in.
    pub tailoring: Option<Vec<String>>,
}

struct Tag {
    language: String,
    region: Option<String>,
}

fn parse_tag(tag: &str) -> Result<Tag, String> {
    let bad = || format!("`{tag}` is not a locale tag (like `en`, `es-CO`)");
    let mut parts = tag.split(['-', '_']);
    let language = parts
        .next()
        .filter(|l| (2..=3).contains(&l.len()) && l.chars().all(|c| c.is_ascii_alphabetic()));
    let language = language.ok_or_else(bad)?.to_ascii_lowercase();
    // Skip a script subtag (`Latn`); the region is two letters or three digits.
    let region = parts
        .find(|p| {
            (p.len() == 2 && p.chars().all(|c| c.is_ascii_alphabetic()))
                || (p.len() == 3 && p.chars().all(|c| c.is_ascii_digit()))
        })
        .map(str::to_ascii_uppercase);
    Ok(Tag { language, region })
}

fn order(s: &str) -> Result<DateOrder, String> {
    match s.to_ascii_lowercase().as_str() {
        "mdy" => Ok(DateOrder::Mdy),
        "dmy" => Ok(DateOrder::Dmy),
        "ymd" => Ok(DateOrder::Ymd),
        _ => Err(format!("date_order must be mdy, dmy or ymd, got `{s}`")),
    }
}

fn fixed<const N: usize>(what: &str, v: &[String]) -> Result<[String; N], String> {
    <[String; N]>::try_from(v.to_vec())
        .map_err(|_| format!("{what} needs {N} names, got {}", v.len()))
}

fn tailoring(spec: &[String]) -> Result<Vec<(char, char)>, String> {
    spec.iter()
        .map(|s| {
            let mut it = s.split('>').map(str::trim);
            match (it.next(), it.next(), it.next()) {
                (Some(a), Some(b), None) if a.chars().count() == 1 && b.chars().count() == 1 => {
                    Ok((
                        a.chars().next().unwrap_or(' '),
                        b.chars().next().unwrap_or(' '),
                    ))
                }
                _ => Err(format!("tailoring `{s}` must look like `ñ>n`")),
            }
        })
        .collect()
}

/// Resolve `tag` and the overrides into a locale. The second element lists what
/// was inherited from the built-in language of a region that differs from it.
pub fn from_tag(tag: &str, over: &LocaleOverride) -> Result<(Locale, Vec<String>), String> {
    let t = parse_tag(tag)?;
    let mut notes = Vec::new();
    let mut missing: Vec<&str> = Vec::new();
    let builtin = match t.language.as_str() {
        "en" => Some(("US", Locale::en_us())),
        "es" => Some(("ES", Locale::es_es())),
        _ => None,
    };
    let mut loc = match &builtin {
        Some((home, base)) => {
            if let Some(r) = t.region.as_deref().filter(|r| r != home) {
                // The conventions of another region are stated, not guessed.
                for (key, set) in [
                    ("decimal", over.decimal.is_some()),
                    ("group", over.group.is_some()),
                    ("date_order", over.date_order.is_some()),
                ] {
                    if !set {
                        missing.push(key);
                    }
                }
                if !missing.is_empty() {
                    return Err(format!(
                        "locale {tag}: the built-in {} locale follows {home} conventions; for region {r} set {} in [locale]",
                        t.language,
                        missing.join(", ")
                    ));
                }
                notes.push(format!(
                    "locale {tag}: month and day names, boolean text and currency are those of the built-in `{}`; the table [locale] can change them",
                    t.language
                ));
            }
            base.clone()
        }
        None => {
            for (key, set) in [
                ("decimal", over.decimal.is_some()),
                ("group", over.group.is_some()),
                ("date_order", over.date_order.is_some()),
                ("months", over.months.is_some()),
                ("days", over.days.is_some()),
                ("true_text", over.true_text.is_some()),
                ("false_text", over.false_text.is_some()),
                ("currency", over.currency.is_some()),
            ] {
                if !set {
                    missing.push(key);
                }
            }
            if !missing.is_empty() {
                return Err(format!(
                    "locale {tag}: only `en` and `es` are built in; define this locale in [locale] (missing: {})",
                    missing.join(", ")
                ));
            }
            Locale::en_us()
        }
    };
    loc.tag = tag.replace('_', "-");
    if builtin.is_none() {
        loc.tailoring.clear();
    }
    if let Some(c) = over.decimal {
        loc.decimal = c;
    }
    if let Some(c) = over.group {
        loc.group = c;
    }
    if let Some(t) = &over.true_text {
        loc.true_text.clone_from(t);
    }
    if let Some(t) = &over.false_text {
        loc.false_text.clone_from(t);
    }
    if let Some(t) = &over.currency {
        loc.currency.clone_from(t);
    }
    if let Some(y) = over.two_digit_year_start {
        loc.two_digit_year_start = y;
    }
    if let Some(o) = &over.date_order {
        loc.date_order = order(o)?;
    }
    if let Some(m) = &over.months {
        loc.months = fixed("months", m)?;
    }
    if let Some(d) = &over.days {
        loc.days = fixed("days", d)?;
    }
    if let Some(t) = &over.tailoring {
        loc.tailoring = tailoring(t)?;
    }
    if loc.decimal == loc.group {
        return Err(format!(
            "locale {tag}: decimal and group separators are both `{}`",
            loc.decimal
        ));
    }
    Ok((loc, notes))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn get(tag: &str) -> Result<Locale, String> {
        from_tag(tag, &LocaleOverride::default()).map(|l| l.0)
    }

    #[test]
    fn en_and_es_are_built_in() {
        let l = get("en").unwrap();
        assert_eq!((l.decimal, l.group), ('.', ','));
        assert_eq!(l.date_order, DateOrder::Mdy);
        let l = get("es_ES").unwrap();
        assert_eq!((l.decimal, l.group, l.tag.as_str()), (',', '.', "es-ES"));
        assert_eq!(l.tailoring, [('ñ', 'n')]);
        assert_eq!(l.true_text, "VERDADERO");
        assert!(get("en-US").is_ok());
    }

    #[test]
    fn other_regions_are_stated_not_guessed() {
        let e = get("es-MX").unwrap_err();
        assert!(e.contains("decimal, group, date_order"), "{e}");
        let over = LocaleOverride {
            decimal: Some('.'),
            group: Some(','),
            date_order: Some("dmy".into()),
            ..LocaleOverride::default()
        };
        let (l, notes) = from_tag("es-MX", &over).unwrap();
        assert_eq!(
            (l.decimal, l.group, l.date_order),
            ('.', ',', DateOrder::Dmy)
        );
        assert_eq!(l.months[0], "enero");
        assert_eq!(l.tailoring, [('ñ', 'n')]);
        assert!(
            notes.iter().any(|n| n.contains("built-in `es`")),
            "{notes:?}"
        );
    }

    #[test]
    fn other_languages_are_defined_in_full() {
        let e = get("de-DE").unwrap_err();
        assert!(
            e.contains("only `en` and `es` are built in") && e.contains("months"),
            "{e}"
        );
        let names = |s: &str| s.split(' ').map(String::from).collect::<Vec<_>>();
        let over = LocaleOverride {
            decimal: Some(','),
            group: Some('.'),
            date_order: Some("dmy".into()),
            true_text: Some("WAHR".into()),
            false_text: Some("FALSCH".into()),
            currency: Some("€".into()),
            months: Some(names(
                "Januar Februar März April Mai Juni Juli August September Oktober November Dezember",
            )),
            days: Some(names(
                "Sonntag Montag Dienstag Mittwoch Donnerstag Freitag Samstag",
            )),
            tailoring: Some(vec!["ß>s".into()]),
            ..LocaleOverride::default()
        };
        let (l, notes) = from_tag("de-DE", &over).unwrap();
        assert_eq!(
            (l.months[2].as_str(), l.days[1].as_str()),
            ("März", "Montag")
        );
        assert_eq!(l.true_text, "WAHR");
        assert_eq!(l.tailoring, [('ß', 's')]);
        assert!(notes.is_empty());
    }

    #[test]
    fn malformed_input_fails() {
        assert!(get("not a tag").is_err());
        assert!(get("").is_err());
        let bad = |f: fn(&mut LocaleOverride)| {
            let mut o = LocaleOverride::default();
            f(&mut o);
            from_tag("en", &o).unwrap_err()
        };
        assert!(bad(|o| o.date_order = Some("xyz".into())).contains("mdy, dmy or ymd"));
        assert!(bad(|o| o.months = Some(vec!["a".into()])).contains("12 names"));
        assert!(bad(|o| o.tailoring = Some(vec!["ñ".into()])).contains("ñ>n"));
        assert!(bad(|o| o.group = Some('.')).contains("both `.`"));
    }
}
