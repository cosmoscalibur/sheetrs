//! Locales: the user's regional settings, resolved at build time.
//!
//! The locales LibreOffice offers are a table of the engine (`locale_data.rs`, generated from
//! LibreOffice by `scripts/gen_locale_table.py`): separators, names, the patterns of a date typed as
//! text, the currency format. A tag the table has (`es-CO`, or a bare language such as `es`)
//! needs nothing else. Any other region or language is stated in the `[locale]` table of the
//! build configuration rather than guessed: what Excel and LibreOffice take from the operating
//! system's regional settings cannot be derived from a tag with certainty, and a wrong decimal
//! separator changes every number read from text.

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
    /// Texts of `AM/PM` in time formats (default `AM`/`PM`).
    pub am: Option<String>,
    pub pm: Option<String>,
    pub currency: Option<String>,
    /// The symbol goes after the amount (with a space), instead of what the locale has.
    pub currency_after: Option<bool>,
    /// Negative amounts in parentheses, instead of what the locale has.
    pub currency_parens: Option<bool>,
    /// First year of the window for two-digit years (default 1930).
    pub two_digit_year_start: Option<u16>,
    /// `mdy`, `dmy` or `ymd`: the order of numeric text dates. A locale of another order reads
    /// dates as the reference locale of that order does (`en-US`, `es-ES`); `ymd` has none and
    /// needs `date_patterns`, and month names are not read as dates in it.
    pub date_order: Option<String>,
    /// The patterns of a date of digits that are accepted (`D/M/Y`, `D.M.`…).
    pub date_patterns: Option<Vec<String>>,
    /// Twelve month names, January first.
    pub months: Option<Vec<String>>,
    /// The abbreviations of the month names.
    pub months_abbr: Option<Vec<String>>,
    /// Seven day names, Sunday first.
    pub days: Option<Vec<String>>,
    /// The abbreviations of the day names.
    pub days_abbr: Option<Vec<String>>,
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

/// The locale whose date conventions a stated `date_order` takes over.
fn reference(order: DateOrder) -> Option<&'static str> {
    match order {
        DateOrder::Mdy => Some("en-US"),
        DateOrder::Dmy => Some("es-ES"),
        DateOrder::Ymd => None,
    }
}

/// Resolve `tag` and the overrides into a locale. The second element lists what
/// was inherited from the language of a region the table does not have.
pub fn from_tag(tag: &str, over: &LocaleOverride) -> Result<(Locale, Vec<String>), String> {
    let t = parse_tag(tag)?;
    let mut notes = Vec::new();
    let mut missing: Vec<&str> = Vec::new();
    let mut loc = if let Some(row) = Locale::row(tag) {
        Locale::from_row(row)
    } else if let Some(home) = Locale::row(&t.language) {
        // A region the table does not have: its conventions are stated, not guessed.
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
                "locale {tag}: the table has no region {} for `{}` (see `locale_data.rs`); set {} in [locale]",
                t.region.as_deref().unwrap_or("?"),
                t.language,
                missing.join(", ")
            ));
        }
        notes.push(format!(
            "locale {tag}: month and day names, boolean text and currency are those of `{}`; the table [locale] can change them",
            home.tag
        ));
        Locale::from_row(home)
    } else {
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
                "locale {tag}: the language `{}` is not in the table of locales; define this locale in [locale] (missing: {})",
                t.language,
                missing.join(", ")
            ));
        }
        Locale::en_us()
    };
    loc.tag = tag.replace('_', "-");
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
    if let Some(t) = &over.am {
        loc.am.clone_from(t);
    }
    if let Some(t) = &over.pm {
        loc.pm.clone_from(t);
    }
    if let Some(t) = &over.currency {
        loc.currency.clone_from(t);
    }
    if over.currency_after.is_some() || over.currency_parens.is_some() {
        let after = over
            .currency_after
            .unwrap_or_else(|| loc.currency_pos.starts_with('#'));
        let parens = over
            .currency_parens
            .unwrap_or_else(|| loc.currency_neg.starts_with('('));
        loc.currency_pos = if after { "# ¤" } else { "¤#" }.into();
        loc.currency_neg = if parens {
            format!("({})", loc.currency_pos)
        } else {
            format!("-{}", loc.currency_pos)
        };
    }
    if let Some(y) = over.two_digit_year_start {
        loc.two_digit_year_start = y;
    }
    if let Some(o) = &over.date_order {
        let o = order(o)?;
        if o != loc.date_order {
            // The patterns and the reading of a month name of that order.
            match reference(o).and_then(Locale::row).map(Locale::from_row) {
                Some(r) => {
                    loc.date_patterns = r.date_patterns;
                    loc.named = r.named;
                    loc.date_sep = r.date_sep;
                }
                None => {
                    loc.date_patterns = Vec::new();
                    loc.named = String::new();
                }
            }
            loc.date_order = o;
        }
    }
    if let Some(p) = &over.date_patterns {
        loc.date_patterns.clone_from(p);
    }
    if let Some(m) = &over.months {
        loc.months = fixed("months", m)?;
        loc.months_abbr = loc.months.clone();
    }
    if let Some(m) = &over.months_abbr {
        loc.months_abbr = fixed("months_abbr", m)?;
    }
    if let Some(d) = &over.days {
        loc.days = fixed("days", d)?;
        loc.days_abbr = loc.days.clone();
    }
    if let Some(d) = &over.days_abbr {
        loc.days_abbr = fixed("days_abbr", d)?;
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
    fn the_table_has_the_locales_of_libreoffice() {
        let l = get("en").unwrap();
        assert_eq!((l.decimal, l.group), ('.', ','));
        assert_eq!(l.date_order, DateOrder::Mdy);
        let l = get("es_ES").unwrap();
        assert_eq!((l.decimal, l.group, l.tag.as_str()), (',', '.', "es-ES"));
        assert_eq!(l.true_text, "VERDADERO");
        // A region of the table needs nothing stated.
        let l = get("es-CO").unwrap();
        assert_eq!((l.decimal, l.currency.as_str()), (',', "$"));
        assert_eq!(l.currency_neg, "(¤#)");
        let l = get("es-MX").unwrap();
        assert_eq!((l.decimal, l.group), ('.', ','));
        let l = get("en-GB").unwrap();
        assert_eq!(l.date_order, DateOrder::Dmy);
        assert_eq!(l.date_patterns, ["D/M", "D-M", "D/M/Y"]);
        // Only the locales checked against LibreOffice are in the table: the other ones are stated.
        for tag in ["de-DE", "fr-FR", "ja-JP", "es-AR", "en-AU"] {
            assert!(get(tag).is_err(), "{tag}");
        }
    }

    #[test]
    fn other_regions_are_stated_not_guessed() {
        let e = get("es-BO").unwrap_err();
        assert!(e.contains("decimal, group, date_order"), "{e}");
        let over = LocaleOverride {
            decimal: Some('.'),
            group: Some(','),
            date_order: Some("mdy".into()),
            ..LocaleOverride::default()
        };
        let (l, notes) = from_tag("es-BO", &over).unwrap();
        assert_eq!(
            (l.decimal, l.group, l.date_order),
            ('.', ',', DateOrder::Mdy)
        );
        // The order stated brings the reading of dates of that order.
        assert_eq!(l.date_patterns, ["M/D", "M/D/Y"]);
        assert_eq!(l.months[0], "enero");
        assert!(notes.iter().any(|n| n.contains("`es-ES`")), "{notes:?}");
    }

    #[test]
    fn other_languages_are_defined_in_full() {
        let e = get("sw-KE").unwrap_err();
        assert!(
            e.contains("not in the table of locales") && e.contains("months"),
            "{e}"
        );
        let names = |s: &str| s.split(' ').map(String::from).collect::<Vec<_>>();
        let over = LocaleOverride {
            decimal: Some('.'),
            group: Some(','),
            date_order: Some("dmy".into()),
            true_text: Some("KWELI".into()),
            false_text: Some("UONGO".into()),
            currency: Some("KSh".into()),
            months: Some(names(
                "Januari Februari Machi Aprili Mei Juni Julai Agosti Septemba Oktoba Novemba Desemba",
            )),
            days: Some(names(
                "Jumapili Jumatatu Jumanne Jumatano Alhamisi Ijumaa Jumamosi",
            )),
            ..LocaleOverride::default()
        };
        let (l, notes) = from_tag("sw-KE", &over).unwrap();
        assert_eq!(
            (l.months[2].as_str(), l.days[1].as_str()),
            ("Machi", "Jumatatu")
        );
        assert_eq!(l.true_text, "KWELI");
        assert_eq!(l.months_abbr[2], "Machi");
        assert!(notes.is_empty());
    }

    #[test]
    fn currency_flags_rebuild_the_formats() {
        let over = LocaleOverride {
            currency_after: Some(true),
            currency_parens: Some(true),
            ..LocaleOverride::default()
        };
        let (l, _) = from_tag("en-US", &over).unwrap();
        assert_eq!(
            (l.currency_pos.as_str(), l.currency_neg.as_str()),
            ("# ¤", "(# ¤)")
        );
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
        assert!(bad(|o| o.group = Some('.')).contains("both `.`"));
    }
}
