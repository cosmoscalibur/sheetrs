//! Locales for any BCP-47 tag, derived from CLDR through ICU4X.
//!
//! The engine only holds the resolved [`Locale`] (separators, month and day
//! names, date order, boolean text), so ICU stays out of the runtime for these;
//! text ordering uses the same tag at run time. `en` and `es` are verified
//! against Excel and LibreOffice; for every other language CLDR supplies the
//! separators, names and date order and the rest keeps the English defaults
//! unless the configuration says otherwise.

use crate::options::{DateOrder, Locale};
use icu_calendar::Date;
use icu_datetime::{DateTimeFormatter, fieldsets};
use icu_decimal::DecimalFormatter;
use icu_decimal::input::Decimal;
use icu_locale::Locale as IcuLocale;
use serde::Deserialize;

/// What the configuration may set on top of the CLDR data: the regional
/// settings an application lets the user change.
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct LocaleOverride {
    /// BCP-47 tag; default `en`.
    pub tag: Option<String>,
    pub decimal: Option<char>,
    pub group: Option<char>,
    pub true_text: Option<String>,
    pub false_text: Option<String>,
    pub currency: Option<String>,
    /// First year of the window for two-digit years (default 1930).
    pub two_digit_year_start: Option<u16>,
    /// `mdy`, `dmy` or `ymd`.
    pub date_order: Option<String>,
}

/// Resolve `tag` into a locale. The second element lists what could not be
/// taken from CLDR and what is unverified.
pub fn from_tag(tag: &str, over: &LocaleOverride) -> Result<(Locale, Vec<String>), String> {
    let norm = tag.replace('_', "-");
    let id: IcuLocale = norm
        .parse()
        .map_err(|e| format!("`{tag}` is not a valid locale tag ({e})"))?;
    let lang = id.id.language.as_str().to_string();
    let verified = matches!(lang.as_str(), "en" | "es");
    let mut notes = Vec::new();
    let mut loc = Locale::by_tag(&lang).unwrap_or_else(Locale::en_us);
    loc.tag = id.to_string();

    match DecimalFormatter::try_new((&id).into(), Default::default()) {
        Ok(f) => {
            let text = f
                .format(&"1234567.5".parse::<Decimal>().map_err(|e| e.to_string())?)
                .to_string();
            let marks: Vec<char> = text.chars().filter(|c| !c.is_numeric()).collect();
            if let (Some(g), Some(d)) = (marks.first(), marks.last()) {
                loc.group = *g;
                loc.decimal = *d;
            }
            if text.chars().any(|c| c.is_numeric() && !c.is_ascii_digit()) {
                notes.push(format!(
                    "locale {}: native digits are not parsed, only 0-9",
                    loc.tag
                ));
            }
        }
        Err(e) => notes.push(format!("locale {}: no separators ({e})", loc.tag)),
    }

    if let Ok(f) = DateTimeFormatter::try_new((&id).into(), fieldsets::M::long()) {
        for (i, m) in loc.months.iter_mut().enumerate() {
            if let Ok(d) = Date::try_new_iso(2024, i as u8 + 1, 1) {
                *m = f.format(&d).to_string();
            }
        }
    }
    if let Ok(f) = DateTimeFormatter::try_new((&id).into(), fieldsets::E::long()) {
        // 2023-01-01 was a Sunday.
        for (i, d) in loc.days.iter_mut().enumerate() {
            if let Ok(date) = Date::try_new_iso(2023, 1, i as u8 + 1) {
                *d = f.format(&date).to_string();
            }
        }
    }
    if let Ok(f) = DateTimeFormatter::try_new((&id).into(), fieldsets::YMD::short())
        && let Ok(date) = Date::try_new_iso(2033, 11, 22)
    {
        let text = f.format(&date).to_string();
        let nums: Vec<&str> = text
            .split(|c: char| !c.is_ascii_digit())
            .filter(|s| !s.is_empty())
            .collect();
        let order = match nums.as_slice() {
            [a, b, c] => {
                let year = |s: &str| s == "33" || s == "2033";
                match (*a, *b, *c) {
                    (y, "11", "22") if year(y) => Some(DateOrder::Ymd),
                    ("11", "22", y) if year(y) => Some(DateOrder::Mdy),
                    ("22", "11", y) if year(y) => Some(DateOrder::Dmy),
                    _ => None,
                }
            }
            _ => None,
        };
        match order {
            Some(o) => loc.date_order = o,
            None => notes.push(format!(
                "locale {}: date order not recognised from `{text}`, keeping month-day-year",
                loc.tag
            )),
        }
    }

    if !verified {
        notes.push(format!(
            "locale {}: separators, month/day names and date order come from CLDR and are not verified \
             against Excel or LibreOffice; boolean text and currency keep the English defaults unless configured",
            loc.tag
        ));
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
        loc.date_order = match o.to_ascii_lowercase().as_str() {
            "mdy" => DateOrder::Mdy,
            "dmy" => DateOrder::Dmy,
            "ymd" => DateOrder::Ymd,
            _ => return Err(format!("date_order must be mdy, dmy or ymd, got `{o}`")),
        };
    }
    Ok((loc, notes))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn get(tag: &str) -> Locale {
        from_tag(tag, &LocaleOverride::default()).unwrap().0
    }

    #[test]
    fn regional_variants_follow_cldr() {
        let l = get("es-CO");
        assert_eq!((l.decimal, l.group), (',', '.'));
        let l = get("es-MX");
        assert_eq!((l.decimal, l.group), ('.', ','));
        assert_eq!(l.months[0], "enero");
        assert_eq!(l.true_text, "VERDADERO");
        let l = get("en-GB");
        assert_eq!(l.date_order, DateOrder::Dmy);
        assert_eq!(l.days[0], "Sunday");
    }

    #[test]
    fn unverified_languages_are_flagged() {
        let (l, notes) = from_tag("de-DE", &LocaleOverride::default()).unwrap();
        assert_eq!(
            (l.decimal, l.group, l.date_order),
            (',', '.', DateOrder::Dmy)
        );
        assert_eq!(l.months[2], "März");
        assert_eq!(l.days[1], "Montag");
        assert!(
            notes.iter().any(|n| n.contains("not verified")),
            "{notes:?}"
        );
        let l = get("ja-JP");
        assert_eq!(l.date_order, DateOrder::Ymd);
        assert!(
            from_tag("en", &LocaleOverride::default())
                .unwrap()
                .1
                .is_empty()
        );
    }

    #[test]
    fn configuration_wins_and_bad_tags_fail() {
        let over = LocaleOverride {
            decimal: Some(';'),
            true_text: Some("WAHR".into()),
            date_order: Some("ymd".into()),
            ..LocaleOverride::default()
        };
        let l = from_tag("de", &over).unwrap().0;
        assert_eq!(
            (l.decimal, l.true_text.as_str(), l.date_order),
            (';', "WAHR", DateOrder::Ymd)
        );
        assert!(from_tag("not a tag", &LocaleOverride::default()).is_err());
    }
}
