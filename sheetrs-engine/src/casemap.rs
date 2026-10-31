//! Case mapping and the equality of text that ignores case, as LibreOffice does them.
//!
//! LibreOffice maps case with ICU's full mappings (`ß` has no one-letter uppercase in Unicode and
//! becomes `SS` there; LibreOffice has `ẞ`), and compares text ignoring case by comparing the
//! uppercase of both. The collator of [`crate::collate`] is only for ordering: `Ñ` written as
//! `N` and a combining tilde is not equal to the single letter, and a soft hyphen is not ignored.
//! The mappings here are ICU4X's with the few characters where LibreOffice 26.8 gives another
//! result (found by comparing them for every character).

use icu_casemap::CaseMapperBorrowed;
use icu_locale_core::LanguageIdentifier;
use std::borrow::Cow;

/// Characters whose uppercase is not ICU4X's, with LibreOffice's.
const UPPER_EXCEPTIONS: &[(char, char)] = &[
    ('\u{df}', '\u{1e9e}'),
    ('\u{277}', '\u{277}'),
    ('\u{27c}', '\u{27c}'),
    ('\u{2ce5}', '\u{2ce4}'),
    ('\u{2ce7}', '\u{2ce6}'),
    ('\u{2ce9}', '\u{2ce8}'),
    ('\u{2ceb}', '\u{2cea}'),
    ('\u{ab4b}', '\u{ab4b}'),
    ('\u{ab4c}', '\u{ab4c}'),
];

/// Characters whose lowercase is not ICU4X's, with LibreOffice's.
const LOWER_EXCEPTIONS: &[(char, char)] = &[
    ('\u{130}', '\u{130}'),
    ('\u{2ce4}', '\u{2ce5}'),
    ('\u{2ce6}', '\u{2ce7}'),
    ('\u{2ce8}', '\u{2ce9}'),
    ('\u{2cea}', '\u{2ceb}'),
];

fn mapper() -> CaseMapperBorrowed<'static> {
    CaseMapperBorrowed::new()
}

fn map(s: &str, exceptions: &[(char, char)], icu: impl Fn(&str) -> String) -> String {
    if !s.chars().any(|c| exceptions.iter().any(|e| e.0 == c)) {
        return icu(s);
    }
    let mut out = String::with_capacity(s.len());
    let mut from = 0;
    for (i, c) in s.char_indices() {
        if let Some(&(_, to)) = exceptions.iter().find(|e| e.0 == c) {
            out.push_str(&icu(&s[from..i]));
            out.push(to);
            from = i + c.len_utf8();
        }
    }
    out.push_str(&icu(&s[from..]));
    out
}

/// `UPPER` of LibreOffice.
pub fn upper(s: &str) -> String {
    if s.is_ascii() {
        return s.to_ascii_uppercase();
    }
    // Accented text is the usual non-ASCII text, and each of its letters maps to one letter,
    // context free: the standard library has that mapping for the plane 0, and ICU4X is asked only
    // when a letter does not map to one letter (the test
    // `simple_uppercase_is_icus_for_every_character` compares them all).
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match simple_upper(c) {
            Some(u) => out.push(u),
            None => return upper_icu(s),
        }
    }
    out
}

/// The uppercase of `c` when it is one character and the same in ICU4X and LibreOffice.
fn simple_upper(c: char) -> Option<char> {
    if c.is_ascii() {
        return Some(c.to_ascii_uppercase());
    }
    // The standard library knows newer characters than ICU4X (the plane 1 additions of Unicode 17);
    // the plane 0 is the one both have the same.
    if c as u32 > 0xffff || UPPER_EXCEPTIONS.iter().any(|e| e.0 == c) {
        return None;
    }
    let mut it = c.to_uppercase();
    let u = it.next()?;
    it.next().is_none().then_some(u)
}

fn upper_icu(s: &str) -> String {
    let und = LanguageIdentifier::UNKNOWN;
    map(s, UPPER_EXCEPTIONS, |t| {
        mapper().uppercase_to_string(t, &und).into_owned()
    })
}

/// `LOWER` of LibreOffice.
pub fn lower(s: &str) -> String {
    let und = LanguageIdentifier::UNKNOWN;
    map(s, LOWER_EXCEPTIONS, |t| {
        mapper().lowercase_to_string(t, &und).into_owned()
    })
}

/// `PROPER` of LibreOffice: the first character of the text and every one that follows a
/// character that is not a letter take the uppercase of the whole text at their position, the others
/// the lowercase. The positions are those of UTF-16 code units, as there.
pub fn proper(s: &str) -> String {
    let orig: Vec<u16> = s.encode_utf16().collect();
    if orig.is_empty() {
        return String::new();
    }
    let up: Vec<u16> = upper(s).encode_utf16().collect();
    let lo: Vec<u16> = lower(s).encode_utf16().collect();
    let mut out = orig.clone();
    let at = |v: &[u16], i: usize| v.get(i).copied().unwrap_or(orig[i]);
    out[0] = at(&up, 0);
    for i in 1..out.len() {
        let letter = char::from_u32(u32::from(out[i - 1])).is_some_and(is_letter);
        out[i] = if letter { at(&lo, i) } else { at(&up, i) };
    }
    String::from_utf16_lossy(&out)
}

/// What LibreOffice's `isLetter` says of a character: the alphabetic ones, except two that its
/// older Unicode data has the other way (found by comparing every character of the plane 0).
fn is_letter(c: char) -> bool {
    match c {
        '\u{2ce4}' => false,
        '\u{2cea}' => true,
        c => c.is_alphabetic(),
    }
}

/// Whether two texts are equal: ignoring case when `case_sensitive` is off (the uppercase of both
/// is compared), else character by character.
pub fn text_eq(a: &str, b: &str, case_sensitive: bool) -> bool {
    if a == b {
        return true;
    }
    if case_sensitive {
        return false;
    }
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    upper(a) == upper(b)
}

/// Whether the text has a letter whose case folding is not one character (`ß` is `ss`, `ﬁ` is
/// `fi`): ICU matches those against the folded text in ways of its own.
pub fn has_multi_fold(s: &str) -> bool {
    !s.is_ascii()
        && s.chars().any(|c| {
            !c.is_ascii() && {
                let mut b = [0u8; 4];
                mapper().fold_string(c.encode_utf8(&mut b)).chars().count() != 1
            }
        })
}

/// The text as LibreOffice's search ignores case (`SEARCH`, a criterion that need not cover the
/// whole cell): each character in its full case folding (`ß` is `ss`; `İ` is left alone), with the
/// position in `text` of the character each folded one came from.
pub fn fold_search(text: &str) -> (String, Vec<usize>) {
    let mut out = String::with_capacity(text.len());
    let mut from = Vec::with_capacity(text.len());
    for (i, c) in text.chars().enumerate() {
        let before = out.len();
        if c.is_ascii() {
            out.push(c.to_ascii_lowercase());
        } else if c == '\u{130}' {
            out.push(c);
        } else {
            let mut b = [0u8; 4];
            out.push_str(&mapper().fold_string(c.encode_utf8(&mut b)));
        }
        from.extend(std::iter::repeat_n(i, out[before..].chars().count()));
    }
    (out, from)
}

/// The text with the case ignored, to compare or hash: its uppercase.
pub fn fold(s: &str) -> Cow<'_, str> {
    if s.is_ascii() {
        if s.bytes().any(|b| b.is_ascii_lowercase()) {
            Cow::Owned(s.to_ascii_uppercase())
        } else {
            Cow::Borrowed(s)
        }
    } else {
        Cow::Owned(upper(s))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn simple_uppercase_is_icus_for_every_character() {
        let mut served = 0;
        for c in (0..=0x10ffff).filter_map(char::from_u32) {
            if let Some(u) = simple_upper(c) {
                served += 1;
                assert_eq!(
                    upper_icu(&c.to_string()),
                    u.to_string(),
                    "U+{:04X}",
                    c as u32
                );
            }
        }
        assert!(served > 60_000);
        let text = "Ñandú, ÁÉÍÓÚ üñ çà straße ΟΔΥΣΣΕΥΣ ǆ ﬁ";
        assert_eq!(upper(text), upper_icu(text));
    }

    #[test]
    fn maps_as_libreoffice() {
        // Full mappings, the capital sharp s, and the final sigma.
        assert_eq!(upper("straße ŉ ﬁsh"), "STRAẞE ʼN FISH");
        assert_eq!(lower("ΟΔΥΣΣΕΥΣ ΑΣ İ"), "οδυσσευς ας İ");
        for (from, to) in [
            ("ŉ", "ʼ"),
            ("ﬁsh", "Fsh"),
            ("ǆemal", "Ǆemal"),
            ("o'neil", "O'Neil"),
            ("ΑΣ", "Ας"),
            ("ß ss", "ẞ Ss"),
            ("İSTANBUL", "İstanbul"),
            ("ǆemal", "Ǆemal"),
            // Characters beyond the plane 0 are two units there, and a lone unit is no letter.
            ("𝒜b", "𝒜B"),
        ] {
            assert_eq!(proper(from), to, "{from}");
        }
    }

    #[test]
    fn equality_ignores_case_but_not_composition() {
        assert!(text_eq("Resume", "rESUME", false));
        assert!(!text_eq("Resume", "rESUME", true));
        assert!(text_eq("É", "é", false));
        assert!(
            !text_eq("ñ", "n\u{303}", false),
            "a combining tilde is not the letter"
        );
        assert!(
            !text_eq("ab", "a\u{ad}b", false),
            "a soft hyphen is not ignored"
        );
        assert!(!text_eq("straße", "STRASSE", false));
        assert!(text_eq("straße", "STRAẞE", false));
    }
}
