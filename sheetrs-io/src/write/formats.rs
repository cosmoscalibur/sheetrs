//! The number formats a replacement gives a cell (`crate::EntryFormat`: what
//! `LibreOffice` assigns to a cell in `General` that is typed a percentage, a date, a boolean...),
//! and how each file writes them.
//!
//! The codes and the data styles are the ones `LibreOffice` 26.8 saves for the same cell.

use super::error::Result;
use crate::EntryFormat;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::fmt::Write as _;

/// The format code in an XLSX (`numFmt/@formatCode`).
pub(crate) fn xlsx_code(format: EntryFormat) -> &'static str {
    match format {
        EntryFormat::Percent => "0.00%",
        EntryFormat::Date => "mm/dd/yy",
        EntryFormat::DateIso => r"yyyy\-mm\-dd",
        EntryFormat::DateTime => r"mm/dd/yy\ hh:mm\ AM/PM",
        EntryFormat::DateTimeIso => r"yyyy\-mm\-dd\ hh:mm:ss",
        EntryFormat::DateTimeIsoT => r"yyyy\-mm\-dd\Thh:mm:ss",
        EntryFormat::Time => r"hh:mm:ss\ AM/PM",
        EntryFormat::TimeTotal => "[hh]:mm:ss",
        EntryFormat::TimeTotalFraction => "[hh]:mm:ss.00",
        EntryFormat::Scientific => "0.00E+00",
        EntryFormat::Currency => r"[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00",
        EntryFormat::Fraction => "# ?/?",
        EntryFormat::Boolean => r#""TRUE";"TRUE";"FALSE""#,
    }
}

/// The code the engine formats a number with, to write the text a cell shows.
#[must_use]
pub fn display_code(format: EntryFormat) -> &'static str {
    match format {
        EntryFormat::DateTime => "mm/dd/yy hh:mm AM/PM",
        EntryFormat::DateIso => "yyyy-mm-dd",
        EntryFormat::DateTimeIso => "yyyy-mm-dd hh:mm:ss",
        EntryFormat::DateTimeIsoT => "yyyy-mm-dd\"T\"hh:mm:ss",
        EntryFormat::Time => "hh:mm:ss AM/PM",
        EntryFormat::Currency => "[$$-409]#,##0.00;-[$$-409]#,##0.00",
        EntryFormat::Boolean => "General",
        other => xlsx_code(other),
    }
}

/// `office:value-type` of a number with the format.
pub(crate) fn ods_value_type(format: EntryFormat) -> &'static str {
    match format {
        EntryFormat::Percent => "percentage",
        EntryFormat::Date
        | EntryFormat::DateIso
        | EntryFormat::DateTime
        | EntryFormat::DateTimeIso
        | EntryFormat::DateTimeIsoT => "date",
        EntryFormat::Time | EntryFormat::TimeTotal | EntryFormat::TimeTotalFraction => "time",
        EntryFormat::Currency => "currency",
        EntryFormat::Scientific | EntryFormat::Fraction | EntryFormat::Boolean => "float",
    }
}

const TWO_DECIMALS: &str =
    r#"number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1""#;

fn hours_minutes(tail: &str) -> String {
    format!(
        r#"<number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/>{tail}"#
    )
}

/// The data style (`number:*-style`) of the format, named `name`.
pub(crate) fn ods_data_style(format: EntryFormat, name: &str) -> String {
    let seconds = r#"<number:text>:</number:text><number:seconds number:style="long"/>"#;
    let iso_date = r#"<number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/>"#;
    let us_date = r#"<number:month number:style="long"/><number:text>/</number:text><number:day number:style="long"/><number:text>/</number:text><number:year/>"#;
    match format {
        EntryFormat::Percent => format!(
            r#"<number:percentage-style style:name="{name}"><number:number {TWO_DECIMALS}/><number:text>%</number:text></number:percentage-style>"#
        ),
        EntryFormat::Date => format!(
            r#"<number:date-style style:name="{name}" number:automatic-order="true">{us_date}</number:date-style>"#
        ),
        EntryFormat::DateIso => {
            format!(r#"<number:date-style style:name="{name}">{iso_date}</number:date-style>"#)
        }
        EntryFormat::DateTime => format!(
            r#"<number:date-style style:name="{name}">{us_date}<number:text> </number:text>{}<number:text> </number:text><number:am-pm/></number:date-style>"#,
            hours_minutes("")
        ),
        EntryFormat::DateTimeIso => format!(
            r#"<number:date-style style:name="{name}">{iso_date}<number:text> </number:text>{}</number:date-style>"#,
            hours_minutes(seconds)
        ),
        EntryFormat::DateTimeIsoT => format!(
            r#"<number:date-style style:name="{name}">{iso_date}<number:text>T</number:text>{}</number:date-style>"#,
            hours_minutes(seconds)
        ),
        EntryFormat::Time => format!(
            r#"<number:time-style style:name="{name}">{}<number:text> </number:text><number:am-pm/></number:time-style>"#,
            hours_minutes(seconds)
        ),
        EntryFormat::TimeTotal => format!(
            r#"<number:time-style style:name="{name}" number:truncate-on-overflow="false">{}</number:time-style>"#,
            hours_minutes(seconds)
        ),
        EntryFormat::TimeTotalFraction => format!(
            r#"<number:time-style style:name="{name}" number:truncate-on-overflow="false">{}<number:text>:</number:text><number:seconds number:style="long" number:decimal-places="2"/></number:time-style>"#,
            hours_minutes("")
        ),
        EntryFormat::Scientific => format!(
            r#"<number:number-style style:name="{name}"><number:scientific-number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2" number:exponent-interval="1" number:forced-exponent-sign="true"/></number:number-style>"#
        ),
        EntryFormat::Fraction => format!(
            r#"<number:number-style style:name="{name}"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1" number:max-denominator-value="9"/></number:number-style>"#
        ),
        EntryFormat::Boolean => format!(
            r#"<number:boolean-style style:name="{name}"><number:boolean/></number:boolean-style>"#
        ),
        EntryFormat::Currency => {
            let symbol = r#"<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>"#;
            let number = format!(r#"<number:number {TWO_DECIMALS} number:grouping="true"/>"#);
            format!(
                r##"<number:currency-style style:name="{name}-positive" style:volatile="true">{symbol}{number}</number:currency-style><number:currency-style style:name="{name}"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>{symbol}{number}<style:map style:condition="value()&gt;=0" style:apply-style-name="{name}-positive"/></number:currency-style>"##
            )
        }
    }
}

/// The text of `office:date-value` / `office:time-value` of the number `x` written with the
/// format, the one `LibreOffice` writes.
#[must_use]
pub fn ods_value(format: EntryFormat, x: f64, sys: sheetrs_formula::DateSystem) -> Option<String> {
    match format {
        EntryFormat::Date | EntryFormat::DateIso => {
            let (y, m, d) = sheetrs_formula::calendar::serial_to_ymd_quiet(x.floor(), sys)?;
            Some(format!("{y:04}-{m:02}-{d:02}"))
        }
        EntryFormat::DateTime | EntryFormat::DateTimeIso | EntryFormat::DateTimeIsoT => {
            let mut whole = x.floor();
            let mut seconds = (x.fract() * 86_400.0).round();
            if seconds >= 86_400.0 {
                whole += 1.0;
                seconds = 0.0;
            }
            let (y, m, d) = sheetrs_formula::calendar::serial_to_ymd_quiet(whole, sys)?;
            Some(format!(
                "{y:04}-{m:02}-{d:02}T{:02}:{:02}:{:02}",
                seconds as u32 / 3600,
                seconds as u32 / 60 % 60,
                seconds as u32 % 60
            ))
        }
        EntryFormat::Time | EntryFormat::TimeTotal | EntryFormat::TimeTotalFraction => {
            let total = (x * 86_400_000.0).round();
            let (hours, rest) = ((total / 3_600_000.0).floor(), total % 3_600_000.0);
            let (minutes, rest) = ((rest / 60_000.0).floor(), rest % 60_000.0);
            let seconds = rest / 1000.0;
            let seconds = if seconds.fract() == 0.0 {
                format!("{seconds:02}")
            } else {
                format!("{seconds:.3}")
                    .trim_end_matches('0')
                    .trim_end_matches('.')
                    .to_string()
            };
            Some(format!("PT{hours:02}H{minutes:02}M{seconds}S"))
        }
        _ => None,
    }
}

/// The styles part of an XLSX, to which the cell formats a replacement needs are added: each is
/// the format `cellXfs/xf` of the cell, with another number format.
pub(crate) struct StyleBook {
    xml: String,
    /// `numFmtId` and `formatCode` of the formats the part defines.
    formats: Vec<(u32, String)>,
    xfs: Vec<String>,
    added_formats: Vec<(u32, String)>,
    added_xfs: Vec<String>,
    derived: HashMap<(u32, EntryFormat), u32>,
}

fn attribute_of(tag: &BytesStart, key: &str) -> Option<String> {
    super::attribute(tag, key).ok().flatten()
}

fn escape_attribute(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
}

impl StyleBook {
    /// The book of the styles part `xml`; `None` when it has no `cellXfs` the way the files that
    /// `Excel` and `LibreOffice` write have it.
    pub(crate) fn new(xml: String) -> Result<Option<Self>> {
        let mut reader = Reader::from_str(&xml);
        let mut formats = Vec::new();
        let mut xfs = Vec::new();
        let mut in_xfs = false;
        let mut depth_in_xf = 0u32;
        let mut xf_start = 0usize;
        let mut buf = Vec::new();
        loop {
            let before = usize::try_from(reader.buffer_position()).unwrap_or(0);
            buf.clear();
            let event = reader.read_event_into(&mut buf)?;
            let after = usize::try_from(reader.buffer_position()).unwrap_or(0);
            match &event {
                Event::Eof => break,
                Event::Empty(e) if e.name().as_ref() == "numFmt" => {
                    if let (Some(id), Some(code)) =
                        (attribute_of(e, "numFmtId"), attribute_of(e, "formatCode"))
                        && let Ok(id) = id.parse()
                    {
                        formats.push((id, code));
                    }
                }
                Event::Start(e) if e.name().as_ref() == "cellXfs" => in_xfs = true,
                Event::End(e) if e.name().as_ref() == "cellXfs" => in_xfs = false,
                Event::Empty(e) if in_xfs && depth_in_xf == 0 && e.name().as_ref() == "xf" => {
                    xfs.push(xml[before..after].to_string());
                }
                Event::Start(e) if in_xfs && depth_in_xf == 0 && e.name().as_ref() == "xf" => {
                    xf_start = before;
                    depth_in_xf = 1;
                }
                Event::Start(_) if depth_in_xf > 0 => depth_in_xf += 1,
                Event::End(_) if depth_in_xf > 0 => {
                    depth_in_xf -= 1;
                    if depth_in_xf == 0 {
                        xfs.push(xml[xf_start..after].to_string());
                    }
                }
                _ => {}
            }
        }
        if xfs.is_empty() {
            return Ok(None);
        }
        Ok(Some(Self {
            xml,
            formats,
            xfs,
            added_formats: Vec::new(),
            added_xfs: Vec::new(),
            derived: HashMap::new(),
        }))
    }

    /// The index of the `xf` that is `base` with the number format of `format`.
    pub(crate) fn derive(&mut self, base: u32, format: EntryFormat) -> u32 {
        if let Some(&index) = self.derived.get(&(base, format)) {
            return index;
        }
        let code = xlsx_code(format);
        let known = self
            .formats
            .iter()
            .chain(&self.added_formats)
            .find(|(_, c)| c == code)
            .map(|&(id, _)| id);
        let id = if let Some(id) = known {
            id
        } else {
            let id = self
                .formats
                .iter()
                .chain(&self.added_formats)
                .map(|&(id, _)| id)
                .fold(163, u32::max)
                + 1;
            self.added_formats.push((id, code.to_string()));
            id
        };
        let original = self.xfs.get(base as usize).unwrap_or(&self.xfs[0]).clone();
        let xf = with_number_format(&original, id);
        self.added_xfs.push(xf);
        let index = u32::try_from(self.xfs.len() + self.added_xfs.len() - 1).unwrap_or(0);
        self.derived.insert((base, format), index);
        index
    }

    /// The styles part with what was added.
    pub(crate) fn render(&self) -> String {
        let mut xml = self.xml.clone();
        // `cellXfs` first: it comes after `numFmts`, so the positions before it stay.
        if let Some(close) = xml.find("</cellXfs>") {
            xml.insert_str(close, &self.added_xfs.concat());
            if let Some(open) = xml.find("<cellXfs") {
                set_count(&mut xml, open, self.xfs.len() + self.added_xfs.len());
            }
        }
        if !self.added_formats.is_empty() {
            let mut entries = String::new();
            for (id, code) in &self.added_formats {
                let _ = write!(
                    entries,
                    r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                    escape_attribute(code)
                );
            }
            if let Some(close) = xml.find("</numFmts>") {
                xml.insert_str(close, &entries);
                if let Some(open) = xml.find("<numFmts") {
                    set_count(
                        &mut xml,
                        open,
                        self.formats.len() + self.added_formats.len(),
                    );
                }
            } else if let Some(root) = xml.find("<styleSheet")
                && let Some(end) = xml[root..].find('>')
            {
                let at = root + end + 1;
                xml.insert_str(
                    at,
                    &format!(
                        r#"<numFmts count="{}">{entries}</numFmts>"#,
                        self.added_formats.len()
                    ),
                );
            }
        }
        xml
    }
}

/// Sets the `count` attribute of the element that opens at `open`.
fn set_count(xml: &mut String, open: usize, count: usize) {
    let Some(tag_end) = xml[open..].find('>') else {
        return;
    };
    let tag = &xml[open..open + tag_end];
    if let Some(at) = tag.find("count=\"") {
        let value_start = open + at + 7;
        if let Some(len) = xml[value_start..].find('"') {
            xml.replace_range(value_start..value_start + len, &count.to_string());
        }
    }
}

/// The `xf` with `numFmtId` and `applyNumberFormat` set.
fn with_number_format(xf: &str, id: u32) -> String {
    let open_end = xf.find('>').unwrap_or(xf.len());
    let (open, rest) = xf.split_at(open_end);
    let mut open = open.to_string();
    if let Some(at) = open.find("numFmtId=\"") {
        let value_start = at + 10;
        if let Some(len) = open[value_start..].find('"') {
            open.replace_range(value_start..value_start + len, &id.to_string());
        }
    } else {
        open.insert_str(3, &format!(r#" numFmtId="{id}""#));
    }
    if let Some(at) = open.find("applyNumberFormat=\"") {
        let value_start = at + 19;
        if let Some(len) = open[value_start..].find('"') {
            open.replace_range(value_start..value_start + len, "1");
        }
    } else {
        let at = open.find("numFmtId=\"").unwrap_or(0);
        let end = open[at + 10..]
            .find('"')
            .map_or(open.len(), |l| at + 10 + l + 1);
        open.insert_str(end, r#" applyNumberFormat="1""#);
    }
    format!("{open}{rest}")
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r#"<?xml version="1.0"?><styleSheet xmlns="x"><fonts count="1"><font/></fonts><cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="4" fontId="0" applyNumberFormat="1"><alignment horizontal="center"/></xf></cellXfs></styleSheet>"#;

    #[test]
    fn a_derived_style_keeps_the_style_and_changes_the_format() {
        let mut book = StyleBook::new(STYLES.to_string()).unwrap().unwrap();
        let a = book.derive(1, EntryFormat::Percent);
        let b = book.derive(1, EntryFormat::Percent);
        let c = book.derive(0, EntryFormat::Percent);
        assert_eq!((a, b, c), (2, 2, 3));
        let xml = book.render();
        assert!(xml.contains(
            r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>"#
        ));
        assert!(xml.contains(r#"<cellXfs count="4">"#));
        assert!(xml.contains(r#"<xf numFmtId="164" fontId="0" applyNumberFormat="1"><alignment horizontal="center"/></xf>"#));
        assert!(xml.contains(r#"<xf numFmtId="164" applyNumberFormat="1" fontId="0"/>"#));
        quick_xml::Reader::from_str(&xml);
    }

    #[test]
    fn a_format_the_part_already_has_is_not_defined_again() {
        let xml = STYLES.replace(
            "<fonts",
            r#"<numFmts count="1"><numFmt numFmtId="170" formatCode="0.00%"/></numFmts><fonts"#,
        );
        let mut book = StyleBook::new(xml).unwrap().unwrap();
        book.derive(0, EntryFormat::Percent);
        let xml = book.render();
        assert!(xml.contains(r#"<numFmts count="1">"#), "{xml}");
        assert!(xml.contains(r#"<xf numFmtId="170" applyNumberFormat="1" fontId="0"/>"#));
    }

    #[test]
    fn the_values_of_dates_and_times_are_the_ones_libreoffice_writes() {
        let value = |f, x| ods_value(f, x, sheetrs_formula::DateSystem::Null1899).unwrap();
        assert_eq!(value(EntryFormat::Date, 45366.0), "2024-03-15");
        assert_eq!(
            value(EntryFormat::DateTimeIso, 45_366.430_555_555_555),
            "2024-03-15T10:20:00"
        );
        assert_eq!(value(EntryFormat::Time, 0.0), "PT00H00M00S");
        assert_eq!(
            value(EntryFormat::TimeTotal, 2.020_833_333_333_333_5),
            "PT48H30M00S"
        );
        assert_eq!(
            value(EntryFormat::TimeTotalFraction, 0.521_359_953_703_703_7),
            "PT12H30M45.5S"
        );
    }
}
