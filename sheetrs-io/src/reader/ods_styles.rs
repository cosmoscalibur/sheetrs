//! The number formats of an ODS document, written as the format codes Excel uses.
//!
//! An ODS keeps a format as a tree (`number:date-style`, `number:time-style`,
//! `number:number-style`, `number:percentage-style`, `number:currency-style`, ...) and names it from
//! the style of the cell (`style:style` with `style:data-style-name`). A cell then reads the same
//! format code in an ODS as in an XLSX: `LibreOffice` writes the XLSX of the same cell with the code
//! this module builds, apart from the locale mark (`[$-409]`) the XLSX adds.
//!
//! The styles are fed event by event while the document is read ([`NumberStyles::feed`]) and
//! resolved once, when the first table starts ([`NumberStyles::resolve`]).

use super::parser_utils::{attr_value, push_general_ref};
use super::workbook::NumFmt;
use crate::error::Result;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

/// What the tree of a data style says, as the sections of one format code.
#[derive(Default)]
struct DataStyle {
    /// The code of this section, without colour.
    body: String,
    /// Excel's name of the colour of the section (`RED`), when the style sets one.
    color: Option<&'static str>,
    /// The styles that apply to the values the conditions of `style:map` pick, in order: they are
    /// the sections that come before this one.
    maps: Vec<String>,
    /// The time of the style is a duration (`[hh]`): its first element does not wrap.
    open_ended: bool,
    /// A time element went by, so the brackets of `open_ended` are spent.
    time_seen: bool,
    /// The style follows the language (`number:format-source="language"`): a day or a month that
    /// does not say how long it is takes two digits, as `LibreOffice` writes the XLSX of it.
    from_language: bool,
}

/// The text of a `number:text` or `number:currency-symbol` being read.
struct Literal {
    text: String,
    symbol: bool,
    /// The character `loext:blank-width-char` asks to leave as wide as it is.
    blank: Option<char>,
}

#[derive(Default)]
pub(super) struct NumberStyles {
    styles: HashMap<String, DataStyle>,
    /// A cell style: its data style and its parent style.
    cells: HashMap<String, (Option<String>, Option<String>)>,
    open: Option<(String, DataStyle)>,
    literal: Option<Literal>,
}

fn attr(e: &BytesStart<'_>, key: &str) -> Result<Option<String>> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == key {
            return Ok(Some(attr_value(&a)?.into_owned()));
        }
    }
    Ok(None)
}

fn count(e: &BytesStart<'_>, key: &str) -> Result<Option<usize>> {
    Ok(attr(e, key)?.and_then(|v| v.parse().ok()))
}

/// Excel's name of a colour `LibreOffice` writes as `#rrggbb`.
fn color_name(hex: &str) -> Option<&'static str> {
    Some(match hex.to_ascii_lowercase().as_str() {
        "#000000" => "BLACK",
        "#0000ff" => "BLUE",
        "#00ffff" => "CYAN",
        "#00ff00" => "GREEN",
        "#ff00ff" => "MAGENTA",
        "#ff0000" => "RED",
        "#ffffff" => "WHITE",
        "#ffff00" => "YELLOW",
        _ => return None,
    })
}

/// A piece of fixed text in a format code: as it is when Excel takes it for itself, and in quotes
/// otherwise.
fn literal_code(text: &str) -> String {
    if text.is_empty() {
        return String::new();
    }
    if text == "%" || text.chars().all(|c| "$-+/():!^&'~{}<>= .".contains(c)) {
        return text.to_string();
    }
    format!("\"{}\"", text.replace('"', "\\\""))
}

/// The digits of a `number:number`: `#,##0.00`.
fn mantissa(e: &BytesStart<'_>) -> Result<String> {
    let min_int = count(e, "number:min-integer-digits")?.unwrap_or(1);
    let places = count(e, "number:decimal-places")?;
    // Neither decimals nor grouping: the number as it is, which is what Excel calls General.
    if places.is_none() && min_int <= 1 && attr(e, "number:grouping")?.is_none() {
        return Ok("General".to_string());
    }
    let decimals = places.unwrap_or(0);
    let min_decimals = count(e, "number:min-decimal-places")?
        .unwrap_or(decimals)
        .min(decimals);
    let digits = if min_int == 0 {
        "#".to_string()
    } else {
        "0".repeat(min_int)
    };
    let mut code = if attr(e, "number:grouping")?.as_deref() == Some("true") {
        format!("#,##{digits}")
    } else {
        digits
    };
    if decimals > 0 {
        code.push('.');
        code.push_str(&"0".repeat(min_decimals));
        code.push_str(&"#".repeat(decimals - min_decimals));
    }
    Ok(code)
}

fn scientific(e: &BytesStart<'_>) -> Result<String> {
    let exponent = count(e, "number:min-exponent-digits")?.unwrap_or(2).max(1);
    let sign = if attr(e, "number:forced-exponent-sign")?.as_deref() == Some("true") {
        '+'
    } else {
        '-'
    };
    Ok(format!("{}E{sign}{}", mantissa(e)?, "0".repeat(exponent)))
}

fn fraction(e: &BytesStart<'_>) -> Result<String> {
    let integer = match count(e, "number:min-integer-digits")?.unwrap_or(0) {
        0 => "#".to_string(),
        n => "0".repeat(n),
    };
    let numerator = count(e, "loext:max-numerator-digits")?
        .or(count(e, "number:min-numerator-digits")?)
        .unwrap_or(1);
    let denominator = if let Some(fixed) = attr(e, "number:denominator-value")? {
        fixed
    } else {
        let digits = attr(e, "number:max-denominator-value")?
            .map(|v| v.len())
            .or(count(e, "number:min-denominator-digits")?)
            .unwrap_or(1);
        "?".repeat(digits)
    };
    Ok(format!("{integer} {}/{denominator}", "?".repeat(numerator)))
}

impl NumberStyles {
    /// Reads what an event of `styles.xml` or of the start of `content.xml` says of the formats.
    ///
    /// # Errors
    ///
    /// When an attribute or a reference of the XML cannot be decoded.
    pub(super) fn feed(&mut self, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(e) => self.start(e, false)?,
            Event::Empty(e) => self.start(e, true)?,
            Event::Text(t) => {
                if let Some(l) = self.literal.as_mut() {
                    l.text.push_str(t);
                }
            }
            Event::GeneralRef(r) => {
                if let Some(l) = self.literal.as_mut() {
                    push_general_ref(&mut l.text, r)?;
                }
            }
            Event::End(e) => self.end(e.name().as_ref()),
            _ => {}
        }
        Ok(())
    }

    fn start(&mut self, e: &BytesStart<'_>, empty: bool) -> Result<()> {
        let name = e.name();
        let name = name.as_ref();
        match name {
            "number:number-style"
            | "number:percentage-style"
            | "number:currency-style"
            | "number:date-style"
            | "number:time-style"
            | "number:text-style"
            | "number:boolean-style" => {
                let style = DataStyle {
                    open_ended: attr(e, "number:truncate-on-overflow")?.as_deref() == Some("false"),
                    from_language: attr(e, "number:format-source")?.as_deref() == Some("language"),
                    ..DataStyle::default()
                };
                if let Some(style_name) = attr(e, "style:name")? {
                    self.open = Some((style_name, style));
                    if empty {
                        self.close();
                    }
                }
            }
            "style:style" => {
                if attr(e, "style:family")?.as_deref() == Some("table-cell")
                    && let Some(name) = attr(e, "style:name")?
                {
                    self.cells.insert(
                        name,
                        (
                            attr(e, "style:data-style-name")?,
                            attr(e, "style:parent-style-name")?,
                        ),
                    );
                }
            }
            "number:text" | "number:currency-symbol" if self.open.is_some() && !empty => {
                self.literal = Some(Literal {
                    text: String::new(),
                    symbol: name == "number:currency-symbol",
                    blank: attr(e, "loext:blank-width-char")?.and_then(|c| c.chars().next()),
                });
            }
            _ if self.open.is_some() => self.element(name, e)?,
            _ => {}
        }
        Ok(())
    }

    fn element(&mut self, name: &str, e: &BytesStart<'_>) -> Result<()> {
        let Some((_, style)) = self.open.as_mut() else {
            return Ok(());
        };
        let long = match attr(e, "number:style")?.as_deref() {
            Some(how) => how == "long",
            None => style.from_language && matches!(name, "number:day" | "number:month"),
        };
        match name {
            "number:number" => style.body.push_str(&mantissa(e)?),
            "number:scientific-number" => style.body.push_str(&scientific(e)?),
            "number:fraction" => style.body.push_str(&fraction(e)?),
            "number:text-content" => style.body.push('@'),
            "number:boolean" => style.body.push_str("BOOLEAN"),
            "number:day" => style.body.push_str(if long { "dd" } else { "d" }),
            "number:day-of-week" => style.body.push_str(if long { "dddd" } else { "ddd" }),
            "number:month" => {
                let textual = attr(e, "number:textual")?.as_deref() == Some("true");
                style.body.push_str(match (textual, long) {
                    (true, true) => "mmmm",
                    (true, false) => "mmm",
                    (false, true) => "mm",
                    (false, false) => "m",
                });
            }
            "number:year" => style.body.push_str(if long { "yyyy" } else { "yy" }),
            "number:am-pm" => style.body.push_str("AM/PM"),
            "number:hours" | "number:minutes" | "number:seconds" => {
                let unit = match name {
                    "number:hours" => 'h',
                    "number:minutes" => 'm',
                    _ => 's',
                };
                let mut code = String::new();
                code.push(unit);
                if long {
                    code.push(unit);
                }
                if style.open_ended && !style.time_seen {
                    code = format!("[{code}]");
                }
                style.time_seen = true;
                style.body.push_str(&code);
                if unit == 's'
                    && let Some(places) = count(e, "number:decimal-places")?
                    && places > 0
                {
                    style.body.push('.');
                    style.body.push_str(&"0".repeat(places));
                }
            }
            "style:text-properties" => {
                if let Some(color) = attr(e, "fo:color")? {
                    style.color = color_name(&color);
                }
            }
            "style:map" => {
                if let Some(apply) = attr(e, "style:apply-style-name")? {
                    style.maps.push(apply);
                }
            }
            _ => {}
        }
        Ok(())
    }

    fn end(&mut self, name: &str) {
        match name {
            "number:text" | "number:currency-symbol" => {
                if let (Some(l), Some((_, style))) = (self.literal.take(), self.open.as_mut()) {
                    if l.symbol {
                        if !l.text.is_empty() {
                            style.body.push_str("[$");
                            style.body.push_str(&l.text);
                            style.body.push(']');
                        }
                    } else if let Some(c) = l.blank {
                        style.body.push('_');
                        style.body.push(c);
                    } else {
                        style.body.push_str(&literal_code(&l.text));
                    }
                }
            }
            "number:number-style"
            | "number:percentage-style"
            | "number:currency-style"
            | "number:date-style"
            | "number:time-style"
            | "number:text-style"
            | "number:boolean-style" => self.close(),
            _ => {}
        }
    }

    fn close(&mut self) {
        if let Some((name, style)) = self.open.take() {
            self.styles.insert(name, style);
        }
    }

    fn section(style: &DataStyle) -> String {
        match style.color {
            Some(color) => format!("[{color}]{}", style.body),
            None => style.body.clone(),
        }
    }

    /// The format code of a data style: the sections of the styles its conditions apply, then its own.
    fn code(&self, style: &DataStyle) -> String {
        if style.maps.is_empty() {
            return Self::section(style);
        }
        let mut sections: Vec<String> = style
            .maps
            .iter()
            .map(|apply| {
                self.styles
                    .get(apply)
                    .map(Self::section)
                    .unwrap_or_default()
            })
            .collect();
        sections.push(Self::section(style));
        sections.join(";")
    }

    /// The format code of every data style and every cell style that has one.
    pub(super) fn resolve(&self) -> HashMap<String, NumFmt> {
        let mut data: HashMap<&str, NumFmt> = HashMap::new();
        for (name, style) in &self.styles {
            let code = self.code(style);
            if !code.is_empty() {
                data.insert(name, NumFmt::from(code.as_str()));
            }
        }
        let mut resolved: HashMap<String, NumFmt> = data
            .iter()
            .map(|(name, fmt)| ((*name).to_string(), fmt.clone()))
            .collect();
        for name in self.cells.keys() {
            // A style without a data style of its own takes the one of its parent.
            let mut at = name.as_str();
            for _ in 0..16 {
                let Some((data_style, parent)) = self.cells.get(at) else {
                    break;
                };
                if let Some(format) = data_style.as_deref().and_then(|d| data.get(d)) {
                    resolved.insert(name.clone(), format.clone());
                    break;
                }
                if data_style.is_some() {
                    break;
                }
                match parent {
                    Some(parent) => at = parent,
                    None => break,
                }
            }
        }
        resolved
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::Reader;

    fn formats(xml: &str) -> HashMap<String, String> {
        let mut styles = NumberStyles::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        loop {
            let event = reader.read_event().unwrap();
            if matches!(event, Event::Eof) {
                break;
            }
            styles.feed(&event).unwrap();
        }
        styles
            .resolve()
            .into_iter()
            .map(|(k, v)| (k, v.to_string()))
            .collect()
    }

    fn one(style: &str) -> String {
        let xml = format!(
            "<r xmlns:number='n' xmlns:style='s' xmlns:fo='f' xmlns:loext='l'>{style}\
             <style:style style:name='c' style:family='table-cell' style:data-style-name='D'/></r>"
        );
        formats(&xml).remove("c").unwrap_or_default()
    }

    #[test]
    fn a_time_with_the_hours_not_wrapped() {
        let code = one(
            "<number:time-style style:name='D'><number:hours/><number:text>:</number:text>\
             <number:minutes number:style='long'/><number:text> </number:text><number:am-pm/>\
             </number:time-style>",
        );
        assert_eq!(code, "h:mm AM/PM");
    }

    #[test]
    fn a_duration_wraps_its_first_element() {
        let code = one(
            "<number:time-style style:name='D' number:truncate-on-overflow='false'>\
             <number:minutes number:style='long'/><number:text>:</number:text>\
             <number:seconds number:style='long' number:decimal-places='1'/></number:time-style>",
        );
        assert_eq!(code, "[mm]:ss.0");
    }

    #[test]
    fn a_percentage_and_a_currency() {
        let percent = one(
            "<number:percentage-style style:name='D'><number:number number:decimal-places='2' \
             number:min-decimal-places='2' number:min-integer-digits='1'/>\
             <number:text>%</number:text></number:percentage-style>",
        );
        assert_eq!(percent, "0.00%");
        let currency = one(
            "<number:currency-style style:name='D'><number:number number:decimal-places='2' \
             number:min-decimal-places='2' number:min-integer-digits='1' number:grouping='true'/>\
             <number:text> </number:text><number:currency-symbol>EUR</number:currency-symbol>\
             </number:currency-style>",
        );
        assert_eq!(currency, "#,##0.00 [$EUR]");
    }

    #[test]
    fn the_sections_of_conditions_come_first() {
        let code = one("<number:number-style style:name='D'>\
             <style:text-properties fo:color='#ff0000'/><number:text>-</number:text>\
             <number:number number:decimal-places='2' number:min-decimal-places='2' \
             number:min-integer-digits='1'/>\
             <style:map style:condition='value()&gt;=0' style:apply-style-name='P'/>\
             </number:number-style>\
             <number:number-style style:name='P'><number:number number:decimal-places='2' \
             number:min-decimal-places='2' number:min-integer-digits='1'/></number:number-style>");
        assert_eq!(code, "0.00;[RED]-0.00");
    }

    #[test]
    fn text_is_quoted_unless_excel_takes_it_as_it_is() {
        let code = one(
            "<number:number-style style:name='D'><number:text>text </number:text>\
             <number:number number:decimal-places='0' number:min-integer-digits='1'/>\
             </number:number-style>",
        );
        assert_eq!(code, "\"text \"0");
        let date = one(
            "<number:date-style style:name='D'><number:month number:style='long' \
             number:textual='true'/><number:text> </number:text><number:day/>\
             <number:text>, </number:text><number:year number:style='long'/></number:date-style>",
        );
        assert_eq!(date, "mmmm d\", \"yyyy");
    }

    #[test]
    fn a_fraction_and_a_scientific_number() {
        let fraction = one(
            "<number:number-style style:name='D'><number:fraction number:min-integer-digits='0' \
             number:min-numerator-digits='1' loext:max-numerator-digits='2' \
             number:min-denominator-digits='1' number:max-denominator-value='99'/>\
             </number:number-style>",
        );
        assert_eq!(fraction, "# ??/??");
        let sci = one(
            "<number:number-style style:name='D'><number:scientific-number \
             number:decimal-places='2' number:min-decimal-places='2' number:min-integer-digits='1' \
             number:min-exponent-digits='2' number:forced-exponent-sign='true'/></number:number-style>",
        );
        assert_eq!(sci, "0.00E+00");
    }

    #[test]
    fn a_cell_style_without_a_data_style_takes_the_one_of_its_parent() {
        let xml = "<r xmlns:number='n' xmlns:style='s'>\
            <number:date-style style:name='D'><number:year number:style='long'/></number:date-style>\
            <style:style style:name='p' style:family='table-cell' style:data-style-name='D'/>\
            <style:style style:name='c' style:family='table-cell' style:parent-style-name='p'/></r>";
        assert_eq!(formats(xml)["c"], "yyyy");
    }
}
