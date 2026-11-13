//! The number format `LibreOffice` 26.8 gives a cell in `General` that is typed a text
//! (`Sheet.replaceAll` into a cell, probed with `sheetrs-edit/tests/assets/replace/make_fixtures.py`
//! and its `typed_*` cases), and what the input line reads the text as.

#![allow(clippy::unreadable_literal)] // the numbers are those LibreOffice reports

use sheetrs_formula::DateSystem;
use sheetrs_query::text::{EntryFormat, TextContext, Typed};

#[derive(Debug)]
enum Kind {
    Text,
    Number(f64),
    /// A number that depends on the current year (a date without a year).
    AnyNumber,
    Boolean(bool),
}

#[test]
fn what_is_typed_into_a_general_cell_takes_the_format_libreoffice_gives_it() {
    let text = TextContext::new(None, DateSystem::default()).unwrap();
    let table: &[(&str, Option<EntryFormat>, Kind)] = &[
        ("50%", Some(EntryFormat::Percent), Kind::Number(0.5)),
        ("03/15/2024", Some(EntryFormat::Date), Kind::Number(45366.0)),
        (
            "2024-03-15",
            Some(EntryFormat::DateIso),
            Kind::Number(45366.0),
        ),
        (
            "12:30",
            Some(EntryFormat::Time),
            Kind::Number(0.5208333333333334),
        ),
        (
            "12:30:45",
            Some(EntryFormat::Time),
            Kind::Number(0.5213541666666667),
        ),
        ("true", Some(EntryFormat::Boolean), Kind::Boolean(true)),
        ("TRUE", Some(EntryFormat::Boolean), Kind::Boolean(true)),
        ("1E3", Some(EntryFormat::Scientific), Kind::Number(1000.0)),
        ("1/2", Some(EntryFormat::Date), Kind::AnyNumber),
        ("$12.50", Some(EntryFormat::Currency), Kind::Number(12.5)),
        ("12.50 USD", None, Kind::Text),
        ("12", None, Kind::Number(12.0)),
        ("12.5", None, Kind::Number(12.5)),
        ("1,234.5", None, Kind::Number(1234.5)),
        ("-5", None, Kind::Number(-5.0)),
        ("(5)", None, Kind::Number(-5.0)),
        ("Mar 5", Some(EntryFormat::Date), Kind::AnyNumber),
        ("15 Mar 2024", None, Kind::Text),
        ("1.5%", Some(EntryFormat::Percent), Kind::Number(0.015)),
        ("0.5", None, Kind::Number(0.5)),
        ("00123", None, Kind::Number(123.0)),
        ("abc", None, Kind::Text),
        ("1e-3", Some(EntryFormat::Scientific), Kind::Number(0.001)),
        ("1/2/2024", Some(EntryFormat::Date), Kind::Number(45293.0)),
        (
            "2024-03-15 10:20",
            Some(EntryFormat::DateTimeIso),
            Kind::Number(45366.430555555555),
        ),
        (
            "10:20 PM",
            Some(EntryFormat::Time),
            Kind::Number(0.9305555555555556),
        ),
        ("€5", None, Kind::Text),
        ("5 EUR", None, Kind::Text),
        ("1.2.2024", None, Kind::Text),
        ("0x10", None, Kind::Text),
        ("1,5", None, Kind::Text),
        (".5", None, Kind::Number(0.5)),
        ("+5", None, Kind::Number(5.0)),
        ("1 1/2", Some(EntryFormat::Fraction), Kind::Number(1.5)),
        ("50 %", Some(EntryFormat::Percent), Kind::Number(0.5)),
        (
            "12-Mar-2024",
            Some(EntryFormat::Date),
            Kind::Number(45363.0),
        ),
        (
            "March 15, 2024",
            Some(EntryFormat::Date),
            Kind::Number(45366.0),
        ),
        ("2024/03/15", None, Kind::Text),
        ("31/12/2024", None, Kind::Text),
        ("13/13/2013", None, Kind::Text),
        (
            "25:00",
            Some(EntryFormat::TimeTotal),
            Kind::Number(1.0416666666666667),
        ),
        (
            "1:2",
            Some(EntryFormat::Time),
            Kind::Number(0.043055555555555555),
        ),
        ("100%", Some(EntryFormat::Percent), Kind::Number(1.0)),
        ("1e3", Some(EntryFormat::Scientific), Kind::Number(1000.0)),
        ("True", Some(EntryFormat::Boolean), Kind::Boolean(true)),
        ("false", Some(EntryFormat::Boolean), Kind::Boolean(false)),
        ("#N/A", None, Kind::Text),
        ("#DIV/0!", None, Kind::Text),
        ("1,234", None, Kind::Number(1234.0)),
        ("12,5", None, Kind::Text),
        ("3.", None, Kind::Number(3.0)),
        ("5-3", None, Kind::Text),
        ("5/3", Some(EntryFormat::Date), Kind::AnyNumber),
        (
            "03/15/2024 10:20",
            Some(EntryFormat::DateTime),
            Kind::Number(45366.430555555555),
        ),
        (
            "03/15/2024 10:20:30",
            Some(EntryFormat::DateTime),
            Kind::Number(45366.43090277778),
        ),
        (
            "03/15/2024 10:20 PM",
            Some(EntryFormat::DateTime),
            Kind::Number(45366.930555555555),
        ),
        (
            "2024-03-15 10:20:30",
            Some(EntryFormat::DateTimeIso),
            Kind::Number(45366.43090277778),
        ),
        (
            "2024-03-15T10:20",
            Some(EntryFormat::DateTimeIsoT),
            Kind::Number(45366.430555555555),
        ),
        ("Mar 5 2024", Some(EntryFormat::Date), Kind::Number(45356.0)),
        ("5 Mar", None, Kind::Text),
        (
            "12:30:45.5",
            Some(EntryFormat::TimeTotalFraction),
            Kind::Number(0.5213599537037037),
        ),
        ("0:00", Some(EntryFormat::Time), Kind::Number(0.0)),
        ("1/2/24", Some(EntryFormat::Date), Kind::Number(45293.0)),
        ("-50%", Some(EntryFormat::Percent), Kind::Number(-0.5)),
        ("+5%", Some(EntryFormat::Percent), Kind::Number(0.05)),
        ("5E2", Some(EntryFormat::Scientific), Kind::Number(500.0)),
        (
            "1.2e+3",
            Some(EntryFormat::Scientific),
            Kind::Number(1200.0),
        ),
        (
            "$1,234.50",
            Some(EntryFormat::Currency),
            Kind::Number(1234.5),
        ),
        ("($5)", Some(EntryFormat::Currency), Kind::Number(-5.0)),
        ("-$5", Some(EntryFormat::Currency), Kind::Number(-5.0)),
        ("5$", Some(EntryFormat::Currency), Kind::Number(5.0)),
        ("$ 5", Some(EntryFormat::Currency), Kind::Number(5.0)),
        ("0 1/2", Some(EntryFormat::Fraction), Kind::Number(0.5)),
        ("-1 1/2", Some(EntryFormat::Fraction), Kind::Number(-1.5)),
        ("2 3/4", Some(EntryFormat::Fraction), Kind::Number(2.75)),
        ("12 1/8", Some(EntryFormat::Fraction), Kind::Number(12.125)),
        ("1,234.5%", Some(EntryFormat::Percent), Kind::Number(12.345)),
        (
            "12:30 am",
            Some(EntryFormat::Time),
            Kind::Number(0.020833333333333332),
        ),
        ("24:00", Some(EntryFormat::TimeTotal), Kind::Number(1.0)),
        (
            "48:30",
            Some(EntryFormat::TimeTotal),
            Kind::Number(2.0208333333333335),
        ),
        (
            "1:30:15 PM",
            Some(EntryFormat::Time),
            Kind::Number(0.5626736111111111),
        ),
        ("March 15", Some(EntryFormat::Date), Kind::AnyNumber),
        (
            "Mar 15 2024",
            Some(EntryFormat::Date),
            Kind::Number(45366.0),
        ),
        ("15-Mar-24", Some(EntryFormat::Date), Kind::Number(45366.0)),
        (
            "1/2/2024 1:00",
            Some(EntryFormat::DateTime),
            Kind::Number(45293.041666666664),
        ),
        ("1E", None, Kind::Text),
        ("e5", None, Kind::Text),
        ("5e", None, Kind::Text),
        ("-1E3", Some(EntryFormat::Scientific), Kind::Number(-1000.0)),
        ("1e+3", Some(EntryFormat::Scientific), Kind::Number(1000.0)),
        ("1e03", Some(EntryFormat::Scientific), Kind::Number(1000.0)),
        (
            "12.5E-2",
            Some(EntryFormat::Scientific),
            Kind::Number(0.125),
        ),
        ("$1e3", None, Kind::Text),
        ("10%%", None, Kind::Text),
        ("5 1/2%", None, Kind::Text),
        ("Jan-24", Some(EntryFormat::Date), Kind::AnyNumber),
        ("Jan 2024", Some(EntryFormat::Date), Kind::Number(45292.0)),
        ("1st", None, Kind::Text),
        ("TRUE/FALSE", None, Kind::Text),
    ];
    let mut wrong = Vec::new();
    for (input, format, kind) in table {
        let typed = text.typed(input, false);
        let reads = match (&typed, kind) {
            (Typed::Text(_), Kind::Text) | (Typed::Number(_), Kind::AnyNumber) => true,
            (Typed::Number(x), Kind::Number(y)) => (x - y).abs() < 1e-6,
            (Typed::Boolean(a), Kind::Boolean(b)) => a == b,
            _ => false,
        };
        let got = text.entry_format(input);
        if !reads || got != *format {
            wrong.push(format!(
                "{input:?}: read {typed:?}, format {got:?}; LibreOffice {kind:?}, {format:?}"
            ));
        }
    }
    assert!(wrong.is_empty(), "{}", wrong.join("\n"));
}
