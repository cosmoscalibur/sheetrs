use serde_json::{Value as Json, json};
use sheetrs_compiler::{BindingSpec, Model, Overrides, Program, Report, Spec, Value, compile};

pub fn spec(ins: &[&str], outs: &[&str]) -> Spec {
    Spec {
        inputs: ins.iter().map(|s| BindingSpec::parse(s)).collect(),
        outputs: outs.iter().map(|s| BindingSpec::parse(s)).collect(),
        ..Spec::default()
    }
}

pub fn build(m: &Model, s: &Spec) -> (Program, Report) {
    compile(m, s).unwrap_or_else(|e| panic!("compile failed: {e}"))
}

/// Evaluate one formula in an empty workbook.
pub fn calc(formula: &str) -> Json {
    let mut m = Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", formula);
    let (p, _) = build(&m, &spec(&[], &["A1"]));
    p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!A1"].clone()
}

/// A workbook with `Sheet1` filled from `(address, value)` pairs, where text
/// starting with `=` is a formula.
pub fn sheet(cells: &[(&str, Json)]) -> Model {
    let mut m = Model::new();
    let s = m.add_sheet("Sheet1");
    for (a, v) in cells {
        match v {
            Json::String(t) if t.starts_with('=') => {
                m.set_formula(s, a, t);
            }
            Json::String(t) => {
                m.set(s, a, Value::text(t));
            }
            Json::Number(n) => {
                m.set(s, a, Value::Number(n.as_f64().unwrap()));
            }
            Json::Bool(b) => {
                m.set(s, a, Value::Bool(*b));
            }
            _ => panic!("unsupported fixture value"),
        }
    }
    m
}

pub fn err(code: &str) -> Json {
    json!({ "error": code })
}

fn near(a: f64, b: f64) -> bool {
    (a - b).abs() <= 1e-9 * a.abs().max(b.abs()).max(1.0)
}

/// Relative difference allowed between the parts of two complex numbers written as text, which
/// carry 15 significant digits: a few units of the last one. The `sin`, `cos`, `sinh` and `cosh`
/// of the system's C library round differently from one version to the next, and the oracle holds
/// what the library of the machine that computed it gave (see `FIDELITY.md`, *C library*).
const COMPLEX_TEXT_TOLERANCE: f64 = 1e-13;

/// The real and imaginary part of a complex number as the `IM` functions write it (`1.5-2i`,
/// `-3j`, `i`), and its suffix.
fn complex_parts(text: &str) -> Option<(f64, f64, char)> {
    let suffix = text.chars().last().filter(|c| matches!(c, 'i' | 'j'))?;
    let body = &text[..text.len() - 1];
    let sign_at = body
        .char_indices()
        .rev()
        .find(|&(i, c)| {
            (c == '+' || c == '-') && i > 0 && !matches!(body.as_bytes()[i - 1], b'e' | b'E')
        })
        .map(|(i, _)| i);
    let (re, im) = sign_at.map_or(("0", body), |i| body.split_at(i));
    let im = match im {
        "" | "+" => "1",
        "-" => "-1",
        other => other,
    };
    Some((re.parse().ok()?, im.parse().ok()?, suffix))
}

fn complex_text_near(expected: &str, got: &str) -> bool {
    let (Some((re_a, im_a, suffix_a)), Some((re_b, im_b, suffix_b))) =
        (complex_parts(expected), complex_parts(got))
    else {
        return false;
    };
    let close = |a: f64, b: f64| (a - b).abs() <= COMPLEX_TEXT_TOLERANCE * a.abs().max(b.abs());
    suffix_a == suffix_b && close(re_a, re_b) && close(im_a, im_b)
}

pub fn matches_oracle(expected: &Value, got: &Json) -> bool {
    match (expected, got) {
        (Value::Number(a), Json::Number(b)) => near(*a, b.as_f64().unwrap()),
        (Value::Text(a), Json::String(b)) => &**a == b || complex_text_near(a, b),
        (Value::Bool(a), Json::Bool(b)) => a == b,
        (Value::Empty, Json::Null) => true,
        // A circular reference is `Err:522` in LibreOffice, which files store as `#VALUE!`.
        (Value::Error(a), Json::Object(o)) => {
            let got = o.get("error").and_then(Json::as_str);
            got == Some(a.as_str()) || (got == Some("#CIRC!") && a.as_str() == "#VALUE!")
        }
        // A formula returning an empty string reads back as empty in ODS.
        (Value::Empty, Json::String(s)) => s.is_empty(),
        _ => false,
    }
}
