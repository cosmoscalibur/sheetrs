//! `INFO` and `CELL("filename")` say what the build fixes, or else the machine and the program.

use sheetrs_compiler::{Model, Spec, Value, compile};
use sheetrs_engine::options::Host;
use sheetrs_engine::run::Overrides;

fn value(formula: &str, ods: bool, host: Host, run: Host) -> serde_json::Value {
    let mut m = Model::new();
    let s = m.add_sheet("Data");
    m.set_formula(s, "A1", formula);
    m.file_name = Some("book.xlsx".into());
    if ods {
        m.from_file = true;
        m.settings.string_ref_excel = Some(false);
    }
    let spec = Spec {
        outputs: vec![sheetrs_compiler::BindingSpec::parse("Data!A1")],
        host,
        ..Spec::default()
    };
    let (program, _) = compile(&m, &spec).unwrap();
    let o = Overrides {
        host: run,
        ..Overrides::default()
    };
    let out = program.execute(&serde_json::json!({}), &o).unwrap();
    out["Data!A1"].clone()
}

fn fixed() -> Host {
    Host {
        system: Some("WNT".into()),
        osversion: Some("Windows (32-bit) NT 5.01".into()),
        release: Some("abc123".into()),
        directory: Some("/opt/my app".into()),
        filename: None,
    }
}

#[test]
fn info_says_what_the_build_fixes() {
    let h = fixed();
    let v = |f| value(f, false, h.clone(), Host::default());
    assert_eq!(v("INFO(\"system\")"), "WNT");
    assert_eq!(v("INFO(\"osversion\")"), "Windows (32-bit) NT 5.01");
    assert_eq!(v("INFO(\"release\")"), "abc123");
    assert_eq!(
        v("INFO(\"directory\")"),
        serde_json::json!({"error": "#N/A"})
    );
    // A run fixes what the build left open, and wins.
    let run = Host {
        system: Some("MACOSX".into()),
        ..Host::default()
    };
    assert_eq!(value("INFO(\"system\")", false, h, run), "MACOSX");
}

#[test]
fn info_without_a_fixed_value_is_the_machine() {
    let v = value("INFO(\"system\")", false, Host::default(), Host::default());
    let want = match std::env::consts::OS {
        "windows" => "WNT",
        "macos" => "MACOSX",
        _ => "LINUX",
    };
    assert_eq!(v, want);
}

#[test]
fn filename_is_the_directory_the_file_name_and_the_sheet_in_the_syntax_of_the_document() {
    let h = fixed();
    assert_eq!(
        value("CELL(\"filename\",A1)", false, h.clone(), Host::default()),
        "file:///opt/my%20app/[book.xlsx]Data"
    );
    assert_eq!(
        value("CELL(\"filename\",A1)", true, h.clone(), Host::default()),
        "'file:///opt/my%20app/book.xlsx'#$Data"
    );
    // Without a directory it is the one of the program.
    let dir = std::env::current_exe().unwrap();
    let dir = dir.parent().unwrap().to_string_lossy().replace(' ', "%20");
    let v = value(
        "CELL(\"filename\",A1)",
        false,
        Host::default(),
        Host::default(),
    );
    assert_eq!(v, format!("file://{dir}/[book.xlsx]Data"));
    let _ = Value::Empty;
}
