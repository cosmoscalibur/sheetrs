//! End-to-end: compile real XLSX/ODS files into executables and run them.

use serde_json::{Value as Json, json};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");
// The compiler binary is also a complete runtime: it runs an embedded program.
const RUNTIME: &str = SHEETC;

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(name)
}

fn run(cmd: &Path, args: &[&str], stdin: &str) -> (i32, String, String) {
    let mut child = Command::new(cmd)
        .args(args)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(stdin.as_bytes())
        .unwrap();
    let out = child.wait_with_output().unwrap();
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

fn build_loan(file: &str, dir: &Path, extra: &[&str]) -> PathBuf {
    let exe = dir.join(format!("loan-{file}"));
    let mut args = vec![
        "build",
        asset(file).to_str().unwrap().to_owned().leak(),
        "--runtime",
        RUNTIME,
        "-o",
        exe.to_str().unwrap().to_owned().leak(),
    ];
    args.extend_from_slice(extra);
    let (code, _, err) = run(Path::new(SHEETC), &args, "");
    assert_eq!(code, 0, "build failed: {err}");
    exe
}

fn pmt(p: f64, annual: f64, years: f64) -> f64 {
    let r = annual / 12.0;
    let n = years * 12.0;
    p * r / (1.0 - (1.0 + r).powf(-n))
}

fn check_loan(exe: &Path) {
    let (code, out, err) = run(
        exe,
        &[],
        r#"{"principal": 250000, "rate": 0.045, "years": 15, "month": 24}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v: Json = serde_json::from_str(&out).unwrap();
    let want = pmt(250_000.0, 0.045, 15.0);
    assert!((v["payment"].as_f64().unwrap() - want).abs() < 1e-6, "{v}");
    assert!((v["interest"].as_f64().unwrap() - (want * 180.0 - 250_000.0)).abs() < 1e-4);
    // Balance after 24 payments (dynamic INDIRECT in the source workbook).
    let r: f64 = 0.045 / 12.0;
    let bal = 250_000.0 * (1.0 + r).powi(24) - want * (((1.0 + r).powi(24) - 1.0) / r);
    assert!((v["balance"].as_f64().unwrap() - bal).abs() < 1e-4, "{v}");

    // Defaults come from the spreadsheet when an input is omitted.
    let (code, out, _) = run(exe, &[], "{}");
    assert_eq!(code, 0);
    let v: Json = serde_json::from_str(&out).unwrap();
    assert!((v["payment"].as_f64().unwrap() - pmt(200_000.0, 0.05, 30.0)).abs() < 1e-6);
}

#[test]
fn xlsx_and_ods_compile_to_equivalent_executables() {
    let dir = tempfile::tempdir().unwrap();
    for f in ["loan.xlsx", "loan.ods"] {
        let exe = build_loan(f, dir.path(), &[]);
        check_loan(&exe);
    }
}

#[test]
fn executable_is_self_describing_and_strict() {
    let dir = tempfile::tempdir().unwrap();
    let exe = build_loan("loan.xlsx", dir.path(), &[]);
    let (code, out, _) = run(&exe, &["--info"], "");
    assert_eq!(code, 0);
    let info: Json = serde_json::from_str(&out).unwrap();
    assert_eq!(info["inputs"].as_array().unwrap().len(), 4);
    assert_eq!(info["outputs"][0]["name"], "payment");

    // The settings sheet's rows, inputs and outputs included.
    let (code, out, _) = run(&exe, &["--meta"], "");
    assert_eq!(code, 0);
    let meta: Json = serde_json::from_str(&out).unwrap();
    assert_eq!(meta["outputs"], "declared");
    let rows = meta["meta"].as_array().unwrap();
    assert_eq!(rows.len(), 8);
    assert_eq!(rows[0]["key"], "principal");
    assert_eq!(rows[0]["value"], "Loan!B1");
    assert_eq!(rows[0]["meta_type"], "input");
    assert_eq!(rows[6]["key"], "balance");
    assert_eq!(rows[6]["description"], "Balance at that month");
    assert_eq!(
        rows[7],
        json!({"key": "version", "value": "1.0", "description": "Model version", "meta_type": "info"})
    );

    let (code, _, err) = run(&exe, &[], r#"{"bogus": 1}"#);
    assert_eq!(code, 1);
    assert!(err.contains("unknown input `bogus`"), "{err}");
    let (code, _, err) = run(&exe, &["--wat"], "{}");
    assert_eq!(code, 2, "{err}");
    let (code, _, err) = run(&exe, &[], "{not json");
    assert_eq!(code, 1);
    assert!(err.contains("invalid JSON"), "{err}");
}

#[test]
fn serve_mode_answers_line_by_line() {
    let dir = tempfile::tempdir().unwrap();
    let exe = build_loan("loan.xlsx", dir.path(), &[]);
    let input = "{\"principal\": 100000}\n\n{\"nope\": 1}\n{\"principal\": 200000}\n";
    let (code, out, _) = run(&exe, &["--serve"], input);
    assert_eq!(code, 0);
    let lines: Vec<Json> = out
        .lines()
        .map(|l| serde_json::from_str(l).unwrap())
        .collect();
    assert_eq!(lines.len(), 3);
    assert!(
        lines[1]["error_message"]
            .as_str()
            .unwrap()
            .contains("unknown input")
    );
    let (a, b) = (
        lines[0]["payment"].as_f64().unwrap(),
        lines[2]["payment"].as_f64().unwrap(),
    );
    assert!((b / a - 2.0).abs() < 1e-9);
}

#[test]
fn run_time_options_are_only_what_a_file_cannot_say() {
    let dir = tempfile::tempdir().unwrap();
    // The locale is fixed when the program is built, with or without --no-fold.
    for extra in [&[][..], &["--no-fold"][..]] {
        let fixed = build_loan("loan.xlsx", dir.path(), extra);
        let (code, _, err) = run(&fixed, &["--locale", "es"], "{}");
        assert_eq!(code, 2, "{err}");
        assert!(err.contains("fixed when the program is built"), "{err}");
    }
    let open = build_loan("loan.ods", dir.path(), &[]);

    // Date system, iteration, search settings and dialect belong to the
    // document: no flag, at build time or run time.
    for flag in [
        ["--date-system", "1904"],
        ["--iterate", "5"],
        ["--dialect", "ods"],
        ["--search", "regex"],
    ] {
        let (code, _, err) = run(&open, &flag, "{}");
        assert_eq!(code, 2, "{flag:?}: {err}");
        assert!(err.contains("defined by the document"), "{err}");
        let (code, _, err) = run(
            Path::new(SHEETC),
            &[
                "build",
                asset("loan.xlsx").to_str().unwrap(),
                flag[0],
                flag[1],
            ],
            "",
        );
        assert_eq!(code, 2, "{flag:?}: {err}");
    }
    // `calc` has no document, so it accepts them.
    let (code, out, _) = run(
        Path::new(SHEETC),
        &["calc", "=DATE(1900,1,1)", "--date-system", "1899"],
        "",
    );
    assert_eq!((code, out.trim()), (0, "2"));
}

#[test]
fn run_command_compiles_in_memory() {
    let (code, out, err) = run(
        Path::new(SHEETC),
        &["run", asset("loan.xlsx").to_str().unwrap()],
        r#"{"principal": 1000}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v: Json = serde_json::from_str(&out).unwrap();
    assert!((v["payment"].as_f64().unwrap() - pmt(1000.0, 0.05, 30.0)).abs() < 1e-9);
}

#[test]
fn calc_command() {
    let (code, out, _) = run(Path::new(SHEETC), &["calc", "=SUM(1,2)*3&\"x\""], "");
    assert_eq!(code, 0);
    assert_eq!(out.trim(), "\"9x\"");
    let (_, out, _) = run(
        Path::new(SHEETC),
        &["calc", "=TEXT(1234.5,\"#,##0.00\")", "--locale", "es"],
        "",
    );
    assert_eq!(out.trim(), "\"1.234,50\"");
    // Only en and es are built in; other locales are defined in the configuration.
    for tag in ["de-CH", "es-MX"] {
        let (code, _, err) = run(Path::new(SHEETC), &["calc", "=1", "--locale", tag], "");
        assert_eq!(code, 2, "{err}");
        assert!(err.contains("[locale]"), "{err}");
    }
    let (code, _, err) = run(
        Path::new(SHEETC),
        &["calc", "=1", "--locale", "not a tag"],
        "",
    );
    assert_eq!(code, 2, "{err}");
    let (_, out, _) = run(Path::new(SHEETC), &["calc", "=1/0"], "");
    assert_eq!(out.trim(), "{\"error\":\"#DIV/0!\"}");
}

#[test]
fn bundle_roundtrip_json_matches() {
    let dir = tempfile::tempdir().unwrap();
    let ir = dir.path().join("loan.json");
    let (code, _, err) = run(
        Path::new(SHEETC),
        &[
            "build",
            asset("loan.xlsx").to_str().unwrap(),
            "--runtime",
            RUNTIME,
            "--emit-ir",
            ir.to_str().unwrap(),
            "-o",
            dir.path().join("m").to_str().unwrap(),
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    // The emitted IR runs through `run` without the original file.
    let (code, out, err) = run(
        Path::new(SHEETC),
        &["run", ir.to_str().unwrap()],
        r#"{"principal": 1000}"#,
    );
    assert_eq!(code, 0, "{err}");
    assert!(serde_json::from_str::<Json>(&out).unwrap()["payment"].is_number());
    let _ = json!(null);
}

#[test]
fn outputs_level_decides_what_can_be_queried() {
    let dir = tempfile::tempdir().unwrap();
    let ask = |exe: &Path, q: &str| run(exe, &["--query", q], "{}");

    // declared (default): the outputs only.
    let exe = build_loan("loan.xlsx", dir.path(), &[]);
    let (code, out, _) = ask(&exe, "Loan!B4");
    assert_eq!(code, 0);
    assert!(serde_json::from_str::<Json>(&out).unwrap()["Loan!B4"].is_number());
    let (code, _, err) = ask(&exe, "Loan!B5");
    assert_eq!(code, 1);
    assert!(err.contains("outputs = declared"), "{err}");

    // any: every range of the workbook, and the settings sheet is not part of it.
    let exe = build_loan("loan.xlsx", dir.path(), &["--outputs", "any"]);
    let (code, out, err) = run(
        &exe,
        &["--query", "Loan!B5", "--query", "Loan!A11:B12"],
        r#"{"principal": 100000}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v: Json = serde_json::from_str(&out).unwrap();
    assert!(v["Loan!B5"].is_number());
    assert_eq!(v["Loan!A11:B12"][1][0], 2);
    assert!((v["payment"].as_f64().unwrap() - pmt(100_000.0, 0.05, 30.0)).abs() < 1e-6);
    // The same through `$query` in the JSON line.
    let (_, out, _) = run(&exe, &["--serve"], "{\"$query\": [\"Loan!B3\"]}\n");
    assert_eq!(serde_json::from_str::<Json>(&out).unwrap()["Loan!B3"], 30);
    let (code, _, _) = ask(&exe, "settings!A1");
    assert_eq!(code, 1);

    // named: only the workbook's defined names (none in this workbook).
    let exe = build_loan("loan.xlsx", dir.path(), &["--outputs", "named"]);
    let (code, _, err) = ask(&exe, "Loan!B5");
    assert_eq!(code, 1);
    assert!(err.contains("outputs = named"), "{err}");

    let (code, _, err) = run(
        Path::new(SHEETC),
        &[
            "build",
            asset("loan.xlsx").to_str().unwrap(),
            "--outputs",
            "all",
        ],
        "",
    );
    assert_eq!(code, 1);
    assert!(err.contains("declared, named or any"), "{err}");
}

#[test]
fn configuration_file_and_command_line_layering() {
    let dir = tempfile::tempdir().unwrap();
    let book = dir.path().join("loan.ods");
    std::fs::copy(asset("loan.ods"), &book).unwrap();
    let exe = dir.path().join("m");
    let build = |extra: &[&str]| {
        let mut args = vec![
            "build",
            book.to_str().unwrap(),
            "--runtime",
            RUNTIME,
            "-o",
            exe.to_str().unwrap(),
        ];
        args.extend_from_slice(extra);
        run(Path::new(SHEETC), &args, "")
    };
    let outputs_of = |exe: &Path| {
        let (_, out, _) = run(exe, &["--meta"], "");
        serde_json::from_str::<Json>(&out).unwrap()["outputs"].clone()
    };

    // `sheetc.toml` next to the workbook is picked up.
    std::fs::write(dir.path().join("sheetc.toml"), "outputs = \"any\"\n").unwrap();
    let (code, _, err) = build(&[]);
    assert_eq!(code, 0, "{err}");
    assert_eq!(outputs_of(&exe), "any");
    // The command line wins.
    let (code, _, err) = build(&["--outputs", "declared"]);
    assert_eq!(code, 0, "{err}");
    assert_eq!(outputs_of(&exe), "declared");
    // An explicit file replaces the discovered one; unknown keys are errors.
    let other = dir.path().join("other.toml");
    std::fs::write(&other, "outputs = \"named\"\n").unwrap();
    let (code, _, _) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 0);
    assert_eq!(outputs_of(&exe), "named");
    std::fs::write(&other, "iterations = 5\n").unwrap();
    let (code, _, err) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 1);
    assert!(err.contains("iterations"), "{err}");

    // A locale beyond en/es is defined in full in the file.
    std::fs::write(
        &other,
        "[locale]\ntag = \"de-DE\"\ndecimal = \",\"\ngroup = \".\"\n",
    )
    .unwrap();
    let (code, _, err) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 1);
    assert!(err.contains("missing: date_order, months, days"), "{err}");
    std::fs::write(
        &other,
        "[locale]\ntag = \"de-DE\"\ndecimal = \",\"\ngroup = \".\"\ndate_order = \"dmy\"\n\
         true_text = \"WAHR\"\nfalse_text = \"FALSCH\"\ncurrency = \"€\"\n\
         months = [\"Jan\",\"Feb\",\"Mär\",\"Apr\",\"Mai\",\"Jun\",\"Jul\",\"Aug\",\"Sep\",\"Okt\",\"Nov\",\"Dez\"]\n\
         days = [\"So\",\"Mo\",\"Di\",\"Mi\",\"Do\",\"Fr\",\"Sa\"]\n",
    )
    .unwrap();
    let (code, _, err) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 0, "{err}");

    // A settings sheet that does not exist is an error once named explicitly.
    let (code, _, err) = build(&["--settings", "nowhere"]);
    assert_eq!(code, 1);
    assert!(err.contains("settings sheet `nowhere` not found"), "{err}");
}
