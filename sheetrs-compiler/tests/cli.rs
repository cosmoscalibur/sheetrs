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
        r#"{"inputs": {"principal": 250000, "rate": 0.045, "years": 15, "month": 24}}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v: Json = serde_json::from_str(&out).unwrap();
    let v = v["outputs"].clone();
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
    assert!((v["outputs"]["payment"].as_f64().unwrap() - pmt(200_000.0, 0.05, 30.0)).abs() < 1e-6);
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
    let all = meta["meta"].as_array().unwrap();
    // The properties of the document come first; with `outputs = declared` no name is listed.
    let documents = all
        .iter()
        .take_while(|r| r["meta_type"] == "document")
        .count();
    assert!(documents > 0, "{meta}");
    assert!(all.iter().all(|r| r["meta_type"] != "name"), "{meta}");
    let rows = &all[documents..];
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

    let (code, _, err) = run(&exe, &[], r#"{"inputs": {"bogus": 1}}"#);
    assert_eq!(code, 1);
    assert!(err.contains("unknown input `bogus`"), "{err}");
    // Flat inputs are a mistake with a hint.
    let (code, _, err) = run(&exe, &[], r#"{"principal": 1}"#);
    assert_eq!(code, 1);
    assert!(err.contains("input values go under `inputs`"), "{err}");
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
    let input = "{\"inputs\": {\"principal\": 100000}}\n\n{\"inputs\": {\"nope\": 1}}\n{\"inputs\": {\"principal\": 200000}}\n";
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
        lines[0]["outputs"]["payment"].as_f64().unwrap(),
        lines[2]["outputs"]["payment"].as_f64().unwrap(),
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

    // Date system, iteration and search settings belong to the document: no
    // flag, at build time or run time.
    for flag in [
        ["--date-system", "1904"],
        ["--iterate", "5"],
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
    // There is no dialect to choose: the behavior is LibreOffice's.
    let (code, _, err) = run(&open, &["--dialect", "excel"], "{}");
    assert_eq!(code, 2, "{err}");
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
        r#"{"inputs": {"principal": 1000}}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v: Json = serde_json::from_str(&out).unwrap();
    assert!((v["outputs"]["payment"].as_f64().unwrap() - pmt(1000.0, 0.05, 30.0)).abs() < 1e-9);
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
    // The locales of the table need nothing else; others are defined in the configuration.
    let (_, out, _) = run(
        Path::new(SHEETC),
        &["calc", "=TEXT(1234.5,\"#,##0.00\")", "--locale", "es-CO"],
        "",
    );
    assert_eq!(out.trim(), "\"1.234,50\"");
    for tag in ["sw-KE", "es-BO", "de-CH"] {
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
        r#"{"inputs": {"principal": 1000}}"#,
    );
    assert_eq!(code, 0, "{err}");
    assert!(serde_json::from_str::<Json>(&out).unwrap()["outputs"]["payment"].is_number());
    let _ = json!(null);
}

#[test]
fn outputs_level_decides_what_can_be_queried() {
    let dir = tempfile::tempdir().unwrap();
    let ask = |exe: &Path, r: &str| run(exe, &["--range", r], "");
    let values = |out: &str| serde_json::from_str::<Json>(out).unwrap()["valueRanges"].clone();

    // declared (default): no ranges, the declared outputs; with ranges, exactly
    // those, by output name or by range, and nothing else.
    let exe = build_loan("loan.xlsx", dir.path(), &[]);
    let (code, out, _) = run(&exe, &[], "");
    assert_eq!(code, 0);
    assert!(serde_json::from_str::<Json>(&out).unwrap()["outputs"]["payment"].is_number());
    let (code, out, _) = ask(&exe, "payment");
    assert_eq!(code, 0);
    let v = values(&out);
    assert_eq!(
        (v[0]["name"].clone(), v[0]["range"].clone()),
        (json!("payment"), json!("Loan!B4"))
    );
    assert!(v[0]["values"][0][0].is_number());
    assert_eq!(values(&ask(&exe, "Loan!B4").1)[0]["range"], "Loan!B4");
    let (code, _, err) = ask(&exe, "Loan!B5");
    assert_eq!(code, 1);
    assert!(err.contains("outputs = declared"), "{err}");

    // any: nothing is answered unless asked; ranges in request order, repeats
    // allowed; the settings sheet is not part of it.
    let exe = build_loan("loan.xlsx", dir.path(), &["--outputs", "any"]);
    let (code, _, err) = run(&exe, &[], "{}");
    assert_eq!(code, 1);
    assert!(err.contains("must list `ranges`"), "{err}");
    let (code, out, err) = run(
        &exe,
        &["--range", "Loan!B5", "--range", "Loan!A11:B12"],
        r#"{"inputs": {"principal": 100000}, "ranges": ["payment", "Loan!B5"]}"#,
    );
    assert_eq!(code, 0, "{err}");
    let v = values(&out);
    assert_eq!(
        v.as_array().unwrap().len(),
        4,
        "request ranges, then --range"
    );
    let pay = v[0]["values"][0][0].as_f64().unwrap();
    assert!((pay - pmt(100_000.0, 0.05, 30.0)).abs() < 1e-6);
    assert_eq!(v[1]["range"], "Loan!B5");
    assert_eq!(v[1], v[2]);
    assert_eq!(v[3]["values"][1][0], 2);
    // The same on a serve line, by columns.
    let (_, out, _) = run(
        &exe,
        &["--serve"],
        "{\"ranges\": [\"Loan!A1:A3\"], \"majorDimension\": \"COLUMNS\"}\n",
    );
    let line: Json = serde_json::from_str(out.trim()).unwrap();
    assert_eq!(
        line["valueRanges"][0]["values"],
        json!([["Principal", "Annual rate", "Years"]])
    );
    let (code, _, _) = ask(&exe, "settings!A1");
    assert_eq!(code, 1);

    // named: the declared outputs plus the workbook's defined names (none here).
    let exe = build_loan("loan.xlsx", dir.path(), &["--outputs", "named"]);
    let (code, _, err) = ask(&exe, "Loan!B5");
    assert_eq!(code, 1);
    assert!(err.contains("outputs = named"), "{err}");
    let (code, out, _) = run(&exe, &[], "");
    assert_eq!(code, 0);
    assert!(out.contains("\"outputs\""));

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

    // A locale the table lacks is defined in full in the file; one it has needs only its tag.
    std::fs::write(&other, "[locale]\ntag = \"en-GB\"\n").unwrap();
    let (code, _, err) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 0, "{err}");
    std::fs::write(
        &other,
        "[locale]\ntag = \"sw-KE\"\ndecimal = \".\"\ngroup = \",\"\n",
    )
    .unwrap();
    let (code, _, err) = build(&["--config", other.to_str().unwrap()]);
    assert_eq!(code, 1);
    assert!(err.contains("missing: date_order, months, days"), "{err}");
    std::fs::write(
        &other,
        "[locale]\ntag = \"sw-KE\"\ndecimal = \".\"\ngroup = \",\"\ndate_order = \"dmy\"\n\
         true_text = \"KWELI\"\nfalse_text = \"UONGO\"\ncurrency = \"KSh\"\n\
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

#[test]
fn workbooks_with_macros_need_no_strict() {
    let dir = tempfile::tempdir().unwrap();
    let book = asset("loan_macro.xlsx");
    let out = dir.path().join("m");
    let args = |extra: &[&str]| {
        let mut a = vec![
            "build".to_string(),
            book.to_str().unwrap().into(),
            "--runtime".into(),
            RUNTIME.into(),
            "-o".into(),
            out.to_str().unwrap().into(),
        ];
        a.extend(extra.iter().map(|s| (*s).to_string()));
        a
    };
    let go = |extra: &[&str]| {
        let a = args(extra);
        let refs: Vec<&str> = a.iter().map(String::as_str).collect();
        run(Path::new(SHEETC), &refs, "")
    };
    let (code, _, err) = go(&[]);
    assert_eq!(code, 1);
    assert!(
        err.contains("contains macros") && err.contains("--no-strict"),
        "{err}"
    );
    let (code, _, err) = go(&["--no-strict"]);
    assert_eq!(code, 0, "{err}");
    assert!(
        err.contains("warning: the workbook contains macros"),
        "{err}"
    );
    check_loan(&out);
    // `--strict` undoes a configuration that turned it off.
    let cfg = dir.path().join("c.toml");
    std::fs::write(&cfg, "strict = false\n").unwrap();
    let (code, _, _) = go(&["--config", cfg.to_str().unwrap()]);
    assert_eq!(code, 0);
    let (code, _, _) = go(&["--config", cfg.to_str().unwrap(), "--strict"]);
    assert_eq!(code, 1);
}

#[test]
fn verify_compares_with_the_cached_values() {
    let sheetc = Path::new(SHEETC);
    let book = asset("loan.xlsx");
    let (code, out, err) = run(sheetc, &["verify", book.to_str().unwrap()], "");
    assert_eq!(code, 0, "{out}{err}");
    assert!(out.contains("1086") && out.contains("Loan"), "{out}");
    let (code, _, err) = run(
        sheetc,
        &["verify", book.to_str().unwrap(), "--sheet", "nope"],
        "",
    );
    assert_eq!(code, 1);
    assert!(err.contains("none of the sheets"), "{err}");
    // LibreOffice and Excel disagree on a few cases; the report says which.
    let sem = asset("semantics.xlsx");
    let (code, out, _) = run(sheetc, &["verify", sem.to_str().unwrap(), "--max", "2"], "");
    assert_eq!(code, 1);
    assert!(
        out.contains("differences by kind") && out.contains("more (--max N"),
        "{out}"
    );
}

#[test]
fn inspect_counts_the_functions_a_workbook_uses() {
    let (code, out, err) = run(
        Path::new(SHEETC),
        &[
            "inspect",
            asset("loan.xlsx").to_str().unwrap(),
            "--top",
            "5",
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    assert!(out.contains("functions used ("), "{out}");
    assert!(out.contains("PMT") && out.contains("INDIRECT"), "{out}");
    assert!(out.contains("not implemented: 0 of"), "{out}");
}

#[test]
fn the_executable_says_its_versions_and_how_it_was_built() {
    let dir = tempfile::tempdir().unwrap();
    let exe = build_loan("loan.xlsx", dir.path(), &[]);
    let version = env!("CARGO_PKG_VERSION");
    let (code, out, _) = run(&exe, &["--version"], "");
    assert_eq!(code, 0);
    assert!(
        out.starts_with(&format!(
            "sheetrs-engine {version} (built by sheetrs-compiler {version})"
        )),
        "{out}"
    );
    let (code, out, _) = run(Path::new(SHEETC), &["--version"], "");
    assert_eq!(
        (code, out.trim()),
        (0, format!("sheetc {version}").as_str())
    );

    let (_, out, _) = run(&exe, &["--info"], "");
    let info: Json = serde_json::from_str(&out).unwrap();
    assert_eq!(info["runtime"], format!("sheetrs-engine {version}"));
    let build = &info["build"];
    assert_eq!(build["compiler"], format!("sheetrs-compiler {version}"));
    assert_eq!(build["options"]["outputs"], "declared");
    assert_eq!(build["source"]["file"], "loan.xlsx");
    assert_eq!(build["source"]["format"], "xlsx");
    let hash = build["source"]["sha256"].as_str().unwrap();
    assert!(
        hash.len() == 64 && hash.bytes().all(|b| b.is_ascii_hexdigit()),
        "{hash}"
    );
    assert!(
        build["built_at"].as_str().unwrap().ends_with('Z'),
        "{build}"
    );
}

#[test]
fn source_date_epoch_fixes_the_build_time_and_a_bad_value_stops_the_build() {
    let dir = tempfile::tempdir().unwrap();
    let build = |epoch: &str| {
        let exe = dir.path().join("loan-dated");
        let out = Command::new(SHEETC)
            .env("SOURCE_DATE_EPOCH", epoch)
            .args(["build", asset("loan.xlsx").to_str().unwrap()])
            .args(["--runtime", RUNTIME, "-o", exe.to_str().unwrap()])
            .output()
            .unwrap();
        (
            out.status.code(),
            String::from_utf8_lossy(&out.stderr).into_owned(),
            exe,
        )
    };
    let (code, err, exe) = build("1793361600");
    assert_eq!(code, Some(0), "{err}");
    let (_, out, _) = run(&exe, &["--info"], "");
    let info: Json = serde_json::from_str(&out).unwrap();
    assert_eq!(info["build"]["built_at"], "2026-10-30T12:00:00Z");
    let (code, err, _) = build("yesterday");
    assert_eq!(code, Some(1));
    assert!(err.contains("SOURCE_DATE_EPOCH"), "{err}");
}
