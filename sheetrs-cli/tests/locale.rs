//! The locale of a workbook on the command line: `--locale`, and the `locale` of a configuration
//! file (`--config`, or `sheetrs.toml` of the working directory) between the flag and the default of
//! the format. `locale_es-CO.{xlsx,ods}` were made by a `LibreOffice` with the `es-CO` regional
//! setting; the ODS says it in its default cell style, the XLSX does not.

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

/// `sheetrs query` of `Numbers!C1` (`1234.5` with `#,##0.00`), in a folder with no `sheetrs.toml`
/// unless `folder` has one; the environment has no locale.
fn query(file: &str, args: &[&str], folder: &Path) -> Output {
    Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .current_dir(folder)
        .env_remove("LC_ALL")
        .env_remove("LC_CTYPE")
        .env_remove("LANG")
        .args(["query", "-i"])
        .arg(asset(file))
        .args(["--range", "Numbers!C1", "-p", "display"])
        .args(args)
        .output()
        .unwrap()
}

fn shown(out: &Output) -> String {
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    let answer: serde_json::Value = serde_json::from_slice(&out.stdout).unwrap();
    answer["cells"][0]["display"].as_str().unwrap().to_string()
}

#[test]
fn the_locale_flag_gives_the_locale_of_an_xlsx_and_an_ods_says_its_own() {
    let folder = tempfile::tempdir().unwrap();
    // An XLSX states none and the environment has none: `en-US`.
    assert_eq!(
        shown(&query("locale_es-CO.xlsx", &[], folder.path())),
        "1,234.50"
    );
    assert_eq!(
        shown(&query(
            "locale_es-CO.xlsx",
            &["--locale", "es-CO"],
            folder.path()
        )),
        "1.234,50"
    );
    // An ODS says it: the file wins over the missing system, the flag over the file.
    assert_eq!(
        shown(&query("locale_es-CO.ods", &[], folder.path())),
        "1.234,50"
    );
    assert_eq!(
        shown(&query(
            "locale_es-CO.ods",
            &["--locale", "en-US"],
            folder.path()
        )),
        "1,234.50"
    );
}

#[test]
fn the_locale_of_the_system_is_the_default_of_an_xlsx() {
    let folder = tempfile::tempdir().unwrap();
    let mut command = Command::new(env!("CARGO_BIN_EXE_sheetrs"));
    command
        .current_dir(folder.path())
        .env_remove("LC_ALL")
        .env("LANG", "es_CO.UTF-8")
        .args(["query", "-i"])
        .arg(asset("locale_es-CO.xlsx"))
        .args(["--range", "Numbers!C1", "-p", "display"]);
    assert_eq!(shown(&command.output().unwrap()), "1.234,50");
}

#[test]
fn the_configuration_gives_the_locale_below_the_flag() {
    let folder = tempfile::tempdir().unwrap();
    std::fs::write(
        folder.path().join("sheetrs.toml"),
        "[global]\nlocale = \"es-CO\"\n",
    )
    .unwrap();
    // `sheetrs.toml` of the working directory.
    assert_eq!(
        shown(&query("locale_es-CO.xlsx", &[], folder.path())),
        "1.234,50"
    );
    assert_eq!(
        shown(&query(
            "locale_es-CO.xlsx",
            &["--locale", "en-US"],
            folder.path()
        )),
        "1,234.50"
    );
    // The file named by `--config`, in a folder that has none.
    let other = tempfile::tempdir().unwrap();
    let config = folder.path().join("sheetrs.toml");
    assert_eq!(
        shown(&query(
            "locale_es-CO.xlsx",
            &["--config", config.to_str().unwrap()],
            other.path()
        )),
        "1.234,50"
    );
}

#[test]
fn a_locale_that_is_not_a_tag_stops_the_command() {
    let folder = tempfile::tempdir().unwrap();
    std::fs::write(
        folder.path().join("sheetrs.toml"),
        "[global]\nlocale = \"es CO\"\n",
    )
    .unwrap();
    let out = query("locale_es-CO.xlsx", &[], folder.path());
    assert!(!out.status.success());
    let message = String::from_utf8_lossy(&out.stderr);
    assert!(message.contains("not a locale tag"), "{message}");
    // The linter refuses it too: the file is one.
    let lint = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .current_dir(folder.path())
        .args(["lint", "-i"])
        .arg(asset("locale_es-CO.xlsx"))
        .output()
        .unwrap();
    assert!(String::from_utf8_lossy(&lint.stderr).contains("locale"));
    assert!(!lint.status.success());
}

#[test]
fn the_edit_takes_the_locale_flag_and_the_configuration() {
    // Deleting the first row moves `=TEXT(1234.5,"#,##0.00")` from C11 to C10 and leaves the results
    // stale: the engine calculates them again, with the locale of the command.
    let edit = |file: &str, args: &[&str], dir: &Path| {
        let out = dir
            .join("out")
            .with_extension(file.rsplit('.').next().unwrap());
        let output = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
            .current_dir(dir)
            .env_remove("LC_ALL")
            .env_remove("LC_CTYPE")
            .env_remove("LANG")
            .args(["edit", "-i"])
            .arg(asset(file))
            .args(["--delete-rows", "Numbers!1", "--output"])
            .arg(&out)
            .args(args)
            .output()
            .unwrap();
        assert!(
            output.status.success(),
            "{}",
            String::from_utf8_lossy(&output.stderr)
        );
        let read = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
            .current_dir(dir)
            .args(["query", "-i"])
            .arg(&out)
            .args(["--range", "Numbers!C10", "-p", "value", "--locale", "en-US"])
            .output()
            .unwrap();
        let answer: serde_json::Value = serde_json::from_slice(&read.stdout).unwrap();
        answer["cells"][0]["value"].as_str().unwrap().to_string()
    };
    let folder = tempfile::tempdir().unwrap();
    for file in ["locale_es-CO.xlsx", "locale_es-CO.ods"] {
        let given = edit(file, &["--locale", "es-CO"], folder.path());
        assert_eq!(given, "1.234,50", "{file} --locale es-CO");
        let given = edit(file, &["--locale", "en-US"], folder.path());
        assert_eq!(given, "1,234.50", "{file} --locale en-US");
    }
    // The ODS says es-CO and needs no flag; an XLSX with no flag has the (missing) system's.
    assert_eq!(edit("locale_es-CO.ods", &[], folder.path()), "1.234,50");
    assert_eq!(edit("locale_es-CO.xlsx", &[], folder.path()), "1,234.50");
    // The configuration is the default below the flag.
    let config = folder.path().join("sheetrs.toml");
    std::fs::write(&config, "[global]\nlocale = \"es-CO\"\n").unwrap();
    assert_eq!(edit("locale_es-CO.xlsx", &[], folder.path()), "1.234,50");
    assert_eq!(
        edit("locale_es-CO.xlsx", &["--locale", "en-US"], folder.path()),
        "1,234.50"
    );
    // A flag that is not a tag is refused.
    let refused = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .current_dir(folder.path())
        .args(["edit", "-i"])
        .arg(asset("locale_es-CO.xlsx"))
        .args([
            "--delete-rows",
            "Numbers!1",
            "--dry-run",
            "--locale",
            "not a tag",
        ])
        .output()
        .unwrap();
    assert!(!refused.status.success());
    assert!(String::from_utf8_lossy(&refused.stderr).contains("not a locale tag"));
}
