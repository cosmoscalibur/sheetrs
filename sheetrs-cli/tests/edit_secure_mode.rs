//! `sheetrs edit --secure-mode`: a workbook with macros is refused by the build that the edit runs, by default, and
//! opened in secure mode, where what calls a macro is an error cell and the report warns (the workbooks are
//! `tests/assets/from_cache` of `sheetrs-edit`: `Calc!A1` calls a macro in `macro.*`).

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn convert_calc(extension: &str, flags: &[&str]) -> Output {
    let dir = tempfile::tempdir().unwrap();
    let book = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/from_cache")
        .join(format!("macro.{extension}"));
    let out: PathBuf = dir.path().join(format!("out.{extension}"));
    Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("edit")
        .arg("-i")
        .arg(book)
        .args(["--calculate-all", "--formula-to-value", "Calc", "-o"])
        .arg(out)
        .args(flags)
        .output()
        .unwrap()
}

#[test]
fn a_workbook_with_macros_is_refused_unless_secure_mode_is_on() {
    for extension in ["xlsx", "ods"] {
        let refused = convert_calc(extension, &[]);
        assert!(!refused.status.success(), "{extension}");
        let stderr = String::from_utf8_lossy(&refused.stderr);
        assert!(
            stderr.contains("contains macros") && stderr.contains("secure mode"),
            "{extension}: {stderr}"
        );
        let secure = convert_calc(extension, &["--secure-mode"]);
        let stderr = String::from_utf8_lossy(&secure.stderr);
        assert!(secure.status.success(), "{extension}: {stderr}");
        assert!(stderr.contains("secure mode"), "{extension}: {stderr}");
    }
}
