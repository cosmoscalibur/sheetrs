//! A full recalculation by LibreOffice, to compare the engine with: the values a
//! file caches may come from another version or from a calculation that never ran.

use std::path::{Path, PathBuf};
use std::process::Command;

/// Convert `file` to ODS with LibreOffice recalculating everything on load, under the
/// regional settings `locale` (a BCP-47 tag such as `es-CO`). The binary is `$SOFFICE`
/// or `soffice`. Returns the path of the ODS, inside `dir`.
pub fn recalculate(file: &Path, locale: &str, dir: &Path) -> Result<PathBuf, String> {
    let profile = dir.join("profile");
    std::fs::create_dir_all(profile.join("user")).map_err(|e| e.to_string())?;
    let setting = |path: &str, name: &str, value: &str| {
        format!(
            "<item oor:path=\"{path}\"><prop oor:name=\"{name}\" oor:op=\"fuse\"><value>{value}</value></prop></item>"
        )
    };
    let xcu = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><oor:items xmlns:oor=\"http://openoffice.org/2001/registry\" \
         xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">{}{}{}</oor:items>",
        setting("/org.openoffice.Setup/L10N", "ooSetupSystemLocale", locale),
        setting(
            "/org.openoffice.Office.Calc/Formula/Load",
            "OOXMLRecalcMode",
            "0"
        ),
        setting(
            "/org.openoffice.Office.Calc/Formula/Load",
            "ODFRecalcMode",
            "0"
        ),
    );
    std::fs::write(profile.join("user").join("registrymodifications.xcu"), xcu)
        .map_err(|e| e.to_string())?;
    let soffice = std::env::var("SOFFICE").unwrap_or_else(|_| "soffice".into());
    let copy = dir.join(file.file_name().ok_or("no file name")?);
    std::fs::copy(file, &copy).map_err(|e| format!("{}: {e}", file.display()))?;
    let out = Command::new(&soffice)
        .arg(format!(
            "-env:UserInstallation=file://{}",
            profile.display()
        ))
        .args(["--headless", "--convert-to", "ods", "--outdir"])
        .arg(dir.join("out"))
        .arg(&copy)
        .output()
        .map_err(|e| format!("cannot run `{soffice}` (set SOFFICE): {e}"))?;
    if !out.status.success() {
        return Err(format!(
            "{soffice} failed: {}",
            String::from_utf8_lossy(&out.stderr)
        ));
    }
    let ods = dir
        .join("out")
        .join(copy.file_stem().ok_or("no file name")?)
        .with_extension("ods");
    ods.exists()
        .then_some(ods)
        .ok_or_else(|| "LibreOffice produced no ODS".to_string())
}
