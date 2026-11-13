//! A full recalculation by `LibreOffice`, to compare the engine with: the values a
//! file caches may come from another version or from a calculation that never ran.

use crate::error::{Error, Result};
use std::path::{Path, PathBuf};
use std::process::Command;

/// Convert `file` to ODS with `LibreOffice` recalculating everything on load, under the
/// regional settings `locale` (a BCP-47 tag such as `es-CO`). The binary is `$SOFFICE`
/// or `soffice`. Returns the path of the ODS, inside `dir`.
///
/// # Errors
///
/// When `LibreOffice` is not available, fails, or does not write the recalculated file.
pub fn recalculate(file: &Path, locale: &str, dir: &Path) -> Result<PathBuf> {
    let profile = dir.join("profile");
    std::fs::create_dir_all(profile.join("user")).map_err(|e| Error::io(&profile, e))?;
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
    let registry = profile.join("user").join("registrymodifications.xcu");
    std::fs::write(&registry, xcu).map_err(|e| Error::io(registry, e))?;
    let soffice = std::env::var("SOFFICE").unwrap_or_else(|_| "soffice".into());
    let no_name = || Error::Recalculation(format!("{} has no file name", file.display()));
    let copy = dir.join(file.file_name().ok_or_else(no_name)?);
    std::fs::copy(file, &copy).map_err(|e| Error::io(file, e))?;
    let out = Command::new(&soffice)
        .arg(format!(
            "-env:UserInstallation=file://{}",
            profile.display()
        ))
        .args(["--headless", "--convert-to", "ods", "--outdir"])
        .arg(dir.join("out"))
        .arg(&copy)
        .output()
        .map_err(|e| Error::Recalculation(format!("cannot run `{soffice}` (set SOFFICE): {e}")))?;
    if !out.status.success() {
        return Err(Error::Recalculation(format!(
            "{soffice} failed: {}",
            String::from_utf8_lossy(&out.stderr)
        )));
    }
    let ods = dir
        .join("out")
        .join(copy.file_stem().ok_or_else(no_name)?)
        .with_extension("ods");
    ods.exists()
        .then_some(ods)
        .ok_or_else(|| Error::Recalculation("LibreOffice produced no ODS".into()))
}
