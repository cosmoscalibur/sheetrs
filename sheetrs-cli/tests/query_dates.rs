//! `sheetrs query --dates`: the serial by default and the dates as dates with the flag.

use std::path::Path;
use std::process::Command;

fn query(args: &[&str]) -> serde_json::Value {
    let file = Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-io/tests/assets/dates.xlsx");
    let out = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("query")
        .arg("-i")
        .arg(file)
        .args(["--range", "Dates!A13:C13"])
        .args(args)
        .output()
        .unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    serde_json::from_slice(&out.stdout).unwrap()
}

#[test]
fn the_serial_is_the_default() {
    let answer = query(&[]);
    let values: Vec<_> = answer["cells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|c| c["value"].clone())
        .collect();
    assert_eq!(values, [45366.4375, 45366.4375, 45366.4375]);
}

#[test]
fn a_date_a_date_and_time_and_a_time_are_typed_with_the_dates_flag() {
    let answer = query(&["--dates"]);
    let values: Vec<_> = answer["cells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|c| c["value"].clone())
        .collect();
    assert_eq!(
        values,
        [
            serde_json::json!({"type": "date", "value": "2024-03-15"}),
            serde_json::json!({"type": "datetime", "value": "2024-03-15T10:30:00"}),
            serde_json::json!({"type": "time", "value": "10:30:00"}),
        ]
    );
}

#[test]
fn the_dates_flag_takes_no_value() {
    let file = Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-io/tests/assets/dates.xlsx");
    let out = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("query")
        .arg("-i")
        .arg(file)
        .args(["--range", "Dates!A13:C13", "--dates", "datetime"])
        .output()
        .unwrap();
    assert!(!out.status.success());
}
