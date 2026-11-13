//! Lists the functions of the registry, one per line: `namespace name min max` (`max` 255 is "any
//! number"). `sheetrs-compiler/scripts/probe_arity.py` reads it to ask `LibreOffice` what each function
//! does with each number of arguments.
fn main() {
    for def in sheetrs_engine::functions::table() {
        println!("{} {} {} {}", def.ns.as_str(), def.name, def.min, def.max);
    }
}
