//! Prints the engine's function registry: `namespace NAME` per line.
fn main() {
    for d in sheetrs_compiler::functions::table() {
        println!("{} {}", d.ns.as_str(), d.name);
    }
}
