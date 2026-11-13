//! The recursion of the engine over run-time references is bounded by the shadow stack of the
//! module (1 MiB by default): the same size as the package of the tools (see `sheetrs-wasm`).

fn main() {
    if std::env::var("CARGO_CFG_TARGET_ARCH").as_deref() == Ok("wasm32") {
        println!("cargo:rustc-link-arg=-zstack-size=67108864");
    }
    println!("cargo:rerun-if-changed=build.rs");
}
