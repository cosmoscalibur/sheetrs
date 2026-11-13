//! A WebAssembly module has two stacks: the call stack of the host (the browser or Node), which
//! bounds how deep the calls can go, and a region of its own linear memory, the *shadow stack*,
//! where the locals whose address is taken live. The second is fixed when the module is linked
//! (1 MiB by default) and a recursion that the host allows to go deep (a larger `--stack-size` in
//! Node, a worker with a larger stack) would fill it first, so it is made larger here (it is address
//! space in the memory of the module, not memory in use).
//!
//! The engine evaluates a chain of formulas whose references are computed at run time recursively,
//! which a command-line program runs on a thread of 512 MiB; in a module the depth the host allows
//! is the limit (about a thousand levels per MiB of host stack).

fn main() {
    if std::env::var("CARGO_CFG_TARGET_ARCH").as_deref() == Ok("wasm32") {
        println!("cargo:rustc-link-arg=-zstack-size=67108864");
    }
    println!("cargo:rerun-if-env-changed=SHEETRS_RT_DIR");
    if let Ok(dir) = std::env::var("SHEETRS_RT_DIR") {
        for file in ["sheetrs-rt.wasm", "sheetrs-rt-template.wasm"] {
            println!("cargo:rerun-if-changed={dir}/{file}");
        }
    }
    println!("cargo:rerun-if-changed=build.rs");
}
