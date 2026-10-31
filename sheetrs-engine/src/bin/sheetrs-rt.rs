//! Slim runtime stub. `sheetc build` appends a compiled program to a copy of
//! this binary (or of the full `sheetc`) to produce a standalone executable.

fn main() {
    let args = sheetrs_engine::cli::Args(std::env::args().skip(1).collect());
    let code = match sheetrs_engine::bundle::read_embedded() {
        Some(e) => sheetrs_engine::cli::with_big_stack(move || {
            sheetrs_engine::cli::run_program(&e.program, e.consts.as_ref(), args)
        }),
        None => {
            eprintln!(
                "sheetrs-rt: no program embedded; this is the runtime stub used by `sheetc build`"
            );
            2
        }
    };
    std::process::exit(code);
}
