FROM rust:1.99

RUN rustup component add clippy rustfmt \
 && rustup target add x86_64-pc-windows-msvc aarch64-apple-darwin
