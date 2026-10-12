# List all recipes
default:
    @just --list

# Run all tests
test:
    cargo test

# Run tests in release mode (for rules-heavy workbooks)
test-release:
    cargo test --release

# Lint with clippy
lint:
    cargo clippy --all-targets -- -D warnings

# Format check
fmt-check:
    cargo fmt --check

# Format fix
fmt:
    cargo fmt

# Full quality check
check: fmt-check lint test

# Generate docs
doc:
    cargo doc --no-deps --open

# Run benchmarks
bench:
    ./scripts/benchmark.sh

# Coverage (requires cargo-tarpaulin)
coverage:
    cargo tarpaulin --out html --output-dir target/coverage

# Security audit (requires cargo-audit)
audit:
    cargo audit
