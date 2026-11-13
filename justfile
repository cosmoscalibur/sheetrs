# List all recipes
default:
    @just --list

# Run all tests
test:
    cargo test --workspace --exclude sheetrs-py

# Run tests in release mode (for rules-heavy workbooks)
test-release:
    cargo test --release --workspace --exclude sheetrs-py

# Lint with clippy
lint:
    cargo clippy --workspace --exclude sheetrs-py --all-targets -- -D warnings

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
    cargo doc --workspace --exclude sheetrs-py --no-deps --open

# Run benchmarks
bench:
    ./scripts/bench.py fixtures target/bench-data && ./scripts/bench.py run target/bench-data

# Coverage (requires cargo-tarpaulin)
coverage:
    cargo tarpaulin --out html --output-dir target/coverage

# Security audit (requires cargo-audit)
audit:
    cargo audit
