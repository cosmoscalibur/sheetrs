#!/bin/bash
set -e

# Benchmark script for sheetlint
# Uses /usr/bin/time -v and perf stat for comprehensive metrics:
# - User time, System time, Elapsed time
# - CPU usage, Peak RAM
# - User-space instructions
# - MD5 of each run's output
# Each file runs twice; the first run's output is the violations report.
# Results are saved to a timestamped file in scripts/
#
# Usage: ./scripts/benchmark.sh [--baseline <git-ref>]
#   --baseline  also builds <git-ref> in a temporary worktree and times it in
#               the same session (metrics only, no violations report).

PROD_FILE_XLSX="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.xlsx"
PROD_FILE_ODS="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.ods"
CONFIG="scripts/benchmark.toml"
RUNS=2
BASELINE_REF=""

if [ "$1" = "--baseline" ]; then
    BASELINE_REF=${2:?"--baseline requires a git ref"}
fi

TIMESTAMP=$(date +%Y%m%d_%H%M%S)
RESULTS_FILE="scripts/benchmark_${TIMESTAMP}.txt"

echo "=== Building sheetlint with the profiling profile ==="
cargo build --profile profiling --package sheetlint

SHEETLINT="./target/profiling/sheetlint"

if [ -n "$BASELINE_REF" ]; then
    BASELINE_SHA=$(git rev-parse --short "$BASELINE_REF")
    BASELINE_DIR=$(mktemp -d)
    trap 'git worktree remove --force "$BASELINE_DIR"' EXIT
    echo "=== Building baseline ${BASELINE_REF} (${BASELINE_SHA}) ==="
    git worktree add --detach "$BASELINE_DIR" "$BASELINE_SHA"
    cargo build --profile profiling --package sheetlint \
        --manifest-path "$BASELINE_DIR/Cargo.toml" --target-dir target/baseline
fi
BASELINE_SHEETLINT="./target/baseline/profiling/sheetlint"

# Times RUNS executions of a sheetlint binary on one file. Prints the first
# run's violations unless the binary is the baseline.
benchmark_file() {
    local label=$1
    local file=$2
    local sheetlint=${3:-$SHEETLINT}

    echo "=== Benchmarking ${label} ==="
    for run in $(seq 1 "$RUNS"); do
        echo "--- Run ${run} ---"
        perf stat -x, -e instructions:u -o perf_output.tmp -- \
            /usr/bin/time -v "$sheetlint" -c "$CONFIG" "$file" > "violations_${run}.tmp" 2> time_output.tmp || true
        grep -E "(User time|System time|Elapsed|Maximum resident|Percent of CPU)" time_output.tmp
        awk -F, '/instructions/ {print "\tInstructions (user): " $1}' perf_output.tmp
        md5sum < "violations_${run}.tmp" | awk '{print "\tOutput MD5: " $1}'
    done
    rm -f time_output.tmp perf_output.tmp
    echo ""

    if [ "$sheetlint" = "$SHEETLINT" ]; then
        echo "=== ${label} Violations ==="
        cat violations_1.tmp
        echo ""
    fi
    rm -f violations_*.tmp
}

{
    echo "=== Benchmark Results (${TIMESTAMP}) ==="
    echo ""

    echo "=== Binary Size ==="
    ls -lh "$SHEETLINT" | awk '{print "Binary Size: " $5}'
    echo ""

    if [ -n "$BASELINE_REF" ]; then
        benchmark_file "Baseline ${BASELINE_SHA} XLSX" "$PROD_FILE_XLSX" "$BASELINE_SHEETLINT"
        benchmark_file "Baseline ${BASELINE_SHA} ODS" "$PROD_FILE_ODS" "$BASELINE_SHEETLINT"
    fi
    benchmark_file XLSX "$PROD_FILE_XLSX"
    benchmark_file ODS "$PROD_FILE_ODS"

    echo "=== Benchmark Complete ==="
} 2>&1 | tee "$RESULTS_FILE"

echo ""
echo "Results saved to: $RESULTS_FILE"
