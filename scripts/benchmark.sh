#!/bin/bash
set -e

# Benchmark script for sheetlint
# Uses /usr/bin/time -v and perf stat for comprehensive metrics:
# - User time, System time, Elapsed time
# - CPU usage, Peak RAM
# - User-space instructions
# Each file runs twice; the first run's output is the violations report.
# Results are saved to a timestamped file in scripts/

PROD_FILE_XLSX="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.xlsx"
PROD_FILE_ODS="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.ods"
RUNS=2

TIMESTAMP=$(date +%Y%m%d_%H%M%S)
RESULTS_FILE="scripts/benchmark_${TIMESTAMP}.txt"

echo "=== Building sheetlint with the profiling profile ==="
cargo build --profile profiling --package sheetlint

SHEETLINT="./target/profiling/sheetlint"

# Times RUNS executions of sheetlint on one file and prints the first run's violations.
benchmark_file() {
    local label=$1
    local file=$2

    echo "=== Benchmarking ${label} ==="
    for run in $(seq 1 "$RUNS"); do
        echo "--- Run ${run} ---"
        perf stat -x, -e instructions:u -o perf_output.tmp -- \
            /usr/bin/time -v "$SHEETLINT" "$file" > "violations_${run}.tmp" 2> time_output.tmp || true
        grep -E "(User time|System time|Elapsed|Maximum resident|Percent of CPU)" time_output.tmp
        awk -F, '/instructions/ {print "\tInstructions (user): " $1}' perf_output.tmp
    done
    rm -f time_output.tmp perf_output.tmp
    echo ""

    echo "=== ${label} Violations ==="
    cat violations_1.tmp
    rm -f violations_*.tmp
    echo ""
}

{
    echo "=== Benchmark Results (${TIMESTAMP}) ==="
    echo ""

    echo "=== Binary Size ==="
    ls -lh "$SHEETLINT" | awk '{print "Binary Size: " $5}'
    echo ""

    benchmark_file XLSX "$PROD_FILE_XLSX"
    benchmark_file ODS "$PROD_FILE_ODS"

    echo "=== Benchmark Complete ==="
} 2>&1 | tee "$RESULTS_FILE"

echo ""
echo "Results saved to: $RESULTS_FILE"
