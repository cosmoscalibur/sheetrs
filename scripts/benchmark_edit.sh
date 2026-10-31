#!/bin/bash
set -eo pipefail

# Benchmark script for sheetedit modifications
# Uses /usr/bin/time -v and perf stat, like benchmark.sh, on each scenario of
# SCENARIOS for the XLSX and ODS production files. Each scenario runs twice.
# The deletion scenarios rewrite every sheet: a formula of any sheet can read the rows and
# columns deleted from INPUTS.
# sheetedit reads a copy of each production file in a temporary directory and
# writes its output there; the originals are checked unchanged at the end.
# Results are saved to a timestamped file in scripts/
#
# Usage: ./scripts/benchmark_edit.sh

PROD_FILE_XLSX="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.xlsx"
PROD_FILE_ODS="tests/production/TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.ods"
RUNS=2
# INPUTS is the largest sheet of the production files
SCENARIOS=(
    "remove-sheets|--remove-sheets Tabla_Exogena"
    "remove-ranges|--remove-ranges deuda"
    "convert-sheet|--convert-to-values INPUTS"
    "convert-columns-rows|--convert-to-values INPUTS!B:D INPUTS!1:1000"
    "plan|--remove-sheets Tabla_Exogena --remove-ranges deuda --convert-to-values INPUTS"
    "delete-rows|--delete-rows INPUTS!10:20"
    "delete-columns|--delete-cols INPUTS!B:D"
    "delete-rows-columns|--delete-rows INPUTS!10:20 --delete-cols INPUTS!B:D --recalc-on-open"
    "structure-plan|--remove-sheets Tabla_Exogena --remove-ranges deuda --delete-rows INPUTS!10:20 --delete-cols INPUTS!B:D"
)

TIMESTAMP=$(date +%Y%m%d_%H%M%S)
RESULTS_FILE="scripts/benchmark_edit_${TIMESTAMP}.txt"
WORK_DIR=$(mktemp -d)
trap 'rm -r "$WORK_DIR"' EXIT

echo "=== Building sheetedit with the profiling profile ==="
cargo build --profile profiling --package sheetrs-edit

SHEETEDIT="./target/profiling/sheetedit"

# Times RUNS executions of every scenario on a copy of one file.
benchmark_file() {
    local label=$1
    local file=$2
    local ext=${file##*.}
    local input="${WORK_DIR}/input.${ext}"
    local output="${WORK_DIR}/output.${ext}"
    cp "$file" "$input"

    for scenario in "${SCENARIOS[@]}"; do
        local name=${scenario%%|*}
        local args
        read -ra args <<< "${scenario#*|}"
        echo "=== ${label} ${name}: ${args[*]} ==="
        for run in $(seq 1 "$RUNS"); do
            echo "--- Run ${run} ---"
            rm -f "$output"
            perf stat -x, -e instructions:u -o "${WORK_DIR}/perf.tmp" -- \
                /usr/bin/time -v "$SHEETEDIT" "$input" "${args[@]}" -o "$output" \
                > /dev/null 2> "${WORK_DIR}/time.tmp"
            grep -E "(User time|System time|Elapsed|Maximum resident|Percent of CPU)" "${WORK_DIR}/time.tmp"
            awk -F, '/instructions/ {print "\tInstructions (user): " $1}' "${WORK_DIR}/perf.tmp"
            # ZIP entries carry their write time, so hash the extracted parts
            unzip -p "$output" | md5sum | awk '{print "\tOutput MD5 (parts): " $1}'
        done
        echo ""
    done
}

INPUT_MD5=$(md5sum "$PROD_FILE_XLSX" "$PROD_FILE_ODS")

{
    echo "=== CLI Benchmark Results (${TIMESTAMP}) ==="
    echo ""

    echo "=== Binary Size ==="
    echo "Binary Size: $(stat -c %s "$SHEETEDIT" | LC_ALL=C numfmt --to=iec)"
    echo ""

    benchmark_file XLSX "$PROD_FILE_XLSX"
    benchmark_file ODS "$PROD_FILE_ODS"

    if [ "$(md5sum "$PROD_FILE_XLSX" "$PROD_FILE_ODS")" != "$INPUT_MD5" ]; then
        echo "ERROR: a production file changed during the benchmark"
        exit 1
    fi
    echo "=== Benchmark Complete ==="
} 2>&1 | tee "$RESULTS_FILE"

echo ""
echo "Results saved to: $RESULTS_FILE"
