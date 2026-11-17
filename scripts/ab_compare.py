#!/usr/bin/env python3
"""Run two or more commands alternately and compare them: median, quartiles and peak memory.

    scripts/ab_compare.py [-n RUNS] [--check] LABEL=COMMAND [LABEL=COMMAND ...]

The commands run in rounds (A, B, A, B, ...), so a drift of the machine reaches all of them; give the same
binary twice to see the noise floor (a pair of copies of one binary is the floor of DESIGN.md, "Measuring a
change"). `COMMAND` is split like a shell line; stdout and stderr go to /dev/null. With `--check` the
standard output of the first run of every command is compared with the first one's and the script fails
when they differ (an executable of a program built by two engines must answer the same bytes).
Standard library only; Linux (peak memory is `ru_maxrss` of the child).
"""
import argparse
import os
import shlex
import statistics
import subprocess
import sys
import time


def run(cmd):
    start = time.perf_counter()
    pid = os.fork()
    if pid == 0:
        null = os.open(os.devnull, os.O_WRONLY)
        os.dup2(null, 1)
        os.dup2(null, 2)
        try:
            os.execvp(cmd[0], cmd)
        finally:
            os._exit(127)
    _, status, usage = os.wait4(pid, 0)
    if status != 0:
        sys.exit(f"command failed ({os.waitstatus_to_exitcode(status)}): {shlex.join(cmd)}")
    return time.perf_counter() - start, usage.ru_maxrss / 1024


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("-n", "--runs", type=int, default=7)
    ap.add_argument("--check", action="store_true", help="the answers must be identical")
    ap.add_argument("commands", nargs="+", metavar="LABEL=COMMAND")
    args = ap.parse_args()
    cmds = {}
    for item in args.commands:
        label, _, line = item.partition("=")
        if not line:
            ap.error(f"`{item}` is not LABEL=COMMAND")
        cmds[label] = shlex.split(line)
    if args.check:
        outs = {k: subprocess.run(c, capture_output=True, check=True).stdout for k, c in cmds.items()}
        first = next(iter(outs.values()))
        for k, o in outs.items():
            if o != first:
                sys.exit(f"the answer of `{k}` differs from the first one")
        print(f"answers identical ({len(first)} bytes)")
    results = {k: [] for k in cmds}
    for _ in range(args.runs):
        for k, c in cmds.items():
            results[k].append(run(c))
    for k, rs in results.items():
        times = [t for t, _ in rs]
        q = statistics.quantiles(times, n=4) if len(times) > 3 else [min(times), 0, max(times)]
        print(f"{k:28s} median {statistics.median(times):8.3f} s  q1 {q[0]:.3f}  q3 {q[2]:.3f}"
              f"  peak {max(m for _, m in rs):6.0f} MB")


if __name__ == "__main__":
    main()
