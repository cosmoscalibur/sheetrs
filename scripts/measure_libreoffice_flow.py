#!/opt/libreoffice26.8/program/python
"""Time and memory of the same flows in SheetRS and in LibreOffice (headless, through UNO).

Run it with the Python that ships with LibreOffice (it has `uno`), on a workbook made by
`sheetrs-compiler/scripts/gen_large_fixture.py` (sheets `Calc` and `Summary`):

    python3 sheetrs-compiler/scripts/gen_large_fixture.py /tmp/f300000.xlsx 300000
    /opt/libreoffice26.8/program/python scripts/measure_libreoffice_flow.py /tmp/f300000.xlsx \
        [--rounds 3] [--only NAME[,NAME]] [--sheetrs target/dist/sheetrs] [--sheetc target/dist/sheetc]

Flows (each round runs all of them in turn, so the shared machine affects them alike):

  sheetrs-edit            sheetrs edit -i FILE --delete-rows "Calc!3:10" --formula-to-value Summary -o OUT
  libreoffice-edit        load; delete the rows (Rows.removeByIndex); select the used area of Summary and
                          run the command Formula to Value (.uno:ConvertFormulaToValue); store as XLSX.
                          LibreOffice calculates the dirty cells when the conversion reads them.
  libreoffice-edit-calc   the same with an explicit calculateAll() between the deletion and the conversion
  sheetrs-delete          sheetrs edit -i FILE --delete-rows "Calc!3:10" -o OUT (a deletion alone: the engine
                          recalculates the results it leaves stale)
  libreoffice-delete-calc load; delete the rows; calculateAll(); store as XLSX
  sheetc-calculate        sheetc calculate -i FILE -o OUT (every formula calculated, the workbook written)
  libreoffice-calculate   load; calculateAll(); store as XLSX
  sheetrs-stats           sheetrs stats -i FILE (reads the whole workbook and counts it)
  libreoffice-load        load only (what reading the whole workbook costs in LibreOffice)
  sheetrs-find            sheetrs query -i FILE --find K00 (cells whose value has the text, every sheet)
  libreoffice-find        load; findAll("K00") over the document (search in values)

  libreoffice-model       the model use: load; set the input cell; calculateAll(); read the output cells; then
                          `--calls`-1 more times set the input (changed by 0.1 %), calculate() and read, the document kept
                          open (`--model-input Sheet!A1 --model-ranges Sheet!A1,Sheet!A3 --model-base 0.05`)

The four before it also run on an ODS (`formulas.ods`): they store nothing.

Wall time, CPU time (user + system) and peak resident memory (ru_maxrss of the child for SheetRS; the
VmHWM of `soffice.bin`, and the highest RSS sampled every 0.2 s, for LibreOffice) per phase.
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
import threading
import time

TICKS = os.sysconf("SC_CLK_TCK")


def run_sheetrs(cmd, out):
    t0 = time.perf_counter()
    p = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    _, status, ru = os.wait4(p.pid, 0)
    wall = time.perf_counter() - t0
    err = p.stderr.read().decode(errors="replace").strip() if p.stderr else ""
    code = os.waitstatus_to_exitcode(status)
    return {
        "wall": wall,
        "cpu": ru.ru_utime + ru.ru_stime,
        "peak_mb": ru.ru_maxrss / 1024,
        "exit": code,
        "error": err[-300:] if code else "",
        "out": out,
    }


def proc_of(port):
    """The pid of soffice.bin that listens on `port`."""
    for pid in os.listdir("/proc"):
        if not pid.isdigit():
            continue
        try:
            with open(f"/proc/{pid}/cmdline", "rb") as f:
                cmd = f.read().decode(errors="replace")
        except OSError:
            continue
        if "soffice.bin" in cmd and f"port={port}" in cmd:
            return int(pid)
    return None


def status_kb(pid, key):
    try:
        with open(f"/proc/{pid}/status") as f:
            for line in f:
                if line.startswith(key):
                    return int(line.split()[1])
    except OSError:
        pass
    return 0


def cpu_of(pid):
    try:
        with open(f"/proc/{pid}/stat") as f:
            parts = f.read().rsplit(")", 1)[1].split()
        return (int(parts[11]) + int(parts[12])) / TICKS
    except OSError:
        return 0.0


def cell_of(doc, address):
    sheet, cell = address.split("!")
    return doc.Sheets.getByName(sheet).getCellRangeByName(cell)


def lo_flow(flow, path, out, port, model=None, calls=100):
    import uno
    from com.sun.star.beans import PropertyValue

    def prop(name, value):
        p = PropertyValue()
        p.Name, p.Value = name, value
        return p

    # a run that stopped before closing the document leaves its lock file, and the next load then fails
    lock = os.path.join(os.path.dirname(os.path.abspath(path)), f".~lock.{os.path.basename(path)}#")
    if os.path.exists(lock):
        os.remove(lock)
    profile = tempfile.mkdtemp(prefix="lomeasure-profile-")
    server = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         f"-env:UserInstallation=file://{profile}", f"--accept=socket,host=localhost,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    result = {"phases": {}, "error": ""}
    peak = {"rss": 0}
    stop = threading.Event()
    pid_box = {"pid": None}

    def current():
        """The soffice.bin of this server: a new profile can make it restart once, so look again if it is gone."""
        pid = pid_box["pid"]
        if not pid or not os.path.exists(f"/proc/{pid}"):
            pid = proc_of(port)
            pid_box["pid"] = pid
        return pid

    def sampler():
        while not stop.is_set():
            pid = current()
            if pid:
                peak["rss"] = max(peak["rss"], status_kb(pid, "VmRSS:"))
                peak["hwm"] = max(peak.get("hwm", 0), status_kb(pid, "VmHWM:"))
            time.sleep(0.2)

    th = threading.Thread(target=sampler, daemon=True)
    th.start()
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            raise RuntimeError("LibreOffice did not answer")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        time.sleep(3)  # a new profile may restart soffice.bin once

        def phase(name, fn):
            t0, c0 = time.perf_counter(), cpu_of(current())
            fn()
            result["phases"][name] = (time.perf_counter() - t0, cpu_of(current()) - c0)

        state = {}

        def load():
            state["doc"] = desktop.loadComponentFromURL(
                "file://" + os.path.abspath(path), "_blank", 0, (prop("Hidden", True),))

        phase("load", load)
        doc = state["doc"]
        sheets = doc.Sheets
        if flow == "libreoffice-load":
            doc.close(True)
            return result
        if flow == "libreoffice-model":
            # The model use: open once, set the input cell, calculate, read the output cells; the first call
            # calculates every formula (the file stores no results), the next ones only what the input moves.
            target, ranges = model["input"], model["ranges"]
            base = model["base"]
            times = []

            def call(i):
                t0 = time.perf_counter()
                cell_of(doc, target).setValue(base * (1 + 0.001 * i))
                doc.calculateAll() if i == 0 else doc.calculate()
                for address in ranges:
                    cell_of(doc, address).getDataArray()
                times.append(time.perf_counter() - t0)

            for i in range(calls):
                call(i)
            result["calls"] = times
            doc.close(True)
            return result
        if flow == "libreoffice-find":
            def find():
                count = 0
                for i in range(sheets.Count):
                    sheet = sheets.getByIndex(i)
                    desc = sheet.createSearchDescriptor()
                    desc.SearchString = "K00"
                    desc.SearchType = 1  # values
                    found = sheet.findAll(desc)
                    if found is not None:
                        for rng in found.RangeAddresses:
                            count += (rng.EndRow - rng.StartRow + 1) * (rng.EndColumn - rng.StartColumn + 1)
                result["found"] = count

            phase("find", find)
            doc.close(True)
            return result
        if flow in ("libreoffice-edit", "libreoffice-edit-calc", "libreoffice-delete-calc"):
            phase("delete", lambda: sheets.getByName("Calc").Rows.removeByIndex(2, 8))
            if flow in ("libreoffice-edit-calc", "libreoffice-delete-calc"):
                phase("calculateAll", doc.calculateAll)
            if flow == "libreoffice-delete-calc":
                phase("store", lambda: doc.storeToURL(
                    "file://" + os.path.abspath(out), (prop("FilterName", "Calc MS Excel 2007 XML"),)))
                doc.close(True)
                return result

            def convert():
                summary = sheets.getByName("Summary")
                cursor = summary.createCursor()
                cursor.gotoEndOfUsedArea(False)
                end = cursor.RangeAddress
                rng = summary.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow)
                ctrl = doc.getCurrentController()
                dispatcher = ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.DispatchHelper", ctx)
                ctrl.select(rng)
                dispatcher.executeDispatch(ctrl.getFrame(), ".uno:ConvertFormulaToValue", "", 0, ())

            phase("convert", convert)
        else:
            phase("calculateAll", doc.calculateAll)
        phase("store", lambda: doc.storeToURL(
            "file://" + os.path.abspath(out), (prop("FilterName", "Calc MS Excel 2007 XML"),)))
        doc.close(True)
    except Exception as e:  # the report says where it stopped
        result["error"] = f"{type(e).__name__}: {e}"[:300]
    finally:
        stop.set()
        th.join(timeout=2)
        result["hwm_mb"] = peak.get("hwm", 0) / 1024
        result["sampled_peak_mb"] = peak["rss"] / 1024
        try:
            server.terminate()
            server.wait(timeout=30)
        except Exception:
            server.kill()
        os.system(f"rm -rf {profile}")
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--rounds", type=int, default=3)
    ap.add_argument("--only", default="")
    ap.add_argument("--sheetrs", default="target/dist/sheetrs")
    ap.add_argument("--sheetc", default="target/dist/sheetc")
    ap.add_argument("--workdir", default=tempfile.gettempdir())
    ap.add_argument("--model-input", default="", help="the input cell of the model flow (Sheet!A1)")
    ap.add_argument("--model-ranges", default="", help="the ranges the model flow reads, comma separated")
    ap.add_argument("--model-base", type=float, default=1.0, help="the value of the input cell, which each call changes by 0.1 %%")
    ap.add_argument("--calls", type=int, default=100, help="calls of the model flow with the document kept open")
    a = ap.parse_args()
    flows = ["sheetrs-edit", "libreoffice-edit", "libreoffice-edit-calc", "sheetrs-delete", "libreoffice-delete-calc",
             "sheetc-calculate", "libreoffice-calculate", "sheetrs-stats", "libreoffice-load", "sheetrs-find",
             "libreoffice-find"]
    if a.only:  # the model flow needs --model-input and --model-ranges, so it runs only when asked for
        flows = [f for f in flows + ["libreoffice-model"] if f in a.only.split(",")]
    out = lambda n: os.path.join(a.workdir, f"measure-out-{n}.xlsx")
    for rnd in range(1, a.rounds + 1):
        for i, flow in enumerate(flows):
            o = out(flow)
            if flow in ("sheetrs-edit", "sheetrs-delete"):
                extra = ["--formula-to-value", "Summary"] if flow == "sheetrs-edit" else []
                r = run_sheetrs([a.sheetrs, "edit", "-i", a.workbook, "--delete-rows", "Calc!3:10",
                                 *extra, "-o", o], o)
                print(f"round {rnd} {flow}: wall {r['wall']:.1f} s, cpu {r['cpu']:.1f} s, peak {r['peak_mb']:.0f} MB"
                      f"{' EXIT ' + str(r['exit']) + ' ' + r['error'] if r['exit'] else ''}", flush=True)
            elif flow in ("sheetrs-stats", "sheetrs-find"):
                cmd = ([a.sheetrs, "stats", "-i", a.workbook, "--format", "json"] if flow == "sheetrs-stats"
                       else [a.sheetrs, "query", "-i", a.workbook, "--find", "K00", "--jsonl"])
                r = run_sheetrs(cmd, o)
                print(f"round {rnd} {flow}: wall {r['wall']:.1f} s, cpu {r['cpu']:.1f} s, peak {r['peak_mb']:.0f} MB"
                      f"{' EXIT ' + str(r['exit']) + ' ' + r['error'] if r['exit'] else ''}", flush=True)
            elif flow == "sheetc-calculate":
                r = run_sheetrs([a.sheetc, "calculate", "-i", a.workbook, "-o", o], o)
                print(f"round {rnd} {flow}: wall {r['wall']:.1f} s, cpu {r['cpu']:.1f} s, peak {r['peak_mb']:.0f} MB"
                      f"{' EXIT ' + str(r['exit']) + ' ' + r['error'] if r['exit'] else ''}", flush=True)
            elif flow == "libreoffice-model":
                model = {"input": a.model_input, "ranges": a.model_ranges.split(","), "base": a.model_base}
                r = lo_flow(flow, a.workbook, o, 2300 + rnd * 10 + i, model, a.calls)
                t = r.get("calls", [])
                load = r["phases"].get("load", (0, 0))[0]
                rest = sorted(t[1:])
                print(f"round {rnd} {flow}: load {load:.2f} s; first call (set, calculateAll, read) {t[0] if t else 0:.3f} s; "
                      f"calls 2-{len(t)} (set, calculate, read) min {rest[0] if rest else 0:.4f} median "
                      f"{rest[len(rest) // 2] if rest else 0:.4f} max {rest[-1] if rest else 0:.4f} s; all {len(t)} calls {sum(t):.2f} s; "
                      f"VmHWM {r.get('hwm_mb', 0):.0f} MB{' ERROR ' + r['error'] if r['error'] else ''}", flush=True)
                print("MODEL_JSON " + json.dumps({"load": load, "calls": t, "hwm_mb": r.get("hwm_mb", 0), "error": r["error"]}), flush=True)
            else:
                r = lo_flow(flow, a.workbook, o, 2300 + rnd * 10 + i)
                ph = ", ".join(f"{k} {v[0]:.1f} s (cpu {v[1]:.1f})" for k, v in r["phases"].items())
                total = sum(v[0] for v in r["phases"].values())
                print(f"round {rnd} {flow}: {ph}; total {total:.1f} s; VmHWM {r.get('hwm_mb', 0):.0f} MB, "
                      f"sampled peak {r['sampled_peak_mb']:.0f} MB{' found ' + str(r['found']) + ' cells' if 'found' in r else ''}"
                      f"{' ERROR ' + r['error'] if r['error'] else ''}",
                      flush=True)
            if os.path.exists(o) and rnd < a.rounds:
                os.remove(o)


if __name__ == "__main__":
    main()
