#!/usr/bin/env python3
"""Black-box oracle: applies test cases to LibreOffice through its public UNO API.

Reads {"cases": [...]} as JSON on stdin and writes {"office": ..., "results": [...]}
on stdout. One LibreOffice process, started with a fresh user profile, serves
every case of a run. The process and everything it spawned is killed when the
run ends or exceeds RUN_TIMEOUT_SECONDS.
"""

import contextlib
import json
import math
import os
import re
import signal
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from com.sun.star.sheet.FormulaResult import STRING as FORMULA_RESULT_STRING
from com.sun.star.table.CellContentType import EMPTY, FORMULA, TEXT, VALUE

CONNECT_TIMEOUT_SECONDS = 90
RUN_TIMEOUT_SECONDS = 300
SHUTDOWN_TIMEOUT_SECONDS = 30
LOG_TAIL_BYTES = 2000
SAVE_FILTERS = {
    "xlsx": "Calc MS Excel 2007 XML",
    "ods": "calc8",
}
ROW_SPAN = re.compile(r"([0-9]+)(?::([0-9]+))?")
COLUMN_SPAN = re.compile(r"([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?")


def property_value(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def start_office(work_dir, pipe):
    command = [
        "soffice",
        f"-env:UserInstallation={(work_dir / 'profile').as_uri()}",
        "--headless",
        "--norestore",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        f"--accept=pipe,name={pipe};urp;StarOffice.ComponentContext",
    ]
    with (work_dir / "soffice.log").open("wb") as log:
        # Own session: the whole process group can be killed, including the
        # soffice.bin child that the launcher starts.
        return subprocess.Popen(command, stdout=log, stderr=log, start_new_session=True)


def kill_office(office):
    with contextlib.suppress(ProcessLookupError):
        os.killpg(office.pid, signal.SIGKILL)


def connect(office, pipe, log_path):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    url = f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext"
    deadline = time.monotonic() + CONNECT_TIMEOUT_SECONDS
    while True:
        try:
            return resolver.resolve(url)
        except NoConnectException:
            if office.poll() is not None:
                tail = log_path.read_bytes()[-LOG_TAIL_BYTES:].decode(errors="replace")
                raise RuntimeError(
                    f"soffice exited with code {office.returncode}: {tail}"
                ) from None
            if time.monotonic() > deadline:
                raise
            time.sleep(0.25)


def office_version(context):
    provider = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context
    )
    node = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess",
        (property_value("nodepath", "/org.openoffice.Setup/Product"),),
    )
    return node.getByName("ooSetupVersionAboutBox")


def new_document(desktop):
    return desktop.loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, (property_value("Hidden", True),)
    )


def populate(document, sheets):
    collection = document.Sheets
    collection.getByIndex(0).Name = sheets[0]["name"]
    for position, spec in enumerate(sheets[1:], start=1):
        collection.insertNewByName(spec["name"], position)
    for spec in sheets:
        sheet = collection.getByName(spec["name"])
        for address, content in spec["cells"].items():
            cell = sheet.getCellRangeByName(address)
            if isinstance(content, str) and content.startswith("="):
                cell.setFormula(content)
            elif isinstance(content, str):
                cell.setString(content)
            else:
                cell.setValue(float(content))
    # Typed cells are calculated on entry in the UI; calculate them all so
    # every cell holds its result before the operations run.
    document.calculateAll()


def row_span(spec):
    match = ROW_SPAN.fullmatch(spec)
    if match is None:
        raise ValueError(f"invalid row span {spec!r}: expected N or N:M")
    start, end = int(match[1]), int(match[2] or match[1])
    if not 1 <= start <= end:
        raise ValueError(
            f"invalid row span {spec!r}: rows are one-based and the span cannot be reversed"
        )
    return start - 1, end - start + 1


def column_span(sheet, spec):
    match = COLUMN_SPAN.fullmatch(spec)
    if match is None:
        raise ValueError(
            f"invalid column span {spec!r}: expected letters as in B or B:C"
        )
    first, last = match[1], match[2] or match[1]
    start = sheet.getCellRangeByName(f"{first}1").RangeAddress.StartColumn
    end = sheet.getCellRangeByName(f"{last}1").RangeAddress.StartColumn
    if end < start:
        raise ValueError(f"invalid column span {spec!r}: the span cannot be reversed")
    return start, end - start + 1


def text_filter_options(options):
    """The token string of the text import filter, which names every option explicitly."""
    flag = lambda value: "true" if value else "false"
    tokens = [
        options["separator"],
        options["delimiter"],
        options["charset"],
        1,
        "",
        options["language"],
        flag(options["quoted_fields"] == "text"),
        flag(options["special_numbers"] == "detected"),
        "false",
        "false",
        flag(options["spaces"] == "trimmed"),
        -1,
        flag(options["formula_fields"] == "evaluated"),
    ]
    return ",".join(str(token) for token in tokens)


def apply_op(desktop, document, dispatcher, op):
    """Applies one operation and returns the document that is current afterwards."""
    kind = op["op"]
    if kind == "open":
        # Loaded as the UI does by default: cached results are trusted, not recalculated.
        properties = [property_value("Hidden", True)]
        text_filter = op.get("filter")
        if text_filter is not None:
            properties += [
                property_value("FilterName", "Text - txt - csv (StarCalc)"),
                property_value("FilterOptions", text_filter_options(text_filter)),
            ]
        loaded = desktop.loadComponentFromURL(
            Path(op["path"]).as_uri(), "_blank", 0, tuple(properties)
        )
        document.close(True)
        return loaded
    if kind == "add_basic_module":
        libraries = document.BasicLibraries
        if not libraries.hasByName(op["library"]):
            libraries.createLibrary(op["library"])
        library = libraries.getByName(op["library"])
        library.insertByName(op["module"], op["source"])
    elif kind == "recalculate":
        document.calculate()
    elif kind == "recalculate_hard":
        document.calculateAll()
    elif kind == "delete_sheet":
        document.Sheets.removeByName(op["sheet"])
    elif kind == "save_as":
        document.storeToURL(
            Path(op["path"]).as_uri(),
            (property_value("FilterName", SAVE_FILTERS[op["format"]]),),
        )
    else:
        sheet = document.Sheets.getByName(op["sheet"])
        if kind == "delete_rows":
            start, count = row_span(op["rows"])
            sheet.Rows.removeByIndex(start, count)
        elif kind == "delete_columns":
            start, count = column_span(sheet, op["columns"])
            sheet.Columns.removeByIndex(start, count)
        elif kind == "to_values":
            # The UI command converts the cached results, which in the UI are
            # always up to date.
            document.calculate()
            controller = document.CurrentController
            controller.select(sheet.getCellRangeByName(op["range"]))
            dispatcher.executeDispatch(
                controller.Frame, ".uno:ConvertFormulaToValue", "", 0, ()
            )
        else:
            raise ValueError(f"unknown operation: {kind}")
    return document


def encode_number(value, formula=None):
    """A number as JSON; infinities and NaN have no JSON form, so they get their own type."""
    if math.isnan(value):
        encoded = {"t": "non_finite", "v": "nan"}
    elif math.isinf(value):
        encoded = {"t": "non_finite", "v": "inf" if value > 0 else "-inf"}
    else:
        encoded = {"t": "number", "v": value}
    if formula is not None:
        encoded["formula"] = formula
    return encoded


def encode_cell(cell):
    kind = cell.getType()
    if kind == EMPTY:
        return {"t": "empty"}
    if kind == VALUE:
        return encode_number(cell.getValue())
    if kind == TEXT:
        return {"t": "string", "v": cell.getString()}
    assert kind == FORMULA, kind
    formula = cell.getFormula()
    error = cell.getError()
    if error != 0:
        return {"t": "error", "code": error, "text": cell.getString(), "formula": formula}
    if cell.FormulaResultType2 == FORMULA_RESULT_STRING:
        return {"t": "string", "v": cell.getString(), "formula": formula}
    return encode_number(cell.getValue(), formula)


def read_range(document, read):
    area = document.Sheets.getByName(read["sheet"]).getCellRangeByName(read["range"])
    address = area.RangeAddress
    rows = []
    for row in range(address.EndRow - address.StartRow + 1):
        rows.append(
            [
                encode_cell(area.getCellByPosition(col, row))
                for col in range(address.EndColumn - address.StartColumn + 1)
            ]
        )
    return {"sheet": read["sheet"], "range": read["range"], "cells": rows}


def external_links(document):
    """The names (URLs) of the documents that formulas of the workbook refer to."""
    try:
        return sorted(document.ExternalDocLinks.ElementNames)
    except Exception:
        return []


def basic_modules(document):
    """The (library, module) pairs of the Basic code the document holds; a library that cannot
    be loaded holds none."""
    found = []
    try:
        libraries = document.BasicLibraries
        names = list(libraries.ElementNames)
    except Exception:
        return found
    for library in names:
        try:
            if not libraries.isLibraryLoaded(library):
                libraries.loadLibrary(library)
            for module in libraries.getByName(library).ElementNames:
                found.append([library, module])
        except Exception:
            continue
    return sorted(found)


def run_case(desktop, dispatcher, case):
    document = new_document(desktop)
    try:
        populate(document, case["sheets"])
        for op in case["ops"]:
            document = apply_op(desktop, document, dispatcher, op)
        return {
            "name": case["name"],
            "reads": [read_range(document, read) for read in case["reads"]],
            "external_links": external_links(document),
            "basic_modules": basic_modules(document),
            "sheets": [
                {"name": sheet.Name, "visible": bool(sheet.IsVisible)}
                for sheet in (
                    document.Sheets.getByIndex(index)
                    for index in range(document.Sheets.Count)
                )
            ],
        }
    finally:
        # A failed close must not hide the case's own error; terminate() disposes what is left.
        with contextlib.suppress(Exception):
            document.close(True)


def main():
    cases = json.load(sys.stdin.buffer)["cases"]
    pipe = f"sheetrs_oracle_{uuid.uuid4().hex}"
    with tempfile.TemporaryDirectory(prefix="sheetrs-oracle-") as work:
        work_dir = Path(work)
        office = start_office(work_dir, pipe)
        timed_out = threading.Event()

        def expire():
            timed_out.set()
            kill_office(office)

        watchdog = threading.Timer(RUN_TIMEOUT_SECONDS, expire)
        watchdog.daemon = True
        watchdog.start()
        try:
            context = connect(office, pipe, work_dir / "soffice.log")
            version = office_version(context)
            services = context.ServiceManager
            desktop = services.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )
            dispatcher = services.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", context
            )
            results = []
            for case in cases:
                try:
                    results.append(run_case(desktop, dispatcher, case))
                except Exception as error:  # reported to the caller per case
                    if timed_out.is_set():
                        raise TimeoutError(
                            f"oracle exceeded {RUN_TIMEOUT_SECONDS}s"
                        ) from error
                    results.append(
                        {
                            "name": case["name"],
                            "error": f"{type(error).__name__}: {error}",
                        }
                    )
            json.dump({"office": version, "results": results}, sys.stdout)
            desktop.terminate()
        finally:
            watchdog.cancel()
            with contextlib.suppress(subprocess.TimeoutExpired):
                office.wait(timeout=SHUTDOWN_TIMEOUT_SECONDS)
            kill_office(office)
            office.wait()


if __name__ == "__main__":
    main()
