#!/usr/bin/env python3
"""Capture how Mac Excel parses typed cell input, per locale, into the test fixture.

Every input listed in the fixture is written into a temporary workbook through
the object model (no synthetic keystrokes), the workbook is saved, and the
stored type, value and number format of each cell are read back from the xlsx.
Locale switching and restoration reuse scripts/excel-locale-probe.py. The
fixture is both the case list and the oracle that
tests/unit/commands/coerce-input-excel-oracle.test.ts checks against.

Run with a Python that has xlwings, while Excel is not running, then format
the fixture with `yarn lint:fix`:
  python3 scripts/excel-input-probe.py [--add INPUT ...] [--locale ja_JP ...]
"""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import json
import re
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "packages" / "formulon-cell" / "tests" / "fixtures" / "excel-input-coercion.json"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
LOCALES = {
    "ja_JP": "ja-JP",
    "en_US": "en-US",
    "de_DE": "de-DE",
    "fr_FR": "fr-FR",
    "zh_CN": "zh-CN",
    "ko_KR": "ko-KR",
    "th_TH": "th-TH",
}
# Built-in ids an xlsx references without writing the code: the ECMA-376 set,
# and the locale-dependent date ids Excel writes for CJK locales.
BUILTIN = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
    11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm",
    17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}
LOCALE_BUILTIN = {
    "ja_JP": {31: 'yyyy"年"m"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"',
              57: "[$-411]ge.m.d", 58: '[$-411]ggge"年"m"月"d"日"'},
    "zh_CN": {31: 'yyyy"年"m"月"d"日"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"'},
    "ko_KR": {31: 'yyyy"년" m"월" d"일"'},
}
EPOCH = dt.date(1899, 12, 30)
# Sandboxed Excel can save without a permission prompt only inside its container.
EXCEL_TMP = Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "tmp"


def _locale_probe() -> Any:
    spec = importlib.util.spec_from_file_location("excel_locale_probe", ROOT / "scripts" / "excel-locale-probe.py")
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _capture(elp: Any, xw: Any, tag: str, inputs: list[str], path: Path) -> None:
    elp._set_locale(elp.LOCALES[tag])
    elp._run(["open", "-g", "-a", "Microsoft Excel"], timeout=15, check=True)
    deadline = time.monotonic() + elp.WAIT_SECONDS
    while not elp._excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    app = book = None
    try:
        app, book = elp._open_workbook(xw)
        sheet = book.sheets[0]
        for row, text in enumerate(inputs, start=1):
            sheet.range((row, 1)).value = text
        book.save(str(path))
        # Saving renames the workbook, so the handle no longer resolves; quit instead.
        book = None
    finally:
        errors = elp._close_excel(app, book)
    if errors:
        raise RuntimeError(f"{tag}: Excel close failed: {errors}")


def read_sheet_cells(path: Path, tag: str) -> dict[tuple[int, int], dict[str, Any]]:
    """Every stored cell of the first sheet as {(row, col): {t, v, code}}, 1-based."""
    z = zipfile.ZipFile(path)
    styles = ET.fromstring(z.read("xl/styles.xml"))
    custom = {int(n.get("numFmtId")): n.get("formatCode") for n in styles.findall("m:numFmts/m:numFmt", NS)}
    xf_ids = [int(x.get("numFmtId", 0)) for x in styles.findall("m:cellXfs/m:xf", NS)]
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t")))
    cells: dict[tuple[int, int], dict[str, Any]] = {}
    for c in ET.fromstring(z.read("xl/worksheets/sheet1.xml")).iter(f"{{{NS['m']}}}c"):
        ref = c.get("r") or ""
        col = re.match(r"[A-Z]+", ref)
        if not col:
            continue
        kind = c.get("t", "n")
        v = c.find("m:v", NS)
        value = v.text if v is not None else None
        if kind == "s" and value is not None:
            value = shared[int(value)]
        inline = c.find("m:is", NS)
        if inline is not None:
            value = "".join(t.text or "" for t in inline.iter(f"{{{NS['m']}}}t"))
        if value is None:
            continue
        num_fmt_id = xf_ids[int(c.get("s", 0))]
        code = custom.get(num_fmt_id) or LOCALE_BUILTIN.get(tag, {}).get(num_fmt_id) or BUILTIN.get(num_fmt_id)
        if code is None:
            raise RuntimeError(f"{tag} {ref}: unknown built-in numFmtId {num_fmt_id}")
        col_index = 0
        for ch in col.group():
            col_index = col_index * 26 + ord(ch) - 64
        cells[(int(ref[len(col.group()):]), col_index)] = {"t": kind, "v": value, "code": code}
    return cells


def _read_cells(path: Path, tag: str) -> dict[int, dict[str, Any]]:
    return {row: cell for (row, col), cell in read_sheet_cells(path, tag).items() if col == 1}


def cell_record(cell: dict[str, Any]) -> dict[str, Any]:
    """A stored cell as {kind, value, format?}; General is left implicit."""
    if cell["t"] == "n":
        record: dict[str, Any] = {"kind": "number", "value": float(cell["v"])}
        if cell["code"] != "General":
            record["format"] = cell["code"]
        return record
    if cell["t"] == "b":
        return {"kind": "bool", "value": cell["v"] == "1"}
    if cell["t"] == "e":
        return {"kind": "error", "value": cell["v"]}
    return {"kind": "text", "value": cell["v"]}


def _case(text: str, cell: dict[str, Any] | None, year: int) -> dict[str, Any]:
    if cell is None:
        raise RuntimeError(f"no cell stored for {text!r}")
    case: dict[str, Any] = {"input": text, **cell_record(cell)}
    value = case["value"]
    if "format" in case and 61 <= value < 2958466 and (EPOCH + dt.timedelta(days=int(value))).year == year:
        case["currentYear"] = True
    return case


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--add", nargs="*", default=[], help="inputs to append to the case list")
    parser.add_argument("--locale", nargs="*", choices=sorted(LOCALES), help="subset; default is all seven")
    args = parser.parse_args()

    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    inputs = [case["input"] for case in next(iter(fixture["locales"].values()))]
    inputs += [text for text in args.add if text not in inputs]
    tags = args.locale or list(LOCALES)

    elp = _locale_probe()
    xw = elp._load_xlwings()
    if elp._excel_pids():
        print("Refusing to run: Excel is already running.", file=sys.stderr)
        return 3
    before = elp._pref_snapshot()
    year = dt.date.today().year
    try:
        with tempfile.TemporaryDirectory(dir=EXCEL_TMP) as tmp:
            for tag in tags:
                path = Path(tmp) / f"{tag}.xlsx"
                _capture(elp, xw, tag, inputs, path)
                cells = _read_cells(path, tag)
                fixture["locales"][LOCALES[tag]] = [
                    _case(text, cells.get(row), year) for row, text in enumerate(inputs, start=1)
                ]
                print(f"[{tag}] captured {len(inputs)} inputs", flush=True)
    finally:
        restore_errors = elp._restore_preferences(before)
        restored = elp._same_pref_state(before, elp._pref_snapshot())
        print(f"preferences restored: {restored} {restore_errors or ''}")
    fixture["captureYear"] = year
    FIXTURE.write_text(json.dumps(fixture, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE}")
    return 0 if restored and not restore_errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
