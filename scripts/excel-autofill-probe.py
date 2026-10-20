#!/usr/bin/env python3
"""Capture how Mac Excel extends a range with AutoFill, per locale, into the test fixture.

Each case's source inputs are written into its own block of a temporary
workbook through the object model (no synthetic keystrokes) and extended with
`Range.autofill`. The application's custom lists, which drive text series, are
recorded alongside. The workbook is saved and the stored type, value and number
format of every source and filled cell are read back from the xlsx. Locale
switching and restoration reuse scripts/excel-locale-probe.py. The fixture is
both the case list and the oracle that
tests/unit/commands/fill-excel-oracle.test.ts checks against.

Run with a Python that has xlwings, while Excel is not running, then format
the fixture with `yarn lint:fix`:
  python3 scripts/excel-autofill-probe.py [--cases SEED.json] [--locale ja_JP ...]
"""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import json
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "packages" / "formulon-cell" / "tests" / "fixtures" / "excel-autofill.json"
# Down/up cases take one column each; right/left cases take one row each below them.
FIRST_ROW_CASE_ROW = 60


def _load(name: str, file: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / file)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _layout(cases: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Source and filled cell coordinates (1-based) of each case, in fill order."""
    out = []
    vertical = horizontal = 0
    for case in cases:
        k, n, d = len(case["source"]), case["count"], case["dir"]
        lead = n if d in ("up", "left") else 0
        src = list(range(lead + 1, lead + k + 1))
        fill = list(range(k + 1, k + n + 1)) if lead == 0 else list(range(n, 0, -1))
        if d in ("down", "up"):
            vertical += 1
            col = vertical
            out.append({"src": [(r, col) for r in src], "fill": [(r, col) for r in fill],
                        "range": ((1, col), (k + n, col))})
        else:
            row = FIRST_ROW_CASE_ROW + horizontal
            horizontal += 1
            out.append({"src": [(row, c) for c in src], "fill": [(row, c) for c in fill],
                        "range": ((row, 1), (row, k + n))})
    return out


def _capture(elp: Any, xw: Any, tag: str, cases: list[dict[str, Any]], layout: list[dict[str, Any]],
             path: Path) -> list[list[str]]:
    """Fill every case, save to `path`, and return the application's custom lists."""
    elp._set_locale(elp.LOCALES[tag])
    elp._run(["open", "-g", "-a", "Microsoft Excel"], timeout=15, check=True)
    deadline = time.monotonic() + elp.WAIT_SECONDS
    while not elp._excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    app = book = None
    try:
        app, book = elp._open_workbook(xw)
        lists = [[str(item) for item in app.api.get_custom_list_contents(list_num=i)]
                 for i in range(1, app.api.custom_list_count.get() + 1)]
        sheet = book.sheets[0]
        for case, place in zip(cases, layout):
            for text, cell in zip(case["source"], place["src"]):
                sheet.range(cell).value = text
            src = sheet.range(place["src"][0], place["src"][-1])
            src.autofill(sheet.range(*place["range"]), case.get("type", "fill_default"))
        book.save(str(path))
        # Saving renames the workbook, so the handle no longer resolves; quit instead.
        book = None
    finally:
        errors = elp._close_excel(app, book)
    if errors:
        raise RuntimeError(f"{tag}: Excel close failed: {errors}")
    return lists


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cases", type=Path, help="seed case list; default is the fixture's own")
    parser.add_argument("--locale", nargs="*", help="subset; default is all seven")
    args = parser.parse_args()

    elp = _load("excel_locale_probe", "excel-locale-probe.py")
    inp = _load("excel_input_probe", "excel-input-probe.py")
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8")) if FIXTURE.exists() else {"locales": {}}
    if args.cases:
        fixture["cases"] = json.loads(args.cases.read_text(encoding="utf-8"))
    cases = fixture["cases"]
    layout = _layout(cases)
    tags = args.locale or list(inp.LOCALES)

    xw = elp._load_xlwings()
    if elp._excel_pids():
        print("Refusing to run: Excel is already running.", file=sys.stderr)
        return 3
    before = elp._pref_snapshot()
    try:
        with tempfile.TemporaryDirectory(dir=inp.EXCEL_TMP) as tmp:
            for tag in tags:
                path = Path(tmp) / f"{tag}.xlsx"
                lists = _capture(elp, xw, tag, cases, layout, path)
                fixture.setdefault("customLists", {})[inp.LOCALES[tag]] = lists
                cells = inp.read_sheet_cells(path, tag)

                def record(at: tuple[int, int]) -> dict[str, Any]:
                    cell = cells.get(at)
                    return inp.cell_record(cell) if cell else {"kind": "blank"}

                fixture["locales"][inp.LOCALES[tag]] = {
                    case["id"]: {"source": [record(at) for at in place["src"]],
                                 "fill": [record(at) for at in place["fill"]]}
                    for case, place in zip(cases, layout)
                }
                print(f"[{tag}] captured {len(cases)} cases", flush=True)
    finally:
        restore_errors = elp._restore_preferences(before)
        restored = elp._same_pref_state(before, elp._pref_snapshot())
        print(f"preferences restored: {restored} {restore_errors or ''}")
    fixture["captureYear"] = dt.date.today().year
    FIXTURE.write_text(json.dumps(fixture, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE}")
    return 0 if restored and not restore_errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
