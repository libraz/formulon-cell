#!/usr/bin/env python3
"""Capture a small seven-locale Mac Excel probe without touching oracle goldens.

The full Formulon oracle capture changes the Excel process and writes repository
goldens. This probe is narrower: it launches a fresh Excel process for each
requested locale, evaluates a temporary workbook, records formula/value/display
observations, and restores the per-app language keys exactly.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import platform
import plistlib
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
UPSTREAM = ROOT.parent / "formulon"
DEFAULT_OUTPUT = ROOT / "backup" / "excel-locale-capture-2026-10-11.json"
EXCEL_DOMAIN = "com.microsoft.Excel"
PREF_KEYS = ("AppleLanguages", "AppleLocale")
WAIT_SECONDS = 35

LOCALES: dict[str, dict[str, Any]] = {
    "en_US": {"languages": ["en-US"], "locale": "en_US", "bcp47": "en-US"},
    "ja_JP": {"languages": ["ja-JP"], "locale": "ja_JP", "bcp47": "ja-JP"},
    "de_DE": {"languages": ["de-DE"], "locale": "de_DE", "bcp47": "de-DE"},
    "fr_FR": {"languages": ["fr-FR"], "locale": "fr_FR", "bcp47": "fr-FR"},
    "zh_CN": {"languages": ["zh-Hans-CN"], "locale": "zh_CN", "bcp47": "zh-CN"},
    "ko_KR": {"languages": ["ko-KR"], "locale": "ko_KR", "bcp47": "ko-KR"},
    "th_TH": {"languages": ["th-TH"], "locale": "th_TH", "bcp47": "th-TH"},
}

CASES: dict[str, tuple[str, dict[str, Any]]] = {
    "sum": ("=SUM(1,2)", {}),
    "true": ("=TRUE()", {}),
    "false": ("=FALSE()", {}),
    "if_true": ("=IF(TRUE,1,0)", {}),
    "if_false": ("=IF(FALSE,1,0)", {}),
    "dollar": ("=DOLLAR(1234.5,2)", {}),
    "usdollar": ("=USDOLLAR(1234.5,2)", {}),
    "number": ("=1234.5", {}),
    "date": ("=DATE(2024,1,5)", {}),
    "date_text": ('=TEXT(DATE(2024,1,5),"mmmm d, yyyy")', {}),
    "div0": ("=1/0", {}),
    "na": ("=NA()", {}),
    "eq_fullwidth_kana": ('="ガ"="ｶﾞ"', {}),
    "eq_katakana_hiragana": ('="カ"="か"', {}),
    "eq_umlaut": ('="ä"="a"', {}),
    "unique_kana": ('=ROWS(UNIQUE({"ガ";"ｶﾞ"}))', {}),
    "countif_kana": ('=COUNTIF(A20:A21,"ガ")', {"A20": "ガ", "A21": "ｶﾞ"}),
    "text_de_weekday": ('=TEXT(DATE(2024,1,2),"[$-de-DE]dddd")', {}),
}


def _run(
    args: list[str], *, timeout: float = 10, check: bool = False
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, capture_output=True, text=True, timeout=timeout, check=check)


def _now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def _json_value(value: Any) -> Any:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [_json_value(v) for v in value]
    if isinstance(value, dict):
        return {str(k): _json_value(v) for k, v in value.items()}
    return repr(value)


def _excel_pids() -> list[int]:
    try:
        result = _run(["pgrep", "-x", "Microsoft Excel"], timeout=3)
    except (OSError, subprocess.SubprocessError):
        return []
    pids: list[int] = []
    for line in result.stdout.splitlines():
        try:
            pids.append(int(line.strip()))
        except ValueError:
            pass
    return pids


def _load_xlwings() -> Any:
    try:
        import xlwings as xw  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"xlwings import failed: {exc}") from exc
    return xw


def _app_records(xw: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    try:
        apps: Iterable[Any] = list(xw.apps)
    except Exception as exc:
        return [{"error": f"xlwings app enumeration failed: {exc}"}]
    for app in apps:
        record: dict[str, Any] = {"repr": repr(app), "books": []}
        try:
            books = list(app.books)
        except Exception as exc:
            record["books_error"] = str(exc)
            books = []
        for book in books:
            item: dict[str, Any] = {"name": _json_value(getattr(book, "name", None))}
            try:
                item["fullname"] = _json_value(getattr(book, "fullname", None))
            except Exception as exc:
                item["fullname_error"] = str(exc)
            record["books"].append(item)
        records.append(record)
    return records


def _pref_export(domain: str) -> dict[str, Any]:
    try:
        result = _run(["defaults", "export", domain, "-"], timeout=5)
    except (OSError, subprocess.SubprocessError) as exc:
        return {"error": str(exc)}
    if result.returncode != 0:
        return {"error": (result.stderr or result.stdout).strip()}
    try:
        data = plistlib.loads(result.stdout.encode())
    except Exception as exc:
        return {"error": f"plist parse failed: {exc}"}
    out: dict[str, Any] = {}
    for key in PREF_KEYS:
        if key in data:
            out[key] = {
                "present": True,
                "value": _json_value(data[key]),
                "type": type(data[key]).__name__,
            }
        else:
            out[key] = {"present": False}
    return out


def _pref_snapshot() -> dict[str, Any]:
    return {"app": _pref_export(EXCEL_DOMAIN), "global": _pref_export("NSGlobalDomain")}


def _write_default(key: str, value: Any) -> None:
    if isinstance(value, list):
        _run(
            ["defaults", "write", EXCEL_DOMAIN, key, "-array", *[str(v) for v in value]],
            timeout=5,
            check=True,
        )
    elif isinstance(value, bool):
        _run(
            ["defaults", "write", EXCEL_DOMAIN, key, "-bool", "true" if value else "false"],
            timeout=5,
            check=True,
        )
    elif isinstance(value, int):
        _run(["defaults", "write", EXCEL_DOMAIN, key, "-int", str(value)], timeout=5, check=True)
    else:
        _run(["defaults", "write", EXCEL_DOMAIN, key, str(value)], timeout=5, check=True)


def _restore_preferences(snapshot: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    app_snapshot = snapshot.get("app", {})
    for key in PREF_KEYS:
        record = app_snapshot.get(key, {}) if isinstance(app_snapshot, dict) else {}
        try:
            if record.get("present"):
                _write_default(key, record.get("value"))
            else:
                result = _run(["defaults", "delete", EXCEL_DOMAIN, key], timeout=5)
                if result.returncode not in (0, 1):
                    raise RuntimeError((result.stderr or result.stdout).strip())
        except Exception as exc:
            errors.append(f"{key}: {exc}")
    return errors


def _set_locale(config: dict[str, Any]) -> None:
    _write_default("AppleLanguages", config["languages"])
    _write_default("AppleLocale", config["locale"])


def _system_events_probe() -> dict[str, Any]:
    script = r'''tell application "System Events"
if not (exists process "Microsoft Excel") then
    return "__NO_PROCESS__"
end if
tell process "Microsoft Excel"
    set menuNames to {}
    set windowNames to {}
    set ribbonNames to {}
    try
        repeat with itemRef in every menu bar item of menu bar 1
            set end of menuNames to (name of itemRef as text)
        end repeat
    end try
    try
        repeat with windowRef in every window
            set end of windowNames to (name of windowRef as text)
        end repeat
    end try
    try
        set allElements to entire contents of window 1
        set ribbonCount to 0
        repeat with elementRef in allElements
            if ribbonCount is greater than or equal to 80 then exit repeat
            try
                set elementRole to role of elementRef as text
                if elementRole is "AXRadioButton" or elementRole is "AXButton" then
                    set elementName to ""
                    try
                        set elementName to name of elementRef as text
                    end try
                    if elementName is not "" then
                        set end of ribbonNames to elementRole & tab & elementName
                        set ribbonCount to ribbonCount + 1
                    end if
                end if
            end try
        end repeat
    end try
end tell
set AppleScript's text item delimiters to "\n"
return "MENUS\n" & (menuNames as text) & "\nWINDOWS\n" & (windowNames as text) & "\nRIBBON\n" & (ribbonNames as text)
end tell'''
    try:
        result = _run(["osascript", "-e", script], timeout=12)
    except (OSError, subprocess.SubprocessError) as exc:
        return {"status": "error", "error": str(exc)}
    if result.returncode != 0:
        message = (result.stderr or result.stdout).strip()
        lowered = message.lower()
        status = (
            "permission_denied"
            if "not allowed" in lowered or "assistive" in lowered
            else "error"
        )
        return {"status": status, "error": message}
    output = result.stdout.strip()
    if output == "__NO_PROCESS__":
        return {"status": "no_process", "menus": [], "windows": [], "ribbon": []}
    section = ""
    menus: list[str] = []
    windows: list[str] = []
    ribbon: list[dict[str, str]] = []
    for line in output.splitlines():
        if line == "MENUS":
            section = "menus"
        elif line == "WINDOWS":
            section = "windows"
        elif line == "RIBBON":
            section = "ribbon"
        elif line and section == "menus":
            menus.append(line)
        elif line and section == "windows":
            windows.append(line)
        elif line and section == "ribbon":
            role, separator, name = line.partition("\t")
            ribbon.append({"role": role, "name": name if separator else line})
    return {"status": "ok", "menus": menus, "windows": windows, "ribbon": ribbon, "raw": output}


def _read_property(obj: Any, attr: str) -> tuple[str, Any]:
    try:
        value = getattr(obj, attr)
    except Exception as exc:
        return "error", str(exc)
    try:
        value = value() if callable(value) else value
    except Exception as exc:
        return "error", str(exc)
    return "value", _json_value(value)


def _read_formula_local(cell: Any) -> dict[str, Any]:
    """Keep blank formula_local distinct from an Excel formula result."""
    try:
        api = cell.api
        reference = getattr(api, "formula_local")
        value = reference.get() if hasattr(reference, "get") else reference
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
    value = _json_value(value)
    if isinstance(value, str) and value:
        return {"status": "read", "value": value}
    return {"status": "blank", "value": value}


def _display_text(cell: Any) -> dict[str, Any]:
    try:
        api = cell.api
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
    for attr in ("string_value", "text", "displayed_value", "formatted_text"):
        try:
            value = getattr(api, attr)
            value = value() if callable(value) else value
        except Exception:
            continue
        if isinstance(value, str):
            return {"status": "read", "source": attr, "value": value}
    return {"status": "blank"}


def _formula_case(app: Any, sheet: Any, cell: Any, formula: str, setup: dict[str, Any]) -> dict[str, Any]:
    record: dict[str, Any] = {"input": formula}
    try:
        cell.clear_contents()
        for addr, value in setup.items():
            sheet.range(addr).value = value
        try:
            cell.number_format = "General"
        except Exception:
            pass
        cell.formula2 = formula
        app.calculate()
    except Exception as exc:
        record["status"] = "assignment_error"
        record["error"] = str(exc)
        return record
    record["status"] = "ok"
    status, value = _read_property(cell, "value")
    record["value_status"] = status
    record["value"] = value
    for attr in ("formula2", "formula"):
        status, value = _read_property(cell, attr)
        record[attr + "_status"] = status
        record[attr] = value
    record["formula_local"] = _read_formula_local(cell)
    record["display_text"] = _display_text(cell)
    return record


def _open_workbook(xw: Any) -> tuple[Any, Any]:
    try:
        apps = list(xw.apps)
    except Exception:
        apps = []
    app = apps[-1] if apps else None
    if app is None:
        app = xw.App(visible=True, add_book=False)
    for _ in range(10):
        try:
            return app, app.books.add()
        except Exception:
            time.sleep(1)
    raise RuntimeError("Excel did not accept a new temporary workbook")


def _version(app: Any) -> str:
    try:
        return str(app.version or "")
    except Exception:
        pass
    try:
        api = app.api
        value = api.version() if callable(api.version) else api.version
        return str(value)
    except Exception:
        return ""


def _observed_locale(app: Any) -> dict[str, Any]:
    result: dict[str, Any] = {}
    try:
        from drivers._locale import COUNTRY_CODE_TO_BCP47, detect_mac_excel_locale  # type: ignore

        result["mac_defaults"] = detect_mac_excel_locale() or ""
        result["country_code_map"] = ""
        api = app.api
        for attr in ("International", "international"):
            try:
                accessor = getattr(api, attr)
                value = accessor(1) if callable(accessor) else accessor[1]
                code = int(value)
                result["country_code"] = code
                result["country_code_map"] = COUNTRY_CODE_TO_BCP47.get(code, "")
                break
            except Exception:
                continue
    except Exception as exc:
        result["error"] = str(exc)
    return result


def _quit_excel_without_bridge() -> list[str]:
    try:
        result = _run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to quit saving no'],
            timeout=15,
        )
        if result.returncode:
            return [(result.stderr or result.stdout).strip()]
    except Exception as exc:
        return [str(exc)]
    return []


def _close_excel(app: Any, book: Any) -> list[str]:
    errors: list[str] = []
    if book is not None:
        try:
            book.close()
        except Exception as exc:
            errors.append(f"temporary workbook close: {exc}")
    if app is not None:
        try:
            app.quit()
        except Exception as exc:
            errors.append(f"Excel quit: {exc}")
    else:
        errors.extend(_quit_excel_without_bridge())
    if _excel_pids():
        errors.extend(_quit_excel_without_bridge())
    deadline = time.monotonic() + WAIT_SECONDS
    while _excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    if _excel_pids():
        errors.append(f"Excel process remained: {_excel_pids()}")
    return errors


def _launch_and_capture(tag: str, config: dict[str, Any], xw: Any) -> dict[str, Any]:
    _set_locale(config)
    opened_at = _now()
    try:
        _run(["open", "-a", "Microsoft Excel"], timeout=15, check=True)
    except Exception as exc:
        return {"tag": tag, "status": "launch_error", "error": str(exc), "opened_at": opened_at}
    deadline = time.monotonic() + WAIT_SECONDS
    while not _excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    if not _excel_pids():
        return {"tag": tag, "status": "launch_timeout", "opened_at": opened_at}
    app = None
    book = None
    result: dict[str, Any] = {
        "tag": tag,
        "requested": {
            "AppleLanguages": config["languages"],
            "AppleLocale": config["locale"],
            "bcp47": config["bcp47"],
        },
        "opened_at": opened_at,
        "status": "started",
    }
    try:
        app, book = _open_workbook(xw)
        sheet = book.sheets[0]
        cell = sheet.range("A1")
        result["excel_version"] = _version(app)
        result["observed_locale"] = _observed_locale(app)
        result["ui"] = _system_events_probe()
        result["cases"] = {
            name: _formula_case(app, sheet, cell, formula, setup)
            for name, (formula, setup) in CASES.items()
        }
        result["status"] = "ok"
    except Exception as exc:
        result["status"] = "capture_error"
        result["error"] = str(exc)
    finally:
        result["close_errors"] = _close_excel(app, book)
    return result


def _same_pref_state(before: dict[str, Any], after: dict[str, Any]) -> bool:
    return before.get("app") == after.get("app")


def _capture_is_complete(item: dict[str, Any]) -> tuple[bool, list[str]]:
    problems: list[str] = []
    if item.get("status") != "ok":
        problems.append(f"status={item.get('status')!r}")
    if not item.get("excel_version"):
        problems.append("excel_version is empty")
    observed = item.get("observed_locale")
    if not isinstance(observed, dict) or not observed.get("mac_defaults"):
        problems.append("observed locale is empty")
    elif observed["mac_defaults"] != LOCALES.get(item.get("tag"), {}).get("bcp47"):
        problems.append("observed locale differs from requested locale")
    ui = item.get("ui")
    if not isinstance(ui, dict) or ui.get("status") != "ok":
        problems.append("UI accessibility probe did not succeed")
    elif not ui.get("menus") or len([e for e in ui.get("ribbon", []) if e.get("role") == "AXRadioButton"]) < 9:
        problems.append("UI probe did not return the core ribbon tabs")
    cases = item.get("cases")
    if not isinstance(cases, dict) or set(cases) != set(CASES):
        problems.append("case set is incomplete")
        cases = cases if isinstance(cases, dict) else {}
    for name in CASES:
        case = cases.get(name)
        if not isinstance(case, dict):
            problems.append(f"{name}: missing case record")
            continue
        if case.get("status") != "ok":
            problems.append(f"{name}: status={case.get('status')!r}")
        for field in ("value_status", "formula2_status", "formula_status"):
            if case.get(field) != "value":
                problems.append(f"{name}: {field}={case.get(field)!r}")
        if not isinstance(case.get("formula_local"), dict) or case["formula_local"].get("status") != "read":
            problems.append(f"{name}: formula_local was not read")
        if not isinstance(case.get("display_text"), dict) or case["display_text"].get("status") != "read":
            problems.append(f"{name}: display text was not read")
    return not problems, problems


def _verify_restored_launch(xw: Any) -> dict[str, Any]:
    if _excel_pids():
        return {"status": "skipped", "reason": "Excel process remained from capture"}
    try:
        _run(["open", "-a", "Microsoft Excel"], timeout=15, check=True)
    except Exception as exc:
        return {"status": "launch_error", "error": str(exc)}
    deadline = time.monotonic() + WAIT_SECONDS
    while not _excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    if not _excel_pids():
        return {"status": "launch_timeout"}
    app = None
    try:
        apps = list(xw.apps)
        app = apps[-1] if apps else xw.App(visible=True, add_book=False)
        return {"status": "ok", "excel_version": _version(app), "observed_locale": _observed_locale(app)}
    except Exception as exc:
        return {"status": "probe_error", "error": str(exc)}
    finally:
        _close_excel(app, None)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--locale", nargs="*", choices=sorted(LOCALES), help="subset; default is all seven")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> int:
    if platform.system() != "Darwin":
        print("Excel locale probe requires macOS", file=sys.stderr)
        return 2
    args = _parse_args()
    tags = args.locale or list(LOCALES)
    xw = _load_xlwings()
    sys.path.insert(0, str(UPSTREAM / "tools" / "oracle"))
    startup_pids = _excel_pids()
    startup_apps = _app_records(xw)
    if startup_pids or startup_apps:
        print("Refusing to change Excel locale: Excel or an open workbook is already present.", file=sys.stderr)
        print(json.dumps({"pids": startup_pids, "apps": startup_apps}, ensure_ascii=False), file=sys.stderr)
        return 3
    prefs_before = _pref_snapshot()
    if "error" in prefs_before.get("app", {}):
        print(f"Could not snapshot {EXCEL_DOMAIN}: {prefs_before['app']['error']}", file=sys.stderr)
        return 4
    data: dict[str, Any] = {
        "schema": "formulon-cell.excel-locale-probe/v1",
        "captured_at": _now(),
        "host": {
            "platform": platform.platform(),
            "python": sys.version,
            "xlwings": getattr(xw, "__version__", ""),
        },
        "startup": {"excel_pids": startup_pids, "xlwings_apps": startup_apps},
        "preferences_before": prefs_before,
        "requested_locales": tags,
        "locales": [],
        "automation_probe": _system_events_probe(),
    }
    restore_errors: list[str] = []
    try:
        for tag in tags:
            print(f"[{tag}] launching Excel", flush=True)
            result = _launch_and_capture(tag, LOCALES[tag], xw)
            data["locales"].append(result)
            print(f"[{tag}] {result.get('status')}", flush=True)
            if result.get("close_errors") and _excel_pids():
                data["aborted_after"] = tag
                break
    finally:
        restore_errors.extend(_restore_preferences(prefs_before))
        data["restoration_launch"] = _verify_restored_launch(xw)
        restore_errors.extend(_restore_preferences(prefs_before))
        data["preferences_after"] = _pref_snapshot()
        data["cleanup"] = {
            "restore_errors": restore_errors,
            "preferences_restored_exactly": _same_pref_state(prefs_before, data["preferences_after"]),
            "excel_pids_after": _excel_pids(),
        }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {args.output}")
    capture_checks: dict[str, list[str]] = {}
    for item in data["locales"]:
        complete, problems = _capture_is_complete(item)
        if item.get("close_errors"):
            problems = [*problems, "Excel close reported an error"]
            complete = False
        if not complete:
            capture_checks[str(item.get("tag", "?"))] = problems
    data["cleanup"]["capture_checks"] = capture_checks
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    captures_ok = len(data["locales"]) == len(tags) and not capture_checks
    restoration_ok = data["restoration_launch"].get("status") == "ok"
    return (
        0
        if captures_ok
        and restoration_ok
        and data["cleanup"]["preferences_restored_exactly"]
        and not restore_errors
        else 1
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)

