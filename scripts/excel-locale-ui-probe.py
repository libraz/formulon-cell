#!/usr/bin/env python3
"""Capture bounded Mac Excel UI, formula, input, and collation evidence.

This probe owns only the Excel processes that it starts.  It refuses to touch
an existing Excel process or workbook, changes only the Excel app's language
defaults, writes one incremental JSON artifact, and never writes oracle
goldens.  It is intentionally bounded: a missing AX surface is recorded with
an explicit status instead of being guessed from a localized label.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import platform
import plistlib
import re
import signal
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
UPSTREAM = ROOT.parent / "formulon"
DEFAULT_OUTPUT = ROOT / "backup" / "excel-locale-ui-capture-2026-10-11.json"
EXCEL_DOMAIN = "com.microsoft.Excel"
PREF_KEYS = ("AppleLanguages", "AppleLocale")
WAIT_SECONDS = 35
AX_FIELD_SEPARATOR = "\x1f"

LOCALES: dict[str, dict[str, Any]] = {
    "en_US": {"languages": ["en-US"], "locale": "en_US", "bcp47": "en-US"},
    "ja_JP": {"languages": ["ja-JP"], "locale": "ja_JP", "bcp47": "ja-JP"},
    "de_DE": {"languages": ["de-DE"], "locale": "de_DE", "bcp47": "de-DE"},
    "fr_FR": {"languages": ["fr-FR"], "locale": "fr_FR", "bcp47": "fr-FR"},
    "zh_CN": {"languages": ["zh-Hans-CN"], "locale": "zh_CN", "bcp47": "zh-CN"},
    "ko_KR": {"languages": ["ko-KR"], "locale": "ko_KR", "bcp47": "ko-KR"},
    "th_TH": {"languages": ["th-TH"], "locale": "th_TH", "bcp47": "th-TH"},
}

CORE_TABS = ("home", "insert", "draw", "page_layout", "formulas", "data", "review", "view", "automate")
TAB_LABELS: dict[str, dict[str, str]] = {
    "en_US": {
        "home": "Home",
        "insert": "Insert",
        "draw": "Draw",
        "page_layout": "Page Layout",
        "formulas": "Formulas",
        "data": "Data",
        "review": "Review",
        "view": "View",
        "automate": "Automate",
    },
    "ja_JP": {
        "home": "ホーム",
        "insert": "挿入",
        "draw": "描画",
        "page_layout": "ページ レイアウト",
        "formulas": "数式",
        "data": "データ",
        "review": "校閲",
        "view": "表示",
        "automate": "自動化",
    },
    "de_DE": {
        "home": "Start",
        "insert": "Einfügen",
        "draw": "Zeichnen",
        "page_layout": "Seitenlayout",
        "formulas": "Formeln",
        "data": "Daten",
        "review": "Überprüfen",
        "view": "Ansicht",
        "automate": "Automatisieren",
    },
    "fr_FR": {
        "home": "Accueil",
        "insert": "Insérer",
        "draw": "Dessin",
        "page_layout": "Mise en page",
        "formulas": "Formules",
        "data": "Données",
        "review": "Révision",
        "view": "Affichage",
        "automate": "Automatisation",
    },
    "zh_CN": {
        "home": "开始",
        "insert": "插入",
        "draw": "绘图",
        "page_layout": "页面布局",
        "formulas": "公式",
        "data": "数据",
        "review": "审阅",
        "view": "视图",
        "automate": "自动化",
    },
    "ko_KR": {
        "home": "홈",
        "insert": "삽입",
        "draw": "그리기",
        "page_layout": "페이지 레이아웃",
        "formulas": "수식",
        "data": "데이터",
        "review": "검토",
        "view": "보기",
        "automate": "자동화",
    },
    "th_TH": {
        "home": "หน้าแรก",
        "insert": "แทรก",
        "draw": "วาด",
        "page_layout": "เค้าโครงหน้ากระดาษ",
        "formulas": "สูตร",
        "data": "ข้อมูล",
        "review": "รีวิว",
        "view": "มุมมอง",
        "automate": "อัตโนมัติ",
    },
}

FONT_STYLE_LABELS: dict[str, dict[str, tuple[str, ...]]] = {
    "en_US": {"bold": ("Bold",), "italic": ("Italic",)},
    "ja_JP": {"bold": ("太字",), "italic": ("斜体",)},
    "de_DE": {"bold": ("Fett",), "italic": ("Kursiv",)},
    "fr_FR": {"bold": ("Gras",), "italic": ("Italique",)},
    "zh_CN": {"bold": ("粗体", "加粗"), "italic": ("斜体", "倾斜")},
    "ko_KR": {"bold": ("굵게", "굵은 글꼴"), "italic": ("기울임", "기울임꼴")},
    "th_TH": {"bold": ("ตัวหนา",), "italic": ("ตัวเอียง",)},
}
FONT_PICKER_TOKENS = (
    "font",
    "police",
    "schrift",
    "フォント",
    "字体",
    "글꼴",
    "แบบอักษร",
)

FORMULA_CASES: dict[str, tuple[str, dict[str, Any]]] = {
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

SORT_GROUPS: dict[str, list[str]] = {
    "kana_width": ["ガ", "ｶﾞ", "カ", "か"],
    "accent_case": ["ä", "a", "Ä", "A"],
    "german_sharp_s": ["ß", "ss", "SS", "s"],
    "zh_common": ["中", "文", "阿", "一"],
    "thai": ["ก", "ข", "ค", "ง", "จ"],
}

MENU_TOKENS: dict[str, dict[str, list[str]]] = {
    "name_manager": {
        "en_US": ["Name Manager"],
        "ja_JP": ["名前の管理"],
        "de_DE": ["Namens-Manager", "Namen-Manager"],
        "fr_FR": ["Gestionnaire de noms"],
        "zh_CN": ["名称管理器"],
        "ko_KR": ["이름 관리자"],
        "th_TH": ["ตัวจัดการชื่อ"],
    },
    "data_validation": {
        "en_US": ["Data Validation"],
        "ja_JP": ["データの入力規則", "入力規則"],
        "de_DE": ["Datenüberprüfung"],
        "fr_FR": ["Validation des données"],
        "zh_CN": ["数据验证"],
        "ko_KR": ["데이터 유효성 검사"],
        "th_TH": ["การตรวจสอบข้อมูล"],
    },
    "conditional_formatting": {
        "en_US": ["Conditional Formatting"],
        "ja_JP": ["条件付き書式"],
        "de_DE": ["Bedingte Formatierung"],
        "fr_FR": ["Mise en forme conditionnelle"],
        "zh_CN": ["条件格式"],
        "ko_KR": ["조건부 서식"],
        "th_TH": ["การจัดรูปแบบตามเงื่อนไข"],
    },
    "page_setup": {
        "en_US": ["Page Setup"],
        "ja_JP": ["ページ設定"],
        "de_DE": ["Seite einrichten", "Seite einrichten"],
        "fr_FR": ["Mise en page"],
        "zh_CN": ["页面设置"],
        "ko_KR": ["페이지 설정"],
        "th_TH": ["ตั้งค่าหน้ากระดาษ"],
    },
    "text_to_columns": {
        "en_US": ["Text to Columns"],
        "ja_JP": ["区切り位置"],
        "de_DE": ["Text in Spalten"],
        "fr_FR": ["Convertir"],
        "zh_CN": ["分列"],
        "ko_KR": ["텍스트 나누기"],
        "th_TH": ["ข้อความเป็นคอลัมน์"],
    },
    "goal_seek": {
        "en_US": ["Goal Seek"],
        "ja_JP": ["ゴール シーク"],
        "de_DE": ["Zielwertsuche"],
        "fr_FR": ["Valeur cible"],
        "zh_CN": ["单变量求解"],
        "ko_KR": ["목표값 찾기"],
        "th_TH": ["แสวงหาผลลัพธ์"],
    },
    "sort": {
        "en_US": ["Sort A to Z", "Sort Ascending", "Sort"],
        "ja_JP": ["昇順", "並べ替え"],
        "de_DE": ["Aufsteigend sortieren", "Sortieren"],
        "fr_FR": ["Trier de A à Z", "Trier"],
        "zh_CN": ["升序", "排序"],
        "ko_KR": ["오름차순", "정렬"],
        "th_TH": ["เรียงลำดับจากน้อยไปมาก", "เรียงลำดับ"],
    },
    "filter": {
        "en_US": ["Filter"],
        "ja_JP": ["フィルター"],
        "de_DE": ["Filter"],
        "fr_FR": ["Filtrer"],
        "zh_CN": ["筛选"],
        "ko_KR": ["필터"],
        "th_TH": ["ตัวกรอง"],
    },
    "paste_special": {
        "en_US": ["Paste Special"],
        "ja_JP": ["形式を選択して貼り付け"],
        "de_DE": ["Inhalte einfügen"],
        "fr_FR": ["Collage spécial"],
        "zh_CN": ["选择性粘贴"],
        "ko_KR": ["선택하여 붙여넣기"],
        "th_TH": ["วางแบบพิเศษ"],
    },
    "print": {
        "en_US": ["Print"],
        "ja_JP": ["印刷"],
        "de_DE": ["Drucken"],
        "fr_FR": ["Imprimer"],
        "zh_CN": ["打印"],
        "ko_KR": ["인쇄"],
        "th_TH": ["พิมพ์"],
    },
    "table": {
        "en_US": ["Table"],
        "ja_JP": ["テーブル"],
        "de_DE": ["Tabelle"],
        "fr_FR": ["Tableau"],
        "zh_CN": ["表格"],
        "ko_KR": ["표"],
        "th_TH": ["ตาราง"],
    },
    "pivot": {
        "en_US": ["PivotTable"],
        "ja_JP": ["ピボットテーブル"],
        "de_DE": ["PivotTable"],
        "fr_FR": ["Tableau croisé dynamique"],
        "zh_CN": ["数据透视表"],
        "ko_KR": ["피벗 테이블"],
        "th_TH": ["PivotTable"],
    },
}

SHORTCUT_DIALOGS = (
    ("dialog/format-cells", "format_cells"),
    ("dialog/find", "find"),
    ("dialog/replace", "replace"),
    ("dialog/go-to", "go_to"),
    ("dialog/function-arguments", "function_arguments"),
)

RIBBON_COMMANDS: dict[str, dict[str, tuple[str, ...]]] = {
    "name_manager": {
        "tab": ("formulas",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["name_manager"].items()},
        "roles": ("AXButton",),
        "action": ("AXPress",),
    },
    "data_validation": {
        "tab": ("data",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["data_validation"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXShowMenu",),
    },
    "conditional_formatting": {
        "tab": ("home",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["conditional_formatting"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXShowMenu",),
    },
    "page_setup": {
        "tab": ("page_layout",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["page_setup"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
    "text_to_columns": {
        "tab": ("data",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["text_to_columns"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
    "goal_seek": {
        "tab": ("data",),
        "tokens": {"en_US": ("What-If Analysis",), "ja_JP": ("What-If 分析", "What-If"), "de_DE": ("Was-wäre-wenn-Analyse",), "fr_FR": ("Analyse de scénarios",), "zh_CN": ("模拟分析",), "ko_KR": ("가상 분석",), "th_TH": ("การวิเคราะห์แบบ What-If",)},
        "roles": ("AXMenuButton", "AXButton"),
        "action": ("AXShowMenu",),
    },
    "sort": {
        "tab": ("data",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["sort"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
    "filter": {
        "tab": ("data",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["filter"].items()},
        "roles": ("AXCheckBox", "AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
    "table": {
        "tab": ("insert",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["table"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
    "pivot": {
        "tab": ("insert",),
        "tokens": {tag: tuple(values) for tag, values in MENU_TOKENS["pivot"].items()},
        "roles": ("AXButton", "AXMenuButton"),
        "action": ("AXPress",),
    },
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
    out: list[int] = []
    for line in result.stdout.splitlines():
        try:
            out.append(int(line.strip()))
        except ValueError:
            pass
    return out


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


def _load_xlwings() -> Any:
    try:
        import xlwings as xw  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"xlwings import failed: {exc}") from exc
    return xw


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


def _same_pref_state(before: dict[str, Any], after: dict[str, Any]) -> bool:
    return before == after


def _norm(value: Any) -> str:
    text = str(value or "").strip().casefold()
    return re.sub(r"[^0-9a-z\u0080-\uffff]+", "", text)


def _safe_name(value: Any, limit: int = 160) -> str:
    text = str(value or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return text[:limit]


def _osascript(script: str, args: list[str] | None = None, *, timeout: float = 12) -> tuple[str, str]:
    command = ["osascript", "-e", script]
    if args:
        command.extend(args)
    result = _run(command, timeout=timeout)
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout).strip() or f"osascript exit {result.returncode}")
    return result.stdout.strip(), result.stderr.strip()


AX_DUMP_SCRIPT = r'''
-- Keep this first pass deliberately small.  Reading AX value, help, and
-- action attributes for every descendant makes Excel's System Events bridge
-- hang on some ribbon versions.  Names and roles are enough to map tabs and
-- controls; targeted dialog probes can request additional attributes later.
tell application "System Events"
    if not (exists process "Microsoft Excel") then return "__NO_PROCESS__"
    tell process "Microsoft Excel"
        set fs to ASCII character 31
        set rs to ASCII character 30
        set linesOut to {}
        try
            repeat with itemRef in every menu bar item of menu bar 1
                try
                    set end of linesOut to "MENU" & fs & (name of itemRef as text)
                end try
            end repeat
        end try
        try
            repeat with w in every window
                try
                    set end of linesOut to "WINDOW" & fs & (name of w as text) & fs & "AXWindow"
                end try
            end repeat
        end try
        try
            set allElements to entire contents of window 1
            set countElements to 0
            repeat with e in allElements
                if countElements is greater than or equal to 160 then exit repeat
                try
                    set r to role of e as text
                    if r is in {"AXButton", "AXRadioButton", "AXMenuButton", "AXPopUpButton", "AXCheckBox", "AXTextField", "AXComboBox", "AXStaticText", "AXMenuBarItem", "AXMenuItem", "AXMenu", "AXSheet", "AXDialog"} then
                        set n to ""
                        try
                            set n to name of e as text
                        end try
                        set end of linesOut to "ELEMENT" & fs & r & fs & "" & fs & n & fs & "" & fs & "" & fs & "" & fs & "" & fs & "" & fs & "" & fs & "" & fs & ""
                        set countElements to countElements + 1
                    end if
                end try
            end repeat
        on error errText
            set end of linesOut to "ERROR" & fs & (errText as text)
        end try
        set AppleScript's text item delimiters to rs
        return linesOut as text
    end tell
end tell
'''


AX_DETAIL_SCRIPT = r'''
on safeText(v)
    try
        if v is missing value then return ""
        set t to v as text
    on error
        set t to ""
    end try
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to {return, linefeed, tab, "|", (ASCII character 30), (ASCII character 31)}
    set pieces to text items of t
    set AppleScript's text item delimiters to " "
    set t to pieces as text
    set AppleScript's text item delimiters to oldDelims
    return t
end safeText

tell application "System Events"
    if not (exists process "Microsoft Excel") then return "__NO_PROCESS__"
    tell process "Microsoft Excel"
        set fs to ASCII character 31
        set rs to ASCII character 30
        set linesOut to {}
        try
            set allElements to entire contents of window 1
            set detailCount to 0
            repeat with e in allElements
                if detailCount is greater than or equal to 40 then exit repeat
                try
                    set r to role of e as text
                    if r is in {"AXComboBox", "AXPopUpButton", "AXCheckBox"} then
                        set n to my safeText(name of e)
                        set d to my safeText(description of e)
                        set h to my safeText(help of e)
                        set v to my safeText(value of e)
                        set end of linesOut to "DETAIL" & fs & r & fs & n & fs & d & fs & h & fs & v
                        set detailCount to detailCount + 1
                    end if
                end try
            end repeat
        on error errText
            set end of linesOut to "ERROR" & fs & (my safeText(errText))
        end try
        set AppleScript's text item delimiters to rs
        return linesOut as text
    end tell
end tell
'''


def _parse_ax_dump(raw: str) -> dict[str, Any]:
    if raw == "__NO_PROCESS__":
        return {"status": "no_process", "windows": [], "menus": [], "elements": [], "errors": []}
    windows: list[dict[str, Any]] = []
    menus: list[str] = []
    elements: list[dict[str, Any]] = []
    errors: list[str] = []
    for line in raw.split("\x1e"):
        fields = line.split(AX_FIELD_SEPARATOR)
        if not fields:
            continue
        if fields[0] == "MENU" and len(fields) >= 2:
            menus.append(fields[1])
        elif fields[0] == "WINDOW" and len(fields) >= 3:
            windows.append({"name": fields[1], "role": fields[2]})
        elif fields[0] == "ELEMENT" and len(fields) >= 12:
            elements.append(
                {
                    "role": fields[1],
                    "subrole": fields[2],
                    "name": fields[3],
                    "description": fields[4],
                    "help": fields[5],
                    "value": fields[6],
                    "enabled": fields[7],
                    "selected": fields[8],
                    "position": fields[9],
                    "size": fields[10],
                    "actions": fields[11],
                }
            )
        elif fields[0] == "ERROR":
            errors.append(" | ".join(fields[1:]))
    return {"status": "ok", "windows": windows, "menus": menus, "elements": elements, "errors": errors, "raw": raw}


def _ax_dump() -> dict[str, Any]:
    try:
        raw, _ = _osascript(AX_DUMP_SCRIPT, timeout=15)
    except Exception as exc:
        message = str(exc)
        lowered = message.casefold()
        status = "permission_denied" if "not allowed" in lowered or "assistive" in lowered else "error"
        return {"status": status, "windows": [], "elements": [], "errors": [message]}
    return _parse_ax_dump(raw)


def _parse_ax_detail_dump(raw: str) -> dict[str, Any]:
    if raw == "__NO_PROCESS__":
        return {"status": "no_process", "elements": [], "errors": []}
    elements: list[dict[str, Any]] = []
    errors: list[str] = []
    for line in raw.split("\x1e"):
        fields = line.split(AX_FIELD_SEPARATOR)
        if fields[0] == "DETAIL" and len(fields) >= 6:
            elements.append(
                {
                    "role": fields[1],
                    "name": fields[2],
                    "description": fields[3],
                    "help": fields[4],
                    "value": fields[5],
                }
            )
        elif fields[0] == "ERROR":
            errors.append(" | ".join(fields[1:]))
    return {"status": "ok", "elements": elements, "errors": errors, "raw": raw}


def _ax_detail_dump() -> dict[str, Any]:
    try:
        raw, _ = _osascript(AX_DETAIL_SCRIPT, timeout=8)
    except Exception as exc:
        message = str(exc)
        lowered = message.casefold()
        status = "permission_denied" if "not allowed" in lowered or "assistive" in lowered else "error"
        return {"status": status, "elements": [], "errors": [message]}
    return _parse_ax_detail_dump(raw)


def _element_records(dump: dict[str, Any], role: str | None = None) -> list[dict[str, Any]]:
    elements = dump.get("elements", [])
    if role is None:
        return list(elements)
    return [element for element in elements if element.get("role") == role]


def _semantic_path(prefix: str, element: dict[str, Any], ordinal: int) -> str:
    label = _norm(element.get("name") or element.get("description") or element.get("help"))
    role = str(element.get("role") or "unknown").replace("AX", "").lower()
    return f"{prefix}/{role}/{label or 'unnamed'}#{ordinal}"


def _controls_for_surface(prefix: str, dump: dict[str, Any]) -> dict[str, Any]:
    controls: dict[str, Any] = {}
    counts: dict[tuple[str, str], int] = {}
    for element in _element_records(dump):
        role = str(element.get("role") or "")
        if role not in {
            "AXButton",
            "AXRadioButton",
            "AXMenuButton",
            "AXPopUpButton",
            "AXCheckBox",
            "AXTextField",
            "AXComboBox",
            "AXStaticText",
            "AXMenuItem",
            "AXSheet",
            "AXDialog",
        }:
            continue
        key = (role, str(element.get("name") or element.get("description") or element.get("help") or ""))
        ordinal = counts.get(key, 0)
        counts[key] = ordinal + 1
        path = _semantic_path(prefix, element, ordinal)
        controls[path] = {
            "path": path,
            "role": role,
            "subrole": element.get("subrole") or "",
            "label": element.get("name") or "",
            "description": element.get("description") or "",
            "help": element.get("help") or "",
            "value": element.get("value"),
            "enabled": element.get("enabled"),
            "selected": element.get("selected"),
            "actions": element.get("actions") or "",
            "position": element.get("position") or "",
            "size": element.get("size") or "",
            "tooltip_status": "read" if element.get("help") else "not_exposed",
        }
    return controls


def _menu_items(dump: dict[str, Any]) -> list[dict[str, Any]]:
    items = []
    for element in _element_records(dump):
        if element.get("role") in {"AXMenu", "AXMenuItem", "AXCheckBox", "AXRadioButton"}:
            if element.get("role") == "AXMenuItem" or element.get("name"):
                items.append(element)
    return items


def _press_control(name: str, role: str = "AXButton", ordinal: int = 0, action: str = "AXPress") -> dict[str, Any]:
    script = r'''
on run argv
    set targetName to item 1 of argv
    set targetRole to item 2 of argv
    set targetOrdinal to (item 3 of argv) as integer
    set actionName to item 4 of argv
    tell application "System Events"
        if not (exists process "Microsoft Excel") then return "no_process"
        tell process "Microsoft Excel"
            try
                set allElements to entire contents of window 1
                set seen to 0
                repeat with e in allElements
                    try
                        if (role of e as text) is targetRole and (name of e as text) is targetName then
                            if seen is targetOrdinal then
                                if actionName is "AXShowMenu" then
                                    perform action "AXShowMenu" of e
                                else
                                    perform action "AXPress" of e
                                end if
                                return "ok"
                            end if
                            set seen to seen + 1
                        end if
                    end try
                end repeat
            on error errText
                return "error|" & (errText as text)
            end try
        end tell
    end tell
    return "not_found"
end run
'''
    try:
        out, _ = _osascript(script, [name, role, str(ordinal), action], timeout=8)
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
    if out == "ok":
        return {"status": "ok", "action": action, "name": name, "role": role, "ordinal": ordinal}
    if out == "not_found":
        return {"status": "unreachable", "reason": "AX control not found", "name": name}
    if out == "no_process":
        return {"status": "unreachable", "reason": "Excel process absent"}
    return {"status": "error", "error": out}


def _select_tab(label: str) -> dict[str, Any]:
    result = _press_control(label, "AXRadioButton", 0, "AXPress")
    if result.get("status") == "ok":
        return result
    fallback = _press_control(label, "AXButton", 0, "AXPress")
    if fallback.get("status") == "ok":
        return fallback
    return {"status": "unreachable", "reason": f"tab not found through AXRadioButton/AXButton: {label}"}


def _escape() -> None:
    script = 'tell application "System Events" to tell process "Microsoft Excel" to key code 53'
    try:
        _osascript(script, timeout=5)
    except Exception:
        pass


def _shortcut(kind: str) -> dict[str, Any]:
    if kind == "format_cells":
        script = 'tell application "System Events" to tell process "Microsoft Excel" to keystroke "1" using command down'
    elif kind == "find":
        script = 'tell application "System Events" to tell process "Microsoft Excel" to keystroke "f" using control down'
    elif kind == "replace":
        script = 'tell application "System Events" to tell process "Microsoft Excel" to keystroke "h" using control down'
    elif kind == "go_to":
        script = 'tell application "System Events" to tell process "Microsoft Excel" to keystroke "g" using control down'
    elif kind == "function_arguments":
        script = 'tell application "System Events" to tell process "Microsoft Excel" to key code 100 using shift down'
    else:
        return {"status": "error", "error": f"unknown shortcut {kind}"}
    try:
        _osascript(script, timeout=8)
        return {"status": "ok", "trigger": f"shortcut:{kind}"}
    except Exception as exc:
        return {"status": "error", "error": str(exc), "trigger": f"shortcut:{kind}"}


def _click_menu_tokens(tokens: list[str]) -> dict[str, Any]:
    token_blob = "|".join(tokens)
    script = r'''
on containsToken(s, tokenBlob)
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to "|"
    set parts to text items of tokenBlob
    set AppleScript's text item delimiters to oldDelims
    repeat with p in parts
        set token to p as text
        if token is not "" and s contains token then return true
    end repeat
    return false
end containsToken

on run argv
    set tokenBlob to item 1 of argv
    tell application "System Events"
        if not (exists process "Microsoft Excel") then return "no_process"
        tell process "Microsoft Excel"
            try
                repeat with mb in every menu bar item of menu bar 1
                    try
                        click mb
                        delay 0.15
                        set allItems to entire contents of menu bar 1
                        repeat with e in allItems
                            try
                                if (role of e as text) is "AXMenuItem" then
                                    set itemName to name of e as text
                                    if containsToken(itemName, tokenBlob) then
                                        perform action "AXPress" of e
                                        return "ok|" & itemName
                                    end if
                                end if
                            end try
                        end repeat
                    end try
                    key code 53
                end repeat
            on error errText
                key code 53
                return "error|" & (errText as text)
            end try
        end tell
    end tell
    return "not_found"
end run
'''
    try:
        out, _ = _osascript(script, [token_blob], timeout=15)
    except Exception as exc:
        return {"status": "error", "error": str(exc), "tokens": tokens}
    if out.startswith("ok|"):
        return {"status": "ok", "matched_label": out.split("|", 1)[1], "tokens": tokens}
    if out == "not_found":
        return {"status": "unreachable", "reason": "menu item token not found", "tokens": tokens}
    return {"status": "error", "error": out, "tokens": tokens}


def _type_and_commit(text: str) -> dict[str, Any]:
    script = r'''
on run argv
    set inputText to item 1 of argv
    tell application "System Events"
        if not (exists process "Microsoft Excel") then return "no_process"
        tell process "Microsoft Excel"
            try
                keystroke inputText
                key code 36
                return "ok"
            on error errText
                return "error|" & (errText as text)
            end try
        end tell
    end tell
end run
'''
    try:
        out, _ = _osascript(script, [text], timeout=10)
    except Exception as exc:
        return {"status": "error", "error": str(exc), "input": text}
    if out == "ok":
        return {"status": "ok", "input": text, "method": "System Events keystroke"}
    if out == "no_process":
        return {"status": "unreachable", "reason": "Excel process absent", "input": text}
    return {"status": "error", "error": out, "input": text}


def _screenshot(path: Path) -> dict[str, Any]:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        result = _run(["screencapture", "-x", str(path)], timeout=12)
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
    if result.returncode != 0 or not path.exists():
        message = (result.stderr or result.stdout).strip() or "screencapture did not create a file"
        status = "permission_denied" if "permission" in message.casefold() or "denied" in message.casefold() else "error"
        return {"status": status, "error": message}
    try:
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        size = path.stat().st_size
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
    return {"status": "ok", "path": str(path), "sha256": digest, "bytes": size}


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
        sys.path.insert(0, str(UPSTREAM / "tools" / "oracle"))
        from drivers._locale import COUNTRY_CODE_TO_BCP47, detect_mac_excel_locale  # type: ignore

        result["mac_defaults"] = detect_mac_excel_locale() or ""
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


def _read_property(obj: Any, attr: str) -> tuple[str, Any]:
    try:
        value = getattr(obj, attr)
        value = value() if callable(value) else value
        return "value", _json_value(value)
    except Exception as exc:
        return "error", str(exc)


def _read_formula_local(cell: Any) -> dict[str, Any]:
    try:
        api = cell.api
        reference = getattr(api, "formula_local")
        value = reference.get() if hasattr(reference, "get") else reference
        value = _json_value(value)
    except Exception as exc:
        return {"status": "error", "error": str(exc)}
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
    record: dict[str, Any] = {"input": formula, "input_method": "xlwings.formula2"}
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
    for attr in ("value", "formula2", "formula"):
        status, value = _read_property(cell, attr)
        record[f"{attr}_status"] = status
        record[attr] = value
    record["formula_local"] = _read_formula_local(cell)
    record["display_text"] = _display_text(cell)
    return record


def _typed_cell_case(app: Any, sheet: Any, addr: str, text: str) -> dict[str, Any]:
    cell = sheet.range(addr)
    record: dict[str, Any] = {
        "input": text,
        "input_method": "System Events keystroke",
        "address": addr,
    }
    try:
        cell.clear_contents()
        cell.select()
        app.activate()
    except Exception as exc:
        record["status"] = "selection_error"
        record["error"] = str(exc)
        return record
    typed = _type_and_commit(text)
    record["typing"] = typed
    if typed.get("status") != "ok":
        record["status"] = typed.get("status", "error")
        return record
    try:
        app.calculate()
    except Exception:
        pass
    record["status"] = "ok"
    for attr in ("value", "formula2", "formula"):
        status, value = _read_property(cell, attr)
        record[f"{attr}_status"] = status
        record[attr] = value
    record["formula_local"] = _read_formula_local(cell)
    record["display_text"] = _display_text(cell)
    try:
        record["number_format"] = _json_value(cell.number_format)
    except Exception as exc:
        record["number_format_error"] = str(exc)
    return record


def _spill_values(sheet: Any, start_row: int, count: int) -> list[Any]:
    try:
        values = sheet.range(f"A{start_row}:A{start_row + count - 1}").value
    except Exception as exc:
        return [{"status": "error", "error": str(exc)}]
    if not isinstance(values, list):
        return [_json_value(values)]
    return [_json_value(item[0] if isinstance(item, list) and len(item) == 1 else item) for item in values]


def _sort_formula_cases(app: Any, sheet: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for index, (group, values) in enumerate(SORT_GROUPS.items(), start=1):
        row = 1 + (index - 1) * 20
        formula = "=SORT({" + ";".join(json.dumps(value, ensure_ascii=False) for value in values) + "})"
        cell = sheet.range(f"A{row}")
        try:
            cell.clear_contents()
            cell.formula2 = formula
            app.calculate()
            observed = _spill_values(sheet, row, len(values))
            out[group] = {
                "status": "ok",
                "input": formula,
                "input_method": "xlwings.formula2",
                "values": observed,
                "shape": [len(values), 1],
                "display_text": [_display_text(sheet.range(f"A{row + i}")) for i in range(len(values))],
            }
        except Exception as exc:
            out[group] = {"status": "assignment_error", "input": formula, "error": str(exc)}
    return out


def _ui_sort_case(sheet: Any, locale_tag: str) -> dict[str, Any]:
    values = [item for group in SORT_GROUPS.values() for item in group]
    source = [[value, index] for index, value in enumerate(values)]
    try:
        sheet.range("H1:I200").clear_contents()
        sheet.range(f"H1:I{len(source)}").value = source
        sheet.range(f"H1:I{len(source)}").select()
    except Exception as exc:
        return {"status": "selection_error", "error": str(exc)}
    tab_label = TAB_LABELS.get(locale_tag, TAB_LABELS["en_US"])["data"]
    tab_trigger = _select_tab(tab_label)
    if tab_trigger.get("status") != "ok":
        _escape()
        return {
            "status": "unreachable",
            "reason": "Data ribbon tab could not be activated",
            "tab_trigger": tab_trigger,
            "before": source,
        }
    time.sleep(0.35)
    dump = _fresh_dump(1.0)
    tokens = MENU_TOKENS["sort"].get(locale_tag, MENU_TOKENS["sort"]["en_US"])
    control = _find_ribbon_control(dump, tokens, ("AXButton", "AXMenuButton"))
    if control is None:
        _escape()
        return {
            "status": "unreachable",
            "reason": "native sort ribbon control not exposed by AX",
            "expected_tokens": tokens,
            "before": source,
        }
    trigger = _press_control(control.get("name") or "", control.get("role") or "AXButton", 0, "AXPress")
    if trigger.get("status") != "ok":
        _escape()
        return {
            "status": trigger.get("status", "error"),
            "reason": "native sort ribbon action failed",
            "trigger": trigger,
            "control": control,
            "before": source,
        }
    time.sleep(0.6)
    try:
        after = _json_value(sheet.range(f"H1:I{len(source)}").value)
    except Exception as exc:
        after = {"error": str(exc)}
    _escape()
    return {
        "status": "ok",
        "input_method": "Excel ribbon UI AXPress",
        "control": {"role": control.get("role"), "label": control.get("name")},
        "trigger": trigger,
        "before": source,
        "after": after,
    }


def _fresh_dump(timeout: float = 2.0) -> dict[str, Any]:
    deadline = time.monotonic() + timeout
    last = _ax_dump()
    while time.monotonic() < deadline:
        if last.get("status") in {"ok", "no_process"}:
            return last
        time.sleep(0.2)
        last = _ax_dump()
    return last


def _capture_dialog(
    key: str,
    trigger: dict[str, Any],
    screenshot_dir: Path,
    before: dict[str, Any] | None = None,
) -> dict[str, Any]:
    baseline = before or _fresh_dump()
    result: dict[str, Any] = {"key": key, "kind": "dialog", "trigger": trigger}
    if trigger.get("status") != "ok":
        result["status"] = trigger.get("status", "error")
        result["reason"] = trigger.get("reason") or trigger.get("error") or "trigger failed"
        return result
    deadline = time.monotonic() + 5
    observed = _fresh_dump(0.5)
    while time.monotonic() < deadline:
        observed = _fresh_dump(0.5)
        names = {item.get("name") for item in observed.get("windows", [])}
        before_names = {item.get("name") for item in baseline.get("windows", [])}
        dialog_roles = {"AXDialog", "AXSheet"}
        if names - before_names or any(item.get("role") in dialog_roles for item in observed.get("elements", [])):
            break
        time.sleep(0.25)
    names = {item.get("name") for item in observed.get("windows", [])}
    before_names = {item.get("name") for item in baseline.get("windows", [])}
    if not (names - before_names) and not any(
        item.get("role") in {"AXDialog", "AXSheet"} for item in observed.get("elements", [])
    ):
        result["status"] = "unreachable"
        result["reason"] = "no new AXWindow/AXSheet/AXDialog appeared after trigger"
        result["after_trigger"] = observed
        _escape()
        return result
    result["status"] = "ok"
    result["windows"] = observed.get("windows", [])
    result["controls"] = _controls_for_surface(key, observed)
    result["errors"] = observed.get("errors", [])
    screenshot = _screenshot(screenshot_dir / (key.replace("/", "_") + ".png"))
    result["screenshot"] = screenshot
    _escape()
    time.sleep(0.3)
    after = _fresh_dump(0.5)
    result["closed"] = not any(item.get("role") in {"AXDialog", "AXSheet"} for item in after.get("elements", []))
    result["close_probe"] = {"windows": after.get("windows", [])}
    return result


def _capture_dropdown(
    prefix: str,
    button: dict[str, Any],
    screenshot_dir: Path,
) -> dict[str, Any]:
    name = button.get("name") or button.get("description") or ""
    role = button.get("role") or "AXButton"
    trigger = _press_control(name, role, 0, "AXShowMenu")
    result: dict[str, Any] = {
        "key": f"{prefix}/dropdown/{_norm(name) or 'unnamed'}",
        "kind": "dropdown",
        "trigger": trigger,
        "button": button,
    }
    if trigger.get("status") != "ok":
        result["status"] = trigger.get("status", "error")
        result["reason"] = trigger.get("reason") or trigger.get("error") or "AXShowMenu failed"
        return result
    time.sleep(0.3)
    observed = _fresh_dump(1.0)
    items = _menu_items(observed)
    if not items:
        result["status"] = "unreachable"
        result["reason"] = "AXShowMenu returned but no AXMenuItem was exposed"
        _escape()
        return result
    result["status"] = "ok"
    result["items"] = {}
    counts: dict[str, int] = {}
    for item in items:
        label = item.get("name") or item.get("description") or ""
        key = f"{result['key']}/item/{_norm(label) or 'unnamed'}#{counts.get(label, 0)}"
        counts[label] = counts.get(label, 0) + 1
        result["items"][key] = {
            "path": key,
            "role": item.get("role"),
            "label": label,
            "description": item.get("description"),
            "help": item.get("help"),
            "value": item.get("value"),
            "enabled": item.get("enabled"),
            "selected": item.get("selected"),
            "actions": item.get("actions"),
        }
    result["screenshot"] = _screenshot(screenshot_dir / (result["key"].replace("/", "_") + ".png"))
    _escape()
    return result


def _capture_ribbons(locale_tag: str, screenshot_dir: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    ui: dict[str, Any] = {
        "status": "ok",
        "windows": [],
        "menu_bar": {},
        "tabs": {},
        "errors": [],
    }
    surfaces: list[dict[str, Any]] = []
    initial = _fresh_dump()
    ui["initial_status"] = initial.get("status")
    ui["windows"] = initial.get("windows", [])
    ui["errors"].extend(initial.get("errors", []))
    if initial.get("status") != "ok":
        ui["status"] = initial.get("status", "error")
        ui["reason"] = "System Events could not dump the Excel AX tree"
        return ui, surfaces
    menu_bar = [
        {
            "role": "AXMenuBarItem",
            "name": label,
            "description": "",
            "help": "",
        }
        for label in initial.get("menus", [])
    ]
    menu_bar.extend(
        item
        for item in initial.get("elements", [])
        if item.get("role") in {"AXMenuBarItem", "AXMenuItem"}
    )
    ui["menu_bar"] = {
        f"menu/{_norm(item.get('name')) or 'unnamed'}#{index}": {
            "label": item.get("name") or "",
            "role": item.get("role"),
            "description": item.get("description"),
            "help": item.get("help"),
        }
        for index, item in enumerate(menu_bar)
    }
    observed_tabs = [
        item for item in initial.get("elements", []) if item.get("role") == "AXRadioButton" and item.get("name")
    ]
    ui["observed_tabs"] = [
        {"label": item.get("name"), "selected": item.get("selected"), "role": item.get("role")}
        for item in observed_tabs
    ]
    labels = TAB_LABELS.get(locale_tag, TAB_LABELS["en_US"])
    found: set[str] = set()
    for tab_key in CORE_TABS:
        label = labels[tab_key]
        candidates = [item for item in observed_tabs if item.get("name") == label]
        if not candidates:
            candidates = [
                item
                for item in observed_tabs
                if _norm(item.get("name")) == _norm(label) or _norm(item.get("name")).startswith(_norm(label))
            ]
        if not candidates:
            ui["tabs"][tab_key] = {
                "status": "unreachable",
                "reason": f"expected AXRadioButton not found: {label}",
                "expected_label": label,
            }
            continue
        press = _select_tab(candidates[0].get("name") or label)
        if press.get("status") != "ok":
            ui["tabs"][tab_key] = {
                "status": "unreachable",
                "reason": press.get("reason") or "AX tab activation failed",
                "expected_label": label,
                "observed_label": candidates[0].get("name"),
            }
            continue
        found.add(tab_key)
        time.sleep(0.35)
        dump = _fresh_dump(1.0)
        prefix = f"ribbon/{tab_key}"
        controls = _controls_for_surface(prefix, dump)
        tab_surface = {
            "key": prefix,
            "kind": "ribbon_tab",
            "status": "ok" if dump.get("status") == "ok" else dump.get("status", "error"),
            "expected_label": label,
            "observed_label": candidates[0].get("name"),
            "controls": controls,
            "errors": dump.get("errors", []),
            "screenshot": _screenshot(screenshot_dir / f"ribbon_{tab_key}.png"),
        }
        ui["tabs"][tab_key] = {
            "status": tab_surface["status"],
            "expected_label": label,
            "observed_label": candidates[0].get("name"),
            "control_count": len(controls),
            "radio_count": len([item for item in dump.get("elements", []) if item.get("role") == "AXRadioButton"]),
        }
        surfaces.append(tab_surface)
    # Trigger dropdowns after the complete tab pass.  Opening a ribbon menu
    # changes the AX tree and can hide later tab buttons on some Excel builds.
    # Keep the trigger set bounded and retain the tab capture even when a menu
    # cannot be exposed through Accessibility.
    home_label = labels["home"]
    if "home" in found:
        _select_tab(home_label)
        time.sleep(0.35)
        home_dump = _fresh_dump(1.0)
        candidate_buttons = [
            item
            for item in home_dump.get("elements", [])
            if item.get("role") in {"AXButton", "AXMenuButton", "AXPopUpButton"}
            and item.get("name")
        ]
        dropdown_attempts = 0
        for button in candidate_buttons:
            actions = str(button.get("actions") or "")
            name_norm = _norm(button.get("name"))
            if (
                "AXShowMenu" not in actions
                and button.get("role") not in {"AXMenuButton", "AXPopUpButton"}
                and not any(token in name_norm for token in ("paste", "format", "sort", "filter", "color", "border", "fill", "clear", "autosum"))
            ):
                continue
            if dropdown_attempts >= 8:
                break
            dropdown_attempts += 1
            dropdown = _capture_dropdown("ribbon/home", button, screenshot_dir)
            if dropdown.get("status") == "ok":
                surfaces.append(dropdown)
    ui["core_tab_count"] = len(found)
    if len(found) != len(CORE_TABS):
        ui["status"] = "error"
        ui["reason"] = f"core ribbon tab mismatch: {sorted(found)}"
    return ui, surfaces


def _match_named_control(elements: list[dict[str, Any]], labels: Iterable[str]) -> dict[str, Any] | None:
    normalized = [_norm(label) for label in labels if _norm(label)]
    for element in elements:
        name = _norm(element.get("name"))
        if name and any(name == label or label in name or name in label for label in normalized):
            return element
    return None


def _capture_home_visual(locale_tag: str, screenshot_dir: Path) -> dict[str, Any]:
    labels = TAB_LABELS.get(locale_tag, TAB_LABELS["en_US"])
    home_label = labels["home"]
    result: dict[str, Any] = {
        "key": "home-visual",
        "kind": "home_visual",
        "locale": locale_tag,
        "status": "unreachable",
        "tab": {"expected_label": home_label},
    }
    tab_trigger = _select_tab(home_label)
    result["tab"]["trigger"] = tab_trigger
    if tab_trigger.get("status") != "ok":
        result["reason"] = "Home tab could not be activated through AX"
        return result
    expected_style_labels = FONT_STYLE_LABELS.get(locale_tag, FONT_STYLE_LABELS["en_US"])
    expected_styles = {
        _norm(label)
        for labels_for_style in expected_style_labels.values()
        for label in labels_for_style
    }
    render_deadline = time.monotonic() + 8
    dump = {"status": "error", "elements": [], "windows": [], "errors": []}
    render_wait = 0.0
    while True:
        dump = _ax_dump()
        observed_names = {_norm(item.get("name")) for item in dump.get("elements", [])}
        command_surface_ready = len(dump.get("elements", [])) >= 75
        style_surface_ready = bool(expected_styles & observed_names)
        if dump.get("status") == "ok" and (command_surface_ready or style_surface_ready):
            break
        if time.monotonic() >= render_deadline:
            break
        time.sleep(0.35)
        render_wait = round(8 - max(0, render_deadline - time.monotonic()), 2)
    if dump.get("status") == "ok" and len(dump.get("elements", [])) < 75:
        # Excel can start with the ribbon collapsed.  Pressing the already
        # selected Home tab once more expands it on builds that expose the
        # toggle through AX; retain the collapsed state when it does not.
        expand_trigger = _select_tab(home_label)
        time.sleep(0.45)
        expanded_dump = _ax_dump()
        if len(expanded_dump.get("elements", [])) > len(dump.get("elements", [])):
            dump = expanded_dump
            result["ribbon_expand_trigger"] = expand_trigger
    result["render_wait_seconds"] = render_wait
    result["tab"]["observed_label"] = home_label
    result["tab"]["dump_status"] = dump.get("status")
    result["window"] = dump.get("windows", [])
    result["screenshot"] = _screenshot(screenshot_dir / "home.png")
    result["screenshot"]["scope"] = "full_desktop"
    result["controls"] = _controls_for_surface("home", dump)
    result["ribbon_state"] = {
        "status": "expanded" if len(dump.get("elements", [])) >= 75 else "collapsed_or_not_ready",
        "ax_element_count": len(dump.get("elements", [])),
    }
    details = _ax_detail_dump()
    result["detail_probe"] = {
        "status": details.get("status"),
        "errors": details.get("errors", []),
        "element_count": len(details.get("elements", [])),
    }
    detail_elements = details.get("elements", [])
    font_controls = [
        item
        for item in detail_elements
        if item.get("role") in {"AXComboBox", "AXPopUpButton"}
        and any(token in _norm(" ".join(str(item.get(field, "")) for field in ("name", "description", "help"))) for token in FONT_PICKER_TOKENS)
    ]
    if not font_controls and result["ribbon_state"]["status"] == "expanded":
        # Thai Excel exposes the font picker as an unnamed AXComboBox.  In an
        # expanded ribbon the first non-name/non-number-format combo is the
        # visible font value; retain only that bounded fallback.
        font_controls = [
            item
            for item in detail_elements
            if item.get("role") in {"AXComboBox", "AXPopUpButton"}
            and _norm(item.get("value")) not in {"", "a1", "general", "มาตรฐาน", "12"}
        ][:1]
    result["font_picker"] = {
        "status": "read" if font_controls else "unreachable",
        "reason": "AX did not expose the visible font picker; other combo boxes may still be present"
        if not font_controls
        else "",
        "controls": [
            {
                "role": item.get("role"),
                "label": item.get("name", ""),
                "description": item.get("description", ""),
                "help": item.get("help", ""),
                "value": item.get("value", ""),
                "value_status": "read" if item.get("value") else "blank",
            }
            for item in font_controls
        ],
    }
    style_labels = expected_style_labels
    style_records: dict[str, Any] = {}
    for style_key, expected in style_labels.items():
        fast = _match_named_control(dump.get("elements", []), expected)
        detailed = _match_named_control(detail_elements, expected)
        observed = detailed or fast
        style_records[style_key] = {
            "status": "read" if observed else "unreachable",
            "expected_labels": list(expected),
            "role": observed.get("role") if observed else "",
            "label": observed.get("name") if observed else "",
            "description": observed.get("description", "") if observed else "",
            "help": observed.get("help", "") if observed else "",
            "value": observed.get("value", "") if observed else "",
            "glyph_visual": "included in home.png",
        }
    result["bold_italic"] = style_records
    result["hover"] = {
        "status": "unavailable",
        "attempted": False,
        "reason": "No bounded pointer automation backend is installed; AX help was read for exposed controls",
        "backend": shutil.which("cliclick") or "none",
    }
    if result["screenshot"].get("status") == "ok" and dump.get("status") == "ok":
        result["status"] = "ok"
    else:
        result["reason"] = "Home screenshot or AX dump failed"
    return result


def _capture_shortcut_dialogs(screenshot_dir: Path) -> list[dict[str, Any]]:
    surfaces: list[dict[str, Any]] = []
    for key, shortcut in SHORTCUT_DIALOGS:
        baseline = _fresh_dump(0.5)
        trigger = _shortcut(shortcut)
        surfaces.append(_capture_dialog(key, trigger, screenshot_dir, baseline))
        _escape()
        time.sleep(0.25)
    return surfaces


def _find_ribbon_control(
    dump: dict[str, Any],
    tokens: Iterable[str],
    roles: Iterable[str],
) -> dict[str, Any] | None:
    role_set = set(roles)
    normalized_tokens = [_norm(token) for token in tokens if _norm(token)]
    candidates = [
        item
        for item in dump.get("elements", [])
        if item.get("role") in role_set and item.get("name")
    ]
    for item in candidates:
        label = _norm(item.get("name"))
        if any(label == token or token in label or label in token for token in normalized_tokens):
            return item
    return None


def _capture_ribbon_command(
    locale_tag: str,
    surface_key: str,
    screenshot_dir: Path,
    sheet: Any,
) -> dict[str, Any]:
    spec = RIBBON_COMMANDS[surface_key]
    tab_key = spec["tab"][0]
    tab_label = TAB_LABELS.get(locale_tag, TAB_LABELS["en_US"])[tab_key]
    try:
        sheet.range("A2").select()
    except Exception:
        pass
    selected = _select_tab(tab_label)
    result: dict[str, Any] = {
        "key": f"ribbon-command/{surface_key}",
        "kind": "ribbon_command",
        "tab": tab_key,
        "tab_label": tab_label,
        "status": "unreachable",
        "tab_trigger": selected,
    }
    if selected.get("status") != "ok":
        result["reason"] = "ribbon tab could not be activated"
        return result
    time.sleep(0.35)
    dump = _fresh_dump(1.0)
    tokens = spec["tokens"].get(locale_tag, spec["tokens"].get("en_US", ()))
    roles = spec["roles"]
    control = _find_ribbon_control(dump, tokens, roles)
    result["observed_controls"] = [
        {
            "role": item.get("role"),
            "label": item.get("name"),
        }
        for item in dump.get("elements", [])
        if item.get("role") in set(roles) and item.get("name")
    ]
    if control is None:
        result["reason"] = "localized ribbon control was not exposed by AX"
        result["expected_tokens"] = list(tokens)
        return result
    result["control"] = {
        "role": control.get("role"),
        "label": control.get("name"),
        "path": _semantic_path(f"ribbon/{tab_key}", control, 0),
    }
    action = "AXShowMenu" if control.get("role") in {"AXMenuButton", "AXPopUpButton"} else spec["action"][0]
    if surface_key == "filter" and control.get("role") == "AXCheckBox":
        result["status"] = "observed"
        result["reason"] = "filter AXCheckBox label recorded without toggling the worksheet"
        result["trigger"] = {"status": "skipped_safe", "action": action}
        return result
    baseline = _fresh_dump(0.5)
    if action == "AXShowMenu":
        surface = _capture_dropdown(f"ribbon-command/{surface_key}", control, screenshot_dir)
        result.update(surface)
        result["key"] = f"ribbon-command/{surface_key}"
        result["kind"] = "ribbon_command"
        return result
    trigger = _press_control(control.get("name") or "", control.get("role") or "AXButton", 0, action)
    captured = _capture_dialog(f"ribbon-command/{surface_key}", trigger, screenshot_dir, baseline)
    result.update(captured)
    result["key"] = f"ribbon-command/{surface_key}"
    result["kind"] = "ribbon_command"
    return result


def _capture_menu_dialogs(
    locale_tag: str,
    screenshot_dir: Path,
    sheet: Any,
) -> list[dict[str, Any]]:
    surfaces: list[dict[str, Any]] = []
    for surface_key in RIBBON_COMMANDS:
        surfaces.append(_capture_ribbon_command(locale_tag, surface_key, screenshot_dir, sheet))
        _escape()
        time.sleep(0.2)
    surfaces.append({
        "key": "dialog/paste-special",
        "kind": "dialog",
        "status": "skipped_safe",
        "reason": "clipboard was not modified because no restorable pasteboard snapshot was available",
    })
    return surfaces


def _capture_ui_cases(app: Any, sheet: Any, locale_tag: str) -> dict[str, Any]:
    cases: dict[str, Any] = {}
    for index, text in enumerate(("1,5", "1.5"), start=1):
        cases[f"ui_number_{index}"] = _typed_cell_case(app, sheet, f"C{index}", text)
    for index, text in enumerate(("2024-01-05", "01/05/2024"), start=1):
        cases[f"ui_date_{index}"] = _typed_cell_case(app, sheet, f"D{index}", text)
    cases["ui_sum_comma"] = _typed_cell_case(app, sheet, "E1", "=SUM(1,2)")
    cases["ui_sum_semicolon"] = _typed_cell_case(app, sheet, "E2", "=SUM(1;2)")
    cases["ui_locale"] = locale_tag
    return cases


def _attach_open_excel(xw: Any, timeout: float = WAIT_SECONDS) -> Any:
    deadline = time.monotonic() + timeout
    last_error = ""
    while time.monotonic() < deadline:
        try:
            apps = list(xw.apps)
            if apps:
                return apps[-1]
        except Exception as exc:
            last_error = str(exc)
        time.sleep(1)
    raise RuntimeError(f"Excel process opened but xlwings could not attach: {last_error}")


def _new_workbook(xw: Any, app: Any | None = None) -> tuple[Any, Any]:
    if app is None:
        try:
            _run(["open", "-a", "Microsoft Excel"], timeout=15, check=True)
        except Exception as exc:
            raise RuntimeError(f"Excel launch failed: {exc}") from exc
        app = _attach_open_excel(xw)
    try:
        app.display_alerts = False
    except Exception:
        pass
    existing = list(app.books)
    for book in existing:
        try:
            book.close()
        except Exception as exc:
            raise RuntimeError(f"Excel exposed an existing book that could not be closed: {exc}") from exc
    for _ in range(15):
        try:
            book = app.books.add()
            return app, book
        except Exception:
            time.sleep(1)
    raise RuntimeError("Excel did not accept a new temporary workbook")


def _prepare_sheet(book: Any) -> Any:
    sheet = book.sheets[0]
    sheet.range("A20").value = "ガ"
    sheet.range("A21").value = "ｶﾞ"
    sheet.range("A2").value = 2
    sheet.range("B2").value = 1
    sheet.range("H1:I20").clear_contents()
    sheet.range("H1:I5").value = [["ガ", 0], ["ｶﾞ", 1], ["カ", 2], ["か", 3], ["ä", 4]]
    return sheet


def _close_owned(app: Any, book: Any, owned_pids: set[int]) -> list[str]:
    errors: list[str] = []
    # Ribbon menus and modal sheets can retain focus after an AX probe.  Give
    # Excel a few explicit cancel keystrokes before asking xlwings to close;
    # this also keeps a failed dialog probe from blocking process restoration.
    for _ in range(3):
        _escape()
        time.sleep(0.2)
    book_close_error: str | None = None
    if book is not None:
        for attempt in range(3):
            try:
                book.close()
                book_close_error = None
                break
            except Exception as exc:
                book_close_error = str(exc)
                time.sleep(0.5 * (attempt + 1))
    if app is not None:
        try:
            app.quit()
        except Exception as exc:
            errors.append(f"Excel quit: {exc}")
    deadline = time.monotonic() + WAIT_SECONDS
    while set(_excel_pids()) & owned_pids and time.monotonic() < deadline:
        time.sleep(1)
    remaining = set(_excel_pids()) & owned_pids
    if remaining:
        try:
            _osascript('tell application "Microsoft Excel" to quit saving no', timeout=15)
        except Exception as exc:
            errors.append(f"AppleScript quit: {exc}")
        deadline = time.monotonic() + 15
        while set(_excel_pids()) & owned_pids and time.monotonic() < deadline:
            time.sleep(1)
    remaining = set(_excel_pids()) & owned_pids
    for pid in sorted(remaining):
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception as exc:
            errors.append(f"owned Excel PID {pid} termination: {exc}")
    if set(_excel_pids()) & owned_pids:
        errors.append(f"owned Excel process remained: {sorted(set(_excel_pids()) & owned_pids)}")
    elif book_close_error:
        # A modal command can invalidate the xlwings book proxy while Excel is
        # still able to quit.  The process and workbook are gone, so retain
        # the restoration result without turning this transient proxy error
        # into a false cleanup failure.
        pass
    unexpected = set(_excel_pids()) - owned_pids
    if unexpected:
        errors.append(f"unexpected Excel process appeared: {sorted(unexpected)}")
    return errors


def _launch_locale(
    tag: str,
    config: dict[str, Any],
    xw: Any,
    artifact_dir: Path,
    *,
    include_extended: bool = True,
    home_only: bool = False,
) -> tuple[dict[str, Any], set[int]]:
    _set_locale(config)
    opened_at = _now()
    try:
        _run(["open", "-a", "Microsoft Excel"], timeout=15, check=True)
    except Exception as exc:
        return (
            {
                "tag": tag,
                "status": "launch_error",
                "requested": config,
                "opened_at": opened_at,
                "error": str(exc),
            },
            set(),
        )
    deadline = time.monotonic() + WAIT_SECONDS
    while not _excel_pids() and time.monotonic() < deadline:
        time.sleep(1)
    pids = set(_excel_pids())
    if not pids:
        return (
            {
                "tag": tag,
                "status": "launch_timeout",
                "requested": config,
                "opened_at": opened_at,
            },
            set(),
        )
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
        "mode": "home-only" if home_only else "full",
        "owned_pids": sorted(pids),
        "surfaces": {},
        "cases": {},
    }
    screenshot_dir = artifact_dir / tag / "screenshots"
    try:
        app, book = _new_workbook(xw, _attach_open_excel(xw))
        pids = set(_excel_pids())
        result["owned_pids"] = sorted(pids)
        sheet = book.sheets[0] if home_only else _prepare_sheet(book)
        if home_only:
            # A blank newly-created workbook can expose only the window chrome
            # through AX until the sheet has focus.  A temporary sample glyph
            # keeps the ribbon's font controls visible for the screenshot.
            sheet.range("A1").value = "Aa あ"
            sheet.range("A1").select()
            try:
                app.activate()
            except Exception:
                pass
        result["excel"] = {"version": _version(app)}
        result["observed"] = _observed_locale(app)
        result["workbook"] = {
            "name": _json_value(getattr(book, "name", None)),
            "fullname": _json_value(getattr(book, "fullname", None)),
            "sheets": [_json_value(getattr(item, "name", None)) for item in list(book.sheets)],
            "temporary": True,
        }
        if result["observed"].get("mac_defaults") != config["bcp47"]:
            result["errors"] = [
                f"observed locale {result['observed'].get('mac_defaults')!r} differs from requested {config['bcp47']!r}"
            ]
            result["status"] = "locale_mismatch"
        if home_only:
            home_surface = _capture_home_visual(tag, screenshot_dir)
            result["ui"] = {
                "mode": "home-only",
                "status": home_surface.get("status"),
                "home_label": home_surface.get("tab", {}).get("observed_label"),
            }
            result["surfaces"][home_surface["key"]] = home_surface
        else:
            formula_cases = {
                name: _formula_case(app, sheet, sheet.range("A1"), formula, setup)
                for name, (formula, setup) in FORMULA_CASES.items()
            }
            result["cases"]["formula"] = formula_cases
            result["cases"]["sort_formula"] = _sort_formula_cases(app, sheet)
            result["cases"]["sort_ui"] = _ui_sort_case(sheet, tag)
            result["cases"]["ui_input"] = _capture_ui_cases(app, sheet, tag)
            ui, ribbon_surfaces = _capture_ribbons(tag, screenshot_dir)
            result["ui"] = ui
            for surface in ribbon_surfaces:
                result["surfaces"][surface["key"]] = surface
            if include_extended:
                for surface in _capture_menu_dialogs(tag, screenshot_dir, sheet):
                    result["surfaces"][surface["key"]] = surface
                for surface in _capture_shortcut_dialogs(screenshot_dir):
                    result["surfaces"][surface["key"]] = surface
        result["status"] = "ok" if result.get("status") == "started" else result["status"]
    except Exception as exc:
        result["status"] = "capture_error"
        result["error"] = str(exc)
    finally:
        result["close_errors"] = _close_owned(app, book, pids)
    return result, pids


def _verify_restored_launch(xw: Any) -> dict[str, Any]:
    if _excel_pids():
        return {"status": "skipped", "reason": "Excel process remained before verification"}
    app = None
    pids: set[int] = set()
    try:
        _run(["open", "-a", "Microsoft Excel"], timeout=15, check=True)
        app = _attach_open_excel(xw)
        pids = set(_excel_pids())
        observed = _observed_locale(app)
        return {
            "status": "ok",
            "excel_version": _version(app),
            "observed_locale": observed,
            "owned_pids": sorted(pids),
        }
    except Exception as exc:
        return {"status": "probe_error", "error": str(exc), "owned_pids": sorted(pids)}
    finally:
        _close_owned(app, None, pids)


def _write_incremental(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(path)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--locale", nargs="*", choices=sorted(LOCALES), help="subset; default is all seven")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--artifact-dir", type=Path, default=None)
    parser.add_argument("--no-extended", action="store_true", help="capture formulas/UI only")
    parser.add_argument(
        "--home-only",
        action="store_true",
        help="capture only the Home ribbon visual, font picker attributes, and Bold/Italic labels",
    )
    return parser.parse_args()


def main() -> int:
    if platform.system() != "Darwin":
        print("Excel UI probe requires macOS", file=sys.stderr)
        return 2
    args = _parse_args()
    tags = args.locale or list(LOCALES)
    xw = _load_xlwings()
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
    capture_id = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    artifact_dir = args.artifact_dir or Path(tempfile.mkdtemp(prefix=f"formulon-cell-excel-ui-{capture_id}-"))
    data: dict[str, Any] = {
        "schema": "formulon-cell.excel-ax-capture/v2",
        "capture_id": capture_id,
        "captured_at": _now(),
        "host": {
            "platform": platform.platform(),
            "python": sys.version,
            "xlwings": getattr(xw, "__version__", ""),
            "artifact_dir": str(artifact_dir),
        },
        "safety": {
            "startup_excel_pids": startup_pids,
            "startup_xlwings_apps": startup_apps,
            "widget_processes_ignored": True,
            "clipboard_modified": False,
            "oracle_written": False,
        },
        "preferences_before": prefs_before,
        "requested_locales": tags,
        "locales": [],
        "errors": [],
    }
    restore_errors: list[str] = []
    try:
        for tag in tags:
            print(f"[{tag}] launching Excel", flush=True)
            item, _ = _launch_locale(
                tag,
                LOCALES[tag],
                xw,
                artifact_dir,
                include_extended=not args.no_extended and not args.home_only,
                home_only=args.home_only,
            )
            data["locales"].append(item)
            _write_incremental(args.output, data)
            if args.home_only:
                _write_incremental(artifact_dir / "partial" / f"{tag}.json", {
                    "schema": data["schema"],
                    "capture_id": capture_id,
                    "requested_locale": tag,
                    "locale": item,
                })
            print(
                f"[{tag}] {item.get('status')} "
                f"tabs={item.get('ui', {}).get('core_tab_count', '?')} "
                f"surfaces={len(item.get('surfaces', {}))}",
                flush=True,
            )
    except KeyboardInterrupt:
        data["errors"].append("interrupted")
        raise
    finally:
        restore_errors.extend(_restore_preferences(prefs_before))
        data["restoration_launch"] = _verify_restored_launch(xw)
        restore_errors.extend(_restore_preferences(prefs_before))
        data["preferences_after"] = _pref_snapshot()
        data["cleanup"] = {
            "restore_errors": restore_errors,
            "preferences_restored_exactly": _same_pref_state(prefs_before, data["preferences_after"]),
            "excel_pids_after": _excel_pids(),
            "artifact_dir": str(artifact_dir),
        }
        _write_incremental(args.output, data)
    checks: dict[str, list[str]] = {}
    for item in data.get("locales", []):
        problems: list[str] = []
        if item.get("status") not in {"ok"}:
            problems.append(f"status={item.get('status')!r}")
        if item.get("observed", {}).get("mac_defaults") != item.get("requested", {}).get("bcp47"):
            problems.append("observed locale mismatch")
        if not item.get("workbook", {}).get("name") or not item.get("workbook", {}).get("sheets"):
            problems.append("fresh workbook identity missing")
        ui = item.get("ui", {})
        if item.get("mode") == "home-only":
            home = item.get("surfaces", {}).get("home-visual", {})
            if ui.get("status") != "ok" or home.get("status") != "ok":
                problems.append("Home ribbon capture incomplete")
            if home.get("screenshot", {}).get("status") != "ok":
                problems.append("Home screenshot unread")
            if home.get("ribbon_state", {}).get("status") != "expanded":
                problems.append("Home ribbon is collapsed or not ready")
            for style_key in ("bold", "italic"):
                if home.get("bold_italic", {}).get(style_key, {}).get("status") != "read":
                    problems.append(f"Home {style_key} control unread")
            if home.get("font_picker", {}).get("status") != "read":
                problems.append("Home font picker unread")
        else:
            if ui.get("status") != "ok" or ui.get("core_tab_count") != len(CORE_TABS):
                problems.append("core ribbon AX capture incomplete")
            formula = item.get("cases", {}).get("formula", {})
            for name in FORMULA_CASES:
                case = formula.get(name, {})
                if case.get("status") != "ok":
                    problems.append(f"{name}: {case.get('status')!r}")
                if case.get("value_status") != "value" or case.get("formula2_status") != "value":
                    problems.append(f"{name}: value/formula retention unread")
                if case.get("display_text", {}).get("status") != "read":
                    problems.append(f"{name}: display text unread")
        if item.get("close_errors"):
            problems.append("close_errors present")
        if problems:
            checks[str(item.get("tag", "?"))] = problems
    data.setdefault("cleanup", {})["capture_checks"] = checks
    _write_incremental(args.output, data)
    captures_ok = len(data.get("locales", [])) == len(tags) and not checks
    restoration = data.get("cleanup", {})
    restoration_ok = (
        data.get("restoration_launch", {}).get("status") == "ok"
        and restoration.get("preferences_restored_exactly") is True
        and not restore_errors
        and not restoration.get("excel_pids_after")
    )
    return 0 if captures_ok and restoration_ok else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
