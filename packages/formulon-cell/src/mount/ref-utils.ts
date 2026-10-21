import {
  colFromLetters,
  colLetter,
  formatA1Cell,
  MAX_COL,
  MAX_ROW,
  parseA1Atom,
} from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isWholeColumnRange, isWholeRowRange } from '../store/selection-geometry.js';

function cellRef(row: number, col: number, r1c1: boolean): string {
  return r1c1 ? `R${row + 1}C${col + 1}` : formatA1Cell(row, col);
}

export function formatSelectionRef(
  range: { r0: number; c0: number; r1: number; c1: number },
  active: { row: number; col: number },
  r1c1: boolean,
): string {
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return cellRef(active.row, active.col, r1c1);
  }
  if (!r1c1) {
    if (isWholeColumnRange(range)) {
      return range.c0 === range.c1
        ? colLetter(range.c0)
        : `${colLetter(range.c0)}:${colLetter(range.c1)}`;
    }
    if (isWholeRowRange(range)) {
      return range.r0 === range.r1 ? `${range.r0 + 1}` : `${range.r0 + 1}:${range.r1 + 1}`;
    }
  }
  return `${cellRef(range.r0, range.c0, r1c1)}:${cellRef(range.r1, range.c1, r1c1)}`;
}

/** Case-insensitive defined-name lookup. Sheet-scoped names for `sheetIndex`
 *  shadow workbook-scoped names. Returns the formula text stripped of any
 *  leading `=`, sheet qualifier, and `$` anchors so it can be parsed by
 *  `parseRangeRef` / `parseCellRef`. */
export function lookupDefinedName(
  wb: WorkbookHandle,
  query: string,
  sheetIndex = -1,
): string | null {
  if (!query) return null;
  const q = query.toLowerCase();
  let workbookScoped: string | null = null;
  for (const dn of wb.definedNames()) {
    if (dn.name.toLowerCase() !== q) continue;
    const eq = dn.formula.replace(/^=/, '');
    const bang = eq.lastIndexOf('!');
    const normalized = (bang >= 0 ? eq.slice(bang + 1) : eq).replace(/\$/g, '');
    if (dn.localSheetId === sheetIndex) return normalized;
    if (dn.localSheetId === -1 && workbookScoped === null) workbookScoped = normalized;
  }
  return workbookScoped;
}

export interface SheetRange {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** Split `Sheet2!A1` / `'My Sheet'!A1` into the unquoted sheet name and the
 *  remaining reference. `sheet` is null when there is no qualifier. */
export function splitSheetQualifier(text: string): { sheet: string | null; ref: string } {
  const t = text.trim();
  if (t.startsWith("'")) {
    let i = 1;
    while (i < t.length) {
      if (t[i] === "'") {
        if (t[i + 1] === "'") i += 2;
        else break;
      } else i += 1;
    }
    if (t[i + 1] === '!') {
      return { sheet: t.slice(1, i).replace(/''/g, "'"), ref: t.slice(i + 2) };
    }
    return { sheet: null, ref: t };
  }
  const bang = t.indexOf('!');
  if (bang < 0) return { sheet: null, ref: t };
  return { sheet: t.slice(0, bang), ref: t.slice(bang + 1) };
}

/** Split `A1,C3:D4` on commas that sit outside a quoted sheet name. */
export function splitAreas(raw: string): string[] {
  const out: string[] = [];
  let quoted = false;
  let start = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i];
    if (ch === "'") quoted = !quoted;
    else if (ch === ',' && !quoted) {
      out.push(raw.slice(start, i));
      start = i + 1;
    }
  }
  out.push(raw.slice(start));
  return out;
}

function sheetIndexByName(wb: WorkbookHandle, name: string): number | null {
  const q = name.toLowerCase();
  for (let i = 0; i < wb.sheetCount; i += 1) {
    if (wb.sheetName(i).toLowerCase() === q) return i;
  }
  return null;
}

/** Resolve a (possibly sheet-qualified) cell or range reference. Unqualified
 *  references land on `fallbackSheet`. Names are not followed. */
export function resolveRangeText(
  wb: WorkbookHandle,
  text: string,
  fallbackSheet: number,
): SheetRange | null {
  const { sheet: sheetName, ref } = splitSheetQualifier(text.replace(/^=/, ''));
  const sheet = sheetName === null ? fallbackSheet : sheetIndexByName(wb, sheetName);
  if (sheet === null) return null;
  const clean = ref.replace(/\$/g, '');
  const range = parseRangeRef(clean);
  if (range) return { sheet, ...range };
  const cell = parseCellRef(clean);
  return cell ? { sheet, r0: cell.row, c0: cell.col, r1: cell.row, c1: cell.col } : null;
}

/** Defined names visible from `sheetIndex` (workbook scope plus that sheet's
 *  own), each with the range it refers to; names that do not refer to a plain
 *  range (constants, formulas) are left out. Sorted case-insensitively. */
export function rangeDefinedNames(
  wb: WorkbookHandle,
  sheetIndex: number,
): { name: string; formula: string; localSheetId: number; range: SheetRange }[] {
  const byKey = new Map<
    string,
    { name: string; formula: string; localSheetId: number; range: SheetRange }
  >();
  for (const dn of wb.definedNames()) {
    if (!dn.name.trim() || !dn.formula.trim()) continue;
    if (dn.localSheetId !== -1 && dn.localSheetId !== sheetIndex) continue;
    const range = resolveRangeText(wb, dn.formula, sheetIndex);
    if (!range) continue;
    const key = dn.name.toLowerCase();
    const existing = byKey.get(key);
    // A name scoped to the active sheet shadows a workbook-scoped twin.
    if (existing && existing.localSheetId === sheetIndex) continue;
    byKey.set(key, { name: dn.name, formula: dn.formula, localSheetId: dn.localSheetId, range });
  }
  return [...byKey.values()].sort((a, b) =>
    a.name.toLowerCase() < b.name.toLowerCase()
      ? -1
      : a.name.toLowerCase() > b.name.toLowerCase()
        ? 1
        : 0,
  );
}

/** Name shown in the Name Box for a selection that matches a named range
 *  exactly: a sheet-scoped name first, then the first workbook name in sorted
 *  order. */
export function nameForRange(
  wb: WorkbookHandle,
  sheetIndex: number,
  range: SheetRange,
): string | null {
  let workbookName: string | null = null;
  for (const dn of rangeDefinedNames(wb, sheetIndex)) {
    const r = dn.range;
    if (
      r.sheet !== range.sheet ||
      r.r0 !== range.r0 ||
      r.c0 !== range.c0 ||
      r.r1 !== range.r1 ||
      r.c1 !== range.c1
    ) {
      continue;
    }
    if (dn.localSheetId === sheetIndex) return dn.name;
    workbookName ??= dn.name;
  }
  return workbookName;
}

/** Resolve one Name Box area: a defined name, then a (sheet-qualified)
 *  reference. Names win over column-like text such as `TAX`. */
export function resolveAreaText(
  wb: WorkbookHandle,
  text: string,
  sheetIndex: number,
): SheetRange | null {
  const t = text.trim();
  if (!t) return null;
  const q = t.toLowerCase();
  const named = rangeDefinedNames(wb, sheetIndex).find((dn) => dn.name.toLowerCase() === q);
  if (named) return named.range;
  return resolveRangeText(wb, t, sheetIndex);
}

export function parseCellRef(raw: string): { row: number; col: number } | null {
  const trimmed = raw.trim().toUpperCase();
  // R1C1 form: e.g. "R5C2"
  const r1c1 = trimmed.match(/^R([1-9][0-9]*)C([1-9][0-9]*)$/);
  if (r1c1) {
    const row = Number.parseInt(r1c1[1] ?? '', 10) - 1;
    const col = Number.parseInt(r1c1[2] ?? '', 10) - 1;
    if (row < 0 || col < 0) return null;
    if (col > MAX_COL || row > MAX_ROW) return null;
    return { row, col };
  }
  return parseA1Atom(trimmed);
}

/** Parse A1:B5 style range. Returns null when the input doesn't match. */
export function parseRangeRef(
  raw: string,
): { r0: number; c0: number; r1: number; c1: number } | null {
  const trimmed = raw.trim().toUpperCase();
  const wholeCol = trimmed.match(/^\$?([A-Z]+)(?::\$?([A-Z]+))?$/);
  if (wholeCol) {
    const c0 = parseColRef(wholeCol[1] ?? '');
    const c1 = parseColRef(wholeCol[2] ?? wholeCol[1] ?? '');
    if (c0 == null || c1 == null) return null;
    return { r0: 0, c0: Math.min(c0, c1), r1: MAX_ROW, c1: Math.max(c0, c1) };
  }
  const wholeRow = trimmed.match(/^\$?([1-9][0-9]*)(?::\$?([1-9][0-9]*))?$/);
  if (wholeRow) {
    const r0 = Number.parseInt(wholeRow[1] ?? '', 10) - 1;
    const r1 = Number.parseInt(wholeRow[2] ?? wholeRow[1] ?? '', 10) - 1;
    if (r0 < 0 || r1 < 0 || r0 > MAX_ROW || r1 > MAX_ROW) return null;
    return { r0: Math.min(r0, r1), c0: 0, r1: Math.max(r0, r1), c1: MAX_COL };
  }
  const parts = trimmed.split(':');
  if (parts.length !== 2) return null;
  const a = parseCellRef(parts[0] ?? '');
  const b = parseCellRef(parts[1] ?? '');
  if (!a || !b) return null;
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  };
}

function parseColRef(raw: string): number | null {
  const col = colFromLetters(raw.trim());
  return col >= 0 && col <= MAX_COL ? col : null;
}
