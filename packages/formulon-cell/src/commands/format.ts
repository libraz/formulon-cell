import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import {
  type CellAlign,
  type CellBorderSide,
  type CellBorderStyle,
  type CellBorders,
  type CellFormat,
  type CellVAlign,
  mutators,
  type NumFmt,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { gateProtection, isCellWritable } from './protection.js';

type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

function eachKey(range: Range, fn: (key: string) => void): void {
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      fn(addrKey({ sheet: range.sheet, row: r, col: c }));
    }
  }
}

/** Apply a partial format patch to every writable cell in `range`. When
 *  the sheet is unprotected this falls back to the bulk
 *  `mutators.setRangeFormat` path (one setState call); when protected it
 *  walks the range cell-by-cell and only writes through the cells that
 *  are explicitly unlocked. Returns false when the entire range is gated
 *  so the caller can short-circuit (e.g. emit a single warning). */
function applyRangePatch(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  patch: Partial<CellFormat> | null,
): boolean {
  const allowed = gateProtection(state, range);
  if (allowed === null) return false;
  // Fast path: sheet unprotected, or entire range writable.
  // We can still use the bulk mutator because per-cell-locked cells inside
  // a partially-unlocked range need to be skipped — fall through to the
  // per-cell loop in that case.
  if (!state.protection.protectedSheets.has(range.sheet)) {
    mutators.setRangeFormat(store, range, patch);
    return true;
  }
  // Per-cell loop respecting individual lock flags.
  const sheet = range.sheet;
  let wroteAny = false;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      mutators.setCellFormat(store, addr, patch);
      wroteAny = true;
    }
  }
  return wroteAny;
}

function allHave(
  state: State,
  range: Range,
  predicate: (f: CellFormat | undefined) => boolean,
): boolean {
  let all = true;
  eachKey(range, (key) => {
    if (!predicate(state.format.formats.get(key))) all = false;
  });
  return all;
}

function anyHave(
  state: State,
  range: Range,
  predicate: (f: CellFormat | undefined) => boolean,
): boolean {
  let any = false;
  eachKey(range, (key) => {
    if (predicate(state.format.formats.get(key))) any = true;
  });
  return any;
}

function toggleFlag(state: State, store: SpreadsheetStore, key: ToggleKey): void {
  const range = state.selection.range;
  const allOn = allHave(state, range, (f) => f?.[key] === true);
  applyRangePatch(state, store, range, { [key]: !allOn } as Partial<CellFormat>);
}

export function toggleBold(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'bold');
}

export function toggleItalic(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'italic');
}

export function toggleUnderline(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'underline');
}

export function toggleStrike(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'strike');
}

export function setAlign(state: State, store: SpreadsheetStore, align: CellAlign): void {
  applyRangePatch(state, store, state.selection.range, { align });
}

export function setVAlign(state: State, store: SpreadsheetStore, vAlign: CellVAlign): void {
  applyRangePatch(state, store, state.selection.range, { vAlign });
}

export function toggleWrap(state: State, store: SpreadsheetStore): void {
  const range = state.selection.range;
  const allOn = allHave(state, range, (f) => f?.wrap === true);
  applyRangePatch(state, store, range, { wrap: !allOn });
}

export function bumpIndent(state: State, store: SpreadsheetStore, delta: 1 | -1): void {
  const range = state.selection.range;
  if (gateProtection(state, range) === null) return;
  const sheet = range.sheet;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      const fmt = state.format.formats.get(addrKey(addr));
      const cur = fmt?.indent ?? 0;
      const next = Math.max(0, Math.min(15, cur + delta));
      mutators.setCellFormat(store, addr, { indent: next });
    }
  }
}

export function setRotation(state: State, store: SpreadsheetStore, deg: number): void {
  const r = Math.max(-90, Math.min(90, Math.round(deg)));
  applyRangePatch(state, store, state.selection.range, { rotation: r });
}

export function setNumFmt(state: State, store: SpreadsheetStore, fmt: NumFmt): void {
  applyRangePatch(state, store, state.selection.range, { numFmt: fmt });
}

export function cycleCurrency(state: State, store: SpreadsheetStore): void {
  const range = state.selection.range;
  const hasCurrency = anyHave(state, range, (f) => f?.numFmt?.kind === 'currency');
  const next: NumFmt = hasCurrency
    ? { kind: 'general' }
    : { kind: 'currency', decimals: 2, symbol: '$' };
  applyRangePatch(state, store, range, { numFmt: next });
}

export function cyclePercent(state: State, store: SpreadsheetStore): void {
  const range = state.selection.range;
  const hasPercent = anyHave(state, range, (f) => f?.numFmt?.kind === 'percent');
  const next: NumFmt = hasPercent ? { kind: 'general' } : { kind: 'percent', decimals: 0 };
  applyRangePatch(state, store, range, { numFmt: next });
}

const clampDecimals = (n: number): number => Math.max(0, Math.min(10, n));

export function bumpDecimals(state: State, store: SpreadsheetStore, delta: 1 | -1): void {
  const range = state.selection.range;
  if (gateProtection(state, range) === null) return;
  const sheet = range.sheet;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      const fmt = state.format.formats.get(addrKey(addr));
      const cur = fmt?.numFmt;
      let nextFmt: NumFmt | undefined;
      if (!cur || cur.kind === 'general') {
        if (delta === 1) nextFmt = { kind: 'fixed', decimals: 2 };
      } else if (cur.kind === 'fixed') {
        nextFmt = { kind: 'fixed', decimals: clampDecimals(cur.decimals + delta) };
      } else if (cur.kind === 'currency') {
        nextFmt = {
          kind: 'currency',
          decimals: clampDecimals(cur.decimals + delta),
          ...(cur.symbol !== undefined ? { symbol: cur.symbol } : {}),
        };
      } else if (cur.kind === 'percent') {
        nextFmt = { kind: 'percent', decimals: clampDecimals(cur.decimals + delta) };
      }
      if (nextFmt) mutators.setCellFormat(store, addr, { numFmt: nextFmt });
    }
  }
}

export function setBorders(state: State, store: SpreadsheetStore, sides: CellBorders): void {
  applyRangePatch(state, store, state.selection.range, { borders: sides });
}

export type BorderPreset =
  | 'none'
  | 'outline'
  | 'all'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom';

export function setBorderPreset(
  state: State,
  store: SpreadsheetStore,
  preset: BorderPreset,
  style: CellBorderStyle = 'thin',
): void {
  const range = state.selection.range;
  if (gateProtection(state, range) === null) return;
  const side: CellBorderSide = { style };
  if (preset === 'all') {
    applyRangePatch(state, store, range, {
      borders: { top: side, right: side, bottom: side, left: side },
    });
    return;
  }
  if (preset === 'none') {
    applyRangePatch(state, store, range, {
      borders: {
        top: false,
        right: false,
        bottom: false,
        left: false,
        diagonalDown: false,
        diagonalUp: false,
      },
    });
    return;
  }
  const sheet = range.sheet;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      const borders: CellBorders = {};
      if (preset === 'outline') {
        if (r === range.r0) borders.top = side;
        if (r === range.r1) borders.bottom = side;
        if (c === range.c0) borders.left = side;
        if (c === range.c1) borders.right = side;
      } else if (preset === 'top' && r === range.r0) borders.top = side;
      else if (preset === 'bottom' && r === range.r1) borders.bottom = side;
      else if (preset === 'left' && c === range.c0) borders.left = side;
      else if (preset === 'right' && c === range.c1) borders.right = side;
      else if (preset === 'doubleBottom' && r === range.r1) borders.bottom = { style: 'double' };
      if (borders.top || borders.right || borders.bottom || borders.left) {
        mutators.setCellFormat(store, addr, { borders });
      }
    }
  }
}

/** Toolbar default: outline if missing, all-borders if outline present, else clear.
 *  Three-step cycle on repeated clicks. */
export function cycleBorders(state: State, store: SpreadsheetStore): void {
  const range = state.selection.range;
  if (gateProtection(state, range) === null) return;
  const hasAny = anyHave(state, range, (f) => {
    const b = f?.borders;
    return !!(b && (b.top || b.right || b.bottom || b.left));
  });
  if (!hasAny) {
    // Outline: paint only the perimeter sides.
    const sheet = range.sheet;
    for (let r = range.r0; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        const addr = { sheet, row: r, col: c };
        if (!isCellWritable(state, addr)) continue;
        const sides: CellBorders = {};
        if (r === range.r0) sides.top = true;
        if (r === range.r1) sides.bottom = true;
        if (c === range.c0) sides.left = true;
        if (c === range.c1) sides.right = true;
        if (sides.top || sides.right || sides.bottom || sides.left) {
          mutators.setCellFormat(store, addr, { borders: sides });
        }
      }
    }
    return;
  }
  const allFour = allHave(
    state,
    range,
    (f) => !!(f?.borders?.top && f.borders.right && f.borders.bottom && f.borders.left),
  );
  if (allFour) {
    applyRangePatch(state, store, range, {
      borders: { top: false, right: false, bottom: false, left: false },
    });
    return;
  }
  applyRangePatch(state, store, range, {
    borders: { top: true, right: true, bottom: true, left: true },
  });
}

export function clearFormat(state: State, store: SpreadsheetStore): void {
  const range = state.selection.range;
  if (gateProtection(state, range) === null) return;
  if (!state.protection.protectedSheets.has(range.sheet)) {
    mutators.setRangeFormat(store, range, null);
    return;
  }
  // Per-cell loop respecting individual lock flags.
  const sheet = range.sheet;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      mutators.setCellFormat(store, addr, null);
    }
  }
}

/** Set or clear the font color across the selection. Pass `null` to clear. */
export function setFontColor(state: State, store: SpreadsheetStore, color: string | null): void {
  if (color === null) {
    applyRangePatch(state, store, state.selection.range, { color: undefined });
    return;
  }
  applyRangePatch(state, store, state.selection.range, { color });
}

/** Set or clear the fill (background) color across the selection. */
export function setFillColor(state: State, store: SpreadsheetStore, color: string | null): void {
  if (color === null) {
    applyRangePatch(state, store, state.selection.range, { fill: undefined });
    return;
  }
  applyRangePatch(state, store, state.selection.range, { fill: color });
}

/** Update font family and/or size across the selection. */
export function setFont(
  state: State,
  store: SpreadsheetStore,
  patch: { fontFamily?: string | null; fontSize?: number | null },
): void {
  const next: Partial<CellFormat> = {};
  if (patch.fontFamily !== undefined) {
    next.fontFamily = patch.fontFamily ?? undefined;
  }
  if (patch.fontSize !== undefined) {
    next.fontSize = patch.fontSize ?? undefined;
  }
  applyRangePatch(state, store, state.selection.range, next);
}

export function formatNumber(value: number, fmt: NumFmt | undefined, locale = 'en-US'): string {
  if (!Number.isFinite(value)) return String(value);
  if (!fmt || fmt.kind === 'general') {
    return new Intl.NumberFormat(locale, { maximumFractionDigits: 12 }).format(value);
  }
  if (fmt.kind === 'text') return String(value);
  if (fmt.kind === 'fixed') {
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: !!fmt.thousands,
    }).format(Math.abs(value));
    return applyNegative(value, body, '', negStyle);
  }
  if (fmt.kind === 'currency') {
    const symbol = fmt.symbol ?? '$';
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    return applyNegative(value, body, symbol, negStyle);
  }
  if (fmt.kind === 'percent') {
    return new Intl.NumberFormat(locale, {
      style: 'percent',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
    }).format(value);
  }
  if (fmt.kind === 'scientific') {
    return value.toExponential(fmt.decimals).replace('e', 'E');
  }
  if (fmt.kind === 'accounting') {
    const symbol = fmt.symbol ?? '$';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    if (value === 0) return `${symbol} -`;
    return value < 0 ? `(${symbol}${body})` : `${symbol}${body} `;
  }
  if (fmt.kind === 'date') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'time') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'datetime') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'custom') {
    return formatCustomPattern(value, fmt.pattern, locale);
  }
  return String(value);
}

function applyNegative(
  value: number,
  body: string,
  symbol: string,
  style: 'minus' | 'parens' | 'red' | 'red-parens',
): string {
  const positive = `${symbol}${body}`;
  if (value >= 0) return positive;
  switch (style) {
    case 'parens':
      return `(${symbol}${body})`;
    case 'red':
      return `-${symbol}${body}`; // color is applied at paint time
    case 'red-parens':
      return `(${symbol}${body})`;
    default:
      return `-${symbol}${body}`;
  }
}

/** Spreadsheet serial date → JS Date. spreadsheet epoch is 1899-12-30 (with Lotus 123
 *  1900-leap-year bug compensation already baked in for serials > 60). */
function spreadsheetSerialToDate(serial: number): Date {
  const ms = (serial - 25569) * 86_400_000;
  return new Date(ms);
}

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/** Spreadsheet-style custom format mini-language. Supports section splitting
 *  (pos;neg;zero;text), `0`/`#`/`?` digit placeholders, `.` decimal, `,`
 *  thousands & scaling, `%`, `\\X` escape, `"text"` literals, `[Red]`-style
 *  color tags (stripped — color is applied at paint time), and date tokens
 *  `yyyy`/`yy`/`mmmm`/`mmm`/`mm`/`m`/`dddd`/`ddd`/`dd`/`d`/`hh`/`h`/`ss`/`s`
 *  plus `am/pm`. Not exhaustive but covers the patterns spreadsheets ship in its
 *  built-in format codes. */
function formatCustomPattern(value: number, pattern: string, locale: string): string {
  // Split into up to four sections on ';' that aren't inside a quoted literal
  //  or a bracketed tag. Spreadsheets allow: positive;negative;zero;text. When any
  //  section carries a [>n]/[<n]/[=n] condition we evaluate those first and
  //  only fall back to the sign-based default when no condition matches.
  const sections = splitSections(pattern);

  let active: string | null = null;
  let useAbs = false;
  // First pass: try condition-bearing sections.
  for (const sec of sections) {
    const cond = parseCondition(sec);
    if (!cond) continue;
    if (cond.test(value)) {
      active = cond.body;
      // Condition matched → caller already wrote the comparator into the
      //  literal so we should NOT show a leading minus from `value` itself.
      useAbs = value < 0;
      break;
    }
  }
  if (active === null) {
    // Second pass: classic sign-based selection over sections that carry
    //  no explicit condition.
    const plain = sections.map((s) => (parseCondition(s) ? null : s));
    const pos = plain[0] ?? null;
    const neg = plain[1] ?? null;
    const zero = plain[2] ?? null;
    if (value < 0 && neg) {
      active = neg;
      useAbs = true;
    } else if (value === 0 && zero) {
      active = zero;
    } else {
      active = pos ?? sections[0] ?? '';
    }
  }

  // Strip color tags. Color application belongs to the painter, not here.
  active = normalizeFormatSection(active);

  // If the section contains date/time tokens, render as date.
  if (/y|m|d|h|s/.test(stripLiterals(active))) {
    return renderDateTimePattern(value, active, locale);
  }

  return renderNumericPattern(useAbs ? Math.abs(value) : value, active, locale);
}

function normalizeFormatSection(section: string): string {
  return (
    section
      // Locale/currency tags: [$¥-411]#,##0 → ¥#,##0; [$-ja-JP] is locale-only.
      .replace(/\[\$([^\]-]+)(?:-[^\]]+)?\]/g, '$1')
      .replace(/\[\$-[^\]]+\]/g, '')
      // Color tags are a style concern; the formatter returns text only.
      .replace(/\[(?:Red|Green|Blue|Black|White|Yellow|Magenta|Cyan|Color\d+)\]/gi, '')
      .replace(/"([^"]*)"/g, '$1')
      // Alignment/fill directives. `_x` reserves one char width; `*x`
      // repeats a fill char. Canvas text output should not show either.
      .replace(/_.|\\ /g, '')
      .replace(/\*./g, '')
  );
}

/** Parse a leading condition tag like `[>100]"big"0` into its predicate and
 *  the remaining body. Returns null when the section has no condition. */
function parseCondition(section: string): { test: (n: number) => boolean; body: string } | null {
  const m = section.match(/^\s*\[(>=|<=|<>|=|>|<)\s*(-?\d+(?:\.\d+)?)\s*\](.*)$/s);
  if (!m) return null;
  const op = m[1] ?? '=';
  const target = Number.parseFloat(m[2] ?? '0');
  const body = m[3] ?? '';
  const test = (n: number): boolean => {
    switch (op) {
      case '>':
        return n > target;
      case '<':
        return n < target;
      case '>=':
        return n >= target;
      case '<=':
        return n <= target;
      case '<>':
        return n !== target;
      default:
        return n === target;
    }
  };
  return { test, body };
}

/** Split a format string on `;` that's not inside `"..."` or `[...]`. */
function splitSections(s: string): string[] {
  const out: string[] = [];
  let buf = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      buf += ch + s[i + 1];
      i += 1;
      continue;
    }
    if (!inBracket && ch === '"') {
      inQuote = !inQuote;
      buf += ch;
      continue;
    }
    if (!inQuote && ch === '[') {
      inBracket = true;
      buf += ch;
      continue;
    }
    if (inBracket && ch === ']') {
      inBracket = false;
      buf += ch;
      continue;
    }
    if (!inQuote && !inBracket && ch === ';') {
      out.push(buf);
      buf = '';
      continue;
    }
    buf += ch;
  }
  out.push(buf);
  return out;
}

/** Strip quoted literals and escapes so the remaining string can be probed
 *  for unescaped tokens (e.g. detecting `m` as a month token). */
function stripLiterals(s: string): string {
  return s.replace(/"[^"]*"/g, '').replace(/\\./g, '');
}

function renderDateTimePattern(serial: number, pattern: string, locale: string): string {
  const d = spreadsheetSerialToDate(serial);
  const yyyy = d.getUTCFullYear();
  const mm = d.getUTCMonth() + 1;
  const dd = d.getUTCDate();
  const dow = d.getUTCDay();
  const hh = d.getUTCHours();
  const mi = d.getUTCMinutes();
  const ss = d.getUTCSeconds();
  const has12h = /a\/?p|am\/pm/i.test(pattern);
  const hh12 = ((hh + 11) % 12) + 1;
  const ampm = hh < 12 ? 'AM' : 'PM';
  const DAYS_LONG = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { weekday: 'long', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, 0, i + 1)),
    ),
  );
  const DAYS_SHORT = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { weekday: 'short', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, 0, i + 1)),
    ),
  );
  const MONTHS_LONG = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, i, 1)),
    ),
  );
  const MONTHS_SHORT = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, i, 1)),
    ),
  );
  let out = '';
  let prevWasH = false;
  for (let i = 0; i < pattern.length; ) {
    const rest = pattern.slice(i);
    // Quoted literal — emit without the quotes.
    if (pattern[i] === '"') {
      const end = pattern.indexOf('"', i + 1);
      if (end < 0) {
        out += pattern.slice(i + 1);
        break;
      }
      out += pattern.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (pattern[i] === '\\' && i + 1 < pattern.length) {
      out += pattern[i + 1];
      i += 2;
      continue;
    }
    // Token matching, longest first. After hours, an "m"/"mm" token is
    //  minutes, not month — track prevWasH to disambiguate.
    let tok = '';
    if (rest.startsWith('yyyy')) tok = 'yyyy';
    else if (rest.startsWith('yy')) tok = 'yy';
    else if (rest.startsWith('mmmm')) tok = 'mmmm';
    else if (rest.startsWith('mmm')) tok = 'mmm';
    else if (rest.startsWith('mm')) tok = 'mm';
    else if (rest.startsWith('dddd')) tok = 'dddd';
    else if (rest.startsWith('ddd')) tok = 'ddd';
    else if (rest.startsWith('dd')) tok = 'dd';
    else if (rest.startsWith('hh')) tok = 'hh';
    else if (rest.startsWith('ss')) tok = 'ss';
    else if (/^am\/pm/i.test(rest)) tok = 'am/pm';
    else if (/^a\/p/i.test(rest)) tok = 'a/p';
    else if (rest[0] === 'm') tok = 'm';
    else if (rest[0] === 'd') tok = 'd';
    else if (rest[0] === 'h') tok = 'h';
    else if (rest[0] === 's') tok = 's';

    if (!tok) {
      out += pattern[i];
      i += 1;
      prevWasH = false;
      continue;
    }
    switch (tok) {
      case 'yyyy':
        out += String(yyyy);
        break;
      case 'yy':
        out += String(yyyy).slice(-2);
        break;
      case 'mmmm':
        out += MONTHS_LONG[mm - 1] ?? '';
        break;
      case 'mmm':
        out += MONTHS_SHORT[mm - 1] ?? '';
        break;
      case 'mm':
        out += prevWasH ? pad2(mi) : pad2(mm);
        break;
      case 'm':
        out += prevWasH ? String(mi) : String(mm);
        break;
      case 'dddd':
        out += DAYS_LONG[dow] ?? '';
        break;
      case 'ddd':
        out += DAYS_SHORT[dow] ?? '';
        break;
      case 'dd':
        out += pad2(dd);
        break;
      case 'd':
        out += String(dd);
        break;
      case 'hh':
        out += pad2(has12h ? hh12 : hh);
        break;
      case 'h':
        out += String(has12h ? hh12 : hh);
        break;
      case 'ss':
        out += pad2(ss);
        break;
      case 's':
        out += String(ss);
        break;
      case 'am/pm':
        out += /AM\/PM/.test(pattern.slice(i, i + 5)) ? ampm : ampm.toLowerCase();
        break;
      case 'a/p':
        out += hh < 12 ? 'A' : 'P';
        break;
    }
    prevWasH = tok === 'h' || tok === 'hh';
    i += tok.length;
  }
  return out;
}

function renderNumericPattern(value: number, pattern: string, locale: string): string {
  // Detect trailing thousand-scaling commas (e.g. "0,," divides by 1e6).
  let scale = 1;
  let body = pattern;
  // Remove trailing commas after the last digit placeholder for scaling.
  const trailingCommas = body.match(/[0#?](,+)\s*[^0#?]*$/);
  if (trailingCommas) {
    const commas = trailingCommas[1] ?? '';
    scale = 10 ** (3 * commas.length);
    // Remove just those commas from the body.
    const i = body.lastIndexOf(commas);
    if (i >= 0) body = body.slice(0, i) + body.slice(i + commas.length);
  }
  const isPercent = body.includes('%');
  let scaled = value / scale;
  if (isPercent) scaled *= 100;

  // Find the digit-placeholder block surrounding (and including) the decimal.
  const placeholderMatch = body.match(/[#0?][#0?,]*(?:\.[#0?]+)?|\.[#0?]+/);
  if (!placeholderMatch) return body;

  const block = placeholderMatch[0];
  const dotIndex = block.indexOf('.');
  const intPart = dotIndex >= 0 ? block.slice(0, dotIndex) : block;
  const fracPart = dotIndex >= 0 ? block.slice(dotIndex + 1) : '';
  const grouping = intPart.includes(',');
  const minIntDigits = (intPart.match(/0/g) ?? []).length;
  const minFracDigits = (fracPart.match(/0/g) ?? []).length;
  const maxFracDigits = (fracPart.match(/[0#?]/g) ?? []).length;

  const formatted = new Intl.NumberFormat(locale, {
    minimumIntegerDigits: Math.max(1, minIntDigits),
    minimumFractionDigits: minFracDigits,
    maximumFractionDigits: maxFracDigits,
    useGrouping: grouping,
  }).format(scaled);

  return body.replace(block, formatted);
}
