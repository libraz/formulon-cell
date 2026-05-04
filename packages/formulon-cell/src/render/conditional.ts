import type { CellValue, Range } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { CellFormat, ConditionalIconSet, ConditionalRule, State } from '../store/store.js';

const inRange = (sheet: number, row: number, col: number, r: Range): boolean =>
  r.sheet === sheet && row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

/** Per-cell visual outputs derived from the active conditional rules. The
 *  renderer consults this for each painted cell to overlay fills, bars, and
 *  font tweaks. */
export interface ConditionalCellOverlay {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Width fraction (0..1) for a horizontal data bar drawn behind the text.
   *  When set, `barColor` is also defined. */
  bar?: number;
  barColor?: string;
  /** Icon-set artwork + slot index. When set, the painter draws a small
   *  glyph in a left gutter inside the cell. `slot` is 0-based and
   *  bounded by the icon family (3 or 5). */
  iconKind?: ConditionalIconSet;
  iconSlot?: number;
}

// Single-slot identity cache. zustand replaces conditional.rules /
// data.cells by reference on every mutation, so a triple reference match
// means the previous evaluation is still valid. Pan, scroll, and selection
// changes leave these references untouched and hit the cache.
let cachedRulesRef: State['conditional']['rules'] | null = null;
let cachedCellsRef: State['data']['cells'] | null = null;
let cachedSheet: number | null = null;
let cachedOverlay: Map<string, ConditionalCellOverlay> | null = null;

/** Test hook — drop the cached overlay so the next call recomputes. */
export function _resetConditionalCache(): void {
  cachedRulesRef = null;
  cachedCellsRef = null;
  cachedSheet = null;
  cachedOverlay = null;
}

/** Number of slots per icon family. `arrows5` is the only 5-slot family;
 *  the rest land on 3 slots with thresholds at 0.33 / 0.67. */
export function iconSetSlotCount(set: ConditionalIconSet): 3 | 5 {
  return set === 'arrows5' ? 5 : 3;
}

/** Classify `t` (a 0..1 percentile) into a slot index for the icon family.
 *  Uses Excel's default thresholds — [0.33, 0.67] for 3-slot families and
 *  [0.20, 0.40, 0.60, 0.80] for 5-slot families. */
export function iconSetSlotFor(set: ConditionalIconSet, t: number): number {
  if (set === 'arrows5') {
    if (t < 0.2) return 0;
    if (t < 0.4) return 1;
    if (t < 0.6) return 2;
    if (t < 0.8) return 3;
    return 4;
  }
  if (t < 0.33) return 0;
  if (t < 0.67) return 1;
  return 2;
}

/** Pick the cells whose values land in the top-N (or bottom-N) of `values`.
 *  Ties at the threshold all qualify so the result count can exceed `n` when
 *  the input has duplicates — Excel parity. Returns the inclusive cutoff. */
export function topBottomThreshold(
  values: readonly number[],
  mode: 'top' | 'bottom',
  n: number,
  percent: boolean,
): number | null {
  if (values.length === 0 || !Number.isFinite(n) || n <= 0) return null;
  const k = percent
    ? Math.max(1, Math.ceil((values.length * n) / 100))
    : Math.min(values.length, Math.floor(n));
  if (k <= 0) return null;
  const sorted = values.slice().sort((a, b) => (mode === 'top' ? b - a : a - b));
  // The k-th element (1-indexed) is the threshold; ties at the threshold
  // still qualify so `Math.min(k, sorted.length) - 1` is the index.
  const idx = Math.min(k, sorted.length) - 1;
  return sorted[idx] ?? null;
}

interface FormulaPredicate {
  /** Evaluate against a cell's value. Returns true to apply the format. */
  test(v: CellValue): boolean;
}

/** Parse a v1 lightweight predicate: a leading comparison operator
 *  followed by a numeric or quoted-string literal. Anything more complex
 *  returns null and the rule becomes a no-op. */
export function parseFormulaPredicate(raw: string): FormulaPredicate | null {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  // Strip leading `=` for the comparator-prefix path; an `=`-prefixed
  // expression that doesn't fit a comparator template is reserved for
  // engine-side `evaluateText` (not implemented in v1) — return null.
  let body = trimmed;
  if (body.startsWith('=')) body = body.slice(1).trim();
  // Match: <op><whitespace?><literal>
  const m = body.match(/^(>=|<=|<>|>|<|=)\s*(.+)$/);
  if (!m) return null;
  const op = m[1] as '>' | '<' | '>=' | '<=' | '=' | '<>';
  const rhs = m[2]?.trim() ?? '';
  if (rhs === '') return null;
  // Quoted string literal.
  if ((rhs.startsWith('"') && rhs.endsWith('"')) || (rhs.startsWith("'") && rhs.endsWith("'"))) {
    const inner = rhs.slice(1, -1);
    return {
      test(v): boolean {
        const text = v.kind === 'text' ? v.value : v.kind === 'number' ? String(v.value) : null;
        if (text === null) return false;
        return op === '<>' ? text !== inner : op === '=' ? text === inner : false;
      },
    };
  }
  // Numeric literal.
  const num = Number.parseFloat(rhs);
  if (Number.isNaN(num)) return null;
  return {
    test(v): boolean {
      if (v.kind !== 'number') return false;
      const x = v.value;
      switch (op) {
        case '>':
          return x > num;
        case '<':
          return x < num;
        case '>=':
          return x >= num;
        case '<=':
          return x <= num;
        case '=':
          return x === num;
        case '<>':
          return x !== num;
        default:
          return false;
      }
    },
  };
}

/** Stable canonical key for a cell value, used by the duplicates / unique
 *  predicates. Blank cells are skipped (returns null). */
function valueKey(v: CellValue): string | null {
  switch (v.kind) {
    case 'blank':
      return null;
    case 'number':
      return `n:${v.value}`;
    case 'bool':
      return v.value ? 'b:1' : 'b:0';
    case 'text':
      return `t:${v.value}`;
    case 'error':
      return `e:${v.text}`;
  }
}

const isErrorValue = (v: CellValue): boolean => v.kind === 'error';
const isBlankValue = (v: CellValue): boolean => v.kind === 'blank';

/**
 * Evaluate conditional formatting rules for the active sheet's cells. We
 * compute per-rule numeric extremes for color-scale / data-bar rules once,
 * then walk the cell entries assigning overlays.
 */
export function evaluateConditional(state: State): Map<string, ConditionalCellOverlay> {
  if (
    cachedOverlay !== null &&
    cachedRulesRef === state.conditional.rules &&
    cachedCellsRef === state.data.cells &&
    cachedSheet === state.data.sheetIndex
  ) {
    return cachedOverlay;
  }
  const out = new Map<string, ConditionalCellOverlay>();
  const rules = state.conditional.rules;
  if (rules.length === 0) {
    cachedRulesRef = rules;
    cachedCellsRef = state.data.cells;
    cachedSheet = state.data.sheetIndex;
    cachedOverlay = out;
    return out;
  }
  const sheet = state.data.sheetIndex;

  for (let ri = 0; ri < rules.length; ri += 1) {
    const rule = rules[ri];
    if (!rule) continue;
    if (rule.range.sheet !== sheet) continue;

    if (rule.kind === 'cell-value') {
      paintCellValue(state, rule, out);
    } else if (rule.kind === 'color-scale') {
      paintColorScale(state, rule, out);
    } else if (rule.kind === 'data-bar') {
      paintDataBar(state, rule, out);
    } else if (rule.kind === 'icon-set') {
      paintIconSet(state, rule, out);
    } else if (rule.kind === 'top-bottom') {
      paintTopBottom(state, rule, out);
    } else if (rule.kind === 'formula') {
      paintFormula(state, rule, out);
    } else if (rule.kind === 'duplicates' || rule.kind === 'unique') {
      paintDupsUnique(state, rule, out);
    } else if (
      rule.kind === 'blanks' ||
      rule.kind === 'non-blanks' ||
      rule.kind === 'errors' ||
      rule.kind === 'no-errors'
    ) {
      paintBlankErrorPredicate(state, rule, out);
    }
  }

  cachedRulesRef = state.conditional.rules;
  cachedCellsRef = state.data.cells;
  cachedSheet = state.data.sheetIndex;
  cachedOverlay = out;
  return out;
}

function paintCellValue(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'cell-value' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell || cell.value.kind !== 'number') continue;
      if (!inRange(sheet, r, c, rule.range)) continue;
      if (testCellValue(cell.value.value, rule.op, rule.a, rule.b)) {
        const overlay = out.get(key) ?? {};
        mergeApply(overlay, rule.apply);
        out.set(key, overlay);
      }
    }
  }
}

function paintColorScale(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'color-scale' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  if (!Number.isFinite(min)) return;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      const overlay = out.get(key) ?? {};
      overlay.fill = pickStop(rule.stops, t);
      out.set(key, overlay);
    }
  }
}

function paintDataBar(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'data-bar' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  if (!Number.isFinite(min)) return;
  const denom = Math.max(Math.abs(min), Math.abs(max), 1e-9);
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      const overlay = out.get(key) ?? {};
      overlay.bar = Math.max(0, Math.min(1, Math.abs(v) / denom));
      overlay.barColor = rule.color;
      out.set(key, overlay);
    }
  }
}

function paintIconSet(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'icon-set' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  if (!Number.isFinite(min)) return;
  const slots = iconSetSlotCount(rule.icons);
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      let slot = iconSetSlotFor(rule.icons, t);
      if (rule.reverseOrder) slot = slots - 1 - slot;
      const overlay = out.get(key) ?? {};
      overlay.iconKind = rule.icons;
      overlay.iconSlot = slot;
      out.set(key, overlay);
    }
  }
}

function paintTopBottom(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'top-bottom' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell && cell.value.kind === 'number') values.push(cell.value.value);
    }
  }
  const cutoff = topBottomThreshold(values, rule.mode, rule.n, rule.percent ?? false);
  if (cutoff === null) return;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell || cell.value.kind !== 'number') continue;
      const v = cell.value.value;
      const passes = rule.mode === 'top' ? v >= cutoff : v <= cutoff;
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintFormula(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'formula' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const predicate = parseFormulaPredicate(rule.formula);
  if (!predicate) return;
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      if (!predicate.test(cell.value)) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintDupsUnique(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'duplicates' | 'unique' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const counts = new Map<string, number>();
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell) continue;
      const k = valueKey(cell.value);
      if (k === null) continue;
      counts.set(k, (counts.get(k) ?? 0) + 1);
    }
  }
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const k = valueKey(cell.value);
      if (k === null) continue;
      const count = counts.get(k) ?? 0;
      const passes = rule.kind === 'duplicates' ? count > 1 : count === 1;
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintBlankErrorPredicate(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'blanks' | 'non-blanks' | 'errors' | 'no-errors' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      const value: CellValue = cell?.value ?? { kind: 'blank' };
      let passes = false;
      if (rule.kind === 'blanks') passes = isBlankValue(value);
      else if (rule.kind === 'non-blanks') passes = !isBlankValue(value);
      else if (rule.kind === 'errors') passes = isErrorValue(value);
      else if (rule.kind === 'no-errors') passes = !isErrorValue(value) && !isBlankValue(value);
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function testCellValue(
  v: number,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between',
  a: number,
  b: number | undefined,
): boolean {
  switch (op) {
    case '>':
      return v > a;
    case '<':
      return v < a;
    case '>=':
      return v >= a;
    case '<=':
      return v <= a;
    case '=':
      return v === a;
    case '<>':
      return v !== a;
    case 'between':
      return b !== undefined && v >= Math.min(a, b) && v <= Math.max(a, b);
    case 'not-between':
      return b !== undefined && (v < Math.min(a, b) || v > Math.max(a, b));
    default:
      return false;
  }
}

function mergeApply(target: ConditionalCellOverlay, patch: Partial<CellFormat>): void {
  if (patch.fill) target.fill = patch.fill;
  if (patch.color) target.color = patch.color;
  if (patch.bold) target.bold = true;
  if (patch.italic) target.italic = true;
  if (patch.underline) target.underline = true;
  if (patch.strike) target.strike = true;
}

function pickStop(stops: readonly string[], t: number): string {
  const s0 = stops[0] ?? '#000000';
  const s1 = stops[1] ?? s0;
  const s2 = stops[2] ?? s1;
  if (stops.length === 2) return interpolate(s0, s1, t);
  // Three-stop: low, mid, high
  if (t <= 0.5) return interpolate(s0, s1, t * 2);
  return interpolate(s1, s2, (t - 0.5) * 2);
}

function interpolate(a: string, b: string, t: number): string {
  const ca = parseColor(a);
  const cb = parseColor(b);
  if (!ca || !cb) return a;
  const r = Math.round(ca[0] + (cb[0] - ca[0]) * t);
  const g = Math.round(ca[1] + (cb[1] - ca[1]) * t);
  const blu = Math.round(ca[2] + (cb[2] - ca[2]) * t);
  return `rgb(${r}, ${g}, ${blu})`;
}

function parseColor(s: string): [number, number, number] | null {
  const m = s.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (m) {
    const hex = m[1] ?? '';
    if (hex.length === 3) {
      const h0 = hex[0] ?? '0';
      const h1 = hex[1] ?? '0';
      const h2 = hex[2] ?? '0';
      return [
        Number.parseInt(h0 + h0, 16),
        Number.parseInt(h1 + h1, 16),
        Number.parseInt(h2 + h2, 16),
      ];
    }
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16),
    ];
  }
  const rgb = s.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (rgb) {
    return [
      Number.parseInt(rgb[1] ?? '0', 10),
      Number.parseInt(rgb[2] ?? '0', 10),
      Number.parseInt(rgb[3] ?? '0', 10),
    ];
  }
  return null;
}

/** Used by ConditionalRule consumer types — re-exported through index. */
export type { ConditionalRule };
