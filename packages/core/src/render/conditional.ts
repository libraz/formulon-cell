import type { Range } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { CellFormat, ConditionalRule, State } from '../store/store.js';

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

  // Pre-compute min/max per rule for color-scale and data-bar rules.
  const stats = rules.map((rule) => {
    if (rule.kind !== 'color-scale' && rule.kind !== 'data-bar') return null;
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
      for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
        const cell = state.data.cells.get(addrKey({ sheet: rule.range.sheet, row: r, col: c }));
        if (!cell || cell.value.kind !== 'number') continue;
        const v = cell.value.value;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    return { min, max };
  });

  for (let ri = 0; ri < rules.length; ri += 1) {
    const rule = rules[ri]!;
    if (rule.range.sheet !== sheet) continue;
    const stat = stats[ri];
    for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
      for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
        const key = addrKey({ sheet, row: r, col: c });
        const cell = state.data.cells.get(key);
        if (!cell) continue;
        if (!inRange(sheet, r, c, rule.range)) continue;
        const v = cell.value.kind === 'number' ? cell.value.value : null;
        if (v === null) continue;
        const overlay = out.get(key) ?? {};
        if (rule.kind === 'cell-value') {
          if (testCellValue(v, rule.op, rule.a, rule.b)) {
            mergeApply(overlay, rule.apply);
          }
        } else if (rule.kind === 'color-scale' && stat && Number.isFinite(stat.min)) {
          const t = stat.max === stat.min ? 0.5 : (v - stat.min) / (stat.max - stat.min);
          overlay.fill = pickStop(rule.stops, t);
        } else if (rule.kind === 'data-bar' && stat && Number.isFinite(stat.min)) {
          const denom = Math.max(Math.abs(stat.min), Math.abs(stat.max), 1e-9);
          overlay.bar = Math.max(0, Math.min(1, Math.abs(v) / denom));
          overlay.barColor = rule.color;
        }
        out.set(key, overlay);
      }
    }
  }
  cachedRulesRef = state.conditional.rules;
  cachedCellsRef = state.data.cells;
  cachedSheet = state.data.sheetIndex;
  cachedOverlay = out;
  return out;
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
  if (stops.length === 2) return interpolate(stops[0]!, stops[1]!, t);
  // Three-stop: low, mid, high
  if (t <= 0.5) return interpolate(stops[0]!, stops[1]!, t * 2);
  return interpolate(stops[1]!, stops[2]!, (t - 0.5) * 2);
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
      return [
        Number.parseInt(hex[0]! + hex[0]!, 16),
        Number.parseInt(hex[1]! + hex[1]!, 16),
        Number.parseInt(hex[2]! + hex[2]!, 16),
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
      Number.parseInt(rgb[1]!, 10),
      Number.parseInt(rgb[2]!, 10),
      Number.parseInt(rgb[3]!, 10),
    ];
  }
  return null;
}

/** Used by ConditionalRule consumer types — re-exported through index. */
export type { ConditionalRule };
