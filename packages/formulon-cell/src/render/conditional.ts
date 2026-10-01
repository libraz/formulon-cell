import { addrKey } from '../engine/address.js';
import type { CellValue, Range } from '../engine/types.js';
import type {
  CellFormat,
  ConditionalIconSet,
  ConditionalRule,
  ConditionalScalePoint,
  State,
} from '../store/store.js';
import {
  compileFormulaCellPredicate,
  parseFormulaPredicate,
} from './conditional-formula/evaluator.js';

export { parseFormulaPredicate } from './conditional-formula/evaluator.js';

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
  /** Zero-axis position (0..1) for signed data bars. Defaults to 0. */
  barAxis?: number;
  /** Direction from the zero-axis. Defaults to right. */
  barDirection?: 'left' | 'right';
  barColor?: string;
  /** Optional sign-specific border colour for the bar rectangle. */
  barBorderColor?: string;
  /** Colour of the zero / middle axis when it is visible. */
  barAxisColor?: string;
  /** Whether the bar axis should be painted. Explicit false clears lower rules. */
  barAxisVisible?: boolean;
  barGradient?: boolean;
  /** Icon-set artwork + slot index. When set, the painter draws a small
   *  glyph in a left gutter inside the cell. `slot` is 0-based and
   *  bounded by the icon family (3 or 5). */
  iconKind?: ConditionalIconSet;
  iconSlot?: number;
  /** False when conditional formatting should hide the underlying cell value. */
  showValue?: boolean;
}

// Single-slot identity cache. zustand replaces conditional.rules /
// data.cells by reference on every mutation, so a triple reference match
// means the previous evaluation is still valid. Pan, scroll, and selection
// changes leave these references untouched and hit the cache.
let cachedRulesRef: State['conditional']['rules'] | null = null;
let cachedCellsRef: State['data']['cells'] | null = null;
let cachedSheet: number | null = null;
let cachedRtl: boolean | null = null;
let cachedOverlay: Map<string, ConditionalCellOverlay> | null = null;

/** Test hook — drop the cached overlay so the next call recomputes. */
export function _resetConditionalCache(): void {
  cachedRulesRef = null;
  cachedCellsRef = null;
  cachedSheet = null;
  cachedRtl = null;
  cachedOverlay = null;
}

/** Number of slots per icon family. `arrows5` is the only 5-slot family;
 *  the rest land on 3 slots with thresholds at 0.33 / 0.67. */
export function iconSetSlotCount(set: ConditionalIconSet): 3 | 5 {
  return set === 'arrows5' ||
    set === 'quarters5' ||
    set === 'ratings5' ||
    set === 'bars5' ||
    set === 'boxes5'
    ? 5
    : 3;
}

/** Classify `t` (a 0..1 percentile) into a slot index for the icon family.
 *  Uses the spreadsheet's default thresholds — [0.33, 0.67] for 3-slot families and
 *  [0.20, 0.40, 0.60, 0.80] for 5-slot families. */
export function iconSetSlotFor(set: ConditionalIconSet, t: number): number {
  if (iconSetSlotCount(set) === 5) {
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
 *  the input has duplicates — spreadsheet parity. Returns the inclusive cutoff. */
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
    cachedSheet === state.data.sheetIndex &&
    cachedRtl === state.ui.rightToLeft
  ) {
    return cachedOverlay;
  }
  const out = new Map<string, ConditionalCellOverlay>();
  const rules = state.conditional.rules;
  if (rules.length === 0) {
    cachedRulesRef = rules;
    cachedCellsRef = state.data.cells;
    cachedSheet = state.data.sheetIndex;
    cachedRtl = state.ui.rightToLeft;
    cachedOverlay = out;
    return out;
  }
  const sheet = state.data.sheetIndex;
  const stopped = new Set<string>();

  for (let ri = 0; ri < rules.length; ri += 1) {
    const rule = rules[ri];
    if (!rule) continue;
    if (rule.range.sheet !== sheet) continue;
    const ruleOverlay = new Map<string, ConditionalCellOverlay>();

    if (rule.kind === 'cell-value') {
      paintCellValue(state, rule, ruleOverlay);
    } else if (rule.kind === 'color-scale') {
      paintColorScale(state, rule, ruleOverlay);
    } else if (rule.kind === 'data-bar') {
      paintDataBar(state, rule, ruleOverlay);
    } else if (rule.kind === 'icon-set') {
      paintIconSet(state, rule, ruleOverlay);
    } else if (rule.kind === 'top-bottom') {
      paintTopBottom(state, rule, ruleOverlay);
    } else if (rule.kind === 'average') {
      paintAverage(state, rule, ruleOverlay);
    } else if (rule.kind === 'text-contains') {
      paintTextContains(state, rule, ruleOverlay);
    } else if (rule.kind === 'date-occurring') {
      paintDateOccurring(state, rule, ruleOverlay);
    } else if (rule.kind === 'formula') {
      paintFormula(state, rule, ruleOverlay);
    } else if (rule.kind === 'duplicates' || rule.kind === 'unique') {
      paintDupsUnique(state, rule, ruleOverlay);
    } else if (
      rule.kind === 'blanks' ||
      rule.kind === 'non-blanks' ||
      rule.kind === 'errors' ||
      rule.kind === 'no-errors'
    ) {
      paintBlankErrorPredicate(state, rule, ruleOverlay);
    }

    for (const [key, overlay] of ruleOverlay) {
      if (stopped.has(key)) continue;
      const target = out.get(key) ?? {};
      mergeOverlayByPriority(target, overlay);
      out.set(key, target);
      if (rule.stopIfTrue === true) stopped.add(key);
    }
  }

  cachedRulesRef = state.conditional.rules;
  cachedCellsRef = state.data.cells;
  cachedSheet = state.data.sheetIndex;
  cachedRtl = state.ui.rightToLeft;
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
      if (!cell) continue;
      if (!inRange(sheet, r, c, rule.range)) continue;
      if (testCellValue(cell.value, rule.op, rule.a, rule.b)) {
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
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      values.push(cell.value.value);
    }
  }
  if (values.length === 0) return;
  const scale = colorScaleThresholds(rule, values);
  if (!scale) return;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const t = colorScalePosition(v, scale);
      const overlay = out.get(key) ?? {};
      overlay.fill = pickStop(rule.stops, t);
      out.set(key, overlay);
    }
  }
}

interface ColorScaleThresholds {
  low: number;
  mid?: number;
  high: number;
}

function colorScaleThresholds(
  rule: Extract<ConditionalRule, { kind: 'color-scale' }>,
  values: readonly number[],
): ColorScaleThresholds | null {
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .slice()
    .sort((a, b) => a - b);
  if (sorted.length === 0) return null;
  const defaultThresholds =
    rule.stops.length === 2
      ? ([{ kind: 'min' }, { kind: 'max' }] as const)
      : ([{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }] as const);
  const thresholds = rule.thresholds ?? defaultThresholds;
  const low = resolveScalePoint(thresholds[0] ?? { kind: 'min' }, sorted);
  const high = resolveScalePoint(thresholds[thresholds.length - 1] ?? { kind: 'max' }, sorted);
  if (rule.stops.length === 2) return { low, high };
  const mid = resolveScalePoint(thresholds[1] ?? { kind: 'percentile', value: 50 }, sorted);
  return { low, mid, high };
}

function resolveScalePoint(point: ConditionalScalePoint, sorted: readonly number[]): number {
  const min = sorted[0] ?? 0;
  const max = sorted[sorted.length - 1] ?? min;
  if (point.kind === 'min') return min;
  if (point.kind === 'max') return max;
  if (point.kind === 'number') return point.value;
  if (point.kind !== 'percent' && point.kind !== 'percentile') return min;
  const pct = Math.max(0, Math.min(100, point.value));
  if (point.kind === 'percent') return min + ((max - min) * pct) / 100;
  const rank = ((sorted.length - 1) * pct) / 100;
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  const a = sorted[lo] ?? min;
  const b = sorted[hi] ?? a;
  return a + (b - a) * (rank - lo);
}

function colorScalePosition(value: number, thresholds: ColorScaleThresholds): number {
  const low = thresholds.low;
  const high = thresholds.high;
  const mid = thresholds.mid;
  if (mid === undefined) {
    if (high === low) return 0.5;
    return Math.max(0, Math.min(1, (value - low) / (high - low)));
  }
  if (high === low) return 0.5;
  if (value <= mid) {
    if (mid === low) return 0.5;
    return Math.max(0, Math.min(0.5, ((value - low) / (mid - low)) * 0.5));
  }
  if (high === mid) return 0.5;
  return Math.max(0.5, Math.min(1, 0.5 + ((value - mid) / (high - mid)) * 0.5));
}

function paintDataBar(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'data-bar' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      if (Number.isFinite(cell.value.value)) values.push(cell.value.value);
    }
  }
  const sorted = values.slice().sort((a, b) => a - b);
  if (sorted.length === 0) return;
  const resolvedMin = resolveScalePoint(rule.min ?? { kind: 'min' }, sorted);
  const resolvedMax = resolveScalePoint(rule.max ?? { kind: 'max' }, sorted);
  const min = Number.isFinite(resolvedMin) ? resolvedMin : (sorted[0] ?? 0);
  const max = Number.isFinite(resolvedMax) ? resolvedMax : (sorted[sorted.length - 1] ?? min);
  // Automatic endpoints retain the legacy data-bar zero baseline: positive
  // ranges start at zero and negative ranges end at zero. Explicit endpoints
  // replace that side of the baseline with the resolved scale point.
  const hasExplicitMin = rule.min !== undefined && rule.min.kind !== 'min';
  const hasExplicitMax = rule.max !== undefined && rule.max.kind !== 'max';
  const axisPosition = rule.axisPosition ?? 'automatic';
  const lowerBound = axisPosition === 'automatic' && !hasExplicitMin ? Math.min(min, 0) : min;
  const upperBound = axisPosition === 'automatic' && !hasExplicitMax ? Math.max(max, 0) : max;
  const low = Math.min(lowerBound, upperBound);
  const high = Math.max(lowerBound, upperBound);
  const span = high - low;
  const hasNegativePopulation = sorted[0] !== undefined && sorted[0] < 0;
  const lastValue = sorted[sorted.length - 1];
  const hasPositivePopulation = lastValue !== undefined && lastValue > 0;
  const axis =
    axisPosition === 'middle'
      ? 0.5
      : axisPosition === 'none'
        ? 0
        : low < 0 && high > 0
          ? Math.max(0, Math.min(1, Math.abs(low) / (Math.abs(low) + high)))
          : high <= 0
            ? 1
            : 0;
  const clampToBounds = (value: number): number => Math.max(low, Math.min(high, value));
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number' || !Number.isFinite(cell.value.value)) continue;
      const originalValue = cell.value.value;
      const v = clampToBounds(originalValue);
      const overlay = out.get(key) ?? {};
      const negative = originalValue < 0;
      const mirror =
        rule.direction === 'right-to-left' ||
        ((rule.direction === undefined || rule.direction === 'context') &&
          state.ui.rightToLeft === true);
      const length =
        axisPosition === 'automatic'
          ? negative && low < 0 && high > 0
            ? Math.abs(v) / Math.max(Math.abs(low), 1e-9)
            : negative
              ? span === 0
                ? 0
                : (high - v) / span
              : low < 0 && high > 0
                ? v / Math.max(high, 1e-9)
                : span === 0
                  ? 0
                  : (v - low) / span
          : span === 0
            ? 1
            : Math.max(0, Math.min(1, (v - low) / span));
      const noneNegative = axisPosition === 'none' && negative;
      const barSide = noneNegative ? false : negative;
      overlay.bar = barSide
        ? Math.max(0, Math.min(axis, length * axis))
        : Math.max(0, Math.min(1 - axis, length * (1 - axis)));
      overlay.barAxis = mirror ? 1 - axis : axis;
      overlay.barDirection = barSide !== mirror ? 'left' : 'right';
      overlay.barColor =
        negative && rule.negativeColor !== undefined ? rule.negativeColor : rule.color;
      overlay.barBorderColor = negative ? rule.negativeBorderColor : rule.borderColor;
      overlay.barAxisColor = rule.axisColor ?? '#000000';
      overlay.barAxisVisible =
        axisPosition === 'middle' ||
        (axisPosition === 'automatic' &&
          hasNegativePopulation &&
          hasPositivePopulation &&
          axis > 0 &&
          axis < 1);
      overlay.barGradient = rule.gradient === true;
      overlay.showValue = rule.showValue !== false;
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
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      values.push(cell.value.value);
    }
  }
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .slice()
    .sort((a, b) => a - b);
  if (sorted.length === 0) return;
  const min = sorted[0] ?? 0;
  const max = sorted[sorted.length - 1] ?? min;
  const slots = iconSetSlotCount(rule.icons);
  const thresholds = iconSetThresholdValues(rule, sorted);
  const floor = rule.floor
    ? { value: resolveScalePoint(rule.floor, sorted), gte: rule.floor.gte !== false }
    : undefined;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      // `floor` is separate from the N-1 bucket boundaries. A value below
      // it does not match the icon-set rule and should leave the cell alone.
      if (floor !== undefined && (floor.gte ? v < floor.value : v <= floor.value)) continue;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      let slot =
        thresholds === null
          ? iconSetSlotFor(rule.icons, t)
          : thresholds.reduce(
              (count, threshold) =>
                (threshold.gte ? v >= threshold.value : v > threshold.value) ? count + 1 : count,
              0,
            );
      slot = Math.max(0, Math.min(slots - 1, slot));
      if (rule.reverseOrder) slot = slots - 1 - slot;
      const overlay = out.get(key) ?? {};
      overlay.iconKind = rule.icons;
      overlay.iconSlot = slot;
      overlay.showValue = rule.showValue !== false;
      out.set(key, overlay);
    }
  }
}

interface IconSetThreshold {
  value: number;
  gte: boolean;
}

function iconSetThresholdValues(
  rule: Extract<ConditionalRule, { kind: 'icon-set' }>,
  sorted: readonly number[],
): IconSetThreshold[] | null {
  const slots = iconSetSlotCount(rule.icons);
  if (!rule.thresholds || rule.thresholds.length === 0) return null;
  return rule.thresholds
    .slice(0, slots - 1)
    .map((point) => ({
      value: resolveScalePoint(point, sorted),
      gte: point.gte !== false,
    }))
    .sort((a, b) => a.value - b.value);
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
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const passes = rule.mode === 'top' ? v >= cutoff : v <= cutoff;
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintAverage(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'average' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind === 'number' && Number.isFinite(cell.value.value)) {
        values.push(cell.value.value);
      }
    }
  }
  if (values.length === 0) return;
  const avg = values.reduce((sum, v) => sum + v, 0) / values.length;
  const variance = values.reduce((sum, v) => sum + (v - avg) ** 2, 0) / values.length;
  const stdDev = Math.sqrt(variance);
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const passes =
        rule.mode === 'above'
          ? v > avg
          : rule.mode === 'below'
            ? v < avg
            : rule.mode === 'equal-or-above'
              ? v >= avg
              : rule.mode === 'equal-or-below'
                ? v <= avg
                : rule.mode === 'above-std-dev'
                  ? v > avg + stdDev * (rule.stdDev ?? 1)
                  : v < avg - stdDev * (rule.stdDev ?? 1);
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function cellText(v: CellValue): string | null {
  if (v.kind === 'text') return v.value;
  if (v.kind === 'number') return String(v.value);
  if (v.kind === 'bool') return v.value ? 'TRUE' : 'FALSE';
  if (v.kind === 'error') return v.text;
  return null;
}

function paintTextContains(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'text-contains' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const needle = rule.caseSensitive ? rule.text : rule.text.toLocaleLowerCase();
  if (needle.length === 0) return;
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const raw = cellText(cell.value);
      if (raw === null) continue;
      const haystack = rule.caseSensitive ? raw : raw.toLocaleLowerCase();
      const matches =
        rule.mode === 'not-contains'
          ? !haystack.includes(needle)
          : rule.mode === 'begins-with'
            ? haystack.startsWith(needle)
            : rule.mode === 'ends-with'
              ? haystack.endsWith(needle)
              : haystack.includes(needle);
      if (!matches) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

const DAY_MS = 86_400_000;

function normalizeDate(d: Date): number {
  return Math.floor(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) / DAY_MS);
}

function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * DAY_MS);
}

function cellDateDay(v: CellValue): number | null {
  if (v.kind === 'number' && Number.isFinite(v.value))
    return normalizeDate(excelSerialToDate(v.value));
  if (v.kind === 'text') {
    const time = Date.parse(v.value);
    if (Number.isFinite(time)) return normalizeDate(new Date(time));
  }
  return null;
}

function weekStart(day: number): number {
  const d = new Date(day * DAY_MS);
  const dow = (d.getUTCDay() + 6) % 7;
  return day - dow;
}

function monthKey(day: number): number {
  const d = new Date(day * DAY_MS);
  return d.getUTCFullYear() * 12 + d.getUTCMonth();
}

function datePeriodMatches(
  day: number,
  period: Extract<ConditionalRule, { kind: 'date-occurring' }>['period'],
): boolean {
  const today = normalizeDate(new Date());
  switch (period) {
    case 'yesterday':
      return day === today - 1;
    case 'today':
      return day === today;
    case 'tomorrow':
      return day === today + 1;
    case 'last7':
      return day >= today - 6 && day <= today;
    case 'last-week':
      return weekStart(day) === weekStart(today) - 7;
    case 'this-week':
      return weekStart(day) === weekStart(today);
    case 'next-week':
      return weekStart(day) === weekStart(today) + 7;
    case 'last-month':
      return monthKey(day) === monthKey(today) - 1;
    case 'this-month':
      return monthKey(day) === monthKey(today);
    case 'next-month':
      return monthKey(day) === monthKey(today) + 1;
    default:
      return false;
  }
}

function paintDateOccurring(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'date-occurring' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const day = cellDateDay(cell.value);
      if (day === null || !datePeriodMatches(day, rule.period)) continue;
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
  const formulaPredicate = compileFormulaCellPredicate(state, rule);
  const predicate = parseFormulaPredicate(rule.formula);
  if (!formulaPredicate && !predicate) return;
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      const passes = formulaPredicate?.test(r, c) ?? (cell ? predicate?.test(cell.value) : false);
      if (!passes) continue;
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
  value: CellValue,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between',
  a: number | string,
  b: number | string | undefined,
): boolean {
  const v = value.kind === 'number' ? value.value : cellText(value);
  if (v === null) return false;
  if (typeof v === 'number' && typeof a === 'number') {
    return testComparableValue(v, op, a, typeof b === 'number' ? b : undefined);
  }
  return testComparableValue(
    String(v).toLocaleLowerCase(),
    op,
    String(a).toLocaleLowerCase(),
    b === undefined ? undefined : String(b).toLocaleLowerCase(),
  );
}

function testComparableValue<T extends number | string>(
  v: T,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between',
  a: T,
  b: T | undefined,
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
      return b !== undefined && v >= (a <= b ? a : b) && v <= (a <= b ? b : a);
    case 'not-between':
      return b !== undefined && (v < (a <= b ? a : b) || v > (a <= b ? b : a));
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

function mergeOverlayByPriority(
  target: ConditionalCellOverlay,
  source: ConditionalCellOverlay,
): void {
  if (target.fill === undefined && source.fill !== undefined) target.fill = source.fill;
  if (target.color === undefined && source.color !== undefined) target.color = source.color;
  if (target.bold === undefined && source.bold !== undefined) target.bold = source.bold;
  if (target.italic === undefined && source.italic !== undefined) target.italic = source.italic;
  if (target.underline === undefined && source.underline !== undefined) {
    target.underline = source.underline;
  }
  if (target.strike === undefined && source.strike !== undefined) target.strike = source.strike;
  if (target.bar === undefined && source.bar !== undefined) {
    target.bar = source.bar;
    target.barAxis = source.barAxis;
    target.barDirection = source.barDirection;
    target.barColor = source.barColor;
    target.barBorderColor = source.barBorderColor;
    target.barAxisColor = source.barAxisColor;
    target.barAxisVisible = source.barAxisVisible;
    target.barGradient = source.barGradient;
  }
  if (target.iconKind === undefined && source.iconKind !== undefined) {
    target.iconKind = source.iconKind;
    target.iconSlot = source.iconSlot;
  }
  if (target.showValue === undefined && source.showValue !== undefined) {
    target.showValue = source.showValue;
  }
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
