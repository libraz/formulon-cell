import type { ConditionalCellOverlay } from '../render/conditional.js';
import { iconSetSlotCount } from '../render/conditional.js';
import type { ConditionalIconSet, SpreadsheetStore } from '../store/store.js';
import type { CellFormat, ConditionalRule, ConditionalScalePoint } from '../store/types.js';
import { addrKey } from './address.js';
import {
  borderRecordToFormat,
  fillRecordToFormat,
  fontRecordToFormat,
  formatCodeToNumFmt,
} from './format-writeback.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** CF match kind ordinals — mirror of `formulon::cf::CFMatchKind`. */
const KIND_COLOR_SCALE = 1;
const KIND_DATA_BAR = 2;
const KIND_ICON_SET = 3;

const ENGINE_ICON_SETS: readonly ConditionalIconSet[] = [
  'arrows3',
  'arrows5',
  'triangles3',
  'traffic3',
  'trafficRim3',
  'symbols3',
  'flags3',
  'stars3',
  'quarters5',
  'ratings5',
  'bars5',
  'boxes5',
];

const ENGINE_RULE_TYPE = {
  expression: 0,
  cellIs: 1,
  colorScale: 2,
  dataBar: 3,
  iconSet: 4,
  top10: 5,
  aboveAverage: 6,
  containsText: 7,
  notContainsText: 8,
  beginsWith: 9,
  endsWith: 10,
  containsBlanks: 11,
  notContainsBlanks: 12,
  containsErrors: 13,
  notContainsErrors: 14,
  duplicateValues: 16,
  uniqueValues: 17,
} as const;

const VALUE_OBJECT_TYPE = {
  number: 0,
  percent: 1,
  percentile: 2,
  min: 3,
  max: 4,
} as const;

const ENGINE_CELL_IS_OP: Record<number, Extract<ConditionalRule, { kind: 'cell-value' }>['op']> = {
  0: '<',
  1: '<=',
  2: '=',
  3: '<>',
  4: '>=',
  5: '>',
  6: 'between',
  7: 'not-between',
};

type ConditionalFormatEntry = ReturnType<WorkbookHandle['getConditionalFormats']>[number];

const rgba = (c: { r: number; g: number; b: number; a: number }): string =>
  c.a >= 255
    ? `rgb(${c.r}, ${c.g}, ${c.b})`
    : `rgba(${c.r}, ${c.g}, ${c.b}, ${(c.a / 255).toFixed(3)})`;

const engineIconSet = (ordinal: number): ConditionalIconSet | null =>
  Number.isInteger(ordinal) && ordinal >= 0 && ordinal < ENGINE_ICON_SETS.length
    ? (ENGINE_ICON_SETS[ordinal] ?? null)
    : null;

function engineScalePoint(
  valueObject: NonNullable<ConditionalFormatEntry['colorScale']>['thresholds'][number] | undefined,
): ConditionalScalePoint {
  if (!valueObject) return { kind: 'min' };
  if (valueObject.type === VALUE_OBJECT_TYPE.min) return { kind: 'min' };
  if (valueObject.type === VALUE_OBJECT_TYPE.max) return { kind: 'max' };
  const raw = Number(valueObject.value ?? '0');
  const value = Number.isFinite(raw) ? raw : 0;
  if (valueObject.type === VALUE_OBJECT_TYPE.percent) return { kind: 'percent', value };
  if (valueObject.type === VALUE_OBJECT_TYPE.percentile) return { kind: 'percentile', value };
  return { kind: 'number', value };
}

const maybeNumber = (raw: string | undefined): number | string => {
  if (raw === undefined) return '';
  const trimmed = raw.trim();
  if (trimmed === '') return raw;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : raw;
};

const rangesOf = (sheet: number, entry: ConditionalFormatEntry): ConditionalRule['range'][] =>
  entry.sqref.map((range) => ({
    sheet,
    r0: range.firstRow,
    c0: range.firstCol,
    r1: range.lastRow,
    c1: range.lastCol,
  }));

function dxfToApply(wb: WorkbookHandle, dxfId: number | undefined): Partial<CellFormat> {
  if (dxfId === undefined || !wb.capabilities.conditionalFormatDxf) return {};
  const dxf = wb.getDxf(dxfId);
  if (!dxf) return {};
  const apply: Partial<CellFormat> = {};
  if (dxf.font) Object.assign(apply, fontRecordToFormat(dxf.font));
  if (dxf.fill) Object.assign(apply, fillRecordToFormat(dxf.fill));
  if (dxf.border) Object.assign(apply, borderRecordToFormat(dxf.border));
  if (dxf.numFmt) {
    const numberFormat = formatCodeToNumFmt(dxf.numFmt.formatCode);
    if (numberFormat) apply.numFmt = numberFormat;
  }
  return apply;
}

function dxfToOverlay(wb: WorkbookHandle, dxfId: number | undefined): ConditionalCellOverlay {
  const apply = dxfToApply(wb, dxfId);
  const overlay: ConditionalCellOverlay = {};
  if (apply.fill) overlay.fill = apply.fill;
  if (apply.color) overlay.color = apply.color;
  if (apply.bold === true) overlay.bold = true;
  if (apply.italic === true) overlay.italic = true;
  if (apply.underline === true) overlay.underline = true;
  if (apply.strike === true) overlay.strike = true;
  return overlay;
}

function engineConditionalFormatToRules(
  wb: WorkbookHandle,
  sheet: number,
  entry: ConditionalFormatEntry,
): ConditionalRule[] {
  const ranges = rangesOf(sheet, entry);
  const apply = dxfToApply(wb, entry.dxfId);
  const common = {
    ...(entry.stopIfTrue ? { stopIfTrue: true } : {}),
    engineId: entry.id,
  };
  const out: ConditionalRule[] = [];
  for (const range of ranges) {
    if (entry.type === ENGINE_RULE_TYPE.cellIs) {
      const op = entry.op === undefined ? undefined : ENGINE_CELL_IS_OP[entry.op];
      if (!op || entry.formula1 === undefined) continue;
      out.push({
        ...common,
        kind: 'cell-value',
        range,
        op,
        a: maybeNumber(entry.formula1),
        ...(op === 'between' || op === 'not-between' ? { b: maybeNumber(entry.formula2) } : {}),
        apply,
      });
    } else if (entry.type === ENGINE_RULE_TYPE.expression) {
      if (!entry.formula1) continue;
      out.push({ ...common, kind: 'formula', range, formula: `=${entry.formula1}`, apply });
    } else if (
      entry.type === ENGINE_RULE_TYPE.containsText ||
      entry.type === ENGINE_RULE_TYPE.notContainsText ||
      entry.type === ENGINE_RULE_TYPE.beginsWith ||
      entry.type === ENGINE_RULE_TYPE.endsWith
    ) {
      if (entry.text === undefined) continue;
      out.push({
        ...common,
        kind: 'text-contains',
        range,
        text: entry.text,
        ...(entry.type === ENGINE_RULE_TYPE.notContainsText
          ? { mode: 'not-contains' as const }
          : entry.type === ENGINE_RULE_TYPE.beginsWith
            ? { mode: 'begins-with' as const }
            : entry.type === ENGINE_RULE_TYPE.endsWith
              ? { mode: 'ends-with' as const }
              : {}),
        apply,
      });
    } else if (entry.type === ENGINE_RULE_TYPE.containsBlanks) {
      out.push({ ...common, kind: 'blanks', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.notContainsBlanks) {
      out.push({ ...common, kind: 'non-blanks', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.containsErrors) {
      out.push({ ...common, kind: 'errors', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.notContainsErrors) {
      out.push({ ...common, kind: 'no-errors', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.duplicateValues) {
      out.push({ ...common, kind: 'duplicates', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.uniqueValues) {
      out.push({ ...common, kind: 'unique', range, apply });
    } else if (entry.type === ENGINE_RULE_TYPE.top10) {
      out.push({
        ...common,
        kind: 'top-bottom',
        range,
        mode: entry.bottom ? 'bottom' : 'top',
        n: entry.rank ?? 10,
        ...(entry.percent ? { percent: true } : {}),
        apply,
      });
    } else if (entry.type === ENGINE_RULE_TYPE.aboveAverage) {
      const stdDev = entry.stdDev;
      out.push({
        ...common,
        kind: 'average',
        range,
        mode:
          stdDev && stdDev >= 1
            ? entry.aboveAverage === false
              ? 'below-std-dev'
              : 'above-std-dev'
            : entry.equalAverage
              ? entry.aboveAverage === false
                ? 'equal-or-below'
                : 'equal-or-above'
              : entry.aboveAverage === false
                ? 'below'
                : 'above',
        ...(stdDev === 1 || stdDev === 2 || stdDev === 3 ? { stdDev } : {}),
        apply,
      });
    } else if (entry.type === ENGINE_RULE_TYPE.colorScale) {
      if (!entry.colorScale) continue;
      const colors = entry.colorScale.colors.map(rgba);
      if (colors.length !== 2 && colors.length !== 3) continue;
      const thresholds = entry.colorScale.thresholds.map(engineScalePoint);
      out.push({
        ...common,
        kind: 'color-scale',
        range,
        stops: colors as [string, string] | [string, string, string],
        ...(thresholds.length === colors.length
          ? {
              thresholds: thresholds as
                | [ConditionalScalePoint, ConditionalScalePoint]
                | [ConditionalScalePoint, ConditionalScalePoint, ConditionalScalePoint],
            }
          : {}),
      });
    } else if (entry.type === ENGINE_RULE_TYPE.dataBar) {
      if (!entry.dataBar) continue;
      out.push({
        ...common,
        kind: 'data-bar',
        range,
        color: rgba(entry.dataBar.fill),
        showValue: entry.dataBar.showValue !== false,
      });
    } else if (entry.type === ENGINE_RULE_TYPE.iconSet) {
      if (!entry.iconSet) continue;
      const icons = engineIconSet(entry.iconSet.name);
      if (!icons) continue;
      const slots = iconSetSlotCount(icons);
      const engineThresholds = entry.iconSet.thresholds.map(engineScalePoint);
      const thresholds =
        engineThresholds.length >= slots
          ? engineThresholds.slice(1, slots)
          : engineThresholds.slice(0, slots - 1);
      out.push({
        ...common,
        kind: 'icon-set',
        range,
        icons,
        showValue: entry.iconSet.showValue !== false,
        ...(entry.iconSet.reverse ? { reverseOrder: true } : {}),
        ...(thresholds.length > 0 ? { thresholds } : {}),
      });
    }
  }
  return out;
}

/**
 * Hydrate engine-authored conditional-format rules into the store
 * so command surfaces can reason about imported predicates without clearing or
 * duplicating them during session writeback. When the engine exposes the
 * differential-format table, non-visual rules also hydrate their `apply`
 * formatting from the referenced `dxfId`.
 */
export function hydrateConditionalRulesFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.conditionalFormatMutate) return;
  const importedRules = wb
    .getConditionalFormats(sheet)
    .flatMap((entry) => engineConditionalFormatToRules(wb, sheet, entry));
  store.setState((state) => {
    const rules = state.conditional.rules.filter(
      (rule) => rule.range.sheet !== sheet || !rule.engineId,
    );
    return { ...state, conditional: { rules: [...rules, ...importedRules] } };
  });
}

/**
 * Evaluate engine-side CF rules over `[(firstRow, firstCol), (lastRow, lastCol)]`
 * on `sheet` and lift the result into `ConditionalCellOverlay` shape so it can
 * be merged with the JS-side overlay map.
 *
 * ColorScale, DataBar, known IconSet ordinals, and dxf-backed font/fill
 * overlays lift cleanly when the corresponding engine capabilities are present.
 *
 * Returns an empty map when the engine doesn't expose `evaluateCfRange`.
 */
export function evaluateCfFromEngine(
  wb: WorkbookHandle,
  sheet: number,
  firstRow: number,
  firstCol: number,
  lastRow: number,
  lastCol: number,
  todaySerial: number = Number.NaN,
): Map<string, ConditionalCellOverlay> {
  const out = new Map<string, ConditionalCellOverlay>();
  if (!wb.capabilities.conditionalFormat) return out;
  const cells = wb.evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
  for (const cell of cells) {
    const key = addrKey({ sheet, row: cell.row, col: cell.col });
    const overlay: ConditionalCellOverlay = out.get(key) ?? {};
    // Iterate matches in priority order — engine returns them sorted by
    // priority, so later writes win for fields like `fill` (regular CF
    // semantics: highest priority match overrides).
    for (const m of cell.matches) {
      if (m.kind === KIND_COLOR_SCALE) {
        overlay.fill = rgba(m.color);
      } else if (m.kind === KIND_DATA_BAR) {
        overlay.bar = Math.max(0, Math.min(1, m.barLengthPct / 100));
        overlay.barAxis = Math.max(0, Math.min(1, m.barAxisPositionPct / 100));
        overlay.barDirection = m.barIsNegative ? 'left' : 'right';
        overlay.barColor = rgba(m.barFill);
        overlay.barGradient = m.barGradient;
      } else if (m.kind === KIND_ICON_SET) {
        const iconKind = engineIconSet(m.iconSetName);
        if (iconKind) {
          overlay.iconKind = iconKind;
          overlay.iconSlot = Math.max(0, Math.min(iconSetSlotCount(iconKind) - 1, m.iconIndex));
        }
      } else if (m.dxfIdEngaged) {
        Object.assign(overlay, dxfToOverlay(wb, m.dxfId));
      }
    }
    if (Object.keys(overlay).length > 0) out.set(key, overlay);
  }
  return out;
}
