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

const ENGINE_DATA_BAR_DIRECTIONS = ['context', 'left-to-right', 'right-to-left'] as const;
type DataBarDirection = (typeof ENGINE_DATA_BAR_DIRECTIONS)[number];

const ENGINE_DATA_BAR_AXIS_POSITIONS = ['automatic', 'middle', 'none'] as const;
type DataBarAxisPosition = (typeof ENGINE_DATA_BAR_AXIS_POSITIONS)[number];

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

const sameColor = (
  a: { r: number; g: number; b: number; a: number },
  b: { r: number; g: number; b: number; a: number },
): boolean => a.r === b.r && a.g === b.g && a.b === b.b && a.a === b.a;

const isOpaqueBlack = (c: { r: number; g: number; b: number; a: number }): boolean =>
  c.r === 0 && c.g === 0 && c.b === 0 && c.a === 255;

const engineIconSet = (ordinal: number): ConditionalIconSet | null =>
  Number.isInteger(ordinal) && ordinal >= 0 && ordinal < ENGINE_ICON_SETS.length
    ? (ENGINE_ICON_SETS[ordinal] ?? null)
    : null;

const engineDataBarDirection = (ordinal: number): DataBarDirection | undefined =>
  Number.isInteger(ordinal) && ordinal >= 0 && ordinal < ENGINE_DATA_BAR_DIRECTIONS.length
    ? ENGINE_DATA_BAR_DIRECTIONS[ordinal]
    : undefined;

const engineDataBarAxisPosition = (ordinal: number): DataBarAxisPosition | undefined =>
  Number.isInteger(ordinal) && ordinal >= 0 && ordinal < ENGINE_DATA_BAR_AXIS_POSITIONS.length
    ? ENGINE_DATA_BAR_AXIS_POSITIONS[ordinal]
    : undefined;

function engineScalePoint(
  valueObject: NonNullable<ConditionalFormatEntry['colorScale']>['thresholds'][number] | undefined,
): ConditionalScalePoint {
  if (!valueObject) return { kind: 'min' };
  const comparison = valueObject.gte === undefined ? {} : { gte: valueObject.gte };
  if (valueObject.type === VALUE_OBJECT_TYPE.min) return { kind: 'min', ...comparison };
  if (valueObject.type === VALUE_OBJECT_TYPE.max) return { kind: 'max', ...comparison };
  const raw = Number(valueObject.value ?? '0');
  const value = Number.isFinite(raw) ? raw : 0;
  if (valueObject.type === VALUE_OBJECT_TYPE.percent)
    return { kind: 'percent', value, ...comparison };
  if (valueObject.type === VALUE_OBJECT_TYPE.percentile) {
    return { kind: 'percentile', value, ...comparison };
  }
  return { kind: 'number', value, ...comparison };
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

type EngineRange = { firstRow: number; firstCol: number; lastRow: number; lastCol: number };

const validEngineRange = (range: unknown): EngineRange | null => {
  if (range === null || typeof range !== 'object') return null;
  const candidate = range as {
    firstRow?: unknown;
    firstCol?: unknown;
    lastRow?: unknown;
    lastCol?: unknown;
  };
  if (
    !Number.isInteger(candidate.firstRow) ||
    !Number.isInteger(candidate.firstCol) ||
    !Number.isInteger(candidate.lastRow) ||
    !Number.isInteger(candidate.lastCol)
  ) {
    return null;
  }
  const firstRow = candidate.firstRow as number;
  const firstCol = candidate.firstCol as number;
  const lastRow = candidate.lastRow as number;
  const lastCol = candidate.lastCol as number;
  return firstRow <= lastRow && firstCol <= lastCol
    ? { firstRow, firstCol, lastRow, lastCol }
    : null;
};

const rangeContainsCell = (range: unknown, row: number, col: number): boolean => {
  const candidate = validEngineRange(range);
  return (
    candidate !== null &&
    row >= candidate.firstRow &&
    row <= candidate.lastRow &&
    col >= candidate.firstCol &&
    col <= candidate.lastCol
  );
};

const entryContainsCell = (entry: ConditionalFormatEntry, row: number, col: number): boolean =>
  Array.isArray(entry.sqref) && entry.sqref.some((range) => rangeContainsCell(range, row, col));

/** Resolve one engine data-bar payload for a match. A priority is preferred,
 * but the containing range is a safe fallback for older match projections
 * that omit or renumber priority. Ambiguous or malformed metadata is ignored
 * so it cannot override authoritative per-cell match colours. */
function dataBarMetadataForCell(
  entries: readonly ConditionalFormatEntry[],
  priority: number,
  row: number,
  col: number,
): ConditionalFormatEntry['dataBar'] | undefined {
  const candidates = entries.filter(
    (entry) => entry.type === ENGINE_RULE_TYPE.dataBar && entry.dataBar,
  );
  const containing = candidates.filter((entry) => entryContainsCell(entry, row, col));
  if (containing.length === 0) return undefined;
  const byPriority = Number.isFinite(priority)
    ? containing.filter((entry) => entry.priority === priority)
    : [];
  if (byPriority.length === 1) return byPriority[0]?.dataBar;
  if (byPriority.length > 1) return undefined;
  return containing.length === 1 ? containing[0]?.dataBar : undefined;
}

function dataBarHasMixedPopulation(
  wb: WorkbookHandle,
  sheet: number,
  entry: ConditionalFormatEntry | undefined,
): boolean | undefined {
  if (!entry || !Array.isArray(entry.sqref)) return undefined;
  const ranges = entry.sqref.map(validEngineRange);
  if (ranges.some((range) => range === null)) return undefined;
  const validRanges = ranges as EngineRange[];
  if (typeof wb.physicalCells === 'function') {
    let hasNegative = false;
    let hasPositive = false;
    for (const cell of wb.physicalCells(sheet)) {
      if (
        cell.value.kind !== 'number' ||
        !Number.isFinite(cell.value.value) ||
        !validRanges.some((range) => rangeContainsCell(range, cell.addr.row, cell.addr.col))
      )
        continue;
      hasNegative ||= cell.value.value < 0;
      hasPositive ||= cell.value.value > 0;
      if (hasNegative && hasPositive) return true;
    }
    return false;
  }
  return undefined;
}

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
      const direction = engineDataBarDirection(entry.dataBar.direction);
      const min = entry.dataBar.min
        ? engineScalePoint(entry.dataBar.min)
        : { kind: 'min' as const };
      const max = entry.dataBar.max
        ? engineScalePoint(entry.dataBar.max)
        : { kind: 'max' as const };
      const fillColor = entry.dataBar.fill;
      const negativeFill = entry.dataBar.negativeFill;
      const border = entry.dataBar.border;
      const negativeBorder = entry.dataBar.negativeBorder;
      const axisPosition = engineDataBarAxisPosition(entry.dataBar.axisPosition ?? 0);
      out.push({
        ...common,
        kind: 'data-bar',
        range,
        color: rgba(fillColor),
        showValue: entry.dataBar.showValue !== false,
        ...(min.kind !== 'min' || min.gte === false ? { min } : {}),
        ...(max.kind !== 'max' || max.gte === false ? { max } : {}),
        ...(axisPosition && axisPosition !== 'automatic' ? { axisPosition } : {}),
        ...(negativeFill && !sameColor(negativeFill, fillColor)
          ? { negativeColor: rgba(negativeFill) }
          : {}),
        ...(border ? { borderColor: rgba(border) } : {}),
        ...(negativeBorder ? { negativeBorderColor: rgba(negativeBorder) } : {}),
        ...(entry.dataBar.axisColor && !isOpaqueBlack(entry.dataBar.axisColor)
          ? { axisColor: rgba(entry.dataBar.axisColor) }
          : {}),
        ...(direction ? { direction } : {}),
        ...(entry.dataBar.gradient !== undefined ? { gradient: entry.dataBar.gradient } : {}),
      });
    } else if (entry.type === ENGINE_RULE_TYPE.iconSet) {
      if (!entry.iconSet) continue;
      const icons = engineIconSet(entry.iconSet.name);
      if (!icons) continue;
      const slots = iconSetSlotCount(icons);
      const engineThresholds = entry.iconSet.thresholds.map(engineScalePoint);
      // formulon 0.12 stores the floor separately; `thresholds` contains
      // only the N-1 boundaries between an N-icon set's buckets.
      const thresholds = engineThresholds.slice(0, slots - 1);
      const floor = entry.iconSet.floor ? engineScalePoint(entry.iconSet.floor) : undefined;
      out.push({
        ...common,
        kind: 'icon-set',
        range,
        icons,
        showValue: entry.iconSet.showValue !== false,
        ...(entry.iconSet.reverse ? { reverseOrder: true } : {}),
        ...(thresholds.length > 0 ? { thresholds } : {}),
        ...(floor ? { floor } : {}),
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
      (rule) => rule.range.sheet !== sheet || rule.engineId === undefined,
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
  const sheetRtl =
    typeof wb.getSheetView === 'function' && wb.getSheetView(sheet)?.rightToLeft === true;
  const metadataEntries =
    typeof wb.getConditionalFormats === 'function' ? wb.getConditionalFormats(sheet) : [];
  const negativePopulationCache = new Map<ConditionalFormatEntry, boolean | undefined>();
  const cells = wb.evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
  for (const cell of cells) {
    const key = addrKey({ sheet, row: cell.row, col: cell.col });
    const overlay: ConditionalCellOverlay = out.get(key) ?? {};
    let dataBarDefined = false;
    // Iterate matches in priority order — engine returns them sorted by
    // priority. The C API uses first-defined semantics: a lower-priority
    // match may fill a visual property the higher-priority match omitted, but
    // it cannot overwrite one that was already set.
    for (const m of cell.matches) {
      if (m.kind === KIND_COLOR_SCALE) {
        if (overlay.fill === undefined) overlay.fill = rgba(m.color);
      } else if (m.kind === KIND_DATA_BAR) {
        // A data bar is one atomic visual property. In particular, an
        // omitted border on the first bar is authoritative and must clear a
        // lower-priority bar's border rather than inherit it.
        if (dataBarDefined) continue;
        dataBarDefined = true;
        // The engine reports length as a fraction of the axis's signed side;
        // the canvas overlay stores a fraction of the full cell width.
        const rawLength = Math.max(0, Math.min(1, m.barLengthPct / 100));
        const axis = Math.max(0, Math.min(1, m.barAxisPositionPct / 100));
        const negative = m.barIsNegative === true;
        const metadata = dataBarMetadataForCell(metadataEntries, m.priority, cell.row, cell.col);
        const axisPosition = metadata
          ? (engineDataBarAxisPosition(metadata.axisPosition ?? 0) ?? 'automatic')
          : undefined;
        const matchingEntry = metadata
          ? metadataEntries.find(
              (entry) => entry.type === ENGINE_RULE_TYPE.dataBar && entry.dataBar === metadata,
            )
          : undefined;
        const noneNegative = axisPosition === 'none' && negative;
        overlay.bar = rawLength * (noneNegative || !negative ? 1 - axis : axis);
        const direction =
          'barDirection' in m && typeof m.barDirection === 'number' ? m.barDirection : 0;
        const mirror = direction === 2 || (direction === 0 && sheetRtl);
        overlay.barAxis = mirror ? 1 - axis : axis;
        const baseDirection = noneNegative || !negative ? 'right' : 'left';
        overlay.barDirection = mirror
          ? baseDirection === 'left'
            ? 'right'
            : 'left'
          : baseDirection;
        overlay.barColor = rgba(m.barFill);
        overlay.barBorderColor = m.barBorderEngaged ? rgba(m.barBorder) : undefined;
        if (metadata) {
          overlay.barAxisColor = metadata.axisColor ? rgba(metadata.axisColor) : '#000000';
          if (axisPosition === 'none') {
            overlay.barAxisVisible = false;
          } else if (axisPosition === 'middle') {
            overlay.barAxisVisible = true;
          } else if (matchingEntry) {
            if (axis > 0 && axis < 1) {
              if (!negativePopulationCache.has(matchingEntry)) {
                negativePopulationCache.set(
                  matchingEntry,
                  dataBarHasMixedPopulation(wb, sheet, matchingEntry),
                );
              }
              const hasMixedPopulation = negativePopulationCache.get(matchingEntry);
              // Keep the local overlay's population-aware flag when the
              // engine cannot inspect the rule's full population.
              if (hasMixedPopulation !== undefined) {
                overlay.barAxisVisible = hasMixedPopulation;
              }
            } else {
              overlay.barAxisVisible = false;
            }
          }
        }
        overlay.barGradient = m.barGradient;
      } else if (m.kind === KIND_ICON_SET) {
        const iconKind = engineIconSet(m.iconSetName);
        if (iconKind && overlay.iconKind === undefined) {
          overlay.iconKind = iconKind;
          overlay.iconSlot = Math.max(0, Math.min(iconSetSlotCount(iconKind) - 1, m.iconIndex));
        }
      } else if (m.dxfIdEngaged) {
        const dxf = dxfToOverlay(wb, m.dxfId);
        if (overlay.fill === undefined && dxf.fill !== undefined) overlay.fill = dxf.fill;
        if (overlay.color === undefined && dxf.color !== undefined) overlay.color = dxf.color;
        if (overlay.bold === undefined && dxf.bold !== undefined) overlay.bold = dxf.bold;
        if (overlay.italic === undefined && dxf.italic !== undefined) overlay.italic = dxf.italic;
        if (overlay.underline === undefined && dxf.underline !== undefined) {
          overlay.underline = dxf.underline;
        }
        if (overlay.strike === undefined && dxf.strike !== undefined) overlay.strike = dxf.strike;
      }
    }
    if (Object.keys(overlay).length > 0) out.set(key, overlay);
  }
  return out;
}
