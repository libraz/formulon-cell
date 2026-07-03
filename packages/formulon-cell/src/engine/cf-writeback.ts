import { iconSetSlotCount } from '../render/conditional.js';
import type { CellFormat } from '../store/store.js';
import type { ConditionalRule } from '../store/types.js';
import {
  borderRecordFromFormat,
  cssColorToArgb,
  fillRecordFromFormat,
  fontRecordFromFormat,
  numFmtToFormatCode,
} from './format-writeback.js';
import type { ConditionalFormatInput, DxfRecord } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** `formulon::cf::RuleType` ordinals. */
const RULE_TYPE = {
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

const ENGINE_ICON_SETS = [
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
] as const;

/** `formulon::cf::CellIsOperator` ordinals. */
const CELL_IS_OP: Record<string, number> = {
  '<': 0,
  '<=': 1,
  '=': 2,
  '<>': 3,
  '>=': 4,
  '>': 5,
  between: 6,
  'not-between': 7,
};

const sqrefOf = (rule: ConditionalRule): ConditionalFormatInput['sqref'] => [
  {
    firstRow: rule.range.r0,
    firstCol: rule.range.c0,
    lastRow: rule.range.r1,
    lastCol: rule.range.c1,
  },
];

const conditionalFormatKey = (sheet: number, input: ConditionalFormatInput): string =>
  JSON.stringify([sheet, input]);

type ConditionalFormatEntry = ReturnType<WorkbookHandle['getConditionalFormats']>[number];
type CfColor = NonNullable<ConditionalFormatInput['colorScale']>['colors'][number];
type CfValueObjectInput = NonNullable<ConditionalFormatInput['colorScale']>['thresholds'][number];
type DesiredConditionalRule = {
  readonly input: ConditionalFormatInput;
  readonly rule: ConditionalRule;
};

export interface SyncedConditionalRule {
  readonly sheet: number;
  index: number;
  readonly id?: string;
  readonly input: ConditionalFormatInput;
}

export type SyncedConditionalRuleMap = Map<string, SyncedConditionalRule>;

function sameSqref(
  a: ConditionalFormatInput['sqref'],
  b: ConditionalFormatEntry['sqref'],
): boolean {
  if (a.length !== b.length) return false;
  return a.every((left, i) => {
    const right = b[i];
    return (
      right !== undefined &&
      left.firstRow === right.firstRow &&
      left.firstCol === right.firstCol &&
      left.lastRow === right.lastRow &&
      left.lastCol === right.lastCol
    );
  });
}

function entryMatchesInput(
  entry: ConditionalFormatEntry | undefined,
  input: ConditionalFormatInput,
): boolean {
  if (!entry) return false;
  return (
    entry.type === input.type &&
    sameSqref(input.sqref, entry.sqref) &&
    entry.formula1 === input.formula1 &&
    entry.formula2 === input.formula2 &&
    entry.op === input.op &&
    entry.rank === input.rank &&
    entry.percent === input.percent &&
    entry.bottom === input.bottom &&
    entry.aboveAverage === input.aboveAverage &&
    entry.equalAverage === input.equalAverage &&
    entry.stdDev === input.stdDev &&
    entry.text === input.text &&
    entry.timePeriod === input.timePeriod &&
    JSON.stringify(entry.colorScale ?? null) === JSON.stringify(input.colorScale ?? null) &&
    JSON.stringify(entry.dataBar ?? null) === JSON.stringify(input.dataBar ?? null) &&
    JSON.stringify(entry.iconSet ?? null) === JSON.stringify(input.iconSet ?? null)
  );
}

function trackedEntryIndex(
  currentFormats: readonly ConditionalFormatEntry[],
  tracked: SyncedConditionalRule,
): number {
  const indexed = currentFormats[tracked.index];
  if (tracked.id && indexed?.id === tracked.id && entryMatchesInput(indexed, tracked.input)) {
    return tracked.index;
  }
  if (tracked.id) {
    const byId = currentFormats.findIndex(
      (entry) => entry.id === tracked.id && entryMatchesInput(entry, tracked.input),
    );
    if (byId >= 0) return byId;
  }
  return entryMatchesInput(indexed, tracked.input) ? tracked.index : -1;
}

function decrementTrackedIndexesAfterRemoval(
  tracked: SyncedConditionalRuleMap,
  sheet: number,
  removedIndex: number,
): void {
  for (const value of tracked.values()) {
    if (value.sheet === sheet && value.index > removedIndex) value.index -= 1;
  }
}

function cssColorToCfColor(color: string): CfColor | null {
  const argb = cssColorToArgb(color);
  if (argb === null) return null;
  return {
    a: (argb >>> 24) & 0xff,
    r: (argb >>> 16) & 0xff,
    g: (argb >>> 8) & 0xff,
    b: argb & 0xff,
  };
}

function scalePointToCfValueObject(
  point: NonNullable<Extract<ConditionalRule, { kind: 'color-scale' }>['thresholds']>[number],
): CfValueObjectInput {
  if (!('value' in point)) return { type: VALUE_OBJECT_TYPE[point.kind] };
  return { type: VALUE_OBJECT_TYPE[point.kind], value: String(point.value) };
}

function defaultColorScaleThresholds(
  stops: Extract<ConditionalRule, { kind: 'color-scale' }>['stops'],
): CfValueObjectInput[] {
  return stops.length === 2
    ? [{ type: VALUE_OBJECT_TYPE.min }, { type: VALUE_OBJECT_TYPE.max }]
    : [
        { type: VALUE_OBJECT_TYPE.min },
        { type: VALUE_OBJECT_TYPE.percentile, value: '50' },
        { type: VALUE_OBJECT_TYPE.max },
      ];
}

function defaultIconThresholds(slots: number): CfValueObjectInput[] {
  return Array.from({ length: slots }, (_, index) => ({
    type: VALUE_OBJECT_TYPE.percent,
    value: String(Math.round((index * 100) / slots)),
  }));
}

function iconSetOrdinal(icons: Extract<ConditionalRule, { kind: 'icon-set' }>['icons']): number {
  return ENGINE_ICON_SETS.indexOf(icons);
}

function ruleApply(rule: ConditionalRule): Partial<CellFormat> {
  return 'apply' in rule ? rule.apply : {};
}

function hasFontDxfFields(apply: Partial<CellFormat>): boolean {
  return (
    apply.fontFamily !== undefined ||
    apply.fontSize !== undefined ||
    apply.bold !== undefined ||
    apply.italic !== undefined ||
    apply.strike !== undefined ||
    apply.underline !== undefined ||
    apply.color !== undefined
  );
}

function dxfRecordFromApply(apply: Partial<CellFormat>): DxfRecord | null {
  const dxf: DxfRecord = {};
  if (hasFontDxfFields(apply)) dxf.font = fontRecordFromFormat(apply);
  if (apply.fill !== undefined) dxf.fill = fillRecordFromFormat(apply);
  if (apply.borders !== undefined) dxf.border = borderRecordFromFormat(apply);
  const formatCode = numFmtToFormatCode(apply.numFmt);
  if (formatCode) dxf.numFmt = { numFmtId: 0, formatCode };
  return Object.keys(dxf).length > 0 ? dxf : null;
}

function inputWithDxf(
  wb: WorkbookHandle,
  input: ConditionalFormatInput,
  rule: ConditionalRule,
): ConditionalFormatInput {
  if (!wb.capabilities.conditionalFormatDxf) return input;
  const dxf = dxfRecordFromApply(ruleApply(rule));
  if (!dxf) return input;
  const dxfId = wb.addDxf(dxf);
  return dxfId >= 0 ? { ...input, dxfId } : input;
}

/**
 * Translate a store conditional-format rule into the engine's
 * `addConditionalFormat` input so its predicate and range round-trip through
 * .xlsx. Returns `null` for rules the engine cannot author:
 *
 * - `date-occurring` — the `timePeriod` ordinal set isn't stable in this API.
 * - `formula` rules that carry a comparator-prefix predicate (`>10`) rather
 *   than a full `=`-expression.
 *
 * This pure translator intentionally leaves `rule.apply` out of the returned
 * input. The sync path attaches a `dxfId` after allocating a differential
 * format record through the workbook handle.
 */
export function conditionalRuleToEngineInput(rule: ConditionalRule): ConditionalFormatInput | null {
  const sqref = sqrefOf(rule);
  switch (rule.kind) {
    case 'cell-value': {
      const op = CELL_IS_OP[rule.op];
      if (op === undefined) return null;
      const input: ConditionalFormatInput = {
        sqref,
        type: RULE_TYPE.cellIs,
        op,
        formula1: String(rule.a),
      };
      if ((rule.op === 'between' || rule.op === 'not-between') && rule.b !== undefined) {
        input.formula2 = String(rule.b);
      }
      return input;
    }
    case 'formula': {
      // Only a full `=`-expression is a valid standalone CF expression.
      if (!rule.formula.startsWith('=')) return null;
      return { sqref, type: RULE_TYPE.expression, formula1: rule.formula.slice(1) };
    }
    case 'text-contains':
      return {
        sqref,
        type:
          rule.mode === 'not-contains'
            ? RULE_TYPE.notContainsText
            : rule.mode === 'begins-with'
              ? RULE_TYPE.beginsWith
              : rule.mode === 'ends-with'
                ? RULE_TYPE.endsWith
                : RULE_TYPE.containsText,
        text: rule.text,
      };
    case 'duplicates':
      return { sqref, type: RULE_TYPE.duplicateValues };
    case 'unique':
      return { sqref, type: RULE_TYPE.uniqueValues };
    case 'blanks':
      return { sqref, type: RULE_TYPE.containsBlanks };
    case 'non-blanks':
      return { sqref, type: RULE_TYPE.notContainsBlanks };
    case 'errors':
      return { sqref, type: RULE_TYPE.containsErrors };
    case 'no-errors':
      return { sqref, type: RULE_TYPE.notContainsErrors };
    case 'top-bottom':
      return {
        sqref,
        type: RULE_TYPE.top10,
        rank: rule.n,
        percent: rule.percent === true,
        bottom: rule.mode === 'bottom',
      };
    case 'average':
      return {
        sqref,
        type: RULE_TYPE.aboveAverage,
        aboveAverage:
          rule.mode === 'above' || rule.mode === 'equal-or-above' || rule.mode === 'above-std-dev',
        equalAverage: rule.mode === 'equal-or-above' || rule.mode === 'equal-or-below',
        ...(rule.mode === 'above-std-dev' || rule.mode === 'below-std-dev'
          ? { stdDev: rule.stdDev ?? 1 }
          : {}),
      };
    case 'color-scale': {
      const colors = rule.stops.map(cssColorToCfColor);
      if (colors.some((color) => color === null)) return null;
      const thresholds = rule.thresholds
        ? rule.thresholds.map(scalePointToCfValueObject)
        : defaultColorScaleThresholds(rule.stops);
      if (thresholds.length !== colors.length) return null;
      return {
        sqref,
        type: RULE_TYPE.colorScale,
        colorScale: {
          thresholds,
          colors: colors as CfColor[],
        },
      };
    }
    case 'data-bar': {
      const fill = cssColorToCfColor(rule.color);
      if (!fill) return null;
      return {
        sqref,
        type: RULE_TYPE.dataBar,
        dataBar: {
          min: { type: VALUE_OBJECT_TYPE.min },
          max: { type: VALUE_OBJECT_TYPE.max },
          fill,
          showValue: rule.showValue !== false,
        },
      };
    }
    case 'icon-set': {
      const name = iconSetOrdinal(rule.icons);
      if (name < 0) return null;
      const slots = iconSetSlotCount(rule.icons);
      const thresholds = rule.thresholds
        ? [
            { type: VALUE_OBJECT_TYPE.percent, value: '0' },
            ...rule.thresholds.map(scalePointToCfValueObject),
          ]
        : defaultIconThresholds(slots);
      return {
        sqref,
        type: RULE_TYPE.iconSet,
        iconSet: {
          name,
          thresholds: thresholds.slice(0, slots),
          reverse: rule.reverseOrder === true,
          showValue: rule.showValue !== false,
          percent: true,
        },
      };
    }
    default:
      // date-occurring
      return null;
  }
}

/**
 * Additively write the engine-representable subset of `rules` onto `sheet` so
 * they round-trip through save. Never clears existing engine rules (the store
 * is not hydrated from engine CF on import, so a clear would drop rules authored
 * elsewhere). Returns the counts written and skipped. No-op — `{written:0,
 * skipped:0}` — when the engine can't author CF rules (stub / older builds).
 *
 * Applied differential formatting is persisted as `dxfId` when the engine
 * exposes the dxf table.
 */
export function syncConditionalRulesToEngine(
  wb: WorkbookHandle,
  rules: readonly ConditionalRule[],
  sheet: number,
  opts: { seenKeys?: Set<string> } = {},
): { written: number; skipped: number } {
  if (!wb.capabilities.conditionalFormatMutate) return { written: 0, skipped: 0 };
  let written = 0;
  let skipped = 0;
  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    if (rule.engineId) continue;
    const input = conditionalRuleToEngineInput(rule);
    if (!input) {
      skipped += 1;
      continue;
    }
    const key = conditionalFormatKey(sheet, input);
    if (opts.seenKeys?.has(key)) continue;
    const writeInput = inputWithDxf(wb, input, rule);
    if (wb.addConditionalFormat(sheet, writeInput) >= 0) written += 1;
    else {
      skipped += 1;
      continue;
    }
    opts.seenKeys?.add(key);
  }
  return { written, skipped };
}

/**
 * Reconcile the engine with the session-created, engine-representable subset
 * of store CF rules. Unlike {@link syncConditionalRulesToEngine}, this removes
 * only rules whose engine index was tracked after this mount added them; it
 * never clears imported/untracked CF blocks.
 */
export function syncTrackedConditionalRulesToEngine(
  wb: WorkbookHandle,
  rules: readonly ConditionalRule[],
  sheet: number,
  opts: { tracked: SyncedConditionalRuleMap },
): { written: number; skipped: number; removed: number } {
  if (!wb.capabilities.conditionalFormatMutate) return { written: 0, skipped: 0, removed: 0 };

  let written = 0;
  let skipped = 0;
  let removed = 0;
  const desired = new Map<string, DesiredConditionalRule>();

  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    if (rule.engineId) continue;
    const input = conditionalRuleToEngineInput(rule);
    if (!input) {
      skipped += 1;
      continue;
    }
    desired.set(conditionalFormatKey(sheet, input), { input, rule });
  }

  const stale = [...opts.tracked.entries()]
    .filter(([key, value]) => value.sheet === sheet && !desired.has(key))
    .sort((a, b) => b[1].index - a[1].index);

  let currentFormats = stale.length > 0 ? wb.getConditionalFormats(sheet) : [];
  for (const [key, value] of stale) {
    const removeIndex = trackedEntryIndex(currentFormats, value);
    if (removeIndex < 0) {
      skipped += 1;
      opts.tracked.delete(key);
      continue;
    }
    if (wb.removeConditionalFormatAt(sheet, removeIndex)) {
      removed += 1;
      opts.tracked.delete(key);
      decrementTrackedIndexesAfterRemoval(opts.tracked, sheet, removeIndex);
      currentFormats = currentFormats.filter((_, index) => index !== removeIndex);
    } else {
      skipped += 1;
    }
  }

  for (const [key, { input, rule }] of desired) {
    if (opts.tracked.has(key)) continue;
    const writeInput = inputWithDxf(wb, input, rule);
    const addedIndex = wb.addConditionalFormat(sheet, writeInput);
    if (addedIndex < 0) {
      skipped += 1;
      continue;
    }
    const afterFormats = wb.getConditionalFormats(sheet);
    const index = Math.min(addedIndex, Math.max(0, afterFormats.length - 1));
    const id = afterFormats[index]?.id;
    written += 1;
    opts.tracked.set(key, {
      sheet,
      index,
      ...(id ? { id } : {}),
      input,
    });
  }

  return { written, skipped, removed };
}
