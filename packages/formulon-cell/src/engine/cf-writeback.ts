import {
  type CellRect,
  CfDataBarAxisPosition,
  CfDataBarDirection,
  CfRuleType,
  type CfValueObjectInput,
  CfValueObjectType,
  type Color,
  ColorKind,
  ComparisonOperator,
  type ConditionalFormat,
  DatePeriod,
} from '@libraz/formulon';
import { sameJsonValue } from '../commands/prepared-change.js';
import { iconSetSlotCount } from '../render/conditional.js';
import type { CellFormat } from '../store/store.js';
import type { ConditionalRule, ConditionalScalePoint } from '../store/types.js';
import { formatA1Cell, rectOfRange } from './address.js';
import { iconSetToEngine } from './cf-icon-sets.js';
import {
  borderRecordFromFormat,
  cssColorToArgb,
  fillRecordFromFormat,
  fontRecordFromFormat,
  numFmtToFormatCode,
} from './format-writeback.js';
import type { ConditionalFormatInput, DxfRecord } from './types.js';
import { formulaTextLiteral } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** `DatePeriod` of each store date-occurring period. */
export const TIME_PERIOD: Record<
  Extract<ConditionalRule, { kind: 'date-occurring' }>['period'],
  DatePeriod
> = {
  today: DatePeriod.Today,
  yesterday: DatePeriod.Yesterday,
  tomorrow: DatePeriod.Tomorrow,
  last7: DatePeriod.Last7Days,
  'this-week': DatePeriod.ThisWeek,
  'last-week': DatePeriod.LastWeek,
  'next-week': DatePeriod.NextWeek,
  'this-month': DatePeriod.ThisMonth,
  'last-month': DatePeriod.LastMonth,
  'next-month': DatePeriod.NextMonth,
};

const VALUE_OBJECT_TYPE = {
  number: CfValueObjectType.Number,
  percent: CfValueObjectType.Percent,
  percentile: CfValueObjectType.Percentile,
  min: CfValueObjectType.Min,
  max: CfValueObjectType.Max,
} as const;

const ENGINE_DATA_BAR_DIRECTIONS = {
  context: CfDataBarDirection.Context,
  'left-to-right': CfDataBarDirection.LeftToRight,
  'right-to-left': CfDataBarDirection.RightToLeft,
} as const;

const ENGINE_DATA_BAR_AXIS_POSITIONS = {
  automatic: CfDataBarAxisPosition.Automatic,
  middle: CfDataBarAxisPosition.Middle,
  none: CfDataBarAxisPosition.None,
} as const;

/** `ComparisonOperator` of each store cell-value operator. */
const CELL_IS_OP: Record<string, ComparisonOperator> = {
  '<': ComparisonOperator.LessThan,
  '<=': ComparisonOperator.LessThanOrEqual,
  '=': ComparisonOperator.Equal,
  '<>': ComparisonOperator.NotEqual,
  '>=': ComparisonOperator.GreaterThanOrEqual,
  '>': ComparisonOperator.GreaterThan,
  between: ComparisonOperator.Between,
  'not-between': ComparisonOperator.NotBetween,
};

const rangesOf = (rule: ConditionalRule): CellRect[] => [rectOfRange(rule.range)];

/**
 * Build the case-sensitive equivalent of the native contains-text predicates.
 * The engine's containsText family has no case-sensitive flag, so these rules
 * are authored as expression rules instead. The top-left cell is deliberately
 * left relative; the CF evaluator shifts it across the authored range.
 */
function caseSensitiveTextFormula(
  rule: Extract<ConditionalRule, { kind: 'text-contains' }>,
): string | null {
  if (rule.text.length === 0) return null;
  const text = formulaTextLiteral(rule.text);
  const anchor = formatA1Cell(rule.range.r0, rule.range.c0);
  switch (rule.mode) {
    case 'not-contains':
      return `NOT(ISNUMBER(FIND(${text},${anchor})))`;
    case 'begins-with':
      return `EXACT(LEFT(${anchor},LEN(${text})),${text})`;
    case 'ends-with':
      return `EXACT(RIGHT(${anchor},LEN(${text})),${text})`;
    default:
      return `ISNUMBER(FIND(${text},${anchor}))`;
  }
}

const conditionalFormatKey = (sheet: number, input: ConditionalFormatInput): string =>
  JSON.stringify([sheet, input]);

/** Key of the next rule translating to `input`: identical rules get distinct
 *  keys by occurrence, so a duplicate stays a second engine rule. */
const nextOccurrenceKey = (
  counts: Map<string, number>,
  sheet: number,
  input: ConditionalFormatInput,
): string => {
  const base = conditionalFormatKey(sheet, input);
  const occurrence = counts.get(base) ?? 0;
  counts.set(base, occurrence + 1);
  return `${base}#${occurrence}`;
};

type ConditionalFormatEntry = ReturnType<WorkbookHandle['getConditionalFormats']>[number];
type DesiredConditionalRule = {
  readonly input: ConditionalFormatInput;
  readonly rule: ConditionalRule;
};

export interface SyncedConditionalRule {
  readonly sheet: number;
  index: number;
  readonly id?: string;
  /** Priority requested when the engine entry was written. */
  readonly priority?: number;
  readonly input: ConditionalFormatInput;
}

export type SyncedConditionalRuleMap = Map<string, SyncedConditionalRule>;

function sameRanges(a: unknown, b: readonly CellRect[]): boolean {
  if (!Array.isArray(a) || a.length !== b.length) return false;
  return a.every((left: CellRect, i) => {
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

/** Rule fields compared field by field; an absent input key matches `null`. */
const COMPARED_FIELDS = [
  'operator',
  'rank',
  'percent',
  'bottom',
  'aboveAverage',
  'equalAverage',
  'stdDev',
  'text',
  'timePeriod',
] as const;

/** Formula text without the leading `=` a read carries. */
/** `formula` without its leading `=`. */
export const stripLeadingEquals = (formula: string): string =>
  formula.startsWith('=') ? formula.slice(1) : formula;

const bareFormula = (formula: unknown): unknown =>
  typeof formula === 'string' ? stripLeadingEquals(formula) : (formula ?? null);

function entryMatchesInput(
  entry: ConditionalFormatEntry | undefined,
  input: ConditionalFormatInput,
): boolean {
  if (!entry) return false;
  const loose = input as unknown as Readonly<Record<string, unknown>>;
  return (
    entry.type === input.type &&
    sameRanges(input.ranges, entry.ranges) &&
    bareFormula(entry.formula1) === bareFormula(loose.formula1) &&
    bareFormula(entry.formula2) === bareFormula(loose.formula2) &&
    COMPARED_FIELDS.every((field) => (entry[field] ?? null) === (loose[field] ?? null)) &&
    sameJsonValue(entry.colorScale ?? null, loose.colorScale ?? null) &&
    sameJsonValue(entry.dataBar ?? null, loose.dataBar ?? null) &&
    sameJsonValue(entry.iconSet ?? null, loose.iconSet ?? null)
  );
}

function trackedEntryIndex(
  currentFormats: readonly ConditionalFormatEntry[],
  tracked: SyncedConditionalRule,
): number {
  const indexed = currentFormats[tracked.index];
  if (tracked.id) {
    if (indexed?.id === tracked.id) return tracked.index;
    return currentFormats.findIndex((entry) => entry.id === tracked.id);
  }
  // A missing id is only safe to resolve against an entry authored by this
  // session. Imported engine entries have ids even when the tracked snapshot
  // predates id support; matching their payload could delete user content.
  if (indexed?.id) return -1;
  return entryMatchesInput(indexed, tracked.input) ? tracked.index : -1;
}

export function getTrackedConditionalRuleIndexes(
  currentFormats: readonly ConditionalFormatEntry[],
  sheet: number,
  tracked: SyncedConditionalRuleMap,
): ReadonlySet<number> {
  const indexes = new Set<number>();
  for (const value of tracked.values()) {
    if (value.sheet !== sheet) continue;
    const index = trackedEntryIndex(currentFormats, value);
    if (index >= 0) indexes.add(index);
  }
  return indexes;
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

function cssColorToCfColor(color: string): Color | null {
  const argb = cssColorToArgb(color);
  if (argb === null) return null;
  return { kind: ColorKind.Rgb, argb, theme: 0, tint: 0, indexed: 0 };
}

function scalePointToCfValueObject(point: ConditionalScalePoint): CfValueObjectInput {
  const comparison = point.gte === undefined ? {} : { gte: point.gte };
  if (!('value' in point)) return { type: VALUE_OBJECT_TYPE[point.kind], ...comparison };
  return { type: VALUE_OBJECT_TYPE[point.kind], value: String(point.value), ...comparison };
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
  // formulon 0.12 stores the floor separately from the N-1 bucket
  // boundaries. The defaults are 33/67 for three icons and 20/40/60/80
  // for five icons.
  return Array.from({ length: Math.max(0, slots - 1) }, (_, index) => ({
    type: VALUE_OBJECT_TYPE.percent,
    value: String(Math.round(((index + 1) * 100) / slots)),
  }));
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
    apply.fontVertAlign !== undefined ||
    apply.color !== undefined
  );
}

export function dxfRecordFromApply(
  wb: WorkbookHandle,
  apply: Partial<CellFormat>,
): DxfRecord | null {
  const dxf: DxfRecord = {};
  if (hasFontDxfFields(apply)) {
    const font = fontRecordFromFormat(apply);
    // A differential font states only what the rule changes.
    dxf.font = {
      ...font,
      sizePt: apply.fontSize === undefined ? null : font.sizePt,
      vertAlign: font.vertAlign ?? 0,
    };
  }
  if (
    apply.fill !== undefined ||
    apply.fillPattern !== undefined ||
    apply.fillPatternColor !== undefined
  ) {
    dxf.fill = fillRecordFromFormat(apply);
  }
  if (apply.borders !== undefined) dxf.border = borderRecordFromFormat(apply);
  const formatCode = numFmtToFormatCode(apply.numFmt, wb);
  if (formatCode) dxf.numberFormat = formatCode;
  return Object.keys(dxf).length > 0 ? dxf : null;
}

export function inputWithDxf(
  wb: WorkbookHandle,
  input: ConditionalFormatInput,
  rule: ConditionalRule,
): ConditionalFormatInput {
  if (!wb.capabilities.conditionalFormatDxf) return input;
  const dxf = dxfRecordFromApply(wb, ruleApply(rule));
  if (!dxf) return input;
  const dxfIndex = wb.addDxf(dxf);
  // Only a format rule takes a differential format.
  if (dxfIndex < 0 || !formatRuleTypes.has(input.type)) return input;
  return { ...input, dxfIndex } as ConditionalFormatInput;
}

/**
 * Translate a store conditional-format rule into the engine's
 * `addConditionalFormat` input so its predicate and range round-trip through
 * .xlsx. Returns `null` for rules the engine cannot author:
 *
 * - `formula` rules that carry a comparator-prefix predicate (`>10`) rather
 *   than a full `=`-expression.
 * - `icon-set` rules with fewer than N-1 thresholds for an N-icon set.
 *
 * This pure translator intentionally leaves `rule.apply` out of the returned
 * input. The sync path attaches a `dxfId` after allocating a differential
 * format record through the workbook handle.
 */
export function conditionalRuleToEngineInput(rule: ConditionalRule): ConditionalFormatInput | null {
  const input = conditionalRuleBodyToEngineInput(rule);
  return input && rule.stopIfTrue === true ? { ...input, stopIfTrue: true } : input;
}

/** Rule kinds that apply a differential format. */
const formatRuleTypes: ReadonlySet<number> = new Set([
  CfRuleType.Expression,
  CfRuleType.CellIs,
  CfRuleType.Top10,
  CfRuleType.AboveAverage,
  CfRuleType.ContainsText,
  CfRuleType.NotContainsText,
  CfRuleType.BeginsWith,
  CfRuleType.EndsWith,
  CfRuleType.ContainsBlanks,
  CfRuleType.NotContainsBlanks,
  CfRuleType.ContainsErrors,
  CfRuleType.NotContainsErrors,
  CfRuleType.TimePeriod,
  CfRuleType.DuplicateValues,
  CfRuleType.UniqueValues,
]);

function conditionalRuleBodyToEngineInput(rule: ConditionalRule): ConditionalFormatInput | null {
  const ranges = rangesOf(rule);
  switch (rule.kind) {
    case 'cell-value': {
      const operator = CELL_IS_OP[rule.op];
      if (operator === undefined) return null;
      return {
        ranges,
        type: CfRuleType.CellIs,
        operator,
        formula1: String(rule.a),
        ...((rule.op === 'between' || rule.op === 'not-between') && rule.b !== undefined
          ? { formula2: String(rule.b) }
          : {}),
      };
    }
    case 'formula': {
      // Only a full `=`-expression is a valid standalone CF expression.
      if (!rule.formula.startsWith('=')) return null;
      return { ranges, type: CfRuleType.Expression, formula1: rule.formula.slice(1) };
    }
    case 'text-contains': {
      if (rule.caseSensitive === true) {
        const formula1 = caseSensitiveTextFormula(rule);
        // An empty local needle is intentionally a no-match predicate. Leave
        // it local rather than asking the engine to interpret an empty FIND.
        if (formula1 === null) return null;
        return { ranges, type: CfRuleType.Expression, formula1 };
      }
      return {
        ranges,
        type:
          rule.mode === 'not-contains'
            ? CfRuleType.NotContainsText
            : rule.mode === 'begins-with'
              ? CfRuleType.BeginsWith
              : rule.mode === 'ends-with'
                ? CfRuleType.EndsWith
                : CfRuleType.ContainsText,
        text: rule.text,
      };
    }
    case 'duplicates':
      return { ranges, type: CfRuleType.DuplicateValues };
    case 'unique':
      return { ranges, type: CfRuleType.UniqueValues };
    case 'blanks':
      return { ranges, type: CfRuleType.ContainsBlanks };
    case 'non-blanks':
      return { ranges, type: CfRuleType.NotContainsBlanks };
    case 'errors':
      return { ranges, type: CfRuleType.ContainsErrors };
    case 'no-errors':
      return { ranges, type: CfRuleType.NotContainsErrors };
    case 'top-bottom':
      return {
        ranges,
        type: CfRuleType.Top10,
        rank: rule.n,
        percent: rule.percent === true,
        bottom: rule.mode === 'bottom',
      };
    case 'average':
      return {
        ranges,
        type: CfRuleType.AboveAverage,
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
        ranges,
        type: CfRuleType.ColorScale,
        colorScale: {
          thresholds,
          colors: colors as Color[],
        },
      };
    }
    case 'data-bar': {
      const fill = cssColorToCfColor(rule.color);
      if (!fill) return null;
      const negativeFill =
        rule.negativeColor === undefined ? undefined : cssColorToCfColor(rule.negativeColor);
      const border =
        rule.borderColor === undefined ? undefined : cssColorToCfColor(rule.borderColor);
      const negativeBorder =
        rule.negativeBorderColor === undefined
          ? undefined
          : cssColorToCfColor(rule.negativeBorderColor);
      const axisColor =
        rule.axisColor === undefined ? undefined : cssColorToCfColor(rule.axisColor);
      if (
        (rule.negativeColor !== undefined && negativeFill === null) ||
        (rule.borderColor !== undefined && border === null) ||
        (rule.negativeBorderColor !== undefined && negativeBorder === null) ||
        (rule.axisColor !== undefined && axisColor === null)
      ) {
        return null;
      }
      const safeNegativeFill = negativeFill ?? undefined;
      const safeBorder = border ?? undefined;
      const safeNegativeBorder = negativeBorder ?? undefined;
      const safeAxisColor = axisColor ?? undefined;
      const direction =
        rule.direction === undefined ? undefined : ENGINE_DATA_BAR_DIRECTIONS[rule.direction];
      const axisPosition =
        rule.axisPosition === undefined
          ? undefined
          : ENGINE_DATA_BAR_AXIS_POSITIONS[rule.axisPosition];
      return {
        ranges,
        type: CfRuleType.DataBar,
        dataBar: {
          min: rule.min ? scalePointToCfValueObject(rule.min) : { type: VALUE_OBJECT_TYPE.min },
          max: rule.max ? scalePointToCfValueObject(rule.max) : { type: VALUE_OBJECT_TYPE.max },
          fill,
          showValue: rule.showValue !== false,
          ...(rule.gradient !== undefined ? { gradient: rule.gradient } : {}),
          ...(direction !== undefined ? { direction } : {}),
          ...(axisPosition !== undefined ? { axisPosition } : {}),
          ...(safeNegativeFill !== undefined ? { negativeFill: safeNegativeFill } : {}),
          ...(safeBorder !== undefined ? { border: safeBorder } : {}),
          ...(safeNegativeBorder !== undefined ? { negativeBorder: safeNegativeBorder } : {}),
          ...(safeAxisColor !== undefined ? { axisColor: safeAxisColor } : {}),
        },
      };
    }
    case 'icon-set': {
      const iconSet = iconSetToEngine(rule.icons, rule.customIcons);
      if (!iconSet) return null;
      const slots = iconSetSlotCount(rule.icons);
      // The engine takes exactly N-1 thresholds; a shorter list renders with
      // fewer buckets locally and has no file representation.
      if (rule.thresholds && rule.thresholds.length > 0 && rule.thresholds.length < slots - 1) {
        return null;
      }
      const thresholds =
        rule.thresholds && rule.thresholds.length > 0
          ? rule.thresholds.slice(0, slots - 1).map(scalePointToCfValueObject)
          : defaultIconThresholds(slots);
      return {
        ranges,
        type: CfRuleType.IconSet,
        iconSet: {
          name: iconSet.name,
          thresholds,
          reverse: rule.reverseOrder === true,
          showValue: rule.showValue !== false,
          percent: true,
          ...(rule.floor ? { floor: scalePointToCfValueObject(rule.floor) } : {}),
          ...(iconSet.icons ? { icons: iconSet.icons } : {}),
        },
      };
    }
    case 'date-occurring':
      return { ranges, type: CfRuleType.TimePeriod, timePeriod: TIME_PERIOD[rule.period] };
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
  const occurrences = new Map<string, number>();
  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    if (rule.engineId !== undefined) continue;
    const input = conditionalRuleToEngineInput(rule);
    if (!input) {
      skipped += 1;
      continue;
    }
    const key = nextOccurrenceKey(occurrences, sheet, input);
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
  const occurrences = new Map<string, number>();

  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    if (rule.engineId !== undefined) continue;
    const input = conditionalRuleToEngineInput(rule);
    if (!input) {
      skipped += 1;
      continue;
    }
    desired.set(nextOccurrenceKey(occurrences, sheet, input), { input, rule });
  }

  // Session rules rank after every rule this mount did not write, in store
  // order, so reordering the store reorders the engine priorities.
  let currentFormats = wb.getConditionalFormats(sheet);
  const trackedIndexes = getTrackedConditionalRuleIndexes(currentFormats, sheet, opts.tracked);
  let basePriority = 0;
  currentFormats.forEach((entry, index) => {
    if (!trackedIndexes.has(index)) basePriority = Math.max(basePriority, entry.priority);
  });
  // A rule keeps its written priority while the order stays ascending, so
  // deleting one rule does not rewrite every rule after it.
  const desiredPriority = new Map<string, number>();
  let previousPriority = basePriority;
  for (const key of desired.keys()) {
    const kept = opts.tracked.get(key)?.priority;
    previousPriority = kept !== undefined && kept > previousPriority ? kept : previousPriority + 1;
    desiredPriority.set(key, previousPriority);
  }

  const stale = [...opts.tracked.entries()]
    .filter(
      ([key, value]) =>
        value.sheet === sheet &&
        (!desired.has(key) ||
          (value.priority !== undefined && value.priority !== desiredPriority.get(key))),
    )
    .sort((a, b) => b[1].index - a[1].index);

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
    const priority = desiredPriority.get(key);
    const writeInput = { ...inputWithDxf(wb, input, rule), ...(priority ? { priority } : {}) };
    const addedIndex = wb.addConditionalFormat(sheet, writeInput);
    if (addedIndex < 0) {
      skipped += 1;
      continue;
    }
    const afterFormats = wb.getConditionalFormats(sheet);
    const index = Math.min(addedIndex, Math.max(0, afterFormats.length - 1));
    const id = afterFormats[index]?.id;
    const readBack = afterFormats[index];
    written += 1;
    opts.tracked.set(key, {
      sheet,
      index,
      ...(id ? { id } : {}),
      ...(priority ? { priority } : {}),
      // The engine fills defaults (gte, min/max lengths, icon floors, etc.)
      // while adding a rule. Keep that normalized snapshot for the no-id
      // fallback; stable non-empty ids remain authoritative on later edits.
      input: readBack ? (readBack as unknown as ConditionalFormatInput) : input,
    });
  }

  return { written, skipped, removed };
}

/**
 * Replace a loaded rule (one hydrated from the file, carrying an `engineId`)
 * by `next` in the engine. The rule keeps its id and priority. Returns false
 * when `next` has no engine form or the engine refuses the rewrite.
 */
export function rewriteLoadedConditionalRule(
  wb: WorkbookHandle,
  previous: ConditionalRule,
  next: ConditionalRule,
): boolean {
  if (!previous.engineId) return false;
  const input = conditionalRuleToEngineInput(next);
  if (!input) return false;
  return rewriteEngineConditionalFormat(wb, previous.range.sheet, previous.engineId, {
    input: inputWithDxf(wb, input, next),
  });
}

/**
 * Replace the engine rule `id` on `sheet` by removing it and adding it back,
 * the only way to change a loaded rule. The id and priority are kept unless
 * `change` overrides them; `change.input` replaces the rule body wholesale.
 * Restores the original entry and returns false when the re-add fails.
 */
export function rewriteEngineConditionalFormat(
  wb: WorkbookHandle,
  sheet: number,
  id: string,
  change: { priority?: number; stopIfTrue?: boolean; input?: ConditionalFormatInput },
): boolean {
  if (!wb.capabilities.conditionalFormatMutate) return false;
  const entries = wb.getConditionalFormats(sheet);
  const index = entries.findIndex((entry) => entry.id === id);
  const entry = entries[index];
  if (!entry) return false;
  const original: ConditionalFormat = entry;
  const body: ConditionalFormatInput | ConditionalFormat = change.input ?? original;
  const next: ConditionalFormatInput | ConditionalFormat = {
    ...body,
    id,
    priority: change.priority ?? entry.priority,
    stopIfTrue: change.stopIfTrue ?? entry.stopIfTrue,
  };
  if (!wb.removeConditionalFormatAt(sheet, index)) return false;
  if (wb.addConditionalFormat(sheet, next) >= 0) return true;
  wb.addConditionalFormat(sheet, original);
  return false;
}
