import { iconSetSlotCount } from '../render/conditional.js';
import type { CellFormat } from '../store/store.js';
import type { ConditionalRule, ConditionalScalePoint } from '../store/types.js';
import { formatA1Cell } from './address.js';
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
  timePeriod: 15,
  duplicateValues: 16,
  uniqueValues: 17,
} as const;

/** `formulon::cf::TimePeriod` ordinals. */
export const TIME_PERIOD: Record<
  Extract<ConditionalRule, { kind: 'date-occurring' }>['period'],
  number
> = {
  today: 0,
  yesterday: 1,
  tomorrow: 2,
  last7: 3,
  'this-week': 4,
  'last-week': 5,
  'next-week': 6,
  'this-month': 7,
  'last-month': 8,
  'next-month': 9,
};

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

const ENGINE_DATA_BAR_DIRECTIONS = {
  context: 0,
  'left-to-right': 1,
  'right-to-left': 2,
} as const;

const ENGINE_DATA_BAR_AXIS_POSITIONS = {
  automatic: 0,
  middle: 1,
  none: 2,
} as const;

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

/** Encode a text-contains operand as an Excel formula string literal. */
const formulaTextLiteral = (text: string): string => `"${text.replaceAll('"', '""')}"`;

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
  /** Priority requested when the engine entry was written. */
  readonly priority?: number;
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
    dxf.font = { ...font, vertAlign: font.vertAlign ?? 0 };
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
  if (formatCode) dxf.numFmt = { numFmtId: 0, formatCode };
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
  const dxfId = wb.addDxf(dxf);
  return dxfId >= 0 ? { ...input, dxfId } : input;
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

function conditionalRuleBodyToEngineInput(rule: ConditionalRule): ConditionalFormatInput | null {
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
    case 'text-contains': {
      if (rule.caseSensitive === true) {
        const formula1 = caseSensitiveTextFormula(rule);
        // An empty local needle is intentionally a no-match predicate. Leave
        // it local rather than asking the engine to interpret an empty FIND.
        if (formula1 === null) return null;
        return { sqref, type: RULE_TYPE.expression, formula1 };
      }
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
    }
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
        sqref,
        type: RULE_TYPE.dataBar,
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
      const name = iconSetOrdinal(rule.icons);
      if (name < 0) return null;
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
        sqref,
        type: RULE_TYPE.iconSet,
        iconSet: {
          name,
          thresholds,
          reverse: rule.reverseOrder === true,
          showValue: rule.showValue !== false,
          percent: true,
          ...(rule.floor ? { floor: scalePointToCfValueObject(rule.floor) } : {}),
        },
      };
    }
    case 'date-occurring':
      return { sqref, type: RULE_TYPE.timePeriod, timePeriod: TIME_PERIOD[rule.period] };
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
  const trackedIndexes = new Set<number>();
  for (const value of opts.tracked.values()) {
    if (value.sheet !== sheet) continue;
    const index = trackedEntryIndex(currentFormats, value);
    if (index >= 0) trackedIndexes.add(index);
  }
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
  const original = entry as unknown as ConditionalFormatInput;
  const body = change.input ?? original;
  const next: ConditionalFormatInput = {
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
