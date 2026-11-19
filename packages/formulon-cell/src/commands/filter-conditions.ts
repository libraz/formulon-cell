import { addrKey } from '../engine/address.js';
import { createProfileTextSemantics } from '../engine/profile-text-semantics.js';
import type { Range } from '../engine/types.js';
import { excelSerialToDateParts } from '../format/date-serial.js';
import type { FilterComparisonOp, FilterConditionSpec, State } from '../store/store.js';
import { parseLocalizedNumericInput } from './coerce-input.js';
import { dateInPeriod, localToday } from './date-periods.js';
import { filterValueKey } from './filter-items.js';

export type ConditionFilterOp = FilterConditionSpec['op'];

export type ConditionFilterOptions = FilterConditionSpec;

export const numericConditionOperator = (op: ConditionFilterOp): boolean =>
  op === 'equals' ||
  op === 'notEquals' ||
  op === 'greaterThan' ||
  op === 'greaterThanOrEqual' ||
  op === 'lessThan' ||
  op === 'lessThanOrEqual' ||
  op === 'between' ||
  op === 'top10' ||
  op === 'bottom10';

const COMPARISON_OPS: ReadonlySet<string> = new Set<FilterComparisonOp>([
  'equals',
  'notEquals',
  'contains',
  'notContains',
  'beginsWith',
  'notBeginsWith',
  'endsWith',
  'notEndsWith',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
]);

/** Whether `op` is a one-operand comparison, which a two-condition custom filter can combine. */
export const isFilterComparisonOp = (op: string): op is FilterComparisonOp =>
  COMPARISON_OPS.has(op);

/** Operators judged against the other numbers of their column. */
const rankedConditionOperator = (op: ConditionFilterOp): boolean =>
  op === 'top10' || op === 'bottom10' || op === 'aboveAverage' || op === 'belowAverage';

/** Store numeric criteria in the workbook's canonical grammar. Keeping the
 * raw text for invalid input preserves the existing no-match behavior while
 * making a valid `1,5` criterion survive a later locale switch as `1.5`. */
export const canonicalNumericCriterion = (raw: string, locale: string): string => {
  const parsed = parseLocalizedNumericInput(raw, locale);
  return parsed === null ? raw : String(parsed);
};

export function cellNumber(value: unknown): number | null {
  if (!value || typeof value !== 'object') return null;
  const cv = value as { kind: string; value?: unknown };
  if (cv.kind !== 'number' || typeof cv.value !== 'number' || !Number.isFinite(cv.value)) {
    return null;
  }
  return cv.value;
}

/** The number a ranked condition compares against: the cut-off value of a
 *  top/bottom set (ties kept, percent counts rounded down, at least one) or
 *  the column average. Null for other operators or a column without numbers. */
export function rankedBound(
  state: State,
  range: Range,
  col: number,
  condition: ConditionFilterOptions,
): number | null {
  if (!rankedConditionOperator(condition.op)) return null;
  const numbers: number[] = [];
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const n = cellNumber(state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col }))?.value);
    if (n !== null) numbers.push(n);
  }
  if (numbers.length === 0) return null;
  if (condition.op === 'aboveAverage' || condition.op === 'belowAverage') {
    return numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
  }
  const amount = Number(condition.value.trim());
  if (!(amount > 0)) return null;
  const raw = condition.unit === 'percent' ? (numbers.length * amount) / 100 : amount;
  const count = Math.max(1, Math.min(Math.floor(raw), numbers.length));
  const sorted = numbers.sort((a, b) => a - b);
  return condition.op === 'top10'
    ? (sorted[sorted.length - count] ?? null)
    : (sorted[count - 1] ?? null);
}

/** Whether a serial falls in a `datePeriod` condition's period, judged against today; null for
 *  any other condition. */
export function periodMatcher(
  condition: ConditionFilterOptions | undefined,
  date1904: boolean,
): ((serial: number) => boolean) | null {
  const period = condition?.op === 'datePeriod' ? condition.period : undefined;
  if (!period) return null;
  const today = localToday();
  return (serial) => {
    const parts = excelSerialToDateParts(serial, date1904);
    return parts !== null && dateInPeriod(period, parts, today);
  };
}

/** A column's condition, its second comparison included. */
export function criteriaConditionMatches(
  cell: { value: unknown; formula: string | null } | undefined,
  condition: ConditionFilterOptions,
  locale: string,
  bound: number | null,
  inPeriod: ((serial: number) => boolean) | null,
): boolean {
  const first = conditionMatches(cell, condition, locale, bound, inPeriod);
  const { second } = condition;
  if (!second) return first;
  const other = conditionMatches(cell, { op: second.op, value: second.value }, locale);
  return condition.join === 'or' ? first || other : first && other;
}

function conditionMatches(
  cell: { value: unknown; formula: string | null } | undefined,
  condition: ConditionFilterOptions,
  locale = 'en-US',
  bound: number | null = null,
  inPeriod: ((serial: number) => boolean) | null = null,
): boolean {
  const needle = condition.value.trim();
  const text = filterValueKey(cell?.value);
  const semantics = createProfileTextSemantics(locale);
  switch (condition.op) {
    case 'notEquals':
      return !semantics.equals(text, needle);
    case 'contains':
      return text.toLocaleLowerCase(locale).includes(needle.toLocaleLowerCase(locale));
    case 'notContains':
      return !text.toLocaleLowerCase(locale).includes(needle.toLocaleLowerCase(locale));
    case 'beginsWith':
      return text.toLocaleLowerCase(locale).startsWith(needle.toLocaleLowerCase(locale));
    case 'notBeginsWith':
      return !text.toLocaleLowerCase(locale).startsWith(needle.toLocaleLowerCase(locale));
    case 'endsWith':
      return text.toLocaleLowerCase(locale).endsWith(needle.toLocaleLowerCase(locale));
    case 'notEndsWith':
      return !text.toLocaleLowerCase(locale).endsWith(needle.toLocaleLowerCase(locale));
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual': {
      const target = Number(needle);
      const value = cellNumber(cell?.value);
      if (!Number.isFinite(target) || value == null) return false;
      if (condition.op === 'greaterThan') return value > target;
      if (condition.op === 'greaterThanOrEqual') return value >= target;
      if (condition.op === 'lessThan') return value < target;
      return value <= target;
    }
    case 'between': {
      const low = Number(needle);
      const high = Number((condition.value2 ?? '').trim());
      const value = cellNumber(cell?.value);
      if (!Number.isFinite(low) || !Number.isFinite(high) || needle === '' || value == null) {
        return false;
      }
      return value >= low && value <= high;
    }
    case 'datePeriod': {
      const value = cellNumber(cell?.value);
      return value !== null && inPeriod?.(value) === true;
    }
    case 'top10':
    case 'bottom10':
    case 'aboveAverage':
    case 'belowAverage': {
      const value = cellNumber(cell?.value);
      if (bound === null || value == null) return false;
      if (condition.op === 'top10') return value >= bound;
      if (condition.op === 'bottom10') return value <= bound;
      return condition.op === 'aboveAverage' ? value > bound : value < bound;
    }
    default:
      return semantics.equals(text, needle);
  }
}
