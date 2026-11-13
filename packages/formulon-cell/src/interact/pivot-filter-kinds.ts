import { PivotFilterKind } from '../engine/types.js';
import type { DatePeriodKey, FilterOperatorKey } from '../i18n/strings/data-objects/types.js';

/** The filter menu a pivot filter kind belongs to. */
export type PivotFilterCategory = 'label' | 'value' | 'date';

/** What a filter kind's operand is: text, a number, a date, a pair of either for a range kind,
 *  an item count for Top 10 by items, or nothing for a relative date period. */
export type PivotFilterOperand =
  | 'text'
  | 'textRange'
  | 'number'
  | 'numberRange'
  | 'count'
  | 'date'
  | 'dateRange'
  | 'none';

const LABEL_KINDS: ReadonlySet<PivotFilterKind> = new Set([
  PivotFilterKind.LabelContains,
  PivotFilterKind.LabelBeginsWith,
  PivotFilterKind.LabelEqual,
  PivotFilterKind.LabelNotEqual,
  PivotFilterKind.LabelNotBeginsWith,
  PivotFilterKind.LabelEndsWith,
  PivotFilterKind.LabelNotEndsWith,
  PivotFilterKind.LabelNotContains,
  PivotFilterKind.LabelGreaterThan,
  PivotFilterKind.LabelGreaterThanOrEqual,
  PivotFilterKind.LabelLessThan,
  PivotFilterKind.LabelLessThanOrEqual,
  PivotFilterKind.LabelBetween,
  PivotFilterKind.LabelNotBetween,
]);

const VALUE_KINDS: ReadonlySet<PivotFilterKind> = new Set([
  PivotFilterKind.ValueTopCount,
  PivotFilterKind.ValueGreaterThan,
  PivotFilterKind.ValueBetween,
  PivotFilterKind.ValueEqual,
  PivotFilterKind.ValueNotEqual,
  PivotFilterKind.ValueGreaterThanOrEqual,
  PivotFilterKind.ValueLessThan,
  PivotFilterKind.ValueLessThanOrEqual,
  PivotFilterKind.ValueNotBetween,
  PivotFilterKind.ValueBottomCount,
  PivotFilterKind.ValueTopPercent,
  PivotFilterKind.ValueBottomPercent,
  PivotFilterKind.ValueTopSum,
  PivotFilterKind.ValueBottomSum,
]);

/** The menu `kind` belongs to; every kind that is neither a label nor a value filter is a date
 *  filter, the relative periods included. */
export const pivotFilterCategory = (kind: PivotFilterKind): PivotFilterCategory =>
  LABEL_KINDS.has(kind) ? 'label' : VALUE_KINDS.has(kind) ? 'value' : 'date';

/** The operand `kind` takes. */
export const pivotFilterOperand = (kind: PivotFilterKind): PivotFilterOperand => {
  switch (kind) {
    case PivotFilterKind.LabelBetween:
    case PivotFilterKind.LabelNotBetween:
      return 'textRange';
    case PivotFilterKind.ValueBetween:
    case PivotFilterKind.ValueNotBetween:
      return 'numberRange';
    case PivotFilterKind.DateBetween:
    case PivotFilterKind.DateNotBetween:
      return 'dateRange';
    case PivotFilterKind.ValueTopCount:
    case PivotFilterKind.ValueBottomCount:
      return 'count';
    case PivotFilterKind.DateEqual:
    case PivotFilterKind.DateNotEqual:
    case PivotFilterKind.DateOlderThan:
    case PivotFilterKind.DateOlderThanOrEqual:
    case PivotFilterKind.DateNewerThan:
    case PivotFilterKind.DateNewerThanOrEqual:
      return 'date';
  }
  const category = pivotFilterCategory(kind);
  return category === 'label' ? 'text' : category === 'value' ? 'number' : 'none';
};

/** Comparisons in the order the desktop's filter dialogs list them, with their operator label. */
export const PIVOT_FILTER_COMPARISONS: Record<
  PivotFilterCategory,
  readonly (readonly [PivotFilterKind, FilterOperatorKey])[]
> = {
  label: [
    [PivotFilterKind.LabelEqual, 'equal'],
    [PivotFilterKind.LabelNotEqual, 'notEqual'],
    [PivotFilterKind.LabelGreaterThan, 'greaterThan'],
    [PivotFilterKind.LabelGreaterThanOrEqual, 'greaterThanOrEqual'],
    [PivotFilterKind.LabelLessThan, 'lessThan'],
    [PivotFilterKind.LabelLessThanOrEqual, 'lessThanOrEqual'],
    [PivotFilterKind.LabelBeginsWith, 'beginsWith'],
    [PivotFilterKind.LabelNotBeginsWith, 'notBeginsWith'],
    [PivotFilterKind.LabelEndsWith, 'endsWith'],
    [PivotFilterKind.LabelNotEndsWith, 'notEndsWith'],
    [PivotFilterKind.LabelContains, 'contains'],
    [PivotFilterKind.LabelNotContains, 'notContains'],
    [PivotFilterKind.LabelBetween, 'between'],
    [PivotFilterKind.LabelNotBetween, 'notBetween'],
  ],
  value: [
    [PivotFilterKind.ValueEqual, 'equal'],
    [PivotFilterKind.ValueNotEqual, 'notEqual'],
    [PivotFilterKind.ValueGreaterThan, 'greaterThan'],
    [PivotFilterKind.ValueGreaterThanOrEqual, 'greaterThanOrEqual'],
    [PivotFilterKind.ValueLessThan, 'lessThan'],
    [PivotFilterKind.ValueLessThanOrEqual, 'lessThanOrEqual'],
    [PivotFilterKind.ValueBetween, 'between'],
    [PivotFilterKind.ValueNotBetween, 'notBetween'],
  ],
  date: [
    [PivotFilterKind.DateEqual, 'equal'],
    [PivotFilterKind.DateNotEqual, 'notEqual'],
    [PivotFilterKind.DateOlderThan, 'before'],
    [PivotFilterKind.DateOlderThanOrEqual, 'beforeOrEqual'],
    [PivotFilterKind.DateNewerThan, 'after'],
    [PivotFilterKind.DateNewerThanOrEqual, 'afterOrEqual'],
    [PivotFilterKind.DateBetween, 'between'],
    [PivotFilterKind.DateNotBetween, 'notBetween'],
  ],
};

/** The relative date periods in the desktop's Date Filters order, with their shared label. */
export const PIVOT_FILTER_PERIODS: readonly (readonly [PivotFilterKind, DatePeriodKey])[] = [
  [PivotFilterKind.Tomorrow, 'tomorrow'],
  [PivotFilterKind.Today, 'today'],
  [PivotFilterKind.Yesterday, 'yesterday'],
  [PivotFilterKind.NextWeek, 'nextWeek'],
  [PivotFilterKind.ThisWeek, 'thisWeek'],
  [PivotFilterKind.LastWeek, 'lastWeek'],
  [PivotFilterKind.NextMonth, 'nextMonth'],
  [PivotFilterKind.ThisMonth, 'thisMonth'],
  [PivotFilterKind.LastMonth, 'lastMonth'],
  [PivotFilterKind.NextQuarter, 'nextQuarter'],
  [PivotFilterKind.ThisQuarter, 'thisQuarter'],
  [PivotFilterKind.LastQuarter, 'lastQuarter'],
  [PivotFilterKind.NextYear, 'nextYear'],
  [PivotFilterKind.ThisYear, 'thisYear'],
  [PivotFilterKind.LastYear, 'lastYear'],
  [PivotFilterKind.YearToDate, 'yearToDate'],
  [PivotFilterKind.Q1, 'quarter1'],
  [PivotFilterKind.Q2, 'quarter2'],
  [PivotFilterKind.Q3, 'quarter3'],
  [PivotFilterKind.Q4, 'quarter4'],
  [PivotFilterKind.M1, 'january'],
  [PivotFilterKind.M2, 'february'],
  [PivotFilterKind.M3, 'march'],
  [PivotFilterKind.M4, 'april'],
  [PivotFilterKind.M5, 'may'],
  [PivotFilterKind.M6, 'june'],
  [PivotFilterKind.M7, 'july'],
  [PivotFilterKind.M8, 'august'],
  [PivotFilterKind.M9, 'september'],
  [PivotFilterKind.M10, 'october'],
  [PivotFilterKind.M11, 'november'],
  [PivotFilterKind.M12, 'december'],
];

const ALL_DATES_START = PIVOT_FILTER_PERIODS.findIndex(([kind]) => kind === PivotFilterKind.Q1);

/** The relative periods Date Filters lists itself, through Year to Date. */
export const PIVOT_LISTED_PERIODS = PIVOT_FILTER_PERIODS.slice(0, ALL_DATES_START);

/** The quarters and months All Dates in the Period picks from. */
export const PIVOT_ALL_DATES_PERIODS = PIVOT_FILTER_PERIODS.slice(ALL_DATES_START);

/** Whether `kind` is a quarter or month of All Dates in the Period. */
export const isPivotAllDatesPeriod = (kind: PivotFilterKind): boolean =>
  PIVOT_ALL_DATES_PERIODS.some(([period]) => period === kind);

/** What a Top 10 filter shows: the top or bottom, by items, percent or sum. */
export type PivotTop10Direction = 'top' | 'bottom';
export type PivotTop10Basis = 'items' | 'percent' | 'sum';

const TOP10_KINDS: Record<PivotTop10Direction, Record<PivotTop10Basis, PivotFilterKind>> = {
  top: {
    items: PivotFilterKind.ValueTopCount,
    percent: PivotFilterKind.ValueTopPercent,
    sum: PivotFilterKind.ValueTopSum,
  },
  bottom: {
    items: PivotFilterKind.ValueBottomCount,
    percent: PivotFilterKind.ValueBottomPercent,
    sum: PivotFilterKind.ValueBottomSum,
  },
};

/** The Top 10 kind for a direction and basis. */
export const pivotTop10Kind = (direction: PivotTop10Direction, basis: PivotTop10Basis) =>
  TOP10_KINDS[direction][basis];

/** The direction and basis of a Top 10 kind; null for any other kind. */
export const pivotTop10Parts = (
  kind: PivotFilterKind,
): { direction: PivotTop10Direction; basis: PivotTop10Basis } | null => {
  for (const direction of ['top', 'bottom'] as const) {
    for (const basis of ['items', 'percent', 'sum'] as const) {
      if (TOP10_KINDS[direction][basis] === kind) return { direction, basis };
    }
  }
  return null;
};
