import { DynamicFilterType } from '@libraz/formulon';
import type { DatePeriodKey } from '../i18n/strings/data-objects/types.js';
import type { NumFmt } from '../store/types/cell-format.js';

/** A relative date period with the dynamic filter type that stores it and the type's
 *  `ST_DynamicFilterType` spelling, in the desktop's Date Filters order. */
export const DATE_PERIODS: readonly (readonly [DatePeriodKey, DynamicFilterType, string])[] = [
  ['tomorrow', DynamicFilterType.Tomorrow, 'tomorrow'],
  ['today', DynamicFilterType.Today, 'today'],
  ['yesterday', DynamicFilterType.Yesterday, 'yesterday'],
  ['nextWeek', DynamicFilterType.NextWeek, 'nextWeek'],
  ['thisWeek', DynamicFilterType.ThisWeek, 'thisWeek'],
  ['lastWeek', DynamicFilterType.LastWeek, 'lastWeek'],
  ['nextMonth', DynamicFilterType.NextMonth, 'nextMonth'],
  ['thisMonth', DynamicFilterType.ThisMonth, 'thisMonth'],
  ['lastMonth', DynamicFilterType.LastMonth, 'lastMonth'],
  ['nextQuarter', DynamicFilterType.NextQuarter, 'nextQuarter'],
  ['thisQuarter', DynamicFilterType.ThisQuarter, 'thisQuarter'],
  ['lastQuarter', DynamicFilterType.LastQuarter, 'lastQuarter'],
  ['nextYear', DynamicFilterType.NextYear, 'nextYear'],
  ['thisYear', DynamicFilterType.ThisYear, 'thisYear'],
  ['lastYear', DynamicFilterType.LastYear, 'lastYear'],
  ['yearToDate', DynamicFilterType.YearToDate, 'yearToDate'],
  ['quarter1', DynamicFilterType.Q1, 'Q1'],
  ['quarter2', DynamicFilterType.Q2, 'Q2'],
  ['quarter3', DynamicFilterType.Q3, 'Q3'],
  ['quarter4', DynamicFilterType.Q4, 'Q4'],
  ['january', DynamicFilterType.M1, 'M1'],
  ['february', DynamicFilterType.M2, 'M2'],
  ['march', DynamicFilterType.M3, 'M3'],
  ['april', DynamicFilterType.M4, 'M4'],
  ['may', DynamicFilterType.M5, 'M5'],
  ['june', DynamicFilterType.M6, 'M6'],
  ['july', DynamicFilterType.M7, 'M7'],
  ['august', DynamicFilterType.M8, 'M8'],
  ['september', DynamicFilterType.M9, 'M9'],
  ['october', DynamicFilterType.M10, 'M10'],
  ['november', DynamicFilterType.M11, 'M11'],
  ['december', DynamicFilterType.M12, 'M12'],
];

/** Where the Date Filters list separates its periods: days, weeks, months, quarters, years,
 *  year to date, then the quarters and months of All Dates in the Period. */
export const DATE_PERIOD_GROUP_STARTS: ReadonlySet<number> = new Set([0, 3, 6, 9, 12, 15, 16, 20]);

const BY_KEY = new Map(DATE_PERIODS.map(([key, type, name]) => [key, { type, name }]));
const BY_TYPE = new Map(DATE_PERIODS.map(([key, type]) => [type, key]));

/** The dynamic filter type that stores `key`. */
export const datePeriodDynamicType = (key: DatePeriodKey): DynamicFilterType =>
  BY_KEY.get(key)?.type ?? DynamicFilterType.Null;

/** The `ST_DynamicFilterType` spelling of `key`. */
export const datePeriodXmlType = (key: DatePeriodKey): string => BY_KEY.get(key)?.name ?? '';

/** The period a dynamic filter type stores; null for a non-date type. */
export const datePeriodOfDynamicType = (type: DynamicFilterType): DatePeriodKey | null =>
  BY_TYPE.get(type) ?? null;

export interface CivilDate {
  year: number;
  month: number;
  day: number;
}

const MS_PER_DAY = 86_400_000;
const dayNumber = (d: CivilDate): number => Date.UTC(d.year, d.month - 1, d.day) / MS_PER_DAY;
const quarterIndex = (d: CivilDate): number => d.year * 4 + Math.floor((d.month - 1) / 3);
const monthIndex = (d: CivilDate): number => d.year * 12 + d.month - 1;

/** Whether `date` falls in `period` relative to `today`. Weeks run Sunday to Saturday; the
 *  quarter and month periods of All Dates in the Period match that quarter or month of any year. */
export function dateInPeriod(period: DatePeriodKey, date: CivilDate, today: CivilDate): boolean {
  const n = dayNumber(date);
  const t = dayNumber(today);
  const weekStart = t - new Date(t * MS_PER_DAY).getUTCDay();
  const week = (offset: number): boolean =>
    n >= weekStart + offset * 7 && n < weekStart + (offset + 1) * 7;
  switch (period) {
    case 'tomorrow':
      return n === t + 1;
    case 'today':
      return n === t;
    case 'yesterday':
      return n === t - 1;
    case 'nextWeek':
      return week(1);
    case 'thisWeek':
      return week(0);
    case 'lastWeek':
      return week(-1);
    case 'nextMonth':
      return monthIndex(date) === monthIndex(today) + 1;
    case 'thisMonth':
      return monthIndex(date) === monthIndex(today);
    case 'lastMonth':
      return monthIndex(date) === monthIndex(today) - 1;
    case 'nextQuarter':
      return quarterIndex(date) === quarterIndex(today) + 1;
    case 'thisQuarter':
      return quarterIndex(date) === quarterIndex(today);
    case 'lastQuarter':
      return quarterIndex(date) === quarterIndex(today) - 1;
    case 'nextYear':
      return date.year === today.year + 1;
    case 'thisYear':
      return date.year === today.year;
    case 'lastYear':
      return date.year === today.year - 1;
    case 'yearToDate':
      return date.year === today.year && n <= t;
    case 'quarter1':
    case 'quarter2':
    case 'quarter3':
    case 'quarter4':
      return Math.floor((date.month - 1) / 3) + 1 === Number(period.slice(-1));
    default:
      return date.month === MONTHS.indexOf(period) + 1;
  }
}

const MONTHS: readonly DatePeriodKey[] = DATE_PERIODS.slice(20).map(([key]) => key);

/** Today in the local calendar. */
export const localToday = (now: Date = new Date()): CivilDate => ({
  year: now.getFullYear(),
  month: now.getMonth() + 1,
  day: now.getDate(),
});

/** Whether a number format shows a calendar date. */
export const numFmtShowsDate = (fmt: NumFmt | undefined): boolean => {
  if (!fmt) return false;
  if (fmt.kind === 'date' || fmt.kind === 'datetime') return true;
  if (fmt.kind !== 'custom') return false;
  const bare = fmt.pattern.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[yd]/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare));
};
