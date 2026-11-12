import { describe, expect, it } from 'vitest';
import {
  type CivilDate,
  DATE_PERIODS,
  dateInPeriod,
  numFmtShowsDate,
} from '../../../src/commands/date-periods.js';

// Thursday 2026-11-12; its week runs Sunday 2026-11-08 to Saturday 2026-11-14.
const TODAY: CivilDate = { year: 2026, month: 11, day: 12 };
const d = (year: number, month: number, day: number): CivilDate => ({ year, month, day });

type Period = (typeof DATE_PERIODS)[number][0];

/** For each period, a date inside it and the nearest date just outside it. */
const CASES: Record<Period, [CivilDate, CivilDate]> = {
  tomorrow: [d(2026, 11, 13), d(2026, 11, 14)],
  today: [d(2026, 11, 12), d(2026, 11, 13)],
  yesterday: [d(2026, 11, 11), d(2026, 11, 10)],
  nextWeek: [d(2026, 11, 15), d(2026, 11, 22)],
  thisWeek: [d(2026, 11, 8), d(2026, 11, 7)],
  lastWeek: [d(2026, 11, 7), d(2026, 11, 8)],
  nextMonth: [d(2026, 12, 31), d(2027, 1, 1)],
  thisMonth: [d(2026, 11, 30), d(2026, 12, 1)],
  lastMonth: [d(2026, 10, 1), d(2026, 9, 30)],
  nextQuarter: [d(2027, 1, 1), d(2027, 4, 1)],
  thisQuarter: [d(2026, 10, 1), d(2026, 9, 30)],
  lastQuarter: [d(2026, 7, 1), d(2026, 6, 30)],
  nextYear: [d(2027, 6, 1), d(2028, 1, 1)],
  thisYear: [d(2026, 1, 1), d(2025, 12, 31)],
  lastYear: [d(2025, 12, 31), d(2026, 1, 1)],
  yearToDate: [d(2026, 11, 12), d(2026, 11, 13)],
  quarter1: [d(2019, 3, 31), d(2019, 4, 1)],
  quarter2: [d(2030, 4, 1), d(2030, 7, 1)],
  quarter3: [d(2026, 9, 30), d(2026, 10, 1)],
  quarter4: [d(2001, 12, 31), d(2001, 9, 30)],
  january: [d(2020, 1, 31), d(2020, 2, 1)],
  february: [d(2024, 2, 29), d(2024, 3, 1)],
  march: [d(2026, 3, 1), d(2026, 2, 28)],
  april: [d(2026, 4, 30), d(2026, 5, 1)],
  may: [d(2026, 5, 15), d(2026, 6, 15)],
  june: [d(2026, 6, 1), d(2026, 7, 1)],
  july: [d(2026, 7, 1), d(2026, 8, 1)],
  august: [d(2026, 8, 1), d(2026, 9, 1)],
  september: [d(2026, 9, 1), d(2026, 10, 1)],
  october: [d(2026, 10, 1), d(2026, 11, 1)],
  november: [d(1999, 11, 1), d(1999, 12, 1)],
  december: [d(2026, 12, 1), d(2026, 1, 1)],
};

describe('relative date periods', () => {
  it('matches a date inside each period and not one just outside', () => {
    for (const [period] of DATE_PERIODS) {
      const [inside, outside] = CASES[period];
      expect(dateInPeriod(period, inside, TODAY), `${period} inside`).toBe(true);
      expect(dateInPeriod(period, outside, TODAY), `${period} outside`).toBe(false);
    }
  });

  it('crosses a year boundary for next and last month', () => {
    const newYearsEve = d(2026, 12, 31);
    expect(dateInPeriod('nextMonth', d(2027, 1, 5), newYearsEve)).toBe(true);
    expect(dateInPeriod('lastMonth', d(2026, 12, 5), d(2027, 1, 2))).toBe(true);
    expect(dateInPeriod('nextQuarter', d(2027, 2, 1), newYearsEve)).toBe(true);
  });

  it('recognizes date number formats', () => {
    expect(numFmtShowsDate({ kind: 'date', pattern: 'm/d/yyyy' })).toBe(true);
    expect(numFmtShowsDate({ kind: 'datetime', pattern: 'm/d/yyyy h:mm' })).toBe(true);
    expect(numFmtShowsDate({ kind: 'custom', pattern: 'yyyy"年"m"月"' })).toBe(true);
    expect(numFmtShowsDate({ kind: 'custom', pattern: 'mmm' })).toBe(true);
    expect(numFmtShowsDate({ kind: 'custom', pattern: 'h:mm' })).toBe(false);
    expect(numFmtShowsDate({ kind: 'custom', pattern: '0.00E+00' })).toBe(false);
    expect(numFmtShowsDate({ kind: 'custom', pattern: '[Red]0' })).toBe(false);
    expect(numFmtShowsDate({ kind: 'time', pattern: 'h:mm' })).toBe(false);
    expect(numFmtShowsDate(undefined)).toBe(false);
  });
});
