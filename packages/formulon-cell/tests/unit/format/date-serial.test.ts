import { describe, expect, it } from 'vitest';
import {
  excelSerialFromDate,
  excelSerialFromParts,
  excelSerialToDate,
  excelSerialToDateParts,
  excelSerialToIsoDate,
} from '../../../src/format/date-serial.js';

describe('Excel 1900 date serial conversion', () => {
  it.each([
    [1, { year: 1900, month: 1, day: 1, fictitious: false }],
    [59, { year: 1900, month: 2, day: 28, fictitious: false }],
    [60, { year: 1900, month: 2, day: 29, fictitious: true }],
    [61, { year: 1900, month: 3, day: 1, fictitious: false }],
  ] as const)('maps serial %s to the Excel civil date', (serial, expected) => {
    expect(excelSerialToDateParts(serial)).toEqual(expected);
  });

  it('round-trips valid civil dates and keeps the fictitious day explicit', () => {
    expect(excelSerialFromParts(1900, 1, 1)).toBe(1);
    expect(excelSerialFromParts(1900, 2, 28)).toBe(59);
    expect(excelSerialFromParts(1900, 2, 29)).toBeNull();
    expect(excelSerialFromParts(1900, 2, 29, { allowFictitious: true })).toBe(60);
    expect(excelSerialFromParts(1900, 3, 1)).toBe(61);
  });

  it('does not expose serial 60 as a native ISO date', () => {
    expect(excelSerialToIsoDate(59)).toBe('1900-02-28');
    expect(excelSerialToIsoDate(60)).toBeNull();
    expect(excelSerialToIsoDate(61)).toBe('1900-03-01');
  });
});

describe('Excel 1904 date serial conversion', () => {
  it('counts from 1904-01-01 = 0 with no placeholder day', () => {
    expect(excelSerialToDateParts(0, true)).toEqual({
      year: 1904,
      month: 1,
      day: 1,
      fictitious: false,
    });
    expect(excelSerialToDateParts(1, true)?.day).toBe(2);
    expect(excelSerialFromParts(1904, 1, 1, { date1904: true })).toBe(0);
    expect(excelSerialFromParts(1904, 1, 2, { date1904: true })).toBe(1);
    expect(excelSerialToIsoDate(60, true)).toBe('1904-03-01');
  });

  it('offsets the 1900 serial of the same calendar date by 1462', () => {
    for (const [y, m, d] of [
      [1904, 1, 1],
      [2000, 2, 29],
      [2024, 1, 1],
    ] as const) {
      const s1900 = excelSerialFromParts(y, m, d) as number;
      expect(excelSerialFromParts(y, m, d, { date1904: true })).toBe(s1900 - 1462);
    }
    expect(excelSerialFromParts(2024, 1, 1, { date1904: true })).toBe(43830);
  });

  it('rejects dates before 1904-01-01 and the 1900 placeholder day', () => {
    expect(excelSerialFromParts(1903, 12, 31, { date1904: true })).toBeNull();
    expect(excelSerialFromParts(1900, 2, 29, { date1904: true, allowFictitious: true })).toBeNull();
    expect(excelSerialToDateParts(-1, true)).toBeNull();
    expect(excelSerialToDate(-1, true)).toBeNull();
    expect(excelSerialToIsoDate(-1, true)).toBeNull();
  });

  it('keeps the time fraction in both directions', () => {
    const date = excelSerialToDate(43830.75, true);
    expect(date?.toISOString()).toBe('2024-01-01T18:00:00.000Z');
    expect(excelSerialFromDate(date as Date, true)).toBe(43830.75);
  });

  it('leaves the 1900 system unchanged by default', () => {
    expect(excelSerialToDate(45292)?.toISOString()).toBe('2024-01-01T00:00:00.000Z');
    expect(excelSerialFromDate(new Date(Date.UTC(2024, 0, 1)))).toBe(45292);
    expect(excelSerialToDateParts(60)?.fictitious).toBe(true);
  });
});
