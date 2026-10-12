import { describe, expect, it } from 'vitest';
import {
  excelSerialFromParts,
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
