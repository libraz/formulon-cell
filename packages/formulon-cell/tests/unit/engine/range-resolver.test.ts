import { describe, expect, it } from 'vitest';
import { parseRangeRef, resolveRangeRef } from '../../../src/engine/range-resolver.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const fakeWb = (cells: Record<string, string | number>): WorkbookHandle => {
  return {
    sheetCount: 2,
    sheetName: (idx: number) => (idx === 0 ? 'Sheet1' : 'Sheet2'),
    getValue: ({ sheet, row, col }: { sheet: number; row: number; col: number }) => {
      const key = `${sheet}:${row}:${col}`;
      const v = cells[key];
      if (v === undefined) return { kind: 'blank' };
      if (typeof v === 'number') return { kind: 'number', value: v };
      return { kind: 'text', value: v };
    },
  } as unknown as WorkbookHandle;
};

describe('parseRangeRef', () => {
  it('parses A1', () => {
    expect(parseRangeRef('A1')).toEqual({ sheetName: null, r0: 0, c0: 0, r1: 0, c1: 0 });
  });
  it('parses absolute range A1:B5', () => {
    expect(parseRangeRef('$A$1:$B$5')).toEqual({ sheetName: null, r0: 0, c0: 0, r1: 4, c1: 1 });
  });
  it('parses sheet-qualified range', () => {
    expect(parseRangeRef('Sheet2!$C$3:$D$4')).toEqual({
      sheetName: 'Sheet2',
      r0: 2,
      c0: 2,
      r1: 3,
      c1: 3,
    });
  });
  it('parses quoted sheet name', () => {
    expect(parseRangeRef("'My Sheet'!A1:A3")).toEqual({
      sheetName: 'My Sheet',
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 0,
    });
  });
  it('strips a leading equals sign', () => {
    expect(parseRangeRef('=Sheet1!A1')).not.toBeNull();
  });
  it('returns null for non-ref strings', () => {
    expect(parseRangeRef('"Yes,No"')).toBeNull();
    expect(parseRangeRef('not-a-ref')).toBeNull();
    expect(parseRangeRef('')).toBeNull();
  });
});

describe('resolveRangeRef', () => {
  it('reads non-blank values from the resolved sheet', () => {
    const wb = fakeWb({
      '0:0:0': 'apple',
      '0:1:0': 'banana',
      '0:2:0': 'cherry',
    });
    expect(resolveRangeRef(wb, '$A$1:$A$3', 0)).toEqual(['apple', 'banana', 'cherry']);
  });

  it('skips blanks and de-duplicates in source order', () => {
    const wb = fakeWb({
      '0:0:0': 'a',
      '0:1:0': 'b',
      '0:3:0': 'a', // duplicate
      '0:4:0': 'c',
    });
    expect(resolveRangeRef(wb, 'A1:A5', 0)).toEqual(['a', 'b', 'c']);
  });

  it('routes through sheetName lookup', () => {
    const wb = fakeWb({
      '1:0:0': 'red',
      '1:1:0': 'blue',
    });
    expect(resolveRangeRef(wb, 'Sheet2!A1:A2', 0)).toEqual(['red', 'blue']);
  });

  it('returns [] for unknown sheet', () => {
    const wb = fakeWb({});
    expect(resolveRangeRef(wb, 'Missing!A1:A3', 0)).toEqual([]);
  });

  it('returns [] for unparseable refs', () => {
    const wb = fakeWb({});
    expect(resolveRangeRef(wb, 'gibberish', 0)).toEqual([]);
  });
});
