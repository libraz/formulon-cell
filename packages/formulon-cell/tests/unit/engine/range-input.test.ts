import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import {
  assertPatternSheet,
  compileRangeInput,
  patternContainsAddr,
  patternOnSheet,
  resolveCellInput,
  resolveRangeInput,
} from '../../../src/engine/range-input.js';
import { parseA1RangeText } from '../../../src/engine/range-resolver.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const sheets = (...names: string[]): WorkbookHandle =>
  ({ sheetCount: names.length, sheetName: (i: number) => names[i] ?? '' }) as WorkbookHandle;

describe('parseA1RangeText', () => {
  it('parses cells, cell ranges, whole columns and whole rows', () => {
    expect(parseA1RangeText('B2')).toEqual({ sheetName: null, r0: 1, c0: 1, r1: 1, c1: 1 });
    expect(parseA1RangeText('$C$9:b2')).toEqual({ sheetName: null, r0: 1, c0: 1, r1: 8, c1: 2 });
    expect(parseA1RangeText('D:B')).toEqual({ sheetName: null, r0: 0, c0: 1, r1: MAX_ROW, c1: 3 });
    expect(parseA1RangeText('$11:2')).toEqual({
      sheetName: null,
      r0: 1,
      c0: 0,
      r1: 10,
      c1: MAX_COL,
    });
  });

  it('keeps plain and quoted sheet qualifiers', () => {
    expect(parseA1RangeText('Input!B2:B11')?.sheetName).toBe('Input');
    expect(parseA1RangeText("'Q1 ''24'!A:A")?.sheetName).toBe("Q1 '24");
  });

  it('rejects text that is not a range', () => {
    for (const text of ['', 'B', '5', 'B2:C3:D4', 'XFE1', 'A0', '1:1048577', '!B2', 'B2 C3']) {
      expect(parseA1RangeText(text), text).toBeNull();
    }
  });
});

describe('range inputs', () => {
  it('matches unqualified text on every sheet and qualified text on its sheet only', () => {
    const wb = sheets('Sheet1', 'Input');
    const anywhere = compileRangeInput('B2:B11', 'test');
    const pinned = compileRangeInput('input!B2:B11', 'test');
    expect(patternContainsAddr(anywhere, { sheet: 0, row: 1, col: 1 }, wb)).toBe(true);
    expect(patternContainsAddr(anywhere, { sheet: 1, row: 10, col: 1 }, wb)).toBe(true);
    expect(patternContainsAddr(anywhere, { sheet: 1, row: 11, col: 1 }, wb)).toBe(false);
    expect(patternContainsAddr(pinned, { sheet: 0, row: 1, col: 1 }, wb)).toBe(false);
    expect(patternContainsAddr(pinned, { sheet: 1, row: 1, col: 1 }, wb)).toBe(true);
  });

  it('resolves a qualified sheet name on each check, so it follows a reorder', () => {
    const pinned = compileRangeInput('Input!A1', 'test');
    expect(patternOnSheet(pinned, 1, sheets('Sheet1', 'Input'))).not.toBeNull();
    expect(patternOnSheet(pinned, 0, sheets('Input', 'Sheet1'))).not.toBeNull();
    expect(patternOnSheet(pinned, 0, sheets('Renamed', 'Sheet1'))).toBeNull();
  });

  it('passes numeric ranges through unchanged', () => {
    const range = { sheet: 1, r0: 0, c0: 0, r1: 2, c1: 2 };
    expect(compileRangeInput(range, 'test')).toEqual({ kind: 'range', range });
    expect(resolveRangeInput(range, sheets('A'), 0, 'test')).toBe(range);
  });

  it('throws on bad text and unknown sheet names, naming the option', () => {
    expect(() => compileRangeInput('B2:', 'policy.editable[0]')).toThrow(
      /policy\.editable\[0\]: "B2:" is not an A1 range/,
    );
    const pattern = compileRangeInput('Missing!B2', 'x');
    expect(() => assertPatternSheet(pattern, sheets('Sheet1'), 'x')).toThrow(
      /sheet "Missing" does not exist/,
    );
  });

  it('resolves actions on the active sheet unless the text names one', () => {
    const wb = sheets('Sheet1', 'Input');
    expect(resolveRangeInput('A1:B2', wb, 1, 't')).toEqual({
      sheet: 1,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    expect(resolveCellInput('Sheet1!C3', wb, 1, 't')).toEqual({ sheet: 0, row: 2, col: 2 });
    expect(resolveCellInput('C3', wb, 1, 't')).toEqual({ sheet: 1, row: 2, col: 2 });
    expect(() => resolveCellInput('C3:C4', wb, 0, 't')).toThrow(/not a single cell/);
  });
});
