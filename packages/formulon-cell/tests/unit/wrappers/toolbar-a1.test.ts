import { describe, expect, it } from 'vitest';
import {
  formatA1Range,
  formatSheetAbsoluteRange,
  parseA1Range,
} from '../../../src/wrappers/toolbar-a1.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

describe('formatSheetAbsoluteRange', () => {
  it('renders the sheet-qualified absolute reference dialogs show', () => {
    expect(formatSheetAbsoluteRange('Sheet1', range(0, 0, 2, 1))).toBe('Sheet1!$A$1:$B$3');
  });

  it('collapses a single cell the way formatA1Range does', () => {
    expect(formatSheetAbsoluteRange('Sheet1', range(4, 2, 4, 2))).toBe('Sheet1!$C$5');
    expect(formatA1Range(range(4, 2, 4, 2))).toBe('C5');
  });

  it('quotes sheet names that are not plain identifiers', () => {
    expect(formatSheetAbsoluteRange('My Sheet', range(0, 0, 0, 0))).toBe("'My Sheet'!$A$1");
    expect(formatSheetAbsoluteRange("Bob's", range(0, 0, 0, 0))).toBe("'Bob''s'!$A$1");
  });

  it('drops the prefix when the workbook reports no sheet name', () => {
    expect(formatSheetAbsoluteRange('', range(0, 0, 2, 1))).toBe('$A$1:$B$3');
  });

  it('round-trips through parseA1Range so dialog validation still accepts it', () => {
    const selection = range(1, 1, 4, 3);
    expect(parseA1Range(formatSheetAbsoluteRange('Sheet1', selection), 0, 'Sheet1')).toEqual(
      selection,
    );
    expect(parseA1Range(formatSheetAbsoluteRange('My Sheet', selection), 0, 'My Sheet')).toEqual(
      selection,
    );
    // Cross-sheet references stay rejected.
    expect(parseA1Range(formatSheetAbsoluteRange('Sheet2', selection), 0, 'Sheet1')).toBeNull();
  });
});
