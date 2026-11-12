import { describe, expect, it } from 'vitest';
import { formatMacRangeAddress } from '../../../src/commands/mac-data-tools.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const workbookWithSheet = (name: string): WorkbookHandle =>
  ({ sheetName: () => name }) as unknown as WorkbookHandle;

const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };

// A sheet name that reads as a cell or R1C1 reference, or as a number, keeps its quotes, as in
// any other formula reference.
describe('Mac data tool range addresses', () => {
  it.each([
    ['A1', "'A1'!A1:B2"],
    ['R1C1', "'R1C1'!A1:B2"],
    ['2024', "'2024'!A1:B2"],
    ['Sales', 'Sales!A1:B2'],
    ["Bob's", "'Bob''s'!A1:B2"],
  ])('quote the sheet %s where a formula would', (name, expected) => {
    expect(formatMacRangeAddress(workbookWithSheet(name), range)).toBe(expected);
  });
});
