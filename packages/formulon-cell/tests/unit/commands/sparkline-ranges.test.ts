import { describe, expect, it } from 'vitest';
import {
  MAX_SPARKLINE_RANGE_CELLS,
  resolveSparklineRange,
  sparklineSourceText,
} from '../../../src/commands/sparkline-ranges.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const wbWithSheets = (names: string[]): WorkbookHandle =>
  ({
    sheetName: (i: number) => names[i],
    sheetIndexByName: (name: string) => {
      const at = names.indexOf(name);
      return at < 0 ? undefined : at;
    },
  }) as unknown as WorkbookHandle;

describe('resolveSparklineRange size cap', () => {
  it('accepts a full-width row and refuses anything larger', () => {
    expect(resolveSparklineRange(`A1:XFD1`, 0, null)).not.toBeNull();
    expect(MAX_SPARKLINE_RANGE_CELLS).toBe(16_384);
    expect(resolveSparklineRange('A1:A16385', 0, null)).toBeNull();
    expect(resolveSparklineRange('A1:Z1048576', 0, null)).toBeNull();
  });
});

describe('sparklineSourceText sheet quoting', () => {
  const source = { sheet: 1, r0: 0, c0: 0, r1: 2, c1: 0, explicitSheet: false };

  it('quotes a sheet named like a cell reference', () => {
    expect(sparklineSourceText(source, 0, wbWithSheets(['Data', 'Q1']))).toBe("'Q1'!A1:A3");
  });

  it('leaves a plain identifier bare and doubles quotes in other names', () => {
    expect(sparklineSourceText(source, 0, wbWithSheets(['Data', 'Sales']))).toBe('Sales!A1:A3');
    expect(sparklineSourceText(source, 0, wbWithSheets(['Data', "it's"]))).toBe("'it''s'!A1:A3");
  });
});
