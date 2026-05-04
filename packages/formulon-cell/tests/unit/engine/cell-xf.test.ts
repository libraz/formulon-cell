import { describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

interface XfCalls {
  getIndex: { sheet: number; row: number; col: number }[];
  setIndex: { sheet: number; row: number; col: number; xfIndex: number }[];
  getXf: number[];
}

const makeFake = (
  opts: { cellFormatting?: boolean; record?: Record<string, unknown> } = {},
): { wb: WorkbookHandle; calls: XfCalls } => {
  const calls: XfCalls = { getIndex: [], setIndex: [], getXf: [] };
  const caps = { cellFormatting: opts.cellFormatting ?? true };
  const fake = {
    capabilities: caps,
    getCellXfIndex(sheet: number, row: number, col: number): number | null {
      if (!caps.cellFormatting) return null;
      calls.getIndex.push({ sheet, row, col });
      return 7;
    },
    setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
      if (!caps.cellFormatting) return false;
      calls.setIndex.push({ sheet, row, col, xfIndex });
      return true;
    },
    getCellXf(xfIndex: number): ReturnType<WorkbookHandle['getCellXf']> {
      if (!caps.cellFormatting) return null;
      calls.getXf.push(xfIndex);
      return {
        fontIndex: 1,
        fillIndex: 2,
        borderIndex: 3,
        numFmtId: 49,
        horizontalAlign: 0,
        verticalAlign: 0,
        wrapText: false,
        ...(opts.record ?? {}),
      } as ReturnType<WorkbookHandle['getCellXf']>;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, calls };
};

describe('cellFormatting wrappers', () => {
  it('getCellXfIndex returns the engine value when capability is on', () => {
    const { wb, calls } = makeFake({ cellFormatting: true });
    expect(wb.getCellXfIndex(0, 1, 2)).toBe(7);
    expect(calls.getIndex).toEqual([{ sheet: 0, row: 1, col: 2 }]);
  });

  it('getCellXfIndex returns null and bypasses the call when capability is off', () => {
    const { wb, calls } = makeFake({ cellFormatting: false });
    expect(wb.getCellXfIndex(0, 0, 0)).toBeNull();
    expect(calls.getIndex).toEqual([]);
  });

  it('setCellXfIndex forwards args when capability is on', () => {
    const { wb, calls } = makeFake({ cellFormatting: true });
    expect(wb.setCellXfIndex(0, 5, 6, 12)).toBe(true);
    expect(calls.setIndex).toEqual([{ sheet: 0, row: 5, col: 6, xfIndex: 12 }]);
  });

  it('setCellXfIndex returns false when capability is off', () => {
    const { wb, calls } = makeFake({ cellFormatting: false });
    expect(wb.setCellXfIndex(0, 0, 0, 0)).toBe(false);
    expect(calls.setIndex).toEqual([]);
  });

  it('getCellXf resolves to a plain JS record', () => {
    const { wb } = makeFake({
      cellFormatting: true,
      record: { numFmtId: 14, wrapText: true },
    });
    expect(wb.getCellXf(3)).toEqual({
      fontIndex: 1,
      fillIndex: 2,
      borderIndex: 3,
      numFmtId: 14,
      horizontalAlign: 0,
      verticalAlign: 0,
      wrapText: true,
    });
  });

  it('getCellXf returns null when capability is off', () => {
    const { wb, calls } = makeFake({ cellFormatting: false });
    expect(wb.getCellXf(0)).toBeNull();
    expect(calls.getXf).toEqual([]);
  });
});
