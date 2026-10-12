import { describe, expect, it } from 'vitest';
import { History } from '../../src/commands/history.js';
import {
  autofitColsWidthForRanges,
  autofitRowsHeightForRanges,
} from '../../src/commands/row-col-layout.js';
import { hydrateLayoutFromEngine } from '../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../src/store/store.js';

const rowLayout = (workbook: WorkbookHandle, row: number) =>
  workbook.getRowLayouts(0).find((layout) => layout.row === row);

const colLayout = (workbook: WorkbookHandle, col: number) =>
  workbook
    .getColumnLayouts(0)
    .find((layout) => layout.first <= col && layout.last >= col && layout.hasWidth !== false);

const hydratedLayout = (workbook: WorkbookHandle) => {
  const store = createSpreadsheetStore();
  hydrateLayoutFromEngine(workbook, store, 0);
  return store.getState().layout;
};

describe('native multi-range AutoFit layout history', () => {
  it('round-trips disjoint rows and columns through native layout, undo/redo, and save/reload', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    store.setState((state) => ({
      ...state,
      data: {
        ...state.data,
        cells: new Map([
          ['0:0:0', { value: { kind: 'text', value: 'row zero\nsecond line' }, formula: null }],
          [
            '0:2:2',
            { value: { kind: 'text', value: 'row two\nsecond\nthird line' }, formula: null },
          ],
        ]),
      },
    }));

    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.capabilities.colRowSize).toBe(true);
      expect(workbook.capabilities.geometry).toBe(true);

      const rowRanges = [
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
      ] as const;
      const colRanges = [
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
      ] as const;

      autofitRowsHeightForRanges(store, history, rowRanges, workbook);
      const rowZeroHeight = store.getState().layout.rowHeights.get(0);
      const rowTwoHeight = store.getState().layout.rowHeights.get(2);
      expect(rowZeroHeight).toBeGreaterThan(store.getState().layout.defaultRowHeight);
      expect(rowTwoHeight).toBeGreaterThan(store.getState().layout.defaultRowHeight);
      expect(rowLayout(workbook, 0)?.height).toBeCloseTo(rowZeroHeight ?? 0, 8);
      expect(rowLayout(workbook, 2)?.height).toBeCloseTo(rowTwoHeight ?? 0, 8);

      autofitColsWidthForRanges(store, history, colRanges, workbook);
      const colZeroWidth = store.getState().layout.colWidths.get(0);
      const colTwoWidth = store.getState().layout.colWidths.get(2);
      expect(colZeroWidth).toBeGreaterThanOrEqual(48);
      expect(colTwoWidth).toBeGreaterThanOrEqual(48);
      expect(colLayout(workbook, 0)?.width).toBeCloseTo(colZeroWidth ?? 0, 8);
      expect(colLayout(workbook, 2)?.width).toBeCloseTo(colTwoWidth ?? 0, 8);

      expect(history.undo()).toBe(true);
      const afterColumnUndo = hydratedLayout(workbook);
      expect(afterColumnUndo.colWidths.get(0)).toBeCloseTo(
        store.getState().layout.defaultColWidth,
        8,
      );
      expect(afterColumnUndo.colWidths.get(2)).toBeCloseTo(
        store.getState().layout.defaultColWidth,
        8,
      );
      expect(afterColumnUndo.rowHeights.get(0)).toBeCloseTo(rowZeroHeight ?? 0, 8);
      expect(afterColumnUndo.rowHeights.get(2)).toBeCloseTo(rowTwoHeight ?? 0, 8);

      expect(history.undo()).toBe(true);
      const afterRowUndo = hydratedLayout(workbook);
      expect(afterRowUndo.colWidths.get(0)).toBeCloseTo(store.getState().layout.defaultColWidth, 8);
      expect(afterRowUndo.colWidths.get(2)).toBeCloseTo(store.getState().layout.defaultColWidth, 8);
      expect(afterRowUndo.rowHeights.has(0)).toBe(false);
      expect(afterRowUndo.rowHeights.has(2)).toBe(false);

      expect(history.redo()).toBe(true);
      expect(rowLayout(workbook, 0)?.height).toBeCloseTo(rowZeroHeight ?? 0, 8);
      expect(rowLayout(workbook, 2)?.height).toBeCloseTo(rowTwoHeight ?? 0, 8);
      expect(history.redo()).toBe(true);
      expect(colLayout(workbook, 0)?.width).toBeCloseTo(colZeroWidth ?? 0, 8);
      expect(colLayout(workbook, 2)?.width).toBeCloseTo(colTwoWidth ?? 0, 8);

      const reloaded = await WorkbookHandle.loadBytes(workbook.save());
      try {
        const reloadedLayout = hydratedLayout(reloaded);
        expect(reloadedLayout.rowHeights.get(0)).toBeCloseTo(rowZeroHeight ?? 0, 8);
        expect(reloadedLayout.rowHeights.get(2)).toBeCloseTo(rowTwoHeight ?? 0, 8);
        expect(reloadedLayout.colWidths.get(0)).toBeCloseTo(colZeroWidth ?? 0, 8);
        expect(reloadedLayout.colWidths.get(2)).toBeCloseTo(colTwoWidth ?? 0, 8);
      } finally {
        reloaded.dispose();
      }
    } finally {
      workbook.dispose();
    }
  });
});
