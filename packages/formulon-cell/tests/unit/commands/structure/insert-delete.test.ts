import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  deleteCols,
  deleteRows,
  insertCols,
  insertRows,
} from '../../../../src/commands/structure.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { cellNumber, cellText, newWb, seedCols, seedRows } from './fixtures.js';

describe('insertRows', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    seedRows(wb);
  });

  it('shifts cells down by count', () => {
    insertRows(store, wb, null, 1, 1);
    expect(cellNumber(wb, 0, 0, 0)).toBe(10);
    expect(cellNumber(wb, 0, 1, 0)).toBeNaN(); // blank
    expect(cellNumber(wb, 0, 2, 0)).toBe(20);
    expect(cellNumber(wb, 0, 3, 0)).toBe(30);
  });

  it('shifts formats', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { bold: true });
    insertRows(store, wb, null, 1, 1);
    expect(store.getState().format.formats.get('0:1:0')).toBeUndefined();
    expect(store.getState().format.formats.get('0:2:0')?.bold).toBe(true);
  });

  it('shifts row heights, hidden rows, and freeze pane', () => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        rowHeights: new Map([[2, 50]]),
        hiddenRows: new Set([2]),
        freezeRows: 2,
      },
    }));
    insertRows(store, wb, null, 1, 1);
    const layout = store.getState().layout;
    expect(layout.rowHeights.get(3)).toBe(50);
    expect(layout.rowHeights.get(2)).toBeUndefined();
    expect(Array.from(layout.hiddenRows)).toEqual([3]);
    expect(layout.freezeRows).toBe(3); // 2 was > atRow=1 → 2+1
  });

  it('does not shift freeze when atRow >= freezeRows', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 1 },
    }));
    insertRows(store, wb, null, 5, 2); // atRow=5, beyond freeze
    expect(store.getState().layout.freezeRows).toBe(1);
  });

  it('is a no-op for count <= 0', () => {
    insertRows(store, wb, null, 0, 0);
    expect(cellNumber(wb, 0, 1, 0)).toBe(20);
  });

  it('rejects non-integer and out-of-grid insertion origins', () => {
    wb.setNumber({ sheet: 0, row: 1_048_575, col: 0 }, 77);
    insertRows(store, wb, null, 1_048_575, 1);
    insertRows(store, wb, null, 1.5, 1);
    insertRows(store, wb, null, Number.NaN, 1);
    expect(cellNumber(wb, 0, 1_048_575, 0)).toBe(77);
  });

  it('rejects an insertion that would push a populated tail row or column off-grid', () => {
    const h = new History();
    wb.setText({ sheet: 0, row: 1_048_575, col: 0 }, 'tail row');
    insertRows(store, wb, h, 1_048_574, 1);
    expect(cellText(wb, 0, 1_048_575, 0)).toBe('tail row');

    wb.setText({ sheet: 0, row: 0, col: 16_383 }, 'tail col');
    insertCols(store, wb, h, 16_382, 1);
    expect(cellText(wb, 0, 0, 16_383)).toBe('tail col');
    expect(h.canUndo()).toBe(false);
  });

  it('round-trips with history', () => {
    const h = new History();
    wb.attachHistory(h);
    insertRows(store, wb, h, 1, 1);
    expect(cellNumber(wb, 0, 2, 0)).toBe(20);

    // One transaction = one undo.
    h.undo();
    expect(cellNumber(wb, 0, 1, 0)).toBe(20);
    expect(cellNumber(wb, 0, 2, 0)).toBe(30);

    h.redo();
    expect(cellNumber(wb, 0, 2, 0)).toBe(20);
    expect(cellNumber(wb, 0, 3, 0)).toBe(30);
  });
});

describe('deleteRows', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    seedRows(wb);
  });

  it('removes the targeted row and shifts subsequent rows up', () => {
    deleteRows(store, wb, null, 1, 1);
    expect(cellNumber(wb, 0, 0, 0)).toBe(10);
    expect(cellNumber(wb, 0, 1, 0)).toBe(30); // was row 2
    expect(cellNumber(wb, 0, 2, 0)).toBeNaN();
  });

  it('drops formats in the deleted band', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { bold: true });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { italic: true });
    deleteRows(store, wb, null, 1, 1);
    expect(store.getState().format.formats.get('0:1:0')?.italic).toBe(true);
    expect(store.getState().format.formats.get('0:2:0')).toBeUndefined();
  });

  it('clamps freezeRows when deleting through the freeze band', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 3 },
    }));
    deleteRows(store, wb, null, 1, 5);
    // freezeRows was 3, atRow=1, n=5 → fr = max(atRow, fr-n) = max(1, -2) = 1
    expect(store.getState().layout.freezeRows).toBe(1);
  });

  it('round-trips with history', () => {
    const h = new History();
    wb.attachHistory(h);
    deleteRows(store, wb, h, 1, 1);
    expect(cellNumber(wb, 0, 1, 0)).toBe(30);

    h.undo();
    expect(cellNumber(wb, 0, 1, 0)).toBe(20);
    expect(cellNumber(wb, 0, 2, 0)).toBe(30);
  });
});

describe('insertCols', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    seedCols(wb);
  });

  it('shifts cells right by count', () => {
    insertCols(store, wb, null, 1, 1);
    expect(cellNumber(wb, 0, 0, 0)).toBe(10);
    expect(cellNumber(wb, 0, 0, 1)).toBeNaN();
    expect(cellNumber(wb, 0, 0, 2)).toBe(20);
    expect(cellNumber(wb, 0, 0, 3)).toBe(30);
  });

  it('shifts colWidths, hiddenCols, and freezeCols', () => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([[2, 200]]),
        hiddenCols: new Set([2]),
        freezeCols: 2,
      },
    }));
    insertCols(store, wb, null, 1, 1);
    const layout = store.getState().layout;
    expect(layout.colWidths.get(3)).toBe(200);
    expect(Array.from(layout.hiddenCols)).toEqual([3]);
    expect(layout.freezeCols).toBe(3);
  });
});

describe('deleteCols', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    seedCols(wb);
  });

  it('removes the targeted column', () => {
    deleteCols(store, wb, null, 1, 1);
    expect(cellNumber(wb, 0, 0, 0)).toBe(10);
    expect(cellNumber(wb, 0, 0, 1)).toBe(30);
    expect(cellNumber(wb, 0, 0, 2)).toBeNaN();
  });

  it('shifts colWidths and drops widths in deleted band', () => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([
          [1, 150],
          [2, 200],
        ]),
      },
    }));
    deleteCols(store, wb, null, 1, 1);
    const widths = store.getState().layout.colWidths;
    // col 1 dropped; col 2 → 1
    expect(widths.get(1)).toBe(200);
    expect(widths.size).toBe(1);
  });
});
