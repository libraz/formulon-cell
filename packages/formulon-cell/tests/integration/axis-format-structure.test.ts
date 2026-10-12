import { describe, expect, it, vi } from 'vitest';
import { applySelectionFormatAction, setFont, setFontColor } from '../../src/commands/format.js';
import { History } from '../../src/commands/history.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../../src/commands/structure.js';
import type { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { WorkbookHandle as Workbook } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

const appearance = (wb: WorkbookHandle, row: number, col: number) => {
  const effective = wb.getEffectiveStyle(0, row, col);
  const xf = effective ? wb.getCellXf(effective.xfIndex) : null;
  const font = xf ? wb.getFontRecord(xf.fontIndex) : null;
  return {
    source: effective?.source,
    direct: wb.getCellXfIndex(0, row, col),
    size: font?.size,
    color: font?.colorArgb,
  };
};

const directKeys = (store: ReturnType<typeof createSpreadsheetStore>): string[] =>
  store
    .getState()
    .format.sheetAxisOverrides?.get(0)
    ?.directXfs?.map((entry) => entry.key) ?? [];

describe('axis format structure history', () => {
  it('remaps row and column overrides across insertion and later unrelated edits', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(insertRows(store, wb, history, 2, 1)).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(11);
      expect(appearance(wb, 3, 0).size).toBe(18);
      mutators.selectCol(store, 6);
      expect(setFont(store.getState(), store, { fontSize: 32 })).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(11);
      expect(appearance(wb, 3, 0).size).toBe(18);

      mutators.selectCol(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(insertCols(store, wb, history, 2, 1)).toBe(true);
      expect(appearance(wb, 0, 2).size).toBe(11);
      expect(appearance(wb, 0, 3).size).toBe(18);
      mutators.selectRow(store, 6);
      expect(setFont(store.getState(), store, { fontSize: 32 })).toBe(true);
      expect(appearance(wb, 0, 2).size).toBe(11);
      expect(appearance(wb, 0, 3).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('drops deleted axis styles and restores them through Undo/Redo', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(deleteRows(store, wb, history, 2, 1)).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(11);
      expect(history.undo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect(history.redo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(11);

      mutators.selectCol(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(deleteCols(store, wb, history, 2, 1)).toBe(true);
      expect(appearance(wb, 0, 2).size).toBe(11);
      expect(history.undo()).toBe(true);
      expect(appearance(wb, 0, 2).size).toBe(18);
      expect(history.redo()).toBe(true);
      expect(appearance(wb, 0, 2).size).toBe(11);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('keeps a legacy full-sheet font override through row insertion and Undo/Redo', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectAll(store);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect(insertRows(store, wb, history, 2, 1)).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect(appearance(wb, 3, 0).size).toBe(18);
      expect(history.undo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect(history.redo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect(appearance(wb, 3, 0).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('checkpoints a directly formatted cell that moves through structure history', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.setActive(store, { sheet: 0, row: 4, col: 2 });
      expect(
        applySelectionFormatAction(
          store.getState(),
          store,
          { patch: { fontSize: 24 } },
          { allowPending: false },
        ),
      ).toBe(true);
      expect(store.getState().format.formats.has('0:4:2')).toBe(true);
      expect(appearance(wb, 4, 2).size).toBe(18);
      expect(insertRows(store, wb, history, 3, 1)).toBe(true);
      expect(appearance(wb, 5, 2).size).toBe(18);
      expect(history.undo()).toBe(true);
      expect(appearance(wb, 4, 2).size).toBe(18);
      expect(history.redo()).toBe(true);
      expect(appearance(wb, 5, 2).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('rolls back a structure undo when an axis restore setter rejects once', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(insertRows(store, wb, history, 2, 1)).toBe(true);
      let rejectOnce = true;
      const setter =
        wb.capabilities.rangeStyleSnapshot === true
          ? vi.spyOn(wb, 'restoreRangeStyle').mockImplementation((sheet, bytes) => {
              if (rejectOnce) {
                rejectOnce = false;
                return false;
              }
              return Workbook.prototype.restoreRangeStyle.call(wb, sheet, bytes);
            })
          : vi.spyOn(wb, 'setRowStyle').mockImplementation((sheet, row, xfIndex) => {
              if (rejectOnce) {
                rejectOnce = false;
                return false;
              }
              return Workbook.prototype.setRowStyle.call(wb, sheet, row, xfIndex);
            });
      expect(() => history.undo()).toThrow();
      expect(appearance(wb, 2, 0).size).toBe(11);
      expect(appearance(wb, 3, 0).size).toBe(18);
      setter.mockRestore();
      expect(history.undo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('restores every store slice when native undo fails before mutation', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(insertRows(store, wb, history, 2, 1)).toBe(true);
      const afterStore = structuredClone(store.getState());
      const afterNative = wb.snapshotRangeStyle({
        sheet: 0,
        r0: 2,
        c0: 0,
        r1: 1048575,
        c1: 16383,
      });
      expect(afterNative).not.toBeNull();

      const reject = vi.spyOn(wb, 'engineDeleteRows').mockReturnValue(false);
      expect(() => history.undo()).toThrow(/native axis undo failed/);
      expect(store.getState()).toEqual(afterStore);
      expect([
        ...(wb.snapshotRangeStyle({
          sheet: 0,
          r0: 2,
          c0: 0,
          r1: 1048575,
          c1: 16383,
        }) ?? []),
      ]).toEqual([...(afterNative ?? [])]);
      reject.mockRestore();

      expect(history.undo()).toBe(true);
      expect(appearance(wb, 2, 0).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('restores the before store and cache when redo style replay fails', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      expect(insertRows(store, wb, history, 2, 1)).toBe(true);
      const afterStore = structuredClone(store.getState());
      expect(history.undo()).toBe(true);
      const beforeStore = structuredClone(store.getState());

      let rejectOnce = true;
      const restore = vi.spyOn(wb, 'restoreRangeStyle').mockImplementation((sheet, bytes) => {
        if (rejectOnce) {
          rejectOnce = false;
          return false;
        }
        return Workbook.prototype.restoreRangeStyle.call(wb, sheet, bytes);
      });
      expect(() => history.redo()).toThrow(/failed to restore native range styles/);
      expect(store.getState()).toEqual(beforeStore);
      expect(store.getState()).not.toEqual(afterStore);
      expect(appearance(wb, 2, 0).size).toBe(18);
      restore.mockRestore();

      expect(history.redo()).toBe(true);
      expect(store.getState()).toEqual(afterStore);
      expect(appearance(wb, 3, 0).size).toBe(18);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('rejects a capable structure edit when its pre-snapshot fails', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      const before = wb.snapshotRangeStyle({ sheet: 0, r0: 1, c0: 0, r1: 1048575, c1: 16383 });
      expect(before).not.toBeNull();
      const snapshot = vi.spyOn(wb, 'snapshotRangeStyle').mockReturnValue(null);
      expect(insertRows(store, wb, history, 1, 1)).toBe(false);
      snapshot.mockRestore();
      expect(history.canUndo()).toBe(false);
      expect(appearance(wb, 2, 0).size).toBe(18);
      expect([
        ...(wb.snapshotRangeStyle({ sheet: 0, r0: 1, c0: 0, r1: 1048575, c1: 16383 }) ?? []),
      ]).toEqual([...(before ?? [])]);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('moves mixed row/column direct corrections through insertion', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      mutators.selectCol(store, 3);
      setFontColor(store.getState(), store, '#ff0000');
      expect(directKeys(store)).toContain('0:2:3');
      expect(insertRows(store, wb, null, 1, 1)).toBe(true);
      expect(insertCols(store, wb, null, 1, 1)).toBe(true);
      expect(directKeys(store)).toContain('0:3:4');
      expect(directKeys(store)).not.toContain('0:2:3');
      expect(appearance(wb, 3, 4).size).toBe(18);
      expect(appearance(wb, 3, 4).color).toBe(0xffff0000);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('restores a deleted mixed direct correction and does not resurrect it later', async () => {
    const wb = await Workbook.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    try {
      mutators.selectRow(store, 2);
      expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
      mutators.selectCol(store, 3);
      setFontColor(store.getState(), store, '#ff0000');
      expect(deleteRows(store, wb, history, 2, 1)).toBe(true);
      expect(directKeys(store)).not.toContain('0:2:3');
      expect(appearance(wb, 2, 3).direct).toBe(0);
      expect(history.undo()).toBe(true);
      expect(directKeys(store)).toContain('0:2:3');
      expect(appearance(wb, 2, 3).size).toBe(18);
      expect(appearance(wb, 2, 3).color).toBe(0xffff0000);
      expect(appearance(wb, 2, 3).direct).toBeGreaterThan(0);
      expect(history.redo()).toBe(true);
      expect(directKeys(store)).not.toContain('0:2:3');
      mutators.selectCol(store, 8);
      expect(setFont(store.getState(), store, { fontSize: 32 })).toBe(true);
      expect(appearance(wb, 2, 3).direct).toBe(0);
      expect(appearance(wb, 2, 3).size).toBe(11);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });
});
