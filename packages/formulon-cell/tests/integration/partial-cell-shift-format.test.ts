import { describe, expect, it, vi } from 'vitest';
import { deleteCells, insertCells } from '../../src/commands/cell-shift.js';
import { insertCopiedCellsFromTSV } from '../../src/commands/clipboard/insert-copied-cells.js';
import { setFont, setFontColor } from '../../src/commands/format.js';
import { History } from '../../src/commands/history.js';
import { recordFormatChange } from '../../src/commands/slice-history.js';
import type { Addr, CellValue } from '../../src/engine/types.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

const appearance = (wb: WorkbookHandle, addr: Addr) => {
  const effective = wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = effective ? wb.getCellXf(effective.xfIndex) : null;
  const font = xf ? wb.getFontRecord(xf.fontIndex) : null;
  return {
    direct: wb.getCellXfIndex(addr.sheet, addr.row, addr.col),
    source: effective?.source,
    size: font?.size,
    color: font?.colorArgb,
  };
};

const seedAxisStyles = (
  wb: WorkbookHandle,
  store: ReturnType<typeof createSpreadsheetStore>,
  history: History,
  source: Addr,
): void => {
  wb.setText(source, 'axis');
  mutators.selectRow(store, source.row);
  recordFormatChange(history, store, () => {
    expect(setFont(store.getState(), store, { fontSize: 24 })).toBe(true);
  });
  mutators.selectCol(store, source.col);
  recordFormatChange(history, store, () => {
    setFontColor(store.getState(), store, '#ff0000');
  });
};

const setActive = (store: ReturnType<typeof createSpreadsheetStore>, addr: Addr): void => {
  mutators.setActive(store, addr);
};

const value = (wb: WorkbookHandle, addr: Addr): CellValue => wb.getValue(addr);

describe('partial cell-shift native style checkpoints', () => {
  it('rejects a capable shift when its pre-snapshot fails before mutation', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    const source = { sheet: 0, row: 2, col: 3 };
    try {
      wb.setText(source, 'axis');
      wb.attachHistory(history);
      const snapshot = vi.spyOn(wb, 'snapshotRangeStyle').mockReturnValue(null);
      expect(
        insertCells(store, wb, history, { ...source, r0: 2, c0: 3, r1: 2, c1: 3 }, 'down'),
      ).toBe(false);
      snapshot.mockRestore();
      expect(value(wb, source)).toEqual({ kind: 'text', value: 'axis' });
      expect(value(wb, { ...source, row: 3 })).toEqual({ kind: 'blank' });
      expect(history.canUndo()).toBe(false);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('restores down-shift styles after undo and reapplies them on redo', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    const source = { sheet: 0, row: 2, col: 3 };
    const target = { ...source, row: 3 };
    try {
      seedAxisStyles(wb, store, history, source);
      wb.attachHistory(history);
      setActive(store, source);

      expect(
        insertCells(store, wb, history, { ...source, r0: 2, c0: 3, r1: 2, c1: 3 }, 'down'),
      ).toBe(true);
      expect(value(wb, target)).toEqual({ kind: 'text', value: 'axis' });

      expect(history.undo()).toBe(true);
      expect(value(wb, target)).toEqual({ kind: 'blank' });
      expect(appearance(wb, target)).toMatchObject({ direct: 0, color: 0xffff0000 });

      expect(history.undo()).toBe(true);
      expect(value(wb, target)).toEqual({ kind: 'blank' });
      // The target must not retain the direct XF materialized by the move.
      // The axis-format history owns the final inherited source transition.
      expect(appearance(wb, target)).toMatchObject({ direct: 0, color: 0xff000000 });

      expect(history.redo()).toBe(true);
      expect(appearance(wb, target)).toMatchObject({ direct: 0, color: 0xffff0000 });
      expect(history.redo()).toBe(true);
      expect(value(wb, target)).toEqual({ kind: 'text', value: 'axis' });
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('restores right-shift styles and supports up/left deletion replay', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    const source = { sheet: 0, row: 2, col: 3 };
    const rightTarget = { ...source, col: 4 };
    try {
      seedAxisStyles(wb, store, history, source);
      wb.attachHistory(history);
      setActive(store, source);

      expect(
        insertCells(store, wb, history, { ...source, r0: 2, c0: 3, r1: 2, c1: 3 }, 'right'),
      ).toBe(true);
      expect(value(wb, rightTarget)).toEqual({ kind: 'text', value: 'axis' });
      expect(history.undo()).toBe(true);
      expect(value(wb, rightTarget)).toEqual({ kind: 'blank' });
      expect(appearance(wb, rightTarget)).toMatchObject({ direct: 0, size: 18 });
      expect(history.redo()).toBe(true);
      expect(value(wb, rightTarget)).toEqual({ kind: 'text', value: 'axis' });

      // A deletion uses the same affected-range checkpoint in the inverse
      // direction. Delete row 3 and move a follower from row 4 to row 3.
      const follower = { ...source, row: 4, col: 5 };
      wb.setText(follower, 'follower');
      setActive(store, { ...follower, row: 3 });
      expect(
        deleteCells(
          store,
          wb,
          history,
          { sheet: 0, r0: 3, c0: follower.col, r1: 3, c1: follower.col },
          'up',
        ),
      ).toBe(true);
      expect(value(wb, { ...follower, row: 3 })).toEqual({ kind: 'text', value: 'follower' });
      expect(value(wb, follower)).toEqual({ kind: 'blank' });
      expect(history.undo()).toBe(true);
      expect(value(wb, follower)).toEqual({ kind: 'text', value: 'follower' });
      expect(value(wb, { ...follower, row: 3 })).toEqual({ kind: 'blank' });
      expect(history.redo()).toBe(true);
      expect(value(wb, { ...follower, row: 3 })).toEqual({ kind: 'text', value: 'follower' });

      // The left variant uses the same coordinate-bound native checkpoint.
      const leftFollower = { sheet: 0, row: 6, col: 6 };
      wb.setText(leftFollower, 'left-follower');
      setActive(store, { sheet: 0, row: leftFollower.row, col: 5 });
      expect(
        deleteCells(
          store,
          wb,
          history,
          { sheet: 0, r0: leftFollower.row, c0: 5, r1: leftFollower.row, c1: 5 },
          'left',
        ),
      ).toBe(true);
      expect(value(wb, { ...leftFollower, col: 5 })).toEqual({
        kind: 'text',
        value: 'left-follower',
      });
      expect(history.undo()).toBe(true);
      expect(value(wb, leftFollower)).toEqual({ kind: 'text', value: 'left-follower' });
      expect(value(wb, { ...leftFollower, col: 5 })).toEqual({ kind: 'blank' });
      expect(history.redo()).toBe(true);
      expect(value(wb, { ...leftFollower, col: 5 })).toEqual({
        kind: 'text',
        value: 'left-follower',
      });
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('keeps copied-cell insertion source styles through undo and redo', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    const source = { sheet: 0, row: 2, col: 3 };
    const inserted = { ...source };
    const shifted = { ...source, row: 3 };
    try {
      seedAxisStyles(wb, store, history, source);
      wb.attachHistory(history);
      setActive(store, source);

      expect(insertCopiedCellsFromTSV(store, wb, history, 'new', 'down')).toEqual({
        writtenRange: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
      });
      expect(value(wb, inserted)).toEqual({ kind: 'text', value: 'new' });
      expect(value(wb, shifted)).toEqual({ kind: 'text', value: 'axis' });

      expect(history.undo()).toBe(true);
      expect(value(wb, inserted)).toEqual({ kind: 'text', value: 'axis' });
      expect(value(wb, shifted)).toEqual({ kind: 'blank' });
      expect(appearance(wb, shifted)).toMatchObject({ direct: 0, color: 0xffff0000 });

      expect(history.redo()).toBe(true);
      expect(value(wb, inserted)).toEqual({ kind: 'text', value: 'new' });
      expect(value(wb, shifted)).toEqual({ kind: 'text', value: 'axis' });
      expect(appearance(wb, shifted)).toMatchObject({ size: 18, color: 0xffff0000 });
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('aborts the TSV insertion transaction when its initial style snapshot fails', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja' });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachStore(store);
    const source = { sheet: 0, row: 2, col: 3 };
    const tail = { ...source, row: 3 };
    try {
      wb.setText(source, 'source');
      wb.setText(tail, 'tail');
      setActive(store, source);
      wb.attachHistory(history);

      let calls = 0;
      const snapshot = vi.spyOn(wb, 'snapshotRangeStyle').mockImplementation((range) => {
        calls += 1;
        if (calls === 2) return null;
        return WorkbookHandle.prototype.snapshotRangeStyle.call(wb, range);
      });
      expect(() => insertCopiedCellsFromTSV(store, wb, history, 'new', 'down')).toThrow(
        /after-snapshot failed/,
      );
      snapshot.mockRestore();

      expect(value(wb, source)).toEqual({ kind: 'text', value: 'source' });
      expect(value(wb, tail)).toEqual({ kind: 'text', value: 'tail' });
      expect(value(wb, { ...tail, row: 1 })).toEqual({ kind: 'blank' });
      expect(history.canUndo()).toBe(false);
      expect(history.canRedo()).toBe(false);
    } finally {
      wb.attachHistory(null);
      wb.detachStore(store);
      wb.dispose();
    }
  });
});
