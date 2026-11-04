import { describe, expect, it, vi } from 'vitest';
import { recordDialogFormatChange } from '../../../src/commands/dialog-format-history.js';
import { History } from '../../../src/commands/history.js';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('recordDialogFormatChange', () => {
  it('accepts an unchanged mutation, seeds repeat, and records no material undo', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const repeat = vi.fn();
    const target = { sheet: 0, row: 0, col: 0 };

    expect(
      recordDialogFormatChange({
        history,
        store,
        workbook: null,
        sheet: 0,
        targets: [target],
        mutate: () => true,
        repeat,
      }),
    ).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(history.repeatLast()).toBe(true);
    expect(repeat).toHaveBeenCalledTimes(1);
  });

  it('does not touch state or history when the mutation is rejected', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    const before = store.getState();
    expect(
      recordDialogFormatChange({
        history,
        store,
        workbook: null,
        sheet: 0,
        targets: [target],
        mutate: () => false,
      }),
    ).toBe(false);
    expect(store.getState().format.formats).toEqual(before.format.formats);
    expect(history.canUndo()).toBe(false);
  });

  it('treats nested format objects with different key order as a no-op', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    mutators.setCellFormat(store, target, {
      validation: { kind: 'list', source: ['A', 'B'] },
    });
    const repeat = vi.fn();

    expect(
      recordDialogFormatChange({
        history,
        store,
        workbook: null,
        sheet: 0,
        targets: [target],
        mutate: () => {
          mutators.setCellFormat(store, target, {
            validation: { source: ['A', 'B'], kind: 'list' },
          });
          return true;
        },
        repeat,
      }),
    ).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(history.repeatLast()).toBe(true);
    expect(repeat).toHaveBeenCalledTimes(1);
  });

  it('flushes one scoped format change and replays it without touching outside edits', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    const outside = { sheet: 0, row: 1, col: 1 };
    workbook.setNumber(target, 1);
    workbook.setNumber(outside, 2);
    mutators.setCellFormat(store, target, { bold: true });
    mutators.setCellFormat(store, outside, { italic: true });
    flushFormatToEngine(workbook, store, 0);

    expect(
      recordDialogFormatChange({
        history,
        store,
        workbook,
        sheet: 0,
        targets: [target],
        mutate: () => {
          mutators.setCellFormat(store, target, { bold: false, fill: '#ffeeaa' });
          return true;
        },
        repeat: vi.fn(),
      }),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
      bold: false,
      fill: '#ffeeaa',
    });
    expect(store.getState().format.formats.get('0:1:1')?.italic).toBe(true);
    mutators.setCellFormat(store, outside, { fill: '#abc' });
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({ bold: true });
    expect(store.getState().format.formats.get('0:0:0')?.fill).toBeUndefined();
    expect(store.getState().format.formats.get('0:1:1')?.fill).toBe('#abc');
    expect(history.redo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.fill).toBe('#ffeeaa');
    expect(store.getState().format.formats.get('0:1:1')?.fill).toBe('#abc');
    workbook.dispose();
  });

  it('restores the scoped map when the forward flush fails', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(target, 1);
    mutators.setCellFormat(store, target, { bold: true });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    flushFormatToEngine(workbook, store, 0);
    const original = workbook.addXfRecord.bind(workbook);
    let injected = false;
    vi.spyOn(workbook, 'addXfRecord').mockImplementation((record) => {
      if (!injected) {
        injected = true;
        throw new Error('injected dialog flush failure');
      }
      return original(record);
    });

    expect(() =>
      recordDialogFormatChange({
        history,
        store,
        workbook,
        sheet: 0,
        targets: [target],
        mutate: () => {
          mutators.setCellFormat(store, target, { bold: false });
          return true;
        },
        repeat: vi.fn(),
      }),
    ).toThrow('injected dialog flush failure');
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(history.canUndo()).toBe(false);
    workbook.dispose();
  });

  it('surfaces a forward rollback failure instead of claiming atomic recovery', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const target = { sheet: 0, row: 0, col: 0 };
    try {
      mutators.setCellFormat(store, target, { bold: true });
      flushFormatToEngine(workbook, store, 0);
      const failing = vi.spyOn(workbook, 'addXfRecord').mockImplementation(() => {
        throw new Error('persistent engine failure');
      });

      expect(() =>
        recordDialogFormatChange({
          history: new History(),
          store,
          workbook,
          sheet: 0,
          targets: [target],
          mutate: () => {
            mutators.setCellFormat(store, target, { bold: false });
            return true;
          },
        }),
      ).toThrow(AggregateError);
      expect(failing).toHaveBeenCalled();
      failing.mockRestore();
    } finally {
      workbook.dispose();
    }
  });

  it('keeps the history position stable and surfaces replay rollback failure', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    try {
      mutators.setCellFormat(store, target, { bold: true });
      flushFormatToEngine(workbook, store, 0);
      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          mutate: () => {
            mutators.setCellFormat(store, target, { bold: false });
            return true;
          },
        }),
      ).toBe(true);
      const failing = vi.spyOn(workbook, 'addXfRecord').mockImplementation(() => {
        throw new Error('persistent replay failure');
      });

      expect(() => history.undo()).toThrow(AggregateError);
      expect(history.canUndo()).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(false);
      expect(failing).toHaveBeenCalled();
      failing.mockRestore();
    } finally {
      workbook.dispose();
    }
  });
});
