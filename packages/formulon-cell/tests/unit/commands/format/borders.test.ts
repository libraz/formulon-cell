import { beforeEach, describe, expect, it } from 'vitest';
import {
  cycleBorders,
  setBorderPreset,
  setBorders,
  toggleBold,
  toggleItalic,
} from '../../../../src/commands/format.js';
import { History, recordRepeatableFormatChange } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../../src/commands/interaction-policy.js';
import { setCellLocked, setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { effectiveFmtAt, fmtAt, setRange, setSelection } from './fixtures.js';

describe('borders', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 1, 1);
  });

  it('setBorders writes the supplied sides', () => {
    setBorders(store.getState(), store, { top: true, bottom: true });
    expect(fmtAt(store, 0, 0)?.borders).toMatchObject({ top: true, bottom: true });
  });

  it('setBorderPreset carries the selected line color into styled border sides', () => {
    setRange(store, 0, 0, 0, 0);
    setBorderPreset(store.getState(), store, 'outline', 'thick', '#c00000');

    expect(effectiveFmtAt(store, 0, 0)?.borders).toEqual({
      top: { style: 'thick', color: '#c00000' },
      bottom: { style: 'thick', color: '#c00000' },
      left: { style: 'thick', color: '#c00000' },
      right: { style: 'thick', color: '#c00000' },
    });
  });

  it('setBorderPreset applies inside borders only between cells in a range', () => {
    setRange(store, 0, 0, 1, 1);
    setBorderPreset(store.getState(), store, 'inside', 'thin', '#00a000');

    expect(fmtAt(store, 0, 0)?.borders).toBeUndefined();
    expect(fmtAt(store, 0, 1)?.borders).toEqual({
      left: { style: 'thin', color: '#00a000' },
    });
    expect(fmtAt(store, 1, 0)?.borders).toEqual({
      top: { style: 'thin', color: '#00a000' },
    });
    expect(fmtAt(store, 1, 1)?.borders).toEqual({
      top: { style: 'thin', color: '#00a000' },
      left: { style: 'thin', color: '#00a000' },
    });
  });

  it('setBorderPreset applies diagonal borders with the selected style and color', () => {
    setRange(store, 0, 0, 0, 0);
    setBorderPreset(store.getState(), store, 'diagonalDown', 'dashed', '#4472c4');
    setBorderPreset(store.getState(), store, 'diagonalUp', 'double', '#c00000');

    expect(effectiveFmtAt(store, 0, 0)?.borders).toEqual({
      diagonalDown: { style: 'dashed', color: '#4472c4' },
      diagonalUp: { style: 'double', color: '#c00000' },
    });
  });

  it('setBorderPreset merges with pending input format for a single empty active cell', () => {
    setRange(store, 0, 0, 0, 0);
    toggleBold(store.getState(), store);
    setBorderPreset(store.getState(), store, 'bottom', 'thin', '#4472c4');

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: {
        bold: true,
        borders: { bottom: { style: 'thin', color: '#4472c4' } },
      },
    });
    expect(effectiveFmtAt(store, 0, 0)?.borders?.bottom).toEqual({
      style: 'thin',
      color: '#4472c4',
    });
  });

  it('cycleBorders uses pending input format for a single empty active cell', () => {
    setRange(store, 0, 0, 0, 0);
    toggleItalic(store.getState(), store);
    cycleBorders(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: {
        italic: true,
        borders: { top: true, right: true, bottom: true, left: true },
      },
    });
  });

  it('cycleBorders paints an outline on first call (perimeter only)', () => {
    cycleBorders(store.getState(), store);
    // Top-left corner: top + left only.
    expect(fmtAt(store, 0, 0)?.borders).toMatchObject({ top: true, left: true });
    expect(fmtAt(store, 0, 0)?.borders?.right).toBeUndefined();
    expect(fmtAt(store, 0, 0)?.borders?.bottom).toBeUndefined();
    // Bottom-right: bottom + right only.
    expect(fmtAt(store, 1, 1)?.borders).toMatchObject({ bottom: true, right: true });
  });

  it('cycleBorders fills all four sides when only an outline exists', () => {
    cycleBorders(store.getState(), store); // outline
    cycleBorders(store.getState(), store); // all
    expect(fmtAt(store, 0, 1)?.borders).toMatchObject({
      top: true,
      right: true,
      bottom: true,
      left: true,
    });
  });

  it('cycleBorders clears all sides when every cell is fully bordered', () => {
    cycleBorders(store.getState(), store); // outline
    cycleBorders(store.getState(), store); // all
    cycleBorders(store.getState(), store); // clear
    const f = fmtAt(store, 0, 0)?.borders;
    expect(f?.top).toBe(false);
    expect(f?.right).toBe(false);
    expect(f?.bottom).toBe(false);
    expect(f?.left).toBe(false);
  });

  it('applies all borders across every selected fragment while leaving a hole untouched', () => {
    const rectangles = [
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
      { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
      { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
    ];
    const [primary, ...extras] = rectangles;
    if (!primary) throw new Error('Missing primary border range');
    setSelection(store, primary, extras);

    setBorderPreset(store.getState(), store, 'all');

    for (const [row, col] of [
      [0, 0],
      [1, 0],
      [2, 0],
      [0, 2],
      [1, 2],
      [2, 2],
      [0, 1],
      [2, 1],
      [4, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toEqual({
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      });
    }
    expect(fmtAt(store, 1, 1)).toBeUndefined();
  });

  it('keeps border edges independent for separate areas and does not bridge the gap', () => {
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    setBorderPreset(store.getState(), store, 'outline');

    expect(fmtAt(store, 0, 1)?.borders?.right).toEqual({ style: 'thin' });
    expect(fmtAt(store, 0, 3)?.borders?.left).toEqual({ style: 'thin' });
    expect(fmtAt(store, 0, 2)).toBeUndefined();

    const inside = createSpreadsheetStore();
    setSelection(inside, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    setBorderPreset(inside.getState(), inside, 'inside');

    expect(fmtAt(inside, 0, 0)).toBeUndefined();
    expect(fmtAt(inside, 1, 1)?.borders).toEqual({
      top: { style: 'thin' },
      left: { style: 'thin' },
    });
    expect(fmtAt(inside, 0, 3)).toBeUndefined();
    expect(fmtAt(inside, 1, 4)?.borders).toEqual({
      top: { style: 'thin' },
      left: { style: 'thin' },
    });
    expect(fmtAt(inside, 0, 2)).toBeUndefined();
  });

  it('cycles the complete disjoint union through outline, all, and clear', () => {
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { borders: { diagonalDown: true } },
    );

    cycleBorders(store.getState(), store);
    expect(fmtAt(store, 0, 1)?.borders?.right).toBe(true);
    expect(fmtAt(store, 0, 3)?.borders?.left).toBe(true);
    cycleBorders(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
      [0, 4],
      [1, 3],
      [1, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toMatchObject({
        top: true,
        right: true,
        bottom: true,
        left: true,
      });
    }
    cycleBorders(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
      [0, 4],
      [1, 3],
      [1, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toMatchObject({
        top: false,
        right: false,
        bottom: false,
        left: false,
      });
    }
    expect(fmtAt(store, 0, 2)).toBeUndefined();
    expect(fmtAt(store, 1, 2)).toBeUndefined();
    expect(fmtAt(store, 0, 0)?.borders?.diagonalDown).toBe(true);
  });

  it('rejects a capped union atomically and publishes one update for a valid multi-area preset', () => {
    const capped = createSpreadsheetStore();
    setSelection(capped, { sheet: 0, r0: 0, c0: 0, r1: 999, c1: 99 }, [
      { sheet: 0, r0: 1000, c0: 0, r1: 1000, c1: 0 },
    ]);
    mutators.setCellFormat(capped, { sheet: 0, row: 0, col: 0 }, { bold: true });
    mutators.setPendingFormat(capped, {
      addr: { sheet: 0, row: 0, col: 0 },
      format: { italic: true },
    });
    const beforeFormats = new Map(capped.getState().format.formats);
    const beforePending = capped.getState().ui.pendingFormat;

    setBorderPreset(capped.getState(), capped, 'outline');

    expect(capped.getState().format.formats).toEqual(beforeFormats);
    expect(capped.getState().ui.pendingFormat).toEqual(beforePending);

    const published = createSpreadsheetStore();
    setSelection(published, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    let updates = 0;
    const unsubscribe = published.subscribe(() => {
      updates += 1;
    });
    setBorderPreset(published.getState(), published, 'outline');
    unsubscribe();
    expect(updates).toBe(1);
  });

  it('rejects restricted mixed selections and skips locked fragments while formatting unlocked cells', async () => {
    const restricted = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store: restricted,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(restricted, controller);
    try {
      setSelection(restricted, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
      ]);
      controller.setPolicy({
        ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]),
        operations: { format: true },
      });
      setBorderPreset(restricted.getState(), restricted, 'outline');
      expect(restricted.getState().format.formats.size).toBe(0);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }

    const protectedStore = createSpreadsheetStore();
    setSelection(protectedStore, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
    ]);
    setCellLocked(protectedStore, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setCellLocked(protectedStore, { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, false);
    setProtectedSheet(protectedStore, 0, true);
    setBorderPreset(protectedStore.getState(), protectedStore, 'outline');
    expect(fmtAt(protectedStore, 0, 0)?.borders).toBeDefined();
    expect(fmtAt(protectedStore, 0, 3)?.borders).toBeDefined();
    expect(fmtAt(protectedStore, 0, 1)).toBeUndefined();
    expect(fmtAt(protectedStore, 0, 4)).toBeUndefined();
  });

  it('rejects partial merge coverage and formats a fully covered merge plus remote area', () => {
    const partial = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(partial, merge);
    setSelection(partial, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
    ]);
    setBorderPreset(partial.getState(), partial, 'outline');
    expect(partial.getState().format.formats.size).toBe(0);

    const complete = createSpreadsheetStore();
    mutators.mergeRange(complete, merge);
    setSelection(complete, merge, [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }]);
    setBorderPreset(complete.getState(), complete, 'outline');
    expect(fmtAt(complete, 0, 0)?.borders?.top).toEqual({ style: 'thin' });
    expect(fmtAt(complete, 1, 1)?.borders?.bottom).toEqual({ style: 'thin' });
    expect(fmtAt(complete, 0, 3)?.borders?.top).toEqual({ style: 'thin' });
  });

  it('keeps one-cell inside borders pending and materializes a selection with extras', () => {
    const single = createSpreadsheetStore();
    setSelection(single, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toggleBold(single.getState(), single);
    const pendingBefore = single.getState().ui.pendingFormat;
    setBorderPreset(single.getState(), single, 'inside');
    expect(single.getState().ui.pendingFormat).toEqual(pendingBefore);
    expect(single.getState().format.formats.size).toBe(0);

    const multi = createSpreadsheetStore();
    setSelection(multi, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);
    setBorderPreset(multi.getState(), multi, 'outline');
    expect(multi.getState().ui.pendingFormat).toBeNull();
    expect(fmtAt(multi, 0, 0)?.borders).toBeDefined();
    expect(fmtAt(multi, 0, 1)?.borders).toBeDefined();
  });

  it('records a multi-area border preset as one undo entry', () => {
    const history = new History();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    recordRepeatableFormatChange(history, store, () => {
      setBorderPreset(store.getState(), store, 'all');
    });

    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.size).toBe(0);
    expect(history.undo()).toBe(false);
  });
});
