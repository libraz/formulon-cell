import { describe, expect, it } from 'vitest';
import {
  applySelectionFormatPatch,
  bumpIndent,
  clearFormat,
  clearVisualFormat,
  toggleBold,
  withSelectionFormatOrigin,
} from '../../../../src/commands/format.js';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../../src/commands/interaction-policy.js';
import { setCellLocked, setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { fmtAt, setSelection } from './fixtures.js';

describe('non-contiguous selection formatting', () => {
  it('toggles the complete union while leaving a hole untouched', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
    ]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });

    toggleBold(store.getState(), store);
    for (const col of [0, 1, 3, 4]) expect(fmtAt(store, 0, col)?.bold).toBe(true);
    expect(fmtAt(store, 0, 2)).toBeUndefined();

    toggleBold(store.getState(), store);
    for (const col of [0, 1, 3, 4]) expect(fmtAt(store, 0, col)?.bold).toBe(false);
    expect(fmtAt(store, 0, 2)).toBeUndefined();
  });

  it('applies relative changes once to overlapping union areas', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
    ]);
    bumpIndent(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.indent).toBe(1);
    expect(fmtAt(store, 1, 1)?.indent).toBe(1);
    expect(fmtAt(store, 2, 2)?.indent).toBe(1);
  });

  it('rejects a union over the materialization cap without a partial write', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 999, c1: 99 }, [
      { sheet: 0, r0: 1000, c0: 0, r1: 1000, c1: 0 },
    ]);

    toggleBold(store.getState(), store);

    expect(store.getState().format.formats.size).toBe(0);
  });

  it('stages one blank cell but materializes a blank extra selection', () => {
    const single = createSpreadsheetStore();
    setSelection(single, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setActive(single, { sheet: 0, row: 0, col: 0 });
    toggleBold(single.getState(), single);
    expect(single.getState().ui.pendingFormat?.format.bold).toBe(true);
    expect(single.getState().format.formats.size).toBe(0);

    const multi = createSpreadsheetStore();
    setSelection(multi, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);
    toggleBold(multi.getState(), multi);
    expect(multi.getState().ui.pendingFormat).toBeNull();
    expect(fmtAt(multi, 0, 0)?.bold).toBe(true);
    expect(fmtAt(multi, 0, 1)?.bold).toBe(true);
  });

  it('requires complete merge coverage before changing a union', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, merge);
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toggleBold(store.getState(), store);
    expect(store.getState().format.formats.size).toBe(0);

    setSelection(store, merge, [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }]);
    toggleBold(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.bold).toBe(true);
    }
  });

  it('keeps locked cells unchanged while formatting unlocked cells', () => {
    const store = createSpreadsheetStore();
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(store, 0, true);
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);

    toggleBold(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(fmtAt(store, 0, 1)?.bold).toBe(true);
  });

  it('keeps the legacy locked-cell skip with a registered controller without policy', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
      setCellLocked(store, { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 }, false);
      setProtectedSheet(store, 0, true);
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
      ]);

      toggleBold(store.getState(), store);

      expect(fmtAt(store, 0, 0)).toBeUndefined();
      expect(fmtAt(store, 0, 1)?.bold).toBe(true);
      expect(fmtAt(store, 0, 3)).toBeUndefined();
      expect(fmtAt(store, 0, 4)?.bold).toBe(true);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('rejects a mixed selection atomically under an explicit format policy', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
      ]);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { italic: true });
      const policy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
      });
      controller.setPolicy({
        ...policy,
        operations: { ...policy.operations, format: true },
      });
      const before = new Map(store.getState().format.formats);
      const pendingBefore = store.getState().ui.pendingFormat;

      toggleBold(store.getState(), store);

      expect(store.getState().format.formats).toEqual(before);
      expect(store.getState().ui.pendingFormat).toBe(pendingBefore);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('clears sparse visual formats across a huge primary and extra range', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      {
        bold: true,
        comment: 'keep',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 8, col: 4 },
      {
        fill: '#ffeeaa',
        validation: { kind: 'list', source: ['A'] },
      },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 3 }, { italic: true });
    mutators.setCellFormat(store, { sheet: 1, row: 4, col: 2 }, { fill: '#000000' });
    setSelection(store, { sheet: 0, r0: 0, c0: 2, r1: 1_048_575, c1: 2 }, [
      { sheet: 0, r0: 8, c0: 4, r1: 8, c1: 4 },
    ]);

    clearVisualFormat(store.getState(), store);

    expect(fmtAt(store, 4, 2)).toEqual({ comment: 'keep' });
    expect(fmtAt(store, 8, 4)).toEqual({ validation: { kind: 'list', source: ['A'] } });
    expect(fmtAt(store, 4, 3)?.italic).toBe(true);
    expect(store.getState().format.formats.get('1:4:2')?.fill).toBe('#000000');
  });

  it('does not clear a protected pending format', () => {
    for (const clear of [clearFormat, clearVisualFormat]) {
      const store = createSpreadsheetStore();
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { italic: true });
      toggleBold(store.getState(), store);
      const pendingBefore = store.getState().ui.pendingFormat;
      const formatsBefore = new Map(store.getState().format.formats);

      setProtectedSheet(store, 0, true);
      clear(store.getState(), store);

      expect(store.getState().ui.pendingFormat).toEqual(pendingBefore);
      expect(store.getState().format.formats).toEqual(formatsBefore);
    }
  });

  it('propagates nested command context and restores it after a throw', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    const seen: string[] = [];
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      controller.setPolicy({
        operations: { format: true },
        editable: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
        restrict: ({ intent }) => {
          seen.push(`${intent.origin}:${intent.commandId ?? ''}`);
          return true;
        },
      });

      withSelectionFormatOrigin(
        store,
        'ribbon',
        () => {
          applySelectionFormatPatch(store.getState(), store, { bold: true });
          withSelectionFormatOrigin(store, 'keyboard', () => {
            applySelectionFormatPatch(
              store.getState(),
              store,
              { italic: true },
              { origin: 'instanceApi', commandId: 'italic' },
            );
          });
        },
        'bold',
      );

      expect(() =>
        withSelectionFormatOrigin(
          store,
          'contextMenu',
          () => {
            throw new Error('restore context');
          },
          'underline',
        ),
      ).toThrow('restore context');
      applySelectionFormatPatch(store.getState(), store, { strike: true });

      expect(seen).toEqual(['ribbon:bold', 'instanceApi:italic', 'instanceApi:']);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});
