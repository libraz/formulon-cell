// @vitest-environment node

import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../../src/commands/built-in-command-policy.js';
import { pasteTSV } from '../../../../src/commands/clipboard/paste.js';
import {
  insertCopiedCellsAction,
  pasteSpecialAction,
  pasteWrite,
} from '../../../../src/commands/clipboard/paste-actions.js';
import type { PasteSpecialOptions } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { formatAsTable } from '../../../../src/commands/format-as-table.js';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import type {
  InteractionPolicy,
  OperationIntent,
} from '../../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { applyRangeDrop } from '../../../../src/commands/range-drop.js';
import { MAX_COL } from '../../../../src/engine/address.js';
import type { Addr, Range } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

const ALL: PasteSpecialOptions = {
  what: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
};
const range = (r0: number, c0: number, r1 = r0, c1 = c0): Range => ({ sheet: 0, r0, c0, r1, c1 });

describe('paste writes under the interaction policy', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, workbook.cells(0));
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      workbook.dispose();
    });
    const seen = new Set<string>();
    const intents: OperationIntent[] = [];
    const record = (deny?: (intent: OperationIntent) => boolean): InteractionPolicy => ({
      defaultOperation: 'allow',
      restrict: ({ intent }) => {
        seen.add(`${intent.operation}:${intent.origin}:${intent.commandId}`);
        intents.push(intent);
        return !deny?.(intent);
      },
    });
    const at = (addr: Addr): void => {
      mutators.setActive(store, addr);
      mutators.setRange(store, range(addr.row, addr.col));
    };
    const value = (row: number, col: number) => workbook.getValue({ sheet: 0, row, col });
    return { store, workbook, history, controller, seen, intents, record, at, value };
  };

  it('authorizes a Paste Special under its entry point and replays it under the same check', async () => {
    const { store, workbook, history, controller, seen, record, at, value } = await setup();
    controller.setPolicy(record());
    const snap = captureSnapshot(store.getState(), range(0, 0, 1, 0));
    if (!snap) throw new Error('expected snapshot');
    at({ sheet: 0, row: 0, col: 2 });
    const ctx = { store, history, origin: 'contextMenu', commandId: 'pasteAll' } as const;
    expect(pasteSpecialAction(ctx, workbook, snap, ALL)).not.toBeNull();
    expect(value(1, 2)).toEqual({ kind: 'number', value: 2 });
    expect(seen.has('paste:contextMenu:pasteAll')).toBe(true);
    history.undo();
    history.redo();
    expect([...seen].sort()).toEqual([
      'paste:contextMenu:pasteAll',
      'paste:redo:pasteAll',
      'paste:undo:pasteAll',
    ]);
  });

  it('refuses a paste its entry point may not run and leaves the destination alone', async () => {
    const { store, workbook, history, controller, record, at, value } = await setup();
    controller.setPolicy(record((intent) => intent.origin === 'ribbon'));
    const snap = captureSnapshot(store.getState(), range(0, 0));
    if (!snap) throw new Error('expected snapshot');
    at({ sheet: 0, row: 0, col: 2 });
    const ctx = { store, history, origin: 'ribbon', commandId: 'pasteValues' } as const;
    expect(pasteSpecialAction(ctx, workbook, snap, { ...ALL, what: 'values' })).toBeNull();
    expect(value(0, 2)).toEqual({ kind: 'blank' });
    expect(history.canUndo()).toBe(false);
  });

  it('declares the cut source beside the destination', async () => {
    const { store, workbook, history, controller, intents, record, at } = await setup();
    controller.setPolicy(record());
    const snap = captureSnapshot(store.getState(), range(0, 0, 1, 0), 'cut');
    if (!snap) throw new Error('expected snapshot');
    at({ sheet: 0, row: 0, col: 3 });
    const ctx = { store, history, origin: 'keyboard', commandId: 'paste' } as const;
    pasteSpecialAction(ctx, workbook, snap, ALL);
    const paste = intents.find((intent) => intent.origin === 'keyboard');
    expect(paste?.effects).toEqual([
      { kind: 'range', range: range(0, 3, 1, 3) },
      { kind: 'range', range: range(0, 0, 1, 0) },
    ]);
  });

  it('refuses an atomic snapshot paste over a locked cell of a protected sheet', async () => {
    const { store, workbook, history, controller, record, at, value } = await setup();
    controller.setPolicy(record());
    const snap = captureSnapshot(store.getState(), range(0, 0, 1, 0));
    if (!snap) throw new Error('expected snapshot');
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 2 }, { locked: false });
    setProtectedSheet(store, 0, true);
    at({ sheet: 0, row: 0, col: 2 });
    const ctx = { store, history, origin: 'ribbon', commandId: 'pasteAll' } as const;
    expect(pasteSpecialAction(ctx, workbook, snap, ALL)).toBeNull();
    expect(value(1, 2)).toEqual({ kind: 'blank' });
    expect(history.canUndo()).toBe(false);
  });

  it('refuses a text paste whose destination touches a locked cell, with or without a policy', async () => {
    const { store, workbook, history, controller, record, at, value } = await setup();
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 2 }, { locked: false });
    setProtectedSheet(store, 0, true);
    at({ sheet: 0, row: 0, col: 2 });
    expect(pasteTSV(store.getState(), store, workbook, '7\n8', history)).toBeNull();
    controller.setPolicy(record());
    const ctx = { store, history, origin: 'clipboard', commandId: 'paste' } as const;
    expect(
      pasteWrite(ctx, { destination: range(0, 2, 1, 2) }, null, () =>
        pasteTSV(store.getState(), store, workbook, '7\n8', history),
      ),
    ).toBeNull();
    expect(value(1, 2)).toEqual({ kind: 'blank' });
    expect(history.canUndo()).toBe(false);
  });

  it('declares the table a paste grows, so a policy denying table changes refuses it', async () => {
    const { store, workbook, history, controller, seen, record, at, value } = await setup();
    expect(formatAsTable(store, range(0, 0, 1, 0))).not.toBeNull();
    controller.setPolicy(record((intent) => intent.operation === 'table'));
    at({ sheet: 0, row: 2, col: 0 });
    const ctx = { store, history, origin: 'keyboard', commandId: 'paste' } as const;
    expect(
      pasteWrite(ctx, { destination: range(2, 0) }, null, () =>
        pasteTSV(store.getState(), store, workbook, '9', history),
      ),
    ).toBeNull();
    expect(seen.has('table:keyboard:paste')).toBe(true);
    expect(value(2, 0)).toEqual({ kind: 'blank' });
  });

  it('decides Paste Column Widths as a column layout change', async () => {
    const { store, workbook, history, controller, seen, record, at } = await setup();
    controller.setPolicy(record());
    const snap = captureSnapshot(store.getState(), range(0, 0));
    if (!snap) throw new Error('expected snapshot');
    at({ sheet: 0, row: 0, col: 2 });
    const ctx = { store, history, origin: 'contextMenu', commandId: 'pasteColumnWidths' } as const;
    pasteSpecialAction(ctx, workbook, snap, { ...ALL, what: 'column-widths' });
    expect([...seen].sort()).toEqual([
      'resizeColumns:contextMenu:pasteColumnWidths',
      'resizeColumns:redo:pasteColumnWidths',
      'resizeColumns:undo:pasteColumnWidths',
    ]);
  });

  it('decides Insert Copied Cells as a cell insert plus the paste', async () => {
    const { store, workbook, history, controller, seen, record, at, value } = await setup();
    controller.setPolicy(record((intent) => intent.operation === 'insertRows'));
    const snap = captureSnapshot(store.getState(), range(0, 0, 1, 0));
    if (!snap) throw new Error('expected snapshot');
    at({ sheet: 0, row: 0, col: 2 });
    const ctx = { store, history, origin: 'contextMenu', commandId: 'insertCopiedCells' } as const;
    expect(insertCopiedCellsAction(ctx, workbook, '', 'down', snap)).toBeNull();
    expect(value(0, 2)).toEqual({ kind: 'blank' });
    expect(seen.has('insertRows:contextMenu:insertCopiedCells')).toBe(true);

    controller.setPolicy(record());
    expect(insertCopiedCellsAction(ctx, workbook, '', 'down', snap)).not.toBeNull();
    expect(value(1, 2)).toEqual({ kind: 'number', value: 2 });
    expect(seen.has('paste:contextMenu:insertCopiedCells')).toBe(true);
  });

  it('projects every paste command as the paste its write declares', async () => {
    const { store, controller, record } = await setup();
    controller.setPolicy(record((intent) => intent.operation === 'paste'));
    for (const id of [
      'paste',
      'pasteAll',
      'pasteValues',
      'pasteFormatsOnly',
      'pasteFormulas',
      'pasteFormulasNumFmt',
      'pasteLink',
      'cut',
    ] as const)
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    controller.setPolicy(record((intent) => intent.operation === 'format'));
    expect(canExecuteBuiltIn(store, 'pasteFormatsOnly', 'ribbon').allowed).toBe(true);
  });

  it('projects Insert Copied Cells from the insertion and the paste it declares', async () => {
    const { store, controller, record, at } = await setup();
    at({ sheet: 0, row: 3, col: 3 });
    mutators.setCopyRange(store, range(0, 0, 1, 0));
    controller.setPolicy(record((intent) => intent.operation === 'insertRows'));
    // The dialog still offers shifting right.
    expect(canExecuteBuiltIn(store, 'insertCopiedCells', 'contextMenu').allowed).toBe(true);
    controller.setPolicy(
      record((intent) => intent.operation === 'insertRows' || intent.operation === 'insertColumns'),
    );
    expect(canExecuteBuiltIn(store, 'insertCopiedCells', 'contextMenu').allowed).toBe(false);

    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: MAX_COL });
    controller.setPolicy(record((intent) => intent.operation === 'insertRows'));
    expect(canExecuteBuiltIn(store, 'insertCopiedCells', 'ribbon').allowed).toBe(false);
    controller.setPolicy(record((intent) => intent.operation === 'paste'));
    expect(canExecuteBuiltIn(store, 'insertCopiedCells', 'ribbon').allowed).toBe(false);
    controller.setPolicy(record());
    expect(canExecuteBuiltIn(store, 'insertCopiedCells', 'ribbon').allowed).toBe(true);
  });

  it('clears the source of a shifting drop move under the drop gesture', async () => {
    const { store, workbook, history, controller, seen, record } = await setup();
    controller.setPolicy(record());
    const moved = await applyRangeDrop(
      { store, wb: workbook, history },
      range(0, 0, 1, 0),
      range(0, 3, 1, 3),
      'moveShiftDown',
    );
    expect(moved).toBe(true);
    expect([...seen].filter((key) => key.endsWith(':editor:moveShiftDown')).sort()).toEqual([
      'clear:editor:moveShiftDown',
      'insertRows:editor:moveShiftDown',
      'paste:editor:moveShiftDown',
    ]);
    expect([...seen].some((key) => key.includes(':ribbon:'))).toBe(false);
  });
});
