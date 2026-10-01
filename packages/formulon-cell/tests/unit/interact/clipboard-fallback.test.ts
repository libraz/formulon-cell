import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { History } from '../../../src/commands/history.js';
import { insertCols, insertRows } from '../../../src/commands/structure.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard, type ClipboardHandle } from '../../../src/interact/clipboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('shortcut paste when system clipboard access is unavailable', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let handle: ClipboardHandle;
  let host: HTMLElement;
  let history: History;
  let readText: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    store = createSpreadsheetStore();
    host = document.createElement('div');
    history = new History();
    readText = vi.fn().mockRejectedValue(new DOMException('Denied', 'NotAllowedError'));
    vi.stubGlobal('navigator', {
      clipboard: { readText, writeText: vi.fn().mockResolvedValue(undefined) },
    });
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Copied');
    mutators.replaceCells(store, wb.cells(0));
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { bold: true, fontFamily: 'Arial' },
    );
    handle = attachClipboard({
      host,
      wb,
      store,
      history,
      onAfterCommit: () => mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex)),
    });
  });

  afterEach(() => {
    handle.detach();
    wb.dispose();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  const valueAt = (row: number) => wb.getValue({ sheet: 0, row, col: 0 });
  const selectRow = (row: number) => mutators.setActive(store, { sheet: 0, row, col: 0 });

  it('pastes the active internal copy with its formats and supports undo and repeat paste', async () => {
    await handle.runShortcut('copy');
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(1)).toEqual({ kind: 'text', value: 'Copied' });
    expect(store.getState().format.formats.get('0:1:0')).toEqual({
      bold: true,
      fontFamily: 'Arial',
    });
    expect(history.undo()).toBe(true);
    expect(valueAt(1)).toEqual({ kind: 'blank' });
    expect(store.getState().format.formats.has('0:1:0')).toBe(false);
    expect(history.redo()).toBe(true);
    expect(valueAt(1)).toEqual({ kind: 'text', value: 'Copied' });
    selectRow(2);
    await handle.runShortcut('paste');
    expect(valueAt(2)).toEqual({ kind: 'text', value: 'Copied' });
  });

  it('uses the internal copy when the Clipboard API is missing', async () => {
    vi.stubGlobal('navigator', {});
    await handle.runShortcut('copy');
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(1)).toEqual({ kind: 'text', value: 'Copied' });
  });

  it('does not resurrect an internal copy canceled with Escape', async () => {
    await handle.runShortcut('copy');
    mutators.setCopyRange(store, null);
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(1)).toEqual({ kind: 'blank' });
  });

  it('consumes an internal cut once even when system clipboard reads fail', async () => {
    await handle.runShortcut('cut');
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(0)).toEqual({ kind: 'blank' });
    expect(valueAt(1)).toEqual({ kind: 'text', value: 'Copied' });
    expect(store.getState().ui.copyMode).toBeNull();
    selectRow(2);
    await handle.runShortcut('paste');
    expect(valueAt(2)).toEqual({ kind: 'blank' });
  });

  it('uses newly copied external text when a system clipboard read succeeds', async () => {
    await handle.runShortcut('copy');
    readText.mockResolvedValue('External');
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(1)).toEqual({ kind: 'text', value: 'External' });
    expect(store.getState().format.formats.has('0:1:0')).toBe(false);
    expect(handle.getSnapshot()).toBeNull();
  });

  it('does not replace an empty successful system read with an older internal copy', async () => {
    await handle.runShortcut('copy');
    readText.mockResolvedValue('');
    selectRow(1);
    await handle.runShortcut('paste');
    expect(valueAt(1)).toEqual({ kind: 'blank' });
  });

  it('pastes at the current destination if selection changes while reading the clipboard', async () => {
    await handle.runShortcut('copy');
    let finishRead: ((text: string) => void) | undefined;
    readText.mockImplementation(
      () =>
        new Promise<string>((resolve) => {
          finishRead = resolve;
        }),
    );
    selectRow(1);
    const pendingPaste = handle.runShortcut('paste');
    selectRow(2);
    finishRead?.('Copied');
    await pendingPaste;
    expect(valueAt(1)).toEqual({ kind: 'blank' });
    expect(valueAt(2)).toEqual({ kind: 'text', value: 'Copied' });
  });

  it('rechecks copy cancellation after a pending system read is rejected', async () => {
    await handle.runShortcut('copy');
    let failRead: ((reason: Error) => void) | undefined;
    readText.mockImplementation(
      () =>
        new Promise<string>((_resolve, reject) => {
          failRead = reject;
        }),
    );
    selectRow(1);
    const pendingPaste = handle.runShortcut('paste');
    mutators.setCopyRange(store, null);
    failRead?.(new DOMException('Denied', 'NotAllowedError'));
    await pendingPaste;
    expect(valueAt(1)).toEqual({ kind: 'blank' });
  });

  it('aborts a pending paste when a newer copy starts before the read settles', async () => {
    await handle.runShortcut('copy');
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'Replacement');
    mutators.replaceCells(store, wb.cells(0));

    let rejectRead: ((reason: unknown) => void) | undefined;
    readText.mockImplementation(
      () =>
        new Promise<string>((_resolve, reject) => {
          rejectRead = reject;
        }),
    );
    selectRow(1);
    const pendingPaste = handle.runShortcut('paste');
    selectRow(2);
    await handle.runShortcut('copy');
    selectRow(3);

    rejectRead?.(new DOMException('Denied', 'NotAllowedError'));
    await pendingPaste;

    expect(valueAt(3)).toEqual({ kind: 'blank' });
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 2,
      c0: 0,
      r1: 2,
      c1: 0,
    });
  });

  it('does not expose a stale structured snapshot after the marquee is replaced', async () => {
    await handle.runShortcut('copy');
    expect(handle.getSnapshot()).not.toBeNull();

    mutators.setCopyRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });

    expect(handle.getSnapshot()).toBeNull();
  });

  it('does not resurrect a stale payload after the public marquee is replaced', async () => {
    await handle.runShortcut('copy');
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    selectRow(1);

    await handle.runShortcut('paste');

    expect(valueAt(1)).toEqual({ kind: 'blank' });
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 1,
      r1: 0,
      c1: 1,
    });
  });

  it('keeps the active payload through row and column structure shifts', async () => {
    await handle.runShortcut('copy');
    const revision = store.getState().ui.copyRevision;

    insertRows(store, wb, null, 0, 1);
    insertCols(store, wb, null, 0, 1);
    mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
    mutators.setRange(store, { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 });

    await handle.runShortcut('paste');

    expect(valueAt(3)).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({
      kind: 'text',
      value: 'Copied',
    });
    expect(store.getState().ui.copyRevision).toBe(revision);
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 1,
      c0: 1,
      r1: 1,
      c1: 1,
    });
  });
});
