import { describe, expect, it, vi } from 'vitest';
import { pasteTSV } from '../../../../src/commands/clipboard/paste.js';
import { coerceInput } from '../../../../src/commands/coerce-input.js';
import { History } from '../../../../src/commands/history.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import type { Strings } from '../../../../src/i18n/strings.js';
import { attachClipboard } from '../../../../src/interact/clipboard.js';
import { createContextMenuClipboard } from '../../../../src/interact/context-menu-clipboard.js';
import { runContextMenuInsertCopiedCells } from '../../../../src/interact/context-menu-insert-copied.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

const firePaste = (host: HTMLElement, text: string): ClipboardEvent => {
  const transfer = new DataTransfer();
  transfer.setData('text/plain', text);
  const event = new ClipboardEvent('paste', {
    clipboardData: transfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return event;
};

describe('TSV paste value batching', () => {
  it('publishes one complete cell map and records one undoable mounted paste', async () => {
    const wb = await WorkbookHandle.createDefault();
    if (wb.isStub) throw new Error('native workbook required');
    let mounted: MountedStubSheet | null = null;
    let detachClipboard: (() => void) | null = null;
    try {
      mounted = await mountStubSheet({ workbook: wb, features: { clipboard: false } });
      const { instance } = mounted;
      const { store } = instance;
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 999 });
      const onAfterCommit = vi.fn();
      const clipboard = attachClipboard({
        host: mounted.host,
        history: instance.history,
        store,
        wb,
        onAfterCommit,
      });
      detachClipboard = clipboard.detach;

      let cellMapPublications = 0;
      const unsubscribe = store.subscribe((next, previous) => {
        if (next.data.cells !== previous.data.cells) cellMapPublications += 1;
      });
      const callbackCellMaps: ReadonlyMap<string, unknown>[] = [];
      instance.on('cellChange', () => callbackCellMaps.push(store.getState().data.cells));
      const historyPush = vi.spyOn(instance.history, 'push');
      const values = Array.from({ length: 1000 }, (_, col) => String(col + 1));

      const event = firePaste(mounted.host, values.join('\t'));
      unsubscribe();

      expect(event.defaultPrevented).toBe(true);
      expect(cellMapPublications).toBe(1);
      expect(callbackCellMaps).toHaveLength(1000);
      expect(callbackCellMaps.every((cells) => cells === callbackCellMaps[0])).toBe(true);
      expect(callbackCellMaps[0]?.size).toBe(1000);
      expect(store.getState().data.cells.get('0:0:999')?.value).toEqual({
        kind: 'number',
        value: 1000,
      });
      expect(historyPush).toHaveBeenCalledTimes(1);
      expect(onAfterCommit).toHaveBeenCalledTimes(1);

      expect(instance.history.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 999 })).toEqual({ kind: 'blank' });
      expect(instance.history.redo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 999 })).toEqual({
        kind: 'number',
        value: 1000,
      });
    } finally {
      detachClipboard?.();
      mounted?.dispose();
      wb.dispose();
    }
  });

  it('keeps localized coercion, Text cells, R1C1 formulas, errors, and ragged range semantics', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      mutators.setR1C1(store, true);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 3 }, { numFmt: { kind: 'text' } });
      const tsv = "1,25\t5%\t31.12.2024\t00123\t'0042\t=SUM(RC[-5]:RC[-5])\t#NV\ntail";

      const result = pasteTSV(store.getState(), wb, tsv, history);
      const expectedDate = coerceInput('31.12.2024', { locale: 'de-DE' });

      expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 6 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1.25 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 0.05 });
      expect(expectedDate.kind).toBe('number');
      if (expectedDate.kind !== 'number') throw new Error('expected a typed date');
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
        kind: 'number',
        value: expectedDate.value,
      });
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
        kind: 'text',
        value: '00123',
      });
      expect(wb.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({
        kind: 'text',
        value: '0042',
      });
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 5 })).toBe('=SUM(A1:A1)');
      expect(wb.getValue({ sheet: 0, row: 0, col: 5 })).toEqual({ kind: 'number', value: 1.25 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({
        kind: 'error',
        code: 6,
        text: '#N/A',
      });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'tail' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.has('0:0:1')).toBe(false);
      expect(store.getState().format.formats.has('0:0:2')).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('skips protected cells but preserves the input rectangle in its result', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      setProtectedSheet(store, 0, true);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { locked: false });

      const result = pasteTSV(store.getState(), wb, 'allowed\tlocked', history);

      expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'text',
        value: 'allowed',
      });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
      expect(history.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      wb.dispose();
    }
  });

  it.each([
    { mode: 0 as const, recalcEvents: 1 },
    { mode: 1 as const, recalcEvents: 0 },
    { mode: 2 as const, recalcEvents: 1 },
  ])('preserves calculation mode $mode', async ({ mode, recalcEvents }) => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      expect(wb.setCalcMode(mode)).toBe(true);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      let recalcs = 0;
      const unsubscribe = wb.subscribe((event) => {
        if (event.kind === 'recalc') recalcs += 1;
      });

      const result = pasteTSV(store.getState(), wb, '=1+1\t3', history);
      unsubscribe();

      expect(result).not.toBeNull();
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=1+1');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 3 });
      expect(recalcs).toBe(recalcEvents);
      if (mode !== 1)
        expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 2 });
    } finally {
      wb.dispose();
    }
  });

  it('rolls back a mid-patch native failure without adding history', async () => {
    const wb = await WorkbookHandle.createDefault();
    if (wb.isStub) throw new Error('native workbook required');
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      const addresses = [0, 1, 2].map((col) => ({ sheet: 0, row: 0, col }));
      addresses.forEach((addr, col) => {
        wb.setNumber(addr, 10 + col);
      });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      const raw = (
        wb as unknown as {
          wb: {
            setNumber: (
              sheet: number,
              row: number,
              col: number,
              input: number,
            ) => { ok: boolean; status: number; message: string; context: string };
          };
        }
      ).wb;
      const originalSetNumber = raw.setNumber.bind(raw);
      let setterCalls = 0;
      const failingSetter = vi
        .spyOn(raw, 'setNumber')
        .mockImplementation((sheet, row, col, value) => {
          setterCalls += 1;
          if (setterCalls === 2)
            return { ok: false, status: 1, message: 'injected TSV failure', context: '' };
          return originalSetNumber(sheet, row, col, value);
        });
      const historyPush = vi.spyOn(history, 'push');

      try {
        expect(() => pasteTSV(store.getState(), wb, '1\t2\t3', history)).toThrow(
          'injected TSV failure',
        );
        expect(addresses.map((addr) => wb.getValue(addr))).toEqual([
          { kind: 'number', value: 10 },
          { kind: 'number', value: 11 },
          { kind: 'number', value: 12 },
        ]);
        expect(historyPush).not.toHaveBeenCalled();
        expect(history.canUndo()).toBe(false);
      } finally {
        failingSetter.mockRestore();
      }
    } finally {
      wb.dispose();
    }
  });

  it('routes context-menu TSV paste through one atomic history entry', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      const readText = vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('7\t8');
      const afterCommit = vi.fn();
      const menuClipboard = createContextMenuClipboard({
        store,
        wb,
        history,
        deps: {},
        afterCommit,
        hasExplicitPolicy: () => false,
        canPasteToRange: () => true,
        executeClipboardText: () => {},
      });
      const applyAtomic = vi.spyOn(wb, 'applyCellPatchAtomic');

      expect(menuClipboard.run('paste', store.getState())).toBe(true);
      await Promise.resolve();
      await Promise.resolve();

      expect(readText).toHaveBeenCalledTimes(1);
      expect(applyAtomic).toHaveBeenCalledTimes(1);
      expect(afterCommit).toHaveBeenCalledTimes(1);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 7 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 8 });
      expect(history.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(history.redo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 8 });
      readText.mockRestore();
    } finally {
      wb.dispose();
    }
  });

  it('uses the history-aware paste path after context-menu row insertion', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const host = document.createElement('div');
    try {
      mutators.setActive(store, { sheet: 0, row: 3, col: 0 });
      mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'copy');
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('99');
      const applyAtomic = vi.spyOn(wb, 'applyCellPatchAtomic');
      const afterCommit = vi.fn();

      runContextMenuInsertCopiedCells(
        {
          host,
          store,
          wb,
          history,
          strings: () => ({}) as Strings,
          afterCommit,
          clipboard: { snapshot: () => null },
        },
        'row',
      );
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();

      expect(
        applyAtomic.mock.calls.some(([patches]) =>
          patches.some(
            (patch) => patch.addr.sheet === 0 && patch.addr.row === 3 && patch.addr.col === 0,
          ),
        ),
      ).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'number', value: 99 });
      expect(history.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'blank' });
      expect(history.redo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'number', value: 99 });
      expect(afterCommit).toHaveBeenCalledTimes(1);
    } finally {
      host.remove();
      wb.dispose();
    }
  });

  it('keeps the public three-argument helper on WorkbookHandle local undo', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    try {
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      expect(pasteTSV(store.getState(), wb, '11\t22')).not.toBeNull();
      expect(wb.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 11 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    } finally {
      wb.dispose();
    }
  });
});
