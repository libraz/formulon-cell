import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import type { ClipboardSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../src/commands/history.js';
import { addrKey, type WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { fireContextMenu, item, newWb, seed, setRange } from './fixtures.js';

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let onAfterCommit: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    detach = null;
    unregisterController = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    unregisterController?.();
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  describe('row structure', () => {
    it('Insert Copied Cells opens whole rows and drops the copied band into them', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [[{ value: { kind: 'text', value: 'a' }, formula: null, format: undefined }]],
      };
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 1, col: 0, value: 'b' },
      ]);
      mutators.selectRow(store, 0);
      mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 16383 });
      mutators.selectRow(store, 1);
      detach = attachContextMenu({
        host,
        store,
        wb,
        strings: en,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 10, 59); // row 1 header
      // The header variant needs no direction prompt, so it carries no ellipsis.
      expect(item('insertCopiedCells')?.textContent).toBe(en.contextMenu.insertCopiedBand);
      // A pending copy replaces the plain insert entries rather than joining them.
      expect(item('rowInsertAbove')).toBeNull();
      expect(item('rowInsertBelow')).toBeNull();
      item('insertCopiedCells')?.click();

      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'b' });
      // The inserted row stays selected and the marquee stays up for a repeat.
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 1,
        c0: 0,
        r1: 1,
        c1: 16383,
      });
      expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 16383 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('inserts a whole-row cut at the header target and preserves surrounding rows', () => {
      const history = new History();
      seed(store, wb, [
        { row: 1, col: 0, value: 'target' },
        { row: 3, col: 0, value: 'source' },
      ]);
      wb.setFormula({ sheet: 0, row: 3, col: 1 }, '=A4');
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));

      mutators.selectRow(store, 3);
      detach = attachContextMenu({
        host,
        store,
        wb,
        strings: en,
        history,
        onAfterCommit,
      });
      fireContextMenu(host, 10, 90); // row 3 header
      item('cut')?.click();

      mutators.selectRow(store, 1);
      const selectionBefore = store.getState().selection;
      const sourceValueBefore = wb.getValue({ sheet: 0, row: 3, col: 0 });
      const targetValueBefore = wb.getValue({ sheet: 0, row: 1, col: 0 });
      const sourceFormulaBefore = wb.cellFormula({ sheet: 0, row: 3, col: 1 });

      fireContextMenu(host, 10, 59); // row 1 header
      const insertItem = item('insertCopiedCells');
      expect(insertItem?.disabled).toBe(false);
      expect(insertItem?.textContent).toBe(en.contextMenu.insertCutCells);
      insertItem?.click();
      wb.recalc();

      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual(sourceValueBefore);
      expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual(targetValueBefore);
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=A2');
      expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toBeNull();
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 1,
        c0: 0,
        r1: 1,
        c1: 16_383,
      });
      expect(store.getState().ui.copyRange).toBeNull();
      expect(store.getState().ui.copyMode).toBeNull();
      expect(history.canUndo()).toBe(true);
      expect(onAfterCommit).toHaveBeenCalledTimes(1);

      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual(targetValueBefore);
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual(sourceValueBefore);
      expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toBe(sourceFormulaBefore);
      expect(store.getState().selection).toEqual(selectionBefore);
      expect(store.getState().ui.copyRange).toBeNull();
      expect(store.getState().ui.copyMode).toBeNull();
    });

    it('shows the plain row insert entries outside copy mode', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 30);
      expect(item('insertCopiedCells')).toBeNull();
      expect(item('rowInsertAbove')).not.toBeNull();
      expect(item('rowInsertBelow')).not.toBeNull();
    });

    it('Insert Above shifts existing rows down', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'a' }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 30); // row 0 header
      item('rowInsertAbove')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('Insert Below shifts subsequent rows down', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 1, col: 0, value: 'b' },
      ]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 30);
      item('rowInsertBelow')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Delete row shifts subsequent rows up', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 1, col: 0, value: 'b' },
      ]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 30);
      item('rowDelete')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Hide Row records the row as hidden, Unhide restores it', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 10, 30);
      item('rowHide')?.click();
      expect(store.getState().layout.hiddenRows.has(0)).toBe(true);
      // Span rows 0-1 so the next contextmenu's promotion-check sees the row
      // band as already-selected. Row 0 is hidden, so y=30 hits row 1 header,
      // but inSel=true keeps the (0..1) band intact.
      setRange(store, 0, 0, 1, 16383);
      fireContextMenu(host, 10, 30);
      item('rowUnhide')?.click();
      expect(store.getState().layout.hiddenRows.has(0)).toBe(false);
    });

    it('Unhide is a no-op when no hidden rows are in the selection', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 10, 30);
      item('rowUnhide')?.click();
      expect(store.getState().layout.hiddenRows.size).toBe(0);
    });
    it('built-in Group / Ungroup Row items set and clear the row outline', () => {
      setRange(store, 0, 0, 1, 16383);
      detach = attachContextMenu({
        host,
        store,
        wb,
        options: {
          mode: 'builtIn',
          transform: (context) => [
            ...context.defaultItems,
            { id: 'rowGroup', label: 'Group', builtIn: 'rowGroup' },
            { id: 'rowUngroup', label: 'Ungroup', builtIn: 'rowUngroup' },
          ],
        },
      });
      fireContextMenu(host, 10, 30);
      item('rowGroup')?.click();
      expect(store.getState().layout.outlineRows.get(0)).toBe(1);
      expect(store.getState().layout.outlineRows.get(1)).toBe(1);
      fireContextMenu(host, 10, 30);
      item('rowUngroup')?.click();
      expect(store.getState().layout.outlineRows.size).toBe(0);
    });
  });

  describe('col structure', () => {
    it('inserts a whole-column cut at the header target and preserves surrounding columns', () => {
      const history = new History();
      seed(store, wb, [
        { row: 0, col: 1, value: 'target' },
        { row: 0, col: 3, value: 'source' },
      ]);
      wb.setFormula({ sheet: 0, row: 1, col: 3 }, '=D1');
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));

      mutators.selectCol(store, 3);
      detach = attachContextMenu({
        host,
        store,
        wb,
        strings: en,
        history,
        onAfterCommit,
      });
      fireContextMenu(host, 252, 10); // col 3 header
      item('cut')?.click();

      mutators.selectCol(store, 1);
      const selectionBefore = store.getState().selection;
      const sourceValueBefore = wb.getValue({ sheet: 0, row: 0, col: 3 });
      const targetValueBefore = wb.getValue({ sheet: 0, row: 0, col: 1 });
      const sourceFormulaBefore = wb.cellFormula({ sheet: 0, row: 1, col: 3 });

      fireContextMenu(host, 124, 10); // col 1 header

      const insertItem = item('insertCopiedCells');
      expect(insertItem?.disabled).toBe(false);
      expect(insertItem?.textContent).toBe(en.contextMenu.insertCutCells);
      insertItem?.click();
      wb.recalc();

      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(sourceValueBefore);
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual(targetValueBefore);
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'blank' });
      expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=B1');
      expect(wb.cellFormula({ sheet: 0, row: 1, col: 3 })).toBeNull();
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 0,
        c0: 1,
        r1: 1_048_575,
        c1: 1,
      });
      expect(store.getState().ui.copyRange).toBeNull();
      expect(store.getState().ui.copyMode).toBeNull();
      expect(history.canUndo()).toBe(true);
      expect(onAfterCommit).toHaveBeenCalledTimes(1);

      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(targetValueBefore);
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual(sourceValueBefore);
      expect(wb.cellFormula({ sheet: 0, row: 1, col: 3 })).toBe(sourceFormulaBefore);
      expect(store.getState().selection).toEqual(selectionBefore);
      expect(store.getState().ui.copyRange).toBeNull();
      expect(store.getState().ui.copyRanges).toBeNull();
      expect(store.getState().ui.copyMode).toBeNull();
    });

    it('Insert Copied Cells opens whole columns and drops the copied band into them', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [[{ value: { kind: 'text', value: 'a' }, formula: null, format: { bold: true } }]],
      };
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 0, col: 1, value: 'b' },
      ]);
      mutators.selectCol(store, 0);
      mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
      mutators.selectCol(store, 1);
      detach = attachContextMenu({
        host,
        store,
        wb,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 124, 10); // col 1 header
      item('insertCopiedCells')?.click();

      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'text', value: 'b' });
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 })),
      ).toMatchObject({ bold: true });
      // The inserted column stays selected and the marquee stays on the source.
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 0,
        c0: 1,
        r1: 1048575,
        c1: 1,
      });
      expect(store.getState().ui.copyRange).toEqual({
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 1048575,
        c1: 0,
      });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('Insert Left shifts subsequent cols right', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'a' }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 60, 10);
      item('colInsertLeft')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'a' });
    });

    it('Insert Right shifts the col after the selection right', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 0, col: 1, value: 'b' },
      ]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 60, 10);
      item('colInsertRight')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Delete col shifts subsequent cols left', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 0, col: 1, value: 'b' },
      ]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 60, 10);
      item('colDelete')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Hide / Unhide Col toggles layout.hiddenCols', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 60, 10);
      item('colHide')?.click();
      expect(store.getState().layout.hiddenCols.has(0)).toBe(true);
      // Span cols 0-1 so the next contextmenu's promotion-check sees the col
      // band as already-selected. Col 0 is hidden, so x=60 hits col 1 header,
      // but inSel=true keeps the (0..1) band intact.
      setRange(store, 0, 0, 1048575, 1);
      fireContextMenu(host, 60, 10);
      item('colUnhide')?.click();
      expect(store.getState().layout.hiddenCols.has(0)).toBe(false);
    });

    it('built-in Group / Ungroup Col items set and clear the column outline', () => {
      setRange(store, 0, 0, 1048575, 1);
      detach = attachContextMenu({
        host,
        store,
        wb,
        options: {
          mode: 'builtIn',
          transform: (context) => [
            ...context.defaultItems,
            { id: 'colGroup', label: 'Group', builtIn: 'colGroup' },
            { id: 'colUngroup', label: 'Ungroup', builtIn: 'colUngroup' },
          ],
        },
      });
      fireContextMenu(host, 60, 10);
      item('colGroup')?.click();
      expect(store.getState().layout.outlineCols.get(0)).toBe(1);
      expect(store.getState().layout.outlineCols.get(1)).toBe(1);
      fireContextMenu(host, 60, 10);
      item('colUngroup')?.click();
      expect(store.getState().layout.outlineCols.size).toBe(0);
    });
  });
});
