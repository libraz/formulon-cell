import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import type { ClipboardSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import type { CellBatchCommand } from '../../../../src/commands/interaction-policy.js';
import { fixedFormPolicy } from '../../../../src/commands/interaction-policy.js';
import { insertRows } from '../../../../src/commands/structure.js';
import type { Range } from '../../../../src/engine/types.js';
import { addrKey, type WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import type { ContextMenuInteractionController } from '../../../../src/interact/context-menu-options.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import {
  fireContextMenu,
  item,
  newWb,
  root,
  seed,
  setFormat,
  setRange,
  setSelectionRanges,
  visibleMenu,
} from './fixtures.js';

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let onAfterCommit: Mock<() => void>;
  let onPasteSpecial: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    onPasteSpecial = vi.fn<() => void>();
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

  describe('clipboard items', () => {
    it('Copy writes TSV to navigator.clipboard', () => {
      const writeText = vi.spyOn(navigator.clipboard, 'writeText').mockResolvedValue();
      seed(store, wb, [{ row: 0, col: 0, value: 'X' }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('copy')?.click();
      expect(writeText).toHaveBeenCalledWith('X');
      expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      expect(visibleMenu()).toBeNull();
    });

    it('Copy clears a stale copy marquee when the selected range cannot be copied', () => {
      mutators.setCopyRanges(store, [
        { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 16383 },
        { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16383 },
      ]);
      setRange(store, 0, 0, 1048575, 16383);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('copy')?.click();

      expect(store.getState().ui.copyRange).toBeNull();
      expect(store.getState().ui.copyRanges).toBeNull();
    });

    it('Cut writes TSV and keeps source contents until paste', () => {
      const writeText = vi.spyOn(navigator.clipboard, 'writeText').mockResolvedValue();
      seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('cut')?.click();
      expect(writeText).toHaveBeenCalled();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(store.getState().ui.copyMode).toBe('cut');
      expect(onAfterCommit).not.toHaveBeenCalled();
    });

    it('Paste reads from navigator.clipboard and writes via pasteTSV', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('foo\t42');
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      // readClipboard is async; wait for the microtask chain.
      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'foo' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'number', value: 42 });
      expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('fallback cut moves merged cells and formats with one undo step', async () => {
      vi.spyOn(navigator.clipboard, 'writeText').mockResolvedValue();
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const history = new History();
      seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
      setFormat(store, 0, 0, { bold: true, fill: '#ffee00' });
      const source = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
      mutators.mergeRange(store, source);
      setRange(store, 0, 0, 1, 1);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('cut')?.click();
      expect(history.canUndo()).toBe(false);
      expect(store.getState().merges.byAnchor.get('0:0:0')).toEqual(source);

      setRange(store, 3, 3, 3, 3);
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 5 });
      expect(store.getState().merges.byAnchor.get('0:3:3')).toEqual({
        sheet: 0,
        r0: 3,
        c0: 3,
        r1: 4,
        c1: 4,
      });
      expect(store.getState().format.formats.get('0:3:3')).toEqual({ bold: true, fill: '#ffee00' });
      expect(store.getState().ui.copyRange).toBeNull();

      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(wb.getValue({ sheet: 0, row: 3, col: 3 }).kind).toBe('blank');
      expect(store.getState().merges.byAnchor.get('0:0:0')).toEqual(source);
      expect(store.getState().merges.byAnchor.has('0:3:3')).toBe(false);
      expect(store.getState().format.formats.get('0:0:0')).toEqual({ bold: true, fill: '#ffee00' });
      expect(store.getState().format.formats.has('0:3:3')).toBe(false);
      expect(history.canUndo()).toBe(false);
      expect(history.redo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 5 });
      expect(store.getState().merges.byAnchor.has('0:3:3')).toBe(true);
    });

    it('fallback copy follows structural shifts before pasting absolute references', async () => {
      vi.spyOn(navigator.clipboard, 'writeText').mockResolvedValue();
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      seed(store, wb, [{ row: 1, col: 1, value: 7 }]);
      wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=B$2');
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));
      setFormat(store, 1, 0, { bold: true });
      setRange(store, 1, 0, 1, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('copy')?.click();
      insertRows(store, wb, null, 0);
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));
      setRange(store, 4, 3, 4, 3);
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();
      expect(wb.cellFormula({ sheet: 0, row: 4, col: 3 })).toBe('=E$3');
      expect(store.getState().format.formats.get('0:4:3')).toMatchObject({ bold: true });
    });

    it('Paste bundles multi-cell writes into one undo step when history is attached', async () => {
      const history = new History();
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('foo\t42\r\nbar\t99');
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();

      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'foo' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 99 });

      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 1, col: 2 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 2, col: 2 }).kind).toBe('blank');
      expect(history.canUndo()).toBe(false);
    });

    it('Paste resolves to empty string and is a no-op', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();
      expect(onAfterCommit).not.toHaveBeenCalled();
    });

    it('Paste uses the internal snapshot when clipboard text is empty', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [
          [
            {
              value: { kind: 'blank' },
              formula: null,
              format: { hyperlink: 'https://example.test', bold: true },
            },
          ],
        ],
      };
      setRange(store, 4, 4, 4, 4);
      detach = attachContextMenu({
        host,
        store,
        wb,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();

      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 4 })),
      ).toMatchObject({
        hyperlink: 'https://example.test',
        bold: true,
      });
      expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('rejects a repeated internal paste outside the designated policy range', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [[{ value: { kind: 'text', value: 'copied' }, formula: null, format: undefined }]],
      };
      const history = new History();
      const controller = new InteractionController({
        store,
        history,
        getWb: () => wb,
      });
      controller.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }] }));
      unregisterController = registerInteractionController(store, controller);
      setRange(store, 0, 0, 0, 2);
      detach = attachContextMenu({
        host,
        store,
        wb,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
      expect(history.canUndo()).toBe(false);
      expect(onAfterCommit).not.toHaveBeenCalled();
    });

    it('preflights the full repeated range for Paste Special quick actions', () => {
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [[{ value: { kind: 'text', value: 'copied' }, formula: null, format: undefined }]],
      };
      const history = new History();
      const controller = new InteractionController({
        store,
        history,
        getWb: () => wb,
      });
      controller.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }] }));
      unregisterController = registerInteractionController(store, controller);
      setRange(store, 0, 0, 0, 2);
      detach = attachContextMenu({
        host,
        store,
        wb,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 200, 70);
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="pasteSpecialMenu"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));
      document
        .querySelector<HTMLButtonElement>('.fc-ctxmenu__sub [data-fc-action="pasteAll"]')
        ?.click();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
      expect(history.canUndo()).toBe(false);
      expect(onAfterCommit).not.toHaveBeenCalled();
    });

    it('Paste Special triggers the onPasteSpecial callback', () => {
      detach = attachContextMenu({ host, store, wb, onPasteSpecial });
      fireContextMenu(host, 200, 70);
      // Paste Special is a submenu — open it, then click the dialog entry.
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="pasteSpecialMenu"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));
      document
        .querySelector<HTMLButtonElement>('.fc-ctxmenu__sub [data-fc-action="pasteSpecial"]')
        ?.click();
      expect(onPasteSpecial).toHaveBeenCalled();
    });

    it('projects a shared disabled reason on Paste Special quick actions without a copied cell snapshot', () => {
      detach = attachContextMenu({
        host,
        store,
        wb,
        strings: en,
        getClipboardSnapshot: () => null,
      });
      fireContextMenu(host, 200, 70);
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="pasteSpecialMenu"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));

      const pasteValues = document.querySelector<HTMLButtonElement>(
        '.fc-ctxmenu__sub [data-fc-action="pasteValues"]',
      );
      expect(pasteValues?.disabled).toBe(true);
      expect(pasteValues?.getAttribute('aria-disabled')).toBe('true');
      expect(pasteValues?.dataset.disabledReason).toBe(
        'Copy or cut cells before using this paste option.',
      );
      expect(pasteValues?.getAttribute('aria-description')).toBe(
        'Copy or cut cells before using this paste option.',
      );
    });

    it('delegates Copy/Cut/Paste to the shared clipboard path when wired', () => {
      const onClipboardShortcut = vi.fn();
      detach = attachContextMenu({ host, store, wb, onAfterCommit, onClipboardShortcut });

      fireContextMenu(host, 200, 70);
      item('copy')?.click();
      fireContextMenu(host, 200, 70);
      item('cut')?.click();
      fireContextMenu(host, 200, 70);
      item('paste')?.click();

      expect(onClipboardShortcut).toHaveBeenNthCalledWith(1, 'copy');
      expect(onClipboardShortcut).toHaveBeenNthCalledWith(2, 'cut');
      expect(onClipboardShortcut).toHaveBeenNthCalledWith(3, 'paste');
      expect(onClipboardShortcut).toHaveBeenCalledTimes(3);
    });

    it('shows Insert Copied Cells only during copy mode and shifts cells after OK', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('new');
      seed(store, wb, [{ row: 1, col: 1, value: 'old' }]);
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);
      expect(item('insertCopiedCells')).toBeNull();
      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));

      mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      fireContextMenu(host, 200, 70);
      item('insertCopiedCells')?.click();
      expect(document.querySelector('.fc-insertcopied')).not.toBeNull();
      document.querySelector<HTMLButtonElement>('.fc-insertcopied__button--primary')?.click();

      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'new' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'old' });
      // The marquee outlives the insert so the same source can be reused.
      expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('opens Insert Cells and Delete Cells direction dialogs from the cell menu', () => {
      seed(store, wb, [
        { row: 1, col: 1, value: 'first' },
        { row: 2, col: 1, value: 'second' },
      ]);
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('insertCells')?.click();
      expect(document.querySelector('.fc-cellshift')).not.toBeNull();
      document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'first' });

      setRange(store, 2, 1, 2, 1);
      fireContextMenu(host, 200, 70);
      item('deleteCells')?.click();
      document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'second' });
      expect(onAfterCommit).toHaveBeenCalledTimes(2);
    });

    it('Insert Copied Cells accepts an internal snapshot when clipboard text is empty', async () => {
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('');
      const snap: ClipboardSnapshot = {
        mode: 'copy',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        rows: 1,
        cols: 1,
        cells: [
          [
            {
              value: { kind: 'blank' },
              formula: null,
              format: { hyperlink: 'https://example.test', bold: true },
            },
          ],
        ],
      };
      seed(store, wb, [{ row: 1, col: 1, value: 'old' }]);
      setRange(store, 1, 1, 1, 1);
      mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      detach = attachContextMenu({
        host,
        store,
        wb,
        onAfterCommit,
        getClipboardSnapshot: () => snap,
      });

      fireContextMenu(host, 200, 70);
      item('insertCopiedCells')?.click();
      document.querySelector<HTMLButtonElement>('.fc-insertcopied__button--primary')?.click();

      await Promise.resolve();
      await Promise.resolve();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'old' });
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
      ).toMatchObject({
        hyperlink: 'https://example.test',
        bold: true,
      });
      expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('keeps the Insert Copied Cells dialog chrome rectangular and compact', () => {
      const css = readFileSync(
        join(root, 'src/styles/core/app/dialog-modules/insert-copied-cells.css'),
        'utf8',
      );

      expect(css).toMatch(
        /\.fc-insertcopied__panel\s*\{[\s\S]*?padding: 10px 12px 12px;[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-bg-elev, Canvas\);/,
      );
      expect(css).toMatch(
        /\.fc-insertcopied__button\s*\{[\s\S]*?min-height: 26px;[\s\S]*?border: 1px solid var\(--fc-fmtdlg-input-hover-border, var\(--fc-rule-strong\)\);[\s\S]*?border-radius: 2px;/,
      );
      expect(css).toMatch(/\.fc-insertcopied__footer\s*\{[\s\S]*?gap: 6px;/);
    });

    it('Clear blanks every populated cell within the selection range', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 5, col: 5, value: 99 }, // outside range
      ]);
      setRange(store, 0, 0, 0, 1);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('clear')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 99 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('restores every cleared cell in one undo', () => {
      const history = new History();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 0, col: 2, value: 3 },
      ]);
      setRange(store, 0, 0, 0, 2);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();
      history.undo();
      wb.recalc();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 2 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 3 });
    });

    it('clears the complete primary and extra selection union in one undo and redo', () => {
      const history = new History();
      seed(store, wb, [
        { row: 0, col: 0, value: 11 },
        { row: 0, col: 1, value: 12 },
        { row: 0, col: 2, value: 13 },
        { row: 1, col: 0, value: 21 },
        { row: 1, col: 1, value: 99 }, // the B2 hole
        { row: 2, col: 0, value: 31 },
        { row: 2, col: 1, value: 32 },
        { row: 2, col: 2, value: 33 },
        { row: 4, col: 4, value: 55 },
        { row: 6, col: 6, value: 77 }, // outside every selected fragment
      ]);
      setSelectionRanges(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }, [
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 },
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 },
        { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
      ]);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();
      wb.recalc();

      for (const [row, col] of [
        [0, 0],
        [0, 1],
        [0, 2],
        [1, 0],
        [2, 0],
        [2, 1],
        [2, 2],
        [4, 4],
      ] as const) {
        expect(wb.getValue({ sheet: 0, row, col })).toEqual({ kind: 'blank' });
      }
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 99 });
      expect(wb.getValue({ sheet: 0, row: 6, col: 6 })).toEqual({ kind: 'number', value: 77 });
      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 12 });
      expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'number', value: 55 });
      expect(history.redo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 99 });
      expect(onAfterCommit).toHaveBeenCalledTimes(1);
    });

    it('clears sparse physical formulas and values from a whole-column selection', () => {
      const history = new History();
      const formula = { sheet: 0, row: 250, col: 0 };
      const sparse = { sheet: 0, row: 900, col: 0 };
      const outside = { sheet: 0, row: 900, col: 1 };
      wb.setFormula(formula, '=""');
      wb.setNumber(sparse, 42);
      wb.setNumber(outside, 88);
      wb.recalc();
      expect(wb.cellFormula(formula)).toBe('=""');
      setRange(store, 0, 0, 1_048_575, 0);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();
      wb.recalc();

      expect(wb.getValue(formula)).toEqual({ kind: 'blank' });
      expect(wb.cellFormula(formula)).toBeNull();
      expect(wb.getValue(sparse)).toEqual({ kind: 'blank' });
      expect(wb.getValue(outside)).toEqual({ kind: 'number', value: 88 });
      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.cellFormula(formula)).toBe('=""');
      expect(wb.getValue(sparse)).toEqual({ kind: 'number', value: 42 });
      expect(history.redo()).toBe(true);
      wb.recalc();
      expect(wb.cellFormula(formula)).toBeNull();
    });

    it('clears fully covered merge anchors while preserving a merge with a selection hole', () => {
      const history = new History();
      const fullMerge: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
      const partialMerge: Range = { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 3 };
      mutators.mergeRange(store, fullMerge);
      mutators.mergeRange(store, partialMerge);
      seed(store, wb, [
        { row: 0, col: 0, value: 10 },
        { row: 0, col: 2, value: 20 },
        { row: 0, col: 4, value: 30 },
      ]);
      setSelectionRanges(store, fullMerge, [
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
      ]);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();
      wb.recalc();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 20 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'blank' });
      expect(store.getState().merges.byAnchor.has('0:0:0')).toBe(true);
      expect(store.getState().merges.byAnchor.has('0:0:2')).toBe(true);
      expect(history.undo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 10 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'number', value: 30 });
      expect(history.redo()).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 20 });
    });

    it('waits for an asynchronous restricted clear before refreshing', async () => {
      let resolve: (() => void) | undefined;
      const pending = new Promise<void>((done) => {
        resolve = done;
      });
      const policy = { batchDenied: 'skipIneligible' as const };
      const execute = vi.fn((_command: CellBatchCommand) => pending);
      const controller: ContextMenuInteractionController = {
        policy,
        canExecute: () => ({ allowed: true }),
        execute,
      };
      seed(store, wb, [
        { row: 0, col: 0, value: 11 },
        { row: 4, col: 4, value: 55 },
      ]);
      setSelectionRanges(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
        { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
      ]);
      detach = attachContextMenu({
        host,
        store,
        wb,
        interactionController: controller,
        onAfterCommit,
      });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();

      expect(execute).toHaveBeenCalledTimes(1);
      const command = execute.mock.calls[0]?.[0];
      expect(command).toMatchObject({
        operation: 'clear',
        origin: 'contextMenu',
        commandId: 'clear',
        denied: policy.batchDenied,
      });
      expect(command?.changes.map((change) => change.addr)).toEqual([
        { sheet: 0, row: 0, col: 0 },
        { sheet: 0, row: 4, col: 4 },
      ]);
      expect(
        command?.changes.every((change) => 'value' in change && change.value.kind === 'blank'),
      ).toBe(true);
      expect(onAfterCommit).not.toHaveBeenCalled();
      resolve?.();
      await pending;
      await Promise.resolve();
      expect(onAfterCommit).toHaveBeenCalledTimes(1);
    });

    it('rejects a registered policy atomically and skips empty unions without dispatch', async () => {
      const history = new History();
      const primary: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
      seed(store, wb, [
        { row: 0, col: 0, value: 11 },
        { row: 4, col: 4, value: 55 },
      ]);
      const controller = new InteractionController({
        store,
        history,
        getWb: () => wb,
      });
      const formPolicy = fixedFormPolicy({ ranges: [primary] });
      controller.setPolicy({ ...formPolicy, batchDenied: 'reject' });
      unregisterController = registerInteractionController(store, controller);
      setSelectionRanges(store, primary, [{ sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 }]);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 11 });
      expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'number', value: 55 });
      expect(history.canUndo()).toBe(false);

      detach?.();
      detach = null;
      unregisterController?.();
      unregisterController = null;
      controller.dispose();
      onAfterCommit.mockClear();

      store = createSpreadsheetStore();
      wb = await newWb();
      const emptyHistory = new History();
      const emptyController = new InteractionController({
        store,
        history: emptyHistory,
        getWb: () => wb,
      });
      const execute = vi.spyOn(emptyController, 'execute');
      unregisterController = registerInteractionController(store, emptyController);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, history: emptyHistory, onAfterCommit });

      fireContextMenu(host, 200, 70);
      item('clear')?.click();

      expect(execute).not.toHaveBeenCalled();
      expect(onAfterCommit).not.toHaveBeenCalled();
      expect(emptyHistory.canUndo()).toBe(false);
      emptyController.dispose();
    });
  });

  describe('paste enablement', () => {
    it('disables the Paste button when navigator.clipboard.readText is unavailable', () => {
      const original = navigator.clipboard.readText;
      Object.defineProperty(navigator.clipboard, 'readText', {
        configurable: true,
        value: undefined,
      });
      detach = attachContextMenu({ host, store, wb, strings: en });
      fireContextMenu(host, 200, 70);
      expect(item('paste')?.disabled).toBe(true);
      expect(item('paste')?.getAttribute('aria-disabled')).toBe('true');
      expect(item('paste')?.dataset.disabledReason).toBe(
        'Clipboard paste is unavailable in this environment.',
      );
      Object.defineProperty(navigator.clipboard, 'readText', {
        configurable: true,
        value: original,
      });
    });

    it('enables the Paste button when navigator.clipboard.readText is available', () => {
      detach = attachContextMenu({ host, store, wb, strings: en });
      fireContextMenu(host, 200, 70);
      expect(item('paste')?.disabled).toBe(false);
      expect(item('paste')?.getAttribute('aria-disabled')).toBe('false');
      expect(item('paste')?.dataset.disabledReason).toBeUndefined();
    });
  });
});
