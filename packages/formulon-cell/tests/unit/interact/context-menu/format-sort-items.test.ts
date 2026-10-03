import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
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
import { fireContextMenu, item, miniItem, newWb, seed, setFormat, setRange } from './fixtures.js';

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let onAfterCommit: Mock<() => void>;
  let onFormatDialog: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    onFormatDialog = vi.fn<() => void>();
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

  describe('format items', () => {
    it('Bold/Italic/Underline toggle the format on the active cell', () => {
      const history = new History();
      detach = attachContextMenu({ host, store, wb, history });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

      fireContextMenu(host, 200, 70);
      miniItem('bold')?.click();
      expect(store.getState().ui.pendingFormat?.format.bold).toBe(true);
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
      ).toBeUndefined();

      fireContextMenu(host, 200, 70);
      miniItem('italic')?.click();
      expect(store.getState().ui.pendingFormat).toEqual({
        addr: { sheet: 0, row: 0, col: 0 },
        format: { bold: true, italic: true },
      });
      expect(history.undo()).toBe(false);
    });

    it('Align Left / Center / Right set the alignment', () => {
      detach = attachContextMenu({ host, store, wb });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      miniItem('alignLeft')?.click();
      expect(store.getState().ui.pendingFormat?.format.align).toBe('left');
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
      ).toBeUndefined();

      fireContextMenu(host, 200, 70);
      miniItem('alignCenter')?.click();
      expect(store.getState().ui.pendingFormat?.format.align).toBe('center');

      fireContextMenu(host, 200, 70);
      miniItem('alignRight')?.click();
      expect(store.getState().ui.pendingFormat?.format.align).toBe('right');
    });

    it('Borders cycles through outline → all → clear', () => {
      detach = attachContextMenu({ host, store, wb });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      miniItem('borders')?.click();
      expect(store.getState().ui.pendingFormat?.format.borders).toBeDefined();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
      ).toBeUndefined();
    });

    it('repeats Borders on the selection active when F4 runs', () => {
      const history = new History();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 2, col: 2, value: 3 },
      ]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, history });

      fireContextMenu(host, 200, 70);
      miniItem('borders')?.click();
      const a1 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(a1?.borders).toBeDefined();

      setRange(store, 2, 2, 2, 2);
      expect(history.repeatLast()).toBe(true);

      expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual(
        a1,
      );
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))?.borders,
      ).toEqual(a1?.borders);
    });

    it('built-in Underline item toggles underline on the active cell', () => {
      detach = attachContextMenu({
        host,
        store,
        wb,
        options: {
          mode: 'builtIn',
          transform: (context) => [
            ...context.defaultItems,
            { id: 'underline', label: 'Underline', builtIn: 'underline' },
          ],
        },
      });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      item('underline')?.click();
      expect(store.getState().ui.pendingFormat?.format.underline).toBe(true);
    });

    it('Edit Phonetic writes the prompted reading through the engine', async () => {
      Object.defineProperty(wb, 'capabilities', {
        value: { ...wb.capabilities, phonetic: true },
      });
      const setCellPhonetic = vi.spyOn(wb, 'setCellPhonetic').mockReturnValue(true);
      vi.spyOn(wb, 'getCellPhoneticRuns').mockReturnValue(null);
      seed(store, wb, [{ row: 0, col: 0, value: '漢字' }]);
      detach = attachContextMenu({ host, store, wb, strings: en, onAfterCommit });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      const at = store.getState().selection.active;
      item('editPhonetic')?.click();

      const input = document.querySelector<HTMLInputElement>('input');
      expect(input).not.toBeNull();
      if (input) input.value = 'かんじ';
      const ok = Array.from(document.querySelectorAll<HTMLButtonElement>('button')).find(
        (b) => b.textContent === en.formatDialog.ok,
      );
      ok?.click();
      await Promise.resolve();
      await Promise.resolve();

      expect(setCellPhonetic).toHaveBeenCalledWith(at.sheet, at.row, at.col, 'かんじ');
      expect(onAfterCommit).toHaveBeenCalledTimes(1);
    });

    it('Format Cells… triggers the onFormatDialog callback', () => {
      detach = attachContextMenu({ host, store, wb, onFormatDialog });
      fireContextMenu(host, 200, 70);
      item('formatCells')?.click();
      expect(onFormatDialog).toHaveBeenCalled();
    });

    it('enables Open Hyperlink only for a safe hyperlink at the active cell', () => {
      const onOpenHyperlink = vi.fn();
      detach = attachContextMenu({ host, store, wb, onOpenHyperlink });
      fireContextMenu(host, 200, 70);
      expect(item('openHyperlink')?.disabled).toBe(true);

      setFormat(store, 0, 0, { hyperlink: 'https://example.test/docs' });
      fireContextMenu(host, 200, 70);
      item('openHyperlink')?.click();
      expect(onOpenHyperlink).toHaveBeenCalledWith('https://example.test/docs');

      setFormat(store, 0, 0, { hyperlink: 'javascript:alert(1)' });
      fireContextMenu(host, 200, 70);
      expect(item('openHyperlink')?.disabled).toBe(true);
      item('openHyperlink')?.click();
      expect(onOpenHyperlink).toHaveBeenCalledTimes(1);
    });

    it('Select All sets a full-sheet selection', () => {
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 200, 70);
      item('selectAll')?.click();
      const r = store.getState().selection.range;
      expect(r.r1).toBe(1048575);
      expect(r.c1).toBe(16383);
    });
  });

  describe('sort/filter items', () => {
    it('bounds whole-column sort ranges to format-only used rows', () => {
      setRange(store, 0, 0, 1048575, 0);
      setFormat(store, 50000, 0, { hyperlink: 'https://example.test' });
      const setBlank = vi.spyOn(wb, 'setBlank');
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="sortMenu"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));
      item('sortAsc')?.click();

      expect(setBlank).toHaveBeenCalledWith({ sheet: 0, row: 50000, col: 0 });
      expect(onAfterCommit).toHaveBeenCalled();
    });

    it('Sort Descending orders the selected column high to low as one undo step', () => {
      const history = new History();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 1, col: 0, value: 3 },
        { row: 2, col: 0, value: 2 },
      ]);
      setRange(store, 0, 0, 2, 0);
      detach = attachContextMenu({ host, store, wb, history, onAfterCommit });

      fireContextMenu(host, 60, 30);
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="sortMenu"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));
      item('sortDesc')?.click();

      const col = [0, 1, 2].map((row) => wb.getValue({ sheet: 0, row, col: 0 }));
      expect(col).toEqual([
        { kind: 'number', value: 3 },
        { kind: 'number', value: 2 },
        { kind: 'number', value: 1 },
      ]);
      expect(onAfterCommit).toHaveBeenCalledTimes(1);
      expect(history.undo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    });

    it('Filter by Value hides other values; Reapply keeps them hidden; Clear shows them', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'key' },
        { row: 1, col: 0, value: 'a' },
        { row: 2, col: 0, value: 'b' },
        { row: 3, col: 0, value: 'a' },
      ]);
      setRange(store, 1, 0, 1, 0);
      mutators.setFilterRange(store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 });
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const openFilterMenu = (): void => {
        fireContextMenu(host, 60, 50);
        document
          .querySelector<HTMLButtonElement>('[data-fc-submenu="filterMenu"]')
          ?.dispatchEvent(new MouseEvent('mouseenter'));
      };

      openFilterMenu();
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
      item('filterByValue')?.click();
      expect([...store.getState().layout.hiddenRows]).toEqual([2]);

      openFilterMenu();
      item('filterReapply')?.click();
      expect([...store.getState().layout.hiddenRows]).toEqual([2]);

      openFilterMenu();
      item('filterClear')?.click();
      expect(store.getState().layout.hiddenRows.size).toBe(0);
      expect(onAfterCommit).toHaveBeenCalledTimes(3);
    });
  });
});
