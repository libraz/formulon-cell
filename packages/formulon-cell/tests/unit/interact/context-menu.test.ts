import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import { attachContextMenu } from '../../../src/interact/context-menu.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const fireContextMenu = (
  host: HTMLElement,
  x: number,
  y: number,
  init: MouseEventInit = {},
): MouseEvent => {
  const e = new MouseEvent('contextmenu', {
    clientX: x,
    clientY: y,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  host.dispatchEvent(e);
  return e;
};

const item = (id: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.fc-ctxmenu__item[data-fc-action="${id}"]`);

const visibleMenu = (): HTMLElement | null => {
  const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
  if (!root) return null;
  return root.style.display === 'none' ? null : root;
};

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;
  let onAfterCommit: ReturnType<typeof vi.fn>;
  let onFormatDialog: ReturnType<typeof vi.fn>;
  let onPasteSpecial: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn();
    onFormatDialog = vi.fn();
    onPasteSpecial = vi.fn();
  });

  afterEach(() => {
    detach?.();
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  describe('menu opening', () => {
    it('right-click on a cell opens the cell menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const e = fireContextMenu(host, 200, 70);
      expect(e.defaultPrevented).toBe(true);
      const menu = visibleMenu();
      expect(menu).not.toBeNull();
      expect(item('copy')).not.toBeNull();
      expect(item('formatCells')).not.toBeNull();
      // No row/col-only items in cell menu.
      expect(item('rowInsertAbove')).toBeNull();
      expect(item('colInsertLeft')).toBeNull();
    });

    it('right-click on the row header opens the row menu and promotes selection', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 40); // row 0 header
      expect(item('rowInsertAbove')).not.toBeNull();
      expect(item('rowDelete')).not.toBeNull();
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 16383 });
    });

    it('right-click on the col header opens the col menu and promotes selection', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 100, 10); // col 0 header
      expect(item('colInsertLeft')).not.toBeNull();
      expect(item('colDelete')).not.toBeNull();
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    });

    it('does not promote selection when the header click falls inside an existing band', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      mutators.selectRow(store, 0);
      const before = store.getState().selection.range;
      fireContextMenu(host, 10, 40);
      // Range unchanged.
      expect(store.getState().selection.range).toEqual(before);
    });

    it('right-click on the formula bar does not open the menu', () => {
      const formulaBar = document.createElement('div');
      formulaBar.className = 'fc-host__formulabar';
      host.appendChild(formulaBar);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const e = new MouseEvent('contextmenu', {
        clientX: 50,
        clientY: 50,
        bubbles: true,
        cancelable: true,
      });
      formulaBar.dispatchEvent(e);
      expect(visibleMenu()).toBeNull();
      expect(e.defaultPrevented).toBe(false);
    });
  });

  describe('dismissal', () => {
    it('clicking outside the menu hides it', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      expect(visibleMenu()).not.toBeNull();
      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      expect(visibleMenu()).toBeNull();
    });

    it('Escape hides the menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
      expect(visibleMenu()).toBeNull();
    });

    it('scroll hides the menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      window.dispatchEvent(new Event('scroll'));
      expect(visibleMenu()).toBeNull();
    });

    it('clicking inside the menu does not dismiss it', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
      if (root) {
        const e = new MouseEvent('mousedown', { bubbles: true });
        root.dispatchEvent(e);
      }
      expect(visibleMenu()).not.toBeNull();
    });
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
      expect(visibleMenu()).toBeNull();
    });

    it('Cut writes TSV, blanks the range, and notifies onAfterCommit', () => {
      const writeText = vi.spyOn(navigator.clipboard, 'writeText').mockResolvedValue();
      seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      item('cut')?.click();
      expect(writeText).toHaveBeenCalled();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(onAfterCommit).toHaveBeenCalled();
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
      expect(onAfterCommit).toHaveBeenCalled();
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

    it('Paste Special triggers the onPasteSpecial callback', () => {
      detach = attachContextMenu({ host, store, wb, onPasteSpecial });
      fireContextMenu(host, 200, 70);
      item('pasteSpecial')?.click();
      expect(onPasteSpecial).toHaveBeenCalled();
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
  });

  describe('format items', () => {
    it('Bold/Italic/Underline toggle the format on the active cell', () => {
      const history = new History();
      detach = attachContextMenu({ host, store, wb, history });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

      fireContextMenu(host, 200, 70);
      item('bold')?.click();
      expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(
        true,
      );

      fireContextMenu(host, 200, 70);
      item('italic')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.italic,
      ).toBe(true);

      // Undo through history reverts both.
      expect(history.undo()).toBe(true);
      expect(history.undo()).toBe(true);
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
      ).toBeUndefined();
    });

    it('Align Left / Center / Right set the alignment', () => {
      detach = attachContextMenu({ host, store, wb });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      item('alignLeft')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align,
      ).toBe('left');

      fireContextMenu(host, 200, 70);
      item('alignCenter')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align,
      ).toBe('center');

      fireContextMenu(host, 200, 70);
      item('alignRight')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align,
      ).toBe('right');
    });

    it('Borders cycles through outline → all → clear', () => {
      detach = attachContextMenu({ host, store, wb });
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireContextMenu(host, 200, 70);
      item('borders')?.click();
      const f = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(f?.borders).toBeDefined();
    });

    it('Clear Format wipes the format entry', () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 200, 70);
      item('clearFormat')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
      ).toBeUndefined();
    });

    it('Format Cells… triggers the onFormatDialog callback', () => {
      detach = attachContextMenu({ host, store, wb, onFormatDialog });
      fireContextMenu(host, 200, 70);
      item('formatCells')?.click();
      expect(onFormatDialog).toHaveBeenCalled();
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

  describe('row structure', () => {
    it('Insert Above shifts existing rows down', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'a' }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 40); // row 0 header
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
      fireContextMenu(host, 10, 40);
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
      fireContextMenu(host, 10, 40);
      item('rowDelete')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Hide Row records the row as hidden, Unhide restores it', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 10, 40);
      item('rowHide')?.click();
      expect(store.getState().layout.hiddenRows.has(0)).toBe(true);
      // Span rows 0-1 so the next contextmenu's promotion-check sees the row
      // band as already-selected. Row 0 is hidden, so y=40 hits row 1 header,
      // but inSel=true keeps the (0..1) band intact.
      setRange(store, 0, 0, 1, 16383);
      fireContextMenu(host, 10, 40);
      item('rowUnhide')?.click();
      expect(store.getState().layout.hiddenRows.has(0)).toBe(false);
    });

    it('Unhide is a no-op when no hidden rows are in the selection', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 10, 40);
      item('rowUnhide')?.click();
      expect(store.getState().layout.hiddenRows.size).toBe(0);
    });
  });

  describe('col structure', () => {
    it('Insert Left shifts subsequent cols right', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'a' }]);
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 100, 10);
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
      fireContextMenu(host, 100, 10);
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
      fireContextMenu(host, 100, 10);
      item('colDelete')?.click();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'b' });
    });

    it('Hide / Unhide Col toggles layout.hiddenCols', () => {
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 100, 10);
      item('colHide')?.click();
      expect(store.getState().layout.hiddenCols.has(0)).toBe(true);
      // Span cols 0-1 so the next contextmenu's promotion-check sees the col
      // band as already-selected. Col 0 is hidden, so x=100 hits col 1 header,
      // but inSel=true keeps the (0..1) band intact.
      setRange(store, 0, 0, 1048575, 1);
      fireContextMenu(host, 100, 10);
      item('colUnhide')?.click();
      expect(store.getState().layout.hiddenCols.has(0)).toBe(false);
    });
  });

  describe('paste enablement', () => {
    it('disables the Paste button when navigator.clipboard.readText is unavailable', () => {
      const original = navigator.clipboard.readText;
      Object.defineProperty(navigator.clipboard, 'readText', {
        configurable: true,
        value: undefined,
      });
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 200, 70);
      expect(item('paste')?.disabled).toBe(true);
      expect(item('paste')?.getAttribute('aria-disabled')).toBe('true');
      Object.defineProperty(navigator.clipboard, 'readText', {
        configurable: true,
        value: original,
      });
    });

    it('enables the Paste button when navigator.clipboard.readText is available', () => {
      detach = attachContextMenu({ host, store, wb });
      fireContextMenu(host, 200, 70);
      expect(item('paste')?.disabled).toBe(false);
      expect(item('paste')?.hasAttribute('aria-disabled')).toBe(false);
    });
  });

  describe('teardown', () => {
    it('detach removes the menu from the DOM', () => {
      detach = attachContextMenu({ host, store, wb });
      expect(document.querySelector('.fc-ctxmenu')).not.toBeNull();
      detach();
      detach = (): void => {};
      expect(document.querySelector('.fc-ctxmenu')).toBeNull();
    });

    it('after detach, contextmenu events do not open the menu', () => {
      detach = attachContextMenu({ host, store, wb });
      detach();
      detach = (): void => {};
      fireContextMenu(host, 200, 70);
      expect(visibleMenu()).toBeNull();
    });
  });
});
