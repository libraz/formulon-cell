import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachWatchPanel } from '../../../src/interact/watch-panel.js';
import { History } from '../../../src/commands/history.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const setCellInStore = (
  store: SpreadsheetStore,
  sheet: number,
  row: number,
  col: number,
  value: CellValue,
  formula: string | null = null,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet, row, col }), { value, formula });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('watch slice mutators', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('addWatch appends an Addr; duplicates are ignored', () => {
    mutators.addWatch(store, { sheet: 0, row: 1, col: 2 });
    mutators.addWatch(store, { sheet: 0, row: 1, col: 2 });
    expect(store.getState().watch.watches).toEqual([{ sheet: 0, row: 1, col: 2 }]);
  });

  it('addWatchRanges appends cells row-major and ignores duplicates', () => {
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
    mutators.addWatchRanges(store, [
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { sheet: 1, r0: 2, c0: 3, r1: 2, c1: 3 },
    ]);
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 1, col: 1 },
      { sheet: 1, row: 2, col: 3 },
    ]);
  });

  it('removeWatch drops the matching Addr', () => {
    mutators.addWatch(store, { sheet: 0, row: 1, col: 2 });
    mutators.addWatch(store, { sheet: 0, row: 3, col: 4 });
    mutators.removeWatch(store, { sheet: 0, row: 1, col: 2 });
    expect(store.getState().watch.watches).toEqual([{ sheet: 0, row: 3, col: 4 }]);
  });

  it('clearWatches empties the list', () => {
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
    mutators.addWatch(store, { sheet: 1, row: 0, col: 0 });
    mutators.clearWatches(store);
    expect(store.getState().watch.watches).toEqual([]);
  });

  it('setWatchPanelOpen flips the UI flag', () => {
    expect(store.getState().ui.watchPanelOpen).toBe(false);
    mutators.setWatchPanelOpen(store, true);
    expect(store.getState().ui.watchPanelOpen).toBe(true);
    mutators.setWatchPanelOpen(store, false);
    expect(store.getState().ui.watchPanelOpen).toBe(false);
  });
});

describe('attachWatchPanel', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  it('mounts hidden when the panel flag is off; open() shows it with empty state', () => {
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    const root = host.querySelector<HTMLElement>('.fc-watch');
    expect(root?.hidden).toBe(true);
    handle.open();
    expect(root?.hidden).toBe(false);
    expect(root?.tabIndex).toBe(-1);
    expect(host.querySelector<HTMLElement>('.fc-watch__empty')?.hidden).toBe(false);
    handle.detach();
  });

  it('open focuses the panel actions and Escape closes back to the opener', async () => {
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    host.focus();
    handle.open();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    const root = host.querySelector<HTMLElement>('.fc-watch');
    const addBtn = host.querySelector<HTMLButtonElement>('.fc-watch__btn');

    expect(document.activeElement).toBe(addBtn);

    root?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(root?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('close button closes the panel and restores focus to the opener', async () => {
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    host.focus();
    handle.open();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    const root = host.querySelector<HTMLElement>('.fc-watch');
    const closeBtn = host.querySelector<HTMLButtonElement>('.fc-watch__close');
    closeBtn?.focus();
    closeBtn?.click();

    expect(root?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('renders one row per watched cell with book/sheet/cell/value/formula', () => {
    setCellInStore(store, 0, 1, 2, { kind: 'number', value: 42 });
    // Seed the engine too so wb.getValue returns the same number.
    wb.setNumber({ sheet: 0, row: 1, col: 2 }, 42);
    mutators.setWatchPanelOpen(store, true);
    mutators.addWatch(store, { sheet: 0, row: 1, col: 2 });

    const handle = attachWatchPanel({ host, store, getWb: () => wb, workbookName: 'Budget.xlsx' });
    const rows = host.querySelectorAll<HTMLTableRowElement>('.fc-watch__row');
    expect(rows.length).toBe(1);
    const cells = rows[0]?.querySelectorAll('td');
    // book | sheet | name | cell | value | formula | remove
    expect(cells?.length).toBe(7);
    expect(cells?.[0]?.textContent).toBe('Budget.xlsx');
    expect(cells?.[3]?.textContent).toBe('C2');
    expect(cells?.[4]?.textContent).toContain('42');
    handle.detach();
  });

  it('marks the row matching the active selection', () => {
    mutators.setWatchPanelOpen(store, true);
    mutators.addWatch(store, { sheet: 0, row: 1, col: 1 });
    mutators.addWatch(store, { sheet: 0, row: 2, col: 2 });
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    let selected = host.querySelectorAll<HTMLTableRowElement>('.fc-watch__row--selected');
    expect(selected.length).toBe(1);
    expect(selected[0]?.dataset.fcRow).toBe('2');

    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });

    selected = host.querySelectorAll<HTMLTableRowElement>('.fc-watch__row--selected');
    expect(selected.length).toBe(1);
    expect(selected[0]?.dataset.fcRow).toBe('1');
    handle.detach();
  });

  it('clicking a row jumps the active selection to that cell', () => {
    mutators.setWatchPanelOpen(store, true);
    mutators.addWatch(store, { sheet: 0, row: 4, col: 5 });
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    const row = host.querySelector<HTMLTableRowElement>('.fc-watch__row');
    expect(row).not.toBeNull();
    row?.click();
    const active = store.getState().selection.active;
    expect(active).toEqual({ sheet: 0, row: 4, col: 5 });
    handle.detach();
  });

  it('per-row × button removes the watch without changing selection', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setWatchPanelOpen(store, true);
    mutators.addWatch(store, { sheet: 0, row: 7, col: 8 });
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    const removeBtn = host.querySelector<HTMLButtonElement>('.fc-watch__remove');
    expect(removeBtn).not.toBeNull();
    removeBtn?.click();
    expect(store.getState().watch.watches).toEqual([]);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(host.querySelector<HTMLElement>('.fc-watch__empty')?.hidden).toBe(false);
    handle.detach();
  });

  it('header Add button watches every cell in the selected ranges', () => {
    mutators.setActive(store, { sheet: 0, row: 9, col: 9 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        extraRanges: [{ sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 }],
      },
    }));
    mutators.setWatchPanelOpen(store, true);
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    const addBtn = Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-watch__btn')).find(
      (b) => !b.classList.contains('fc-watch__close'),
    );
    addBtn?.click();
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 2 },
      { sheet: 0, row: 3, col: 3 },
    ]);
    handle.detach();
  });

  it('records panel Add, row remove, and Clear All as undoable watch changes', () => {
    const history = new History();
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    mutators.setWatchPanelOpen(store, true);
    const handle = attachWatchPanel({ host, store, getWb: () => wb, history });
    const [addBtn, clearBtn] = Array.from(
      host.querySelectorAll<HTMLButtonElement>('.fc-watch__btn:not(.fc-watch__close)'),
    );

    addBtn?.click();
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 2 },
    ]);
    expect(history.undo()).toBe(true);
    expect(store.getState().watch.watches).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 2 },
    ]);

    host.querySelector<HTMLButtonElement>('.fc-watch__remove')?.click();
    expect(store.getState().watch.watches).toEqual([{ sheet: 0, row: 1, col: 2 }]);
    expect(history.undo()).toBe(true);
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 2 },
    ]);

    clearBtn?.click();
    expect(store.getState().watch.watches).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 2 },
    ]);
    handle.detach();
  });

  it('refresh() re-reads engine values; cell-store update triggers a repaint', () => {
    mutators.setWatchPanelOpen(store, true);
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    setCellInStore(store, 0, 0, 0, { kind: 'number', value: 1 });

    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    let valueCell = host.querySelector<HTMLTableCellElement>('.fc-watch__value');
    expect(valueCell?.textContent).toContain('1');

    // Drive a recalc-style change: engine value updates + store mirror updates.
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 99);
    setCellInStore(store, 0, 0, 0, { kind: 'number', value: 99 });
    // The store subscription should have repainted, but call refresh() too
    // to mirror the mount.ts recalc-event path.
    handle.refresh();
    valueCell = host.querySelector<HTMLTableCellElement>('.fc-watch__value');
    expect(valueCell?.textContent).toContain('99');
    handle.detach();
  });

  it('detach removes the panel node and unsubscribes', () => {
    mutators.setWatchPanelOpen(store, true);
    const handle = attachWatchPanel({ host, store, getWb: () => wb });
    expect(host.querySelector('.fc-watch')).not.toBeNull();
    handle.detach();
    expect(host.querySelector('.fc-watch')).toBeNull();
    // Mutating state after detach should not crash.
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
  });
});
