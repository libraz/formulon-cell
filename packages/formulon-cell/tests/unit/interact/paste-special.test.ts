import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPasteSpecial } from '../../../src/interact/paste-special.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string; formula?: string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (c.formula) {
        wb.setFormula(addr, c.formula);
        map.set(addrKey(addr), {
          value:
            typeof c.value === 'number'
              ? { kind: 'number', value: c.value }
              : { kind: 'text', value: c.value },
          formula: c.formula,
        });
      } else if (typeof c.value === 'number') {
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

const setActive = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row, col },
      anchor: { sheet: 0, row, col },
      range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

describe('attachPasteSpecial', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts a hidden overlay on attach', () => {
    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => null,
      onAfterCommit,
    });
    const overlay = document.querySelector<HTMLElement>('.fc-pastesp');
    expect(overlay).not.toBeNull();
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('open() is a no-op when there is no snapshot', () => {
    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => null,
      onAfterCommit,
    });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-pastesp');
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('open() reveals the overlay with default radios selected', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    const overlay = document.querySelector<HTMLElement>('.fc-pastesp');
    expect(overlay?.hidden).toBe(false);

    // Default: what='all', operation='none', no skipBlanks/transpose.
    const allRadio = document.querySelector<HTMLInputElement>('input[type="radio"][value="all"]');
    const noneRadio = document.querySelector<HTMLInputElement>('input[type="radio"][value="none"]');
    expect(allRadio?.checked).toBe(true);
    expect(noneRadio?.checked).toBe(true);
    const checks = document.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    for (const c of checks) expect(c.checked).toBe(false);

    handle.detach();
  });

  it('OK button applies pasteSpecial with the selected options and notifies onAfterCommit', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    // Switch to "values" mode and submit.
    document.querySelector<HTMLInputElement>('input[type="radio"][value="values"]')?.click();
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn--primary')[0]?.click();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 7 });
    expect(onAfterCommit).toHaveBeenCalled();
    expect(document.querySelector<HTMLElement>('.fc-pastesp')?.hidden).toBe(true);

    handle.detach();
  });

  it('arithmetic operations apply on top of the destination', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 5, col: 5, value: 100 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();
    document.querySelector<HTMLInputElement>('input[type="radio"][value="values"]')?.click();
    document.querySelector<HTMLInputElement>('input[type="radio"][value="add"]')?.click();
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn--primary')[0]?.click();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 107 });
    handle.detach();
  });

  it('history bundles the paste into a single undoable transaction', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 9 }]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 4, 4);

    const history = new History();
    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      history,
      onAfterCommit,
    });
    handle.open();
    // 'all' default: copies value + format.
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn--primary')[0]?.click();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'number', value: 9 });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 4 }))?.bold).toBe(
      true,
    );

    expect(history.undo()).toBe(true);
    // Format reverts on undo.
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 4 })),
    ).toBeUndefined();
    handle.detach();
  });

  it('Cancel closes without applying', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    // Cancel is the first non-primary footer button.
    const cancelBtn = document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn')[0];
    cancelBtn?.click();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 }).kind).toBe('blank');
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(document.querySelector<HTMLElement>('.fc-pastesp')?.hidden).toBe(true);

    handle.detach();
  });

  it('Escape closes the dialog', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    const e = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true, bubbles: true });
    document.dispatchEvent(e);
    expect(document.querySelector<HTMLElement>('.fc-pastesp')?.hidden).toBe(true);
    handle.detach();
  });

  it('Enter applies the dialog', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 11 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 6, 6);

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    const e = new KeyboardEvent('keydown', { key: 'Enter', cancelable: true, bubbles: true });
    document.dispatchEvent(e);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 6, col: 6 })).toEqual({ kind: 'number', value: 11 });
    expect(onAfterCommit).toHaveBeenCalled();
    handle.detach();
  });

  it('keydown is ignored while the overlay is hidden', () => {
    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => null,
      onAfterCommit,
    });
    // Overlay never opened — Escape/Enter must not throw or fire onAfterCommit.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('clicking the overlay backdrop closes the dialog', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    const overlay = document.querySelector<HTMLElement>('.fc-pastesp');
    overlay?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('clicking inside the panel does not close the dialog', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();

    const panel = document.querySelector<HTMLElement>('.fc-pastesp__panel');
    panel?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(document.querySelector<HTMLElement>('.fc-pastesp')?.hidden).toBe(false);
    handle.detach();
  });

  it('detach removes the overlay and unregisters the global keydown listener', () => {
    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => null,
      onAfterCommit,
    });
    expect(document.querySelector('.fc-pastesp')).not.toBeNull();
    handle.detach();
    expect(document.querySelector('.fc-pastesp')).toBeNull();
    // No leftover listener: pressing Escape after detach is harmless.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
  });

  it('apply with no snapshot at apply-time falls through to close', () => {
    // Open with a snapshot, then have getSnapshot return null at apply-time.
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    let snap: ReturnType<typeof captureSnapshot> | null = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();
    snap = null; // The clipboard was wiped between open and apply.
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn--primary')[0]?.click();
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(document.querySelector<HTMLElement>('.fc-pastesp')?.hidden).toBe(true);
    handle.detach();
  });

  it('skipBlanks and transpose options are read from the checkboxes at apply-time', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    setActive(store, 5, 5);

    const handle = attachPasteSpecial({
      host,
      store,
      wb,
      getSnapshot: () => snap,
      onAfterCommit,
    });
    handle.open();
    document.querySelector<HTMLInputElement>('input[type="radio"][value="values"]')?.click();
    // Toggle transpose.
    const transpose = document.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')[1];
    if (transpose) {
      transpose.checked = true;
      transpose.dispatchEvent(new Event('change', { bubbles: true }));
    }
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn--primary')[0]?.click();
    wb.recalc();
    // 1×3 transposed to 3×1.
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 6, col: 5 })).toEqual({ kind: 'number', value: 2 });
    expect(wb.getValue({ sheet: 0, row: 7, col: 5 })).toEqual({ kind: 'number', value: 3 });
    handle.detach();
  });
});
