import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachGoToDialog } from '../../../src/interact/goto-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-goto');
const kindRadios = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-goto__kinds input[type="radio"]'));
const scopeRadios = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-goto__scope input[type="radio"]'));
const okBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-goto .fc-fmtdlg__btn--primary');
const status = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-goto__status');

const sync = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  store.setState((s) => {
    const cells = new Map<
      string,
      { value: ReturnType<WorkbookHandle['getValue']>; formula: string | null }
    >();
    for (const e of wb.cells(0)) {
      cells.set(addrKey(e.addr), { value: e.value, formula: e.formula });
    }
    return { ...s, data: { ...s.data, cells } };
  });
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

describe('attachGoToDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('renders kind radios in the canonical order with constants pre-checked', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const radios = kindRadios();
    expect(radios.map((r) => r.value)).toEqual([
      'blanks',
      'non-blanks',
      'formulas',
      'constants',
      'numbers',
      'text',
      'errors',
      'data-validation',
      'conditional-format',
    ]);
    const checked = radios.find((r) => r.checked);
    expect(checked?.value).toBe('constants');
    handle.detach();
  });

  it('disables the selection-scope radio when the current selection is a single cell', () => {
    setRange(store, 0, 0, 0, 0);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const [sheetRadio, selectionRadio] = scopeRadios();
    expect(sheetRadio?.checked).toBe(true);
    expect(selectionRadio?.disabled).toBe(true);
    handle.detach();
  });

  it('enables the selection-scope radio when the selection covers more than one cell', () => {
    setRange(store, 0, 0, 4, 4);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const [, selectionRadio] = scopeRadios();
    expect(selectionRadio?.disabled).toBe(false);
    handle.detach();
  });

  it('OK with no matches keeps the dialog open and sets the no-results status', () => {
    // Sheet has no formula cells.
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const radios = kindRadios();
    const formulaRadio = radios.find((r) => r.value === 'formulas');
    if (formulaRadio) {
      formulaRadio.checked = true;
    }
    okBtn()?.click();
    expect(status()?.textContent).not.toBe('');
    expect(overlay()?.hidden).toBe(false);
    handle.detach();
  });

  it('OK with 1+ matches closes the dialog and moves the selection to the bounding range', () => {
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 11);
    wb.setNumber({ sheet: 0, row: 5, col: 4 }, 22);
    wb.recalc();
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const numRadio = kindRadios().find((r) => r.value === 'numbers');
    if (numRadio) numRadio.checked = true;
    okBtn()?.click();
    expect(overlay()?.hidden).toBe(true);
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 2, col: 1 });
    expect(sel.range).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 5, c1: 4 });
    handle.detach();
  });

  it('Escape closes the dialog without mutating the selection', () => {
    setRange(store, 1, 1, 1, 1);
    const before = { ...store.getState().selection.active };
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    overlay()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay()?.hidden).toBe(true);
    expect(store.getState().selection.active).toEqual(before);
    handle.detach();
  });

  it('detach removes the overlay node', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.detach();
    expect(overlay()).toBeNull();
  });
});
