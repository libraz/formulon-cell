import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachFindReplace } from '../../../src/interact/find-replace.js';
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

const $ = <T extends Element>(host: HTMLElement, sel: string): T => {
  const el = host.querySelector(sel);
  if (!el) throw new Error(`selector not found: ${sel}`);
  return el as T;
};

const setInputValue = (input: HTMLInputElement, value: string): void => {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
};

const fireKey = (el: Element, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const e = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
  el.dispatchEvent(e);
  return e;
};

describe('attachFindReplace', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;

  beforeEach(async () => {
    host = document.createElement('div');
    // host.focus() requires a focusable element.
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts a hidden overlay on attach and reveals it on open()', () => {
    const handle = attachFindReplace({ host, store, wb, onAfterCommit });
    const overlay = $<HTMLElement>(host, '.fc-find');
    expect(overlay.hidden).toBe(true);
    handle.open();
    expect(overlay.hidden).toBe(false);
    handle.detach();
  });

  it('open() resets currentMatch and shows the empty pill', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 'foo' }]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const pill = $<HTMLElement>(host, '.fc-find__pill');
    expect(pill.textContent).toBe('0 / 0');
    handle.detach();
  });

  it('typing into the find input updates the pill with total matches', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 'foo' },
      { row: 1, col: 0, value: 'bar' },
      { row: 2, col: 0, value: 'foobar' },
    ]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'foo');
    const pill = $<HTMLElement>(host, '.fc-find__pill');
    expect(pill.textContent).toBe('0 / 2');
    handle.detach();
  });

  it('Enter steps to the next match and updates the active cell + pill index', () => {
    // Active starts at (0, 0). findNext requires strict-after, so the first
    // match must live below row 0 to be reachable on the first Enter.
    seed(store, wb, [
      { row: 1, col: 0, value: 'apple' },
      { row: 5, col: 0, value: 'banana' },
      { row: 10, col: 0, value: 'apple' },
    ]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'apple');
    fireKey(findInput, 'Enter');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('1 / 2');

    fireKey(findInput, 'Enter');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 10, col: 0 });
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('2 / 2');
    handle.detach();
  });

  it('Shift+Enter steps backward', () => {
    seed(store, wb, [
      { row: 1, col: 0, value: 'a' },
      { row: 2, col: 0, value: 'a' },
    ]);
    mutators.setActive(store, { sheet: 0, row: 5, col: 0 });
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'a');
    fireKey(findInput, 'Enter', { shiftKey: true });
    // Backward from (5, 0): nearest match before active is (2, 0).
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });
    handle.detach();
  });

  it('Prev / Next buttons step in the corresponding direction', () => {
    seed(store, wb, [
      { row: 1, col: 0, value: 'a' },
      { row: 3, col: 0, value: 'a' },
    ]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'a');
    const buttons = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn');
    // Order in source: Prev, Next, ReplaceOne, ReplaceAll, Close (the case
    // toggle is a label, not a button).
    const [prev, next] = buttons;
    next?.click();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    next?.click();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 3, col: 0 });
    prev?.click();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    handle.detach();
  });

  it('Escape closes the overlay (from either input)', () => {
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    fireKey(findInput, 'Escape');
    expect($<HTMLElement>(host, '.fc-find').hidden).toBe(true);

    handle.open();
    // Find by aria-label since both inputs share the input[type="text"] selector.
    const replaceInput = host.querySelector<HTMLInputElement>(
      '.fc-find input[aria-label]:nth-of-type(1)',
    );
    // Escape on any input keydown closes too.
    fireKey(replaceInput ?? findInput, 'Escape');
    expect($<HTMLElement>(host, '.fc-find').hidden).toBe(true);
    handle.detach();
  });

  it('case toggle invalidates the current match (forces a fresh step)', () => {
    seed(store, wb, [
      { row: 1, col: 0, value: 'Foo' },
      { row: 2, col: 0, value: 'foo' },
    ]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'foo');
    fireKey(findInput, 'Enter'); // case-insensitive: 2 matches; first after (0,0) is (1, 0).
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });

    const caseToggle = host.querySelector<HTMLInputElement>('.fc-find input[type="checkbox"]');
    if (caseToggle) {
      caseToggle.checked = true;
      caseToggle.dispatchEvent(new Event('change', { bubbles: true }));
    }
    // Case-sensitive: only (2, 0) matches now; pill drops to "0 / 1".
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('0 / 1');
    handle.detach();
  });

  it('replaceOne writes through wb and advances to the next match', () => {
    seed(store, wb, [
      { row: 1, col: 0, value: 'aaa' },
      { row: 2, col: 0, value: 'aaa' },
    ]);
    const handle = attachFindReplace({ host, store, wb, onAfterCommit });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    const replaceInput = host.querySelectorAll<HTMLInputElement>('.fc-find input[type="text"]')[1];
    setInputValue(findInput, 'a');
    if (replaceInput) setInputValue(replaceInput, 'X');
    // Enter sets currentMatch to (1, 0) (first match strictly after the
    // default active (0, 0)). The Replace button then writes there and steps
    // forward to (2, 0).
    fireKey(findInput, 'Enter');
    const replaceBtn = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn')[2];
    replaceBtn?.click();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'XXX' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'aaa' });
    expect(onAfterCommit).toHaveBeenCalled();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });
    handle.detach();
  });

  it('replaceOne skips formula cells', () => {
    // Place the formula match at (1, 0) so step('next') from default (0, 0)
    // can land on it.
    seed(store, wb, [{ row: 1, col: 0, value: 'aaa', formula: '="aaa"' }]);
    const handle = attachFindReplace({ host, store, wb, onAfterCommit });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    const replaceInput = host.querySelectorAll<HTMLInputElement>('.fc-find input[type="text"]')[1];
    setInputValue(findInput, 'a');
    if (replaceInput) setInputValue(replaceInput, 'X');
    fireKey(findInput, 'Enter');
    const replaceBtn = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn')[2];
    replaceBtn?.click();
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('="aaa"');
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('replaceAll reports the count in the pill', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 'foo' },
      { row: 1, col: 0, value: 'foo' },
      { row: 2, col: 0, value: 'foo' },
    ]);
    const handle = attachFindReplace({ host, store, wb, onAfterCommit });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    const replaceInput = host.querySelectorAll<HTMLInputElement>('.fc-find input[type="text"]')[1];
    setInputValue(findInput, 'foo');
    if (replaceInput) setInputValue(replaceInput, 'bar');
    const replaceAllBtn = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn')[3];
    replaceAllBtn?.click();
    wb.recalc();
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('3 replaced');
    expect(onAfterCommit).toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'bar' });
    handle.detach();
  });

  it('replaceAll on empty query is a no-op', () => {
    const handle = attachFindReplace({ host, store, wb, onAfterCommit });
    handle.open();
    const replaceAllBtn = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn')[3];
    replaceAllBtn?.click();
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('0 / 0');
    handle.detach();
  });

  it('close button closes the overlay', () => {
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    // Close is the 5th button (after prev/next/replace/replaceAll).
    const closeBtn = host.querySelectorAll<HTMLButtonElement>('.fc-find__btn')[4];
    closeBtn?.click();
    expect($<HTMLElement>(host, '.fc-find').hidden).toBe(true);
    handle.detach();
  });

  it('close() preserves the last query for the next open', () => {
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    setInputValue(findInput, 'remembered');
    handle.close();
    handle.open();
    expect($<HTMLInputElement>(host, '.fc-find input[type="text"]').value).toBe('remembered');
    handle.detach();
  });

  it('detach removes the overlay from the DOM', () => {
    const handle = attachFindReplace({ host, store, wb });
    expect(host.querySelector('.fc-find')).not.toBeNull();
    handle.detach();
    expect(host.querySelector('.fc-find')).toBeNull();
  });

  it('Enter with empty query clears currentMatch and pill', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 'a' }]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const findInput = $<HTMLInputElement>(host, '.fc-find input[type="text"]');
    fireKey(findInput, 'Enter');
    expect($<HTMLElement>(host, '.fc-find__pill').textContent).toBe('0 / 0');
    handle.detach();
  });

  it('overlay keydown stops propagation so grid handlers are not triggered', () => {
    const handle = attachFindReplace({ host, store, wb });
    handle.open();
    const overlay = $<HTMLElement>(host, '.fc-find');
    let bubbled = false;
    host.addEventListener('keydown', () => {
      bubbled = true;
    });
    const e = new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true });
    overlay.dispatchEvent(e);
    expect(bubbled).toBe(false);
    handle.detach();
  });
});
