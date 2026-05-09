import { describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachErrorMenu } from '../../../src/interact/error-menu.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const fakeWb = () => ({}) as WorkbookHandle;

describe('attachErrorMenu', () => {
  it('invokes trace hook and still emits the host event', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };
    const trace = vi.fn();
    const event = vi.fn();
    host.addEventListener('fc:traceerror', event);

    const handle = attachErrorMenu({
      host,
      store,
      getWb: fakeWb,
      onTraceError: trace,
    });
    handle.open(addr, 10, 10, 'error');

    document
      .querySelector<HTMLButtonElement>('.fc-errmenu__item[data-fc-action="traceError"]')
      ?.click();

    expect(trace).toHaveBeenCalledWith(addr, 'error');
    expect(event).toHaveBeenCalledOnce();
    expect(document.querySelector<HTMLElement>('.fc-errmenu')?.style.display).toBe('none');
    handle.detach();
  });

  it('falls back to selecting the cell when edit-cell has no hook', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 4, col: 3 };
    const handle = attachErrorMenu({ host, store, getWb: fakeWb });

    handle.open(addr, 10, 10, 'error');
    document
      .querySelector<HTMLButtonElement>('.fc-errmenu__item[data-fc-action="editCell"]')
      ?.click();

    expect(store.getState().selection.active).toEqual(addr);
    handle.detach();
  });

  it('ignore suppresses the error indicator for the session', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    const handle = attachErrorMenu({ host, store, getWb: fakeWb });

    handle.open(addr, 10, 10, 'error');
    document
      .querySelector<HTMLButtonElement>('.fc-errmenu__item[data-fc-action="ignore"]')
      ?.click();

    expect(store.getState().errorIndicators.ignoredErrors.has('0:0:0')).toBe(true);
    handle.detach();
  });
});
