import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { attachErrorMenu } from '../../../src/interact/error-menu.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const fakeWb = () => ({}) as WorkbookHandle;

describe('attachErrorMenu', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

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
      strings: en,
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

  it('opens as an accessible keyboard menu and restores focus on Escape', () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };
    const handle = attachErrorMenu({ host, store, getWb: fakeWb, strings: en });

    host.focus();
    handle.open(addr, 10, 10, 'validation');
    const root = document.querySelector<HTMLElement>('.fc-errmenu');
    const items = document.querySelectorAll<HTMLButtonElement>('.fc-errmenu__item');

    expect(root?.getAttribute('role')).toBe('menu');
    expect(root?.getAttribute('aria-label')).toBe('Validation issue');
    expect(document.activeElement).toBe(items[0]);

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', cancelable: true }));
    expect(document.activeElement).toBe(items[items.length - 1]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', cancelable: true }));
    expect(document.activeElement).toBe(items[0]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));

    expect(root?.style.display).toBe('none');
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('clamps the menu inside the viewport through the shared overlay helper', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };
    const handle = attachErrorMenu({ host, store, getWb: fakeWb, strings: en });
    const root = document.querySelector<HTMLElement>('.fc-errmenu');
    expect(root).toBeTruthy();
    if (root) {
      Object.defineProperty(root, 'offsetWidth', { configurable: true, value: 240 });
      Object.defineProperty(root, 'offsetHeight', { configurable: true, value: 120 });
    }

    handle.open(addr, 310, 170, 'error');

    expect(root?.style.left).toBe('76px');
    expect(root?.style.top).toBe('56px');
    handle.detach();
  });

  it('Enter invokes the focused menu item', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };
    const trace = vi.fn();
    const handle = attachErrorMenu({
      host,
      store,
      getWb: fakeWb,
      strings: en,
      onTraceError: trace,
    });

    handle.open(addr, 10, 10, 'error');
    const traceBtn = document.querySelector<HTMLButtonElement>(
      '.fc-errmenu__item[data-fc-action="traceError"]',
    );
    traceBtn?.focus();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(trace).toHaveBeenCalledWith(addr, 'error');
    expect(document.querySelector<HTMLElement>('.fc-errmenu')?.style.display).toBe('none');
    handle.detach();
  });

  it('falls back to selecting the cell when edit-cell has no hook', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 4, col: 3 };
    const handle = attachErrorMenu({ host, store, getWb: fakeWb, strings: en });

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
    const handle = attachErrorMenu({ host, store, getWb: fakeWb, strings: en });

    handle.open(addr, 10, 10, 'error');
    document
      .querySelector<HTMLButtonElement>('.fc-errmenu__item[data-fc-action="ignore"]')
      ?.click();

    expect(store.getState().errorIndicators.ignoredErrors.has('0:0:0')).toBe(true);
    handle.detach();
  });

  it('keeps error menu item button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/error-menu.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const createErrorMenuItemButton');
    expect(source).toContain('const btn = createErrorMenuItemButton(entry)');
    expect(source).toContain("className: 'fc-errmenu__item'");
    expect(source).toContain("role: 'menuitem'");
    expect(source).not.toContain("document.createElement('button')");
  });
});
