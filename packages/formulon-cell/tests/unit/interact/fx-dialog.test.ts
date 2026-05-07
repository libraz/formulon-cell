import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachFxDialog, FUNCTION_DESCRIPTIONS } from '../../../src/interact/fx-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('attachFxDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts a hidden overlay until open() is called', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    const overlay = host.querySelector<HTMLElement>('.fc-fxdialog');
    expect(overlay?.hidden).toBe(true);
    handle.open();
    expect(overlay?.hidden).toBe(false);
    handle.detach();
  });

  it('renders the function picker step on open without a seed', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const picker = host.querySelector<HTMLElement>('.fc-fxdialog__picker');
    const args = host.querySelector<HTMLElement>('.fc-fxdialog__args');
    expect(picker?.hidden).toBe(false);
    expect(args?.hidden).toBe(true);
    expect(host.querySelectorAll('.fc-fxdialog__item').length).toBeGreaterThan(0);
    handle.detach();
  });

  it('jumps straight to argument entry when open() is given a known function', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open('SUM');
    const picker = host.querySelector<HTMLElement>('.fc-fxdialog__picker');
    const args = host.querySelector<HTMLElement>('.fc-fxdialog__args');
    expect(picker?.hidden).toBe(true);
    expect(args?.hidden).toBe(false);
    const argName = host.querySelector<HTMLElement>('.fc-fxdialog__args-name');
    expect(argName?.textContent).toMatch(/^SUM\(/);
    handle.detach();
  });

  it('filters the picker list by case-insensitive prefix', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const search = host.querySelector<HTMLInputElement>('.fc-fxdialog__search');
    expect(search).toBeTruthy();
    if (!search) return;
    search.value = 'vlo';
    search.dispatchEvent(new Event('input'));
    const items = host.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name');
    const names = Array.from(items).map((i) => i.textContent ?? '');
    expect(names.every((n) => n.includes('VLO'))).toBe(true);
    expect(names).toContain('VLOOKUP');
    handle.detach();
  });

  it('assembles the formula and fires onInsert on confirm', () => {
    const inserted: string[] = [];
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: (f) => inserted.push(f),
    });
    handle.open('SUM');
    const inputs = host.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(inputs.length).toBeGreaterThan(0);
    const input = inputs[0];
    expect(input).toBeDefined();
    if (!input) throw new Error('expected function argument input');
    input.value = 'A1:A5';
    input.dispatchEvent(new Event('input'));
    const insertBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    insertBtn?.click();
    expect(inserted).toEqual(['=SUM(A1:A5)']);
    const overlay = host.querySelector<HTMLElement>('.fc-fxdialog');
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('exposes Excel-style descriptions for the common functions', () => {
    expect(FUNCTION_DESCRIPTIONS.SUM?.en).toMatch(/sum|add/i);
    expect(FUNCTION_DESCRIPTIONS.IF?.en).toMatch(/condition|true|false/i);
    expect(FUNCTION_DESCRIPTIONS.VLOOKUP?.en).toMatch(/lookup|column/i);
  });
});
