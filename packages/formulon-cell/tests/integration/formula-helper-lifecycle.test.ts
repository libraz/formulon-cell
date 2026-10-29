import { afterEach, describe, expect, it, vi } from 'vitest';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

describe('integration: formula helper lifecycle', () => {
  let sheet: MountedStubSheet | null = null;

  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('accepts a mouse candidate on double-click and synchronizes the formula bar once', async () => {
    sheet = await mountStubSheet();
    const input = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    if (!input) throw new Error('Formula bar input was not mounted.');
    input.focus();
    input.value = '=SU';
    input.dispatchEvent(new Event('input', { bubbles: true }));

    const sum = [...document.querySelectorAll<HTMLElement>('[role="option"]')].find(
      (option) => option.querySelector('.fc-autocomplete__name')?.textContent === 'SUM',
    );
    expect(sum).toBeDefined();
    const inputSync = vi.fn();
    input.addEventListener('input', inputSync);

    sum?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    expect(input.value).toBe('=SU');
    sum?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));

    expect(input.value).toBe('=SUM(');
    expect(inputSync).toHaveBeenCalledOnce();
    expect(document.querySelector('.fc-autocomplete__list')).toBeNull();
    expect(sheet.host.querySelector('.fc-host__formulabar')?.getAttribute('data-fc-editing')).toBe(
      '1',
    );
  });

  it('commits the typed formula on Enter instead of inserting the highlighted function', async () => {
    sheet = await mountStubSheet();
    const anchor = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, anchor);
    const input = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    if (!input) throw new Error('Formula bar input was not mounted.');
    input.focus();
    input.value = '=SU';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector('.fc-autocomplete__list')).not.toBeNull();

    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(sheet.workbook.cellFormula(anchor)).toBe('=SU');
    expect(input.value).not.toBe('=SUM(');
    expect(document.querySelector('.fc-autocomplete__list')).toBeNull();
  });

  it('hides and resumes formula bar helpers at the IME composition boundary', async () => {
    sheet = await mountStubSheet();
    const input = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    if (!input) throw new Error('Formula bar input was not mounted.');
    input.focus();
    input.value = '=SU';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector('.fc-autocomplete__list')).not.toBeNull();

    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    input.value = '=SUB';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector('.fc-autocomplete__list')).toBeNull();

    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    expect(document.querySelector('.fc-autocomplete__list')).not.toBeNull();
  });
});
