import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

/**
 * Unit: formula-bar controller — drives editing lifecycle (focus → input →
 * commit / cancel) on the fxInput textarea. Exercised through the mounted
 * sheet because the controller relies on the wider host + autocomplete +
 * arg-helper plumbing that mount.ts wires together.
 */
describe('mount/formula-bar — edit lifecycle', () => {
  let sheet: MountedStubSheet;
  let fxInput: HTMLTextAreaElement;
  let fxCancel: HTMLButtonElement;
  let fxAccept: HTMLButtonElement;
  let formulabar: HTMLDivElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    fxCancel = sheet.host.querySelector('.fc-host__formulabar-action--cancel') as HTMLButtonElement;
    fxAccept = sheet.host.querySelector('.fc-host__formulabar-action--accept') as HTMLButtonElement;
    formulabar = sheet.host.querySelector('.fc-host__formulabar') as HTMLDivElement;
  });

  afterEach(() => sheet.dispose());

  it('idle: cancel + accept buttons are disabled, fcEditing=0', () => {
    expect(fxCancel.disabled).toBe(true);
    expect(fxAccept.disabled).toBe(true);
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('focus starts editing — cancel enables, accept stays disabled until dirty', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    expect(formulabar.dataset.fcEditing).toBe('1');
    expect(fxCancel.disabled).toBe(false);
    expect(fxAccept.disabled).toBe(true);
  });

  it('changing the value flips accept on (dirty); Escape rolls it back', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=A1+1';
    fxInput.dispatchEvent(new Event('input'));
    expect(fxAccept.disabled).toBe(false);

    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(fxInput.value).toBe('');
    expect(formulabar.dataset.fcEditing).toBe('0');
    expect(fxAccept.disabled).toBe(true);
  });

  it('Enter commits and advances the active cell down', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '42';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 42,
    });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 1,
      col: 0,
    });
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('Tab commits and advances the active cell right', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'hello';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab' }));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'hello',
    });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 0,
      col: 1,
    });
  });

  it('clicking the cancel button reverts a dirty edit', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=BAD';
    fxInput.dispatchEvent(new Event('input'));
    expect(fxAccept.disabled).toBe(false);

    fxCancel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(fxInput.value).toBe('');
    expect(fxCancel.disabled).toBe(true);
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('clicking the accept button commits without changing the active cell', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '7';
    fxInput.dispatchEvent(new Event('input'));

    fxAccept.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 7,
    });
    // No advance for "accept" — selection stays where it was.
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 1,
      col: 1,
    });
  });

  it('blur with a dirty edit commits (matches click-elsewhere behaviour)', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'blur-commit';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new FocusEvent('blur'));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'blur-commit',
    });
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('blur without changes just leaves edit-mode (no commit)', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.dispatchEvent(new FocusEvent('blur'));
    expect(formulabar.dataset.fcEditing).toBe('0');
    // Active cell stayed blank.
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
  });

  it('F4 rotates the ref under the caret', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=A1';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.setSelectionRange(3, 3);
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4' }));
    // First rotation: A1 → $A$1.
    expect(fxInput.value).toBe('=$A$1');
  });
});
